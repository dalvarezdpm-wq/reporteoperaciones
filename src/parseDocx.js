import mammoth from "mammoth";
import { PREVIO_FIELDS, DESPACHO_FIELDS } from "./fields.js";

function cellTextOf(el) {
  return (el.textContent || "").replace(/\s+/g, " ").trim();
}
function normLabel(l) {
  return l
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

const MESES = {
  enero: "01", febrero: "02", marzo: "03", abril: "04", mayo: "05", junio: "06",
  julio: "07", agosto: "08", septiembre: "09", setiembre: "09", octubre: "10",
  noviembre: "11", diciembre: "12",
};

export function parseFechaLoose(fecha) {
  const m = (fecha || "").match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return "";
  let [_, d, mo, y] = m;
  if (y.length === 2) y = "20" + y;
  return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

/**
 * Intenta adivinar una fecha a partir del nombre de archivo, ej. "luis_06_agosto.docx"
 * o "06-08-2026.pdf". Es solo un valor por defecto — siempre editable en la revisión.
 */
export function guessFechaFromFileName(fileName, defaultYear) {
  const name = (fileName || "").toLowerCase().replace(/\.[a-z0-9]+$/, "");
  const numeric = name.match(/(\d{1,2})[\/\-_](\d{1,2})[\/\-_](\d{2,4})/);
  if (numeric) {
    let [_, a, b, y] = numeric;
    if (y.length === 2) y = "20" + y;
    return `${y}-${a.padStart(2, "0")}-${b.padStart(2, "0")}`;
  }
  const monthWord = Object.keys(MESES).find((m) => name.includes(m));
  if (monthWord) {
    const dayMatch = name.match(/(\d{1,2})/);
    const day = dayMatch ? dayMatch[1].padStart(2, "0") : "01";
    const year = defaultYear || new Date().getFullYear();
    return `${year}-${MESES[monthWord]}-${day}`;
  }
  return "";
}

/**
 * Lee un archivo .docx con la plantilla "Reporte Operativo Diario de Previos y Despachos"
 * (formato simplificado: encabezado = Tramitador + Aduana; sin columnas de horario).
 */
export async function parseDocxTemplate(file) {
  const buf = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer: buf });
  const doc = new DOMParser().parseFromString(result.value, "text/html");
  const tables = Array.from(doc.querySelectorAll("table"));
  if (tables.length < 3) {
    throw new Error(
      "El documento no tiene la estructura esperada (se encontraron " +
        tables.length +
        " tablas de 3 esperadas: encabezado, previos y despachos)."
    );
  }

  const tableRows = (table) =>
    Array.from(table.querySelectorAll("tr")).map((tr) =>
      Array.from(tr.querySelectorAll("td,th")).map(cellTextOf)
    );

  // --- encabezado: Tramitador / Aduana (y opcionalmente Total previos/despachos, que ignoramos) ---
  const headerMap = {};
  tableRows(tables[0]).forEach((row) =>
    row.forEach((cell) => {
      const m = cell.match(/^([^:]{2,30}):\s*(.*)$/s);
      if (m) {
        const key = normLabel(m[1]);
        if (key === "tramitador") headerMap.tramitador = m[2].trim();
        if (key === "aduana") headerMap.aduana = m[2].trim();
      }
    })
  );

  // --- tabla de previos: usa la fila de encabezado real para saber cuántas columnas trae ---
  const previosAllRows = tableRows(tables[1]);
  const previosHeaderRow = previosAllRows[0] || [];
  const previosDataRows = previosAllRows.slice(1);
  const previos = previosDataRows
    .filter((row) => row.some((c) => c))
    .map((row) => mapRowByHeaderNames(row, previosHeaderRow, PREVIO_FIELDS));

  // --- tabla de despachos ---
  const despachosAllRows = tableRows(tables[2]);
  const despachosHeaderRow = despachosAllRows[0] || [];
  const despachosDataRows = despachosAllRows.slice(1);
  const despachos = despachosDataRows
    .filter((row) => row.some((c) => c))
    .map((row) => mapRowByHeaderNames(row, despachosHeaderRow, DESPACHO_FIELDS));

  // --- otras actividades: párrafos/lista después de la tabla de despachos ---
  const bodyChildren = Array.from(doc.body.children);
  const idxDespachosTable = bodyChildren.indexOf(tables[2]);
  let otrasActividades = [];
  if (idxDespachosTable >= 0) {
    bodyChildren.slice(idxDespachosTable + 1).forEach((el) => {
      if (el.tagName === "OL" || el.tagName === "UL") {
        Array.from(el.querySelectorAll("li")).forEach((li) => {
          const t = cellTextOf(li);
          if (t) otrasActividades.push(t);
        });
      } else if (el.tagName === "TABLE") {
        // llegamos a una tabla de firmas u otra sección; detener
        return;
      } else {
        const t = cellTextOf(el).replace(/^\d+[\.\)]\s*/, "").trim();
        if (t && !/otras\s*actividades/i.test(t)) otrasActividades.push(t);
      }
    });
  }
  if (otrasActividades.length === 0) otrasActividades = ["", "", "", ""];

  return {
    fecha: "",
    tramitador: headerMap.tramitador || "",
    aduana: headerMap.aduana || "",
    previos,
    despachos,
    otrasActividades,
  };
}

/**
 * Empareja los valores de una fila con los campos esperados usando los nombres reales
 * de la fila de encabezado de la tabla (para tolerar columnas de más/menos, ej. si
 * el documento trae Inicio/Fin y nuestro esquema ya no los usa).
 */
function mapRowByHeaderNames(row, headerRow, fields) {
  const normHeaders = headerRow.map((h) => normLabel(h));
  const FIELD_LABEL_ALIASES = {
    ref: ["referencia", "ref"],
    guia: ["guia"],
    cliente: ["cliente"],
    almacen: ["almacen", "almacenes"],
    partidas: ["partidas"],
    tipoPrevio: ["tipodeprevio", "tipoprevio"],
    pedimento: ["pedimento"],
    tipoDespacho: ["tipodespacho"],
    resultado: ["resultado"],
  };
  const out = {};
  fields.forEach((f) => {
    const aliases = FIELD_LABEL_ALIASES[f.k] || [normLabel(f.label)];
    const idx = normHeaders.findIndex((h) => aliases.includes(h));
    out[f.k] = idx >= 0 ? row[idx] || "" : "";
  });
  return out;
}
