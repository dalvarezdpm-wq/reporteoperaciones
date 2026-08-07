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
function cleanSignaturePlaceholder(v) {
  return /^_+$/.test((v || "").replace(/\s/g, "")) ? "" : (v || "").trim();
}
export function parseFechaLoose(fecha) {
  const m = (fecha || "").match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return "";
  let [_, d, mo, y] = m;
  if (y.length === 2) y = "20" + y;
  return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

const HEADER_LABEL_MAP = {
  fecha: "fecha",
  coordinadorop: "coordinadorOp",
  coordinador: "coordinador",
  tramitador: "tramitador",
  aduana: "aduana",
  horainicio: "horaInicio",
  horacierre: "horaCierre",
};

/**
 * Lee un archivo .docx con la plantilla "Reporte Operativo Diario de Previos y Despachos"
 * y devuelve los datos estructurados, sin usar ningún servicio de IA.
 */
export async function parseDocxTemplate(file) {
  const buf = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer: buf });
  const doc = new DOMParser().parseFromString(result.value, "text/html");
  const tables = Array.from(doc.querySelectorAll("table"));
  if (tables.length < 4) {
    throw new Error(
      "El documento no tiene la estructura esperada (se encontraron " +
        tables.length +
        " tablas de 4 esperadas). ¿Es la plantilla 'Reporte Operativo Diario'?"
    );
  }
  const tableRows = (table) =>
    Array.from(table.querySelectorAll("tr")).map((tr) =>
      Array.from(tr.querySelectorAll("td,th")).map(cellTextOf)
    );

  const headerMap = {};
  tableRows(tables[0]).forEach((row) =>
    row.forEach((cell) => {
      const m = cell.match(/^([^:]{2,30}):\s*(.*)$/s);
      if (m) {
        const key = HEADER_LABEL_MAP[normLabel(m[1])];
        if (key) headerMap[key] = m[2].trim();
      }
    })
  );

  const previosRows = tableRows(tables[1]).slice(1);
  const previos = previosRows
    .filter((row) => row.some((c) => c))
    .map((row) => {
      const o = {};
      PREVIO_FIELDS.forEach((f, i) => (o[f.k] = row[i] || ""));
      return o;
    });

  const despachosRows = tableRows(tables[2]).slice(1);
  const despachos = despachosRows
    .filter((row) => row.some((c) => c))
    .map((row) => {
      const o = {};
      DESPACHO_FIELDS.forEach((f, i) => (o[f.k] = row[i] || ""));
      return o;
    });

  const bodyChildren = Array.from(doc.body.children);
  const idx2 = bodyChildren.indexOf(tables[2]);
  const idx3 = bodyChildren.indexOf(tables[3]);
  let otrasActividades = [];
  if (idx2 >= 0 && idx3 > idx2) {
    bodyChildren.slice(idx2 + 1, idx3).forEach((el) => {
      if (el.tagName === "OL" || el.tagName === "UL") {
        Array.from(el.querySelectorAll("li")).forEach((li) => {
          const t = cellTextOf(li);
          if (t) otrasActividades.push(t);
        });
      } else {
        const t = cellTextOf(el);
        if (t && !/otras\s*actividades/i.test(t)) otrasActividades.push(t);
      }
    });
  }
  if (otrasActividades.length === 0) otrasActividades = ["", "", "", ""];

  const footerRows = tableRows(tables[3]);
  const footerVals = footerRows[1] || [];
  const elaboro = cleanSignaturePlaceholder(footerVals[0]);
  const reviso = cleanSignaturePlaceholder(footerVals[1]);
  const autorizo = cleanSignaturePlaceholder(footerVals[2]);

  return {
    fecha: parseFechaLoose(headerMap.fecha),
    coordinadorOp: headerMap.coordinadorOp || "",
    coordinador: headerMap.coordinador || "",
    tramitador: headerMap.tramitador || "",
    horaInicio: headerMap.horaInicio || "",
    horaCierre: headerMap.horaCierre || "",
    aduana: headerMap.aduana || "",
    previos,
    despachos,
    otrasActividades,
    elaboro,
    reviso,
    autorizo,
  };
}
