import * as pdfjsLib from "pdfjs-dist/build/pdf";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.js?url";
import { PREVIO_FIELDS, DESPACHO_FIELDS } from "./fields.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

function normLabel(l) {
  return l
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

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

/**
 * Lee un PDF con texto real (exportado desde la plantilla Word) reconstruyendo
 * las tablas por posición del texto en la página. No funciona con PDFs escaneados.
 * Tolera columnas de más/menos comparando por el nombre real del encabezado de cada tabla.
 */
export async function parsePdfTemplate(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  let lines = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items
      .map((it) => ({ str: it.str, x: it.transform[4], y: it.transform[5] }))
      .filter((it) => it.str && it.str.trim());
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const TOL = 4;
    let current = null;
    items.forEach((it) => {
      if (current && Math.abs(it.y - current.y) <= TOL) current.items.push(it);
      else {
        current = { y: it.y, items: [it] };
        lines.push(current);
      }
    });
  }
  lines.forEach((l) => l.items.sort((a, b) => a.x - b.x));
  if (lines.length < 5) {
    throw new Error(
      "El PDF no tiene texto legible extraíble (¿es una hoja escaneada?). Usa el archivo Word en su lugar."
    );
  }

  const lineText = (l) => l.items.map((it) => it.str).join(" ").replace(/\s+/g, " ").trim();

  const idxPrevios = lines.findIndex((l) => /control\s*de\s*previos/i.test(lineText(l)));
  const idxDespachos = lines.findIndex((l) => /control\s*de\s*despachos/i.test(lineText(l)));
  const idxActividades = lines.findIndex((l) => /otras\s*actividades/i.test(lineText(l)));

  if (idxPrevios < 0 || idxDespachos < 0) {
    throw new Error(
      "No se reconoció la estructura del PDF (no se encontraron las secciones de previos/despachos). Verifica que sea la plantilla correcta, o usa el Word."
    );
  }

  const headerText = lines.slice(0, idxPrevios).map(lineText).join(" ");
  let tramitador = "";
  let aduana = "";
  const mTram = headerText.match(/tramitador\s*:?\s*([^:]*?)(?=(aduana)\s*:|$)/i);
  if (mTram) tramitador = mTram[1].trim();
  const mAduana = headerText.match(/aduana\s*:?\s*([^:]*?)(?=(total)\s*:|$)/i);
  if (mAduana) aduana = mAduana[1].trim();

  function extractTable(startIdx, endIdx, fields) {
    let headerLine = lines[startIdx + 1];
    if (!headerLine) return [];

    // 1) Fusionar palabras del encabezado que están muy cerca entre sí en la misma línea
    //    (ej. "Tipo" "de" "previo" -> "Tipo de previo"), para no crear columnas de más.
    const MERGE_THRESH = 35;
    const sortedHeaderItems = [...headerLine.items].sort((a, b) => a.x - b.x);
    let headerItems = [{ ...sortedHeaderItems[0] }];
    let lastX = sortedHeaderItems[0].x;
    for (let i = 1; i < sortedHeaderItems.length; i++) {
      const it = sortedHeaderItems[i];
      if (it.x - lastX < MERGE_THRESH) {
        headerItems[headerItems.length - 1].str += " " + it.str;
      } else {
        headerItems.push({ ...it });
      }
      lastX = it.x;
    }

    // 2) La línea inmediatamente después del encabezado puede ser en realidad la
    //    continuación de una palabra de columna que se partió en dos renglones
    //    (ej. "Tipo" en una línea y "despacho" en la siguiente). Una fila de datos
    //    real siempre trae varias columnas llenas; una continuación de encabezado
    //    trae solo una o dos palabras sueltas. Si detectamos ese patrón, la fusionamos
    //    al encabezado y avanzamos el inicio de los datos una línea más.
    let dataStartOffset = 2;
    const maybeContinuation = lines[startIdx + 2];
    if (maybeContinuation && maybeContinuation.items.length <= 2 && maybeContinuation.items.length < Math.ceil(fields.length / 2)) {
      maybeContinuation.items.forEach((it) => {
        let nearest = headerItems[0];
        let bestDist = Math.abs(nearest.x - it.x);
        headerItems.forEach((h) => {
          const dist = Math.abs(h.x - it.x);
          if (dist < bestDist) {
            bestDist = dist;
            nearest = h;
          }
        });
        nearest.str += " " + it.str;
      });
      dataStartOffset = 3;
    }

    const combinedHeaderNames = headerItems.map((it) => normLabel(it.str));
    const bounds = [];
    for (let i = 0; i < headerItems.length - 1; i++) bounds.push((headerItems[i].x + headerItems[i + 1].x) / 2);

    const dataLines = lines.slice(startIdx + dataStartOffset, endIdx);
    const rows = [];
    dataLines.forEach((line) => {
      const cols = Array.from({ length: headerItems.length }, () => []);
      line.items.forEach((it) => {
        let colIdx = 0;
        for (let i = 0; i < bounds.length; i++) {
          if (it.x >= bounds[i]) colIdx = i + 1;
        }
        if (colIdx >= headerItems.length) colIdx = headerItems.length - 1;
        cols[colIdx].push(it.str);
      });
      const rawRow = cols.map((c) => c.join(" ").trim());
      if (!rawRow.some((v) => v)) return;
      const row = {};
      fields.forEach((f) => {
        const aliases = FIELD_LABEL_ALIASES[f.k] || [normLabel(f.label)];
        const idx = combinedHeaderNames.findIndex((h) => aliases.includes(h));
        row[f.k] = idx >= 0 ? rawRow[idx] || "" : "";
      });
      rows.push(row);
    });
    return rows;
  }

  const previos = extractTable(idxPrevios, idxDespachos, PREVIO_FIELDS);
  const despachosEnd = idxActividades >= 0 ? idxActividades : lines.length;
  const despachos = extractTable(idxDespachos, despachosEnd, DESPACHO_FIELDS);

  let otrasActividades = [];
  if (idxActividades >= 0) {
    lines.slice(idxActividades + 1).forEach((l) => {
      const t = lineText(l).replace(/^\d+[\.\)]\s*/, "").trim();
      if (t) otrasActividades.push(t);
    });
  }
  if (otrasActividades.length === 0) otrasActividades = ["", "", "", ""];

  return {
    fecha: "",
    tramitador,
    aduana,
    previos,
    despachos,
    otrasActividades,
  };
}
