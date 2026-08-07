import * as pdfjsLib from "pdfjs-dist/build/pdf";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.js?url";
import { PREVIO_FIELDS, DESPACHO_FIELDS } from "./fields.js";
import { parseFechaLoose } from "./parseDocx.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

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
 * Lee un PDF con texto real (exportado desde la plantilla Word) reconstruyendo
 * las tablas por posición del texto en la página. No funciona con PDFs escaneados.
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
    const TOL = 3;
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
  const idxElaboro = lines.findIndex((l) => /elabor[oó]/i.test(lineText(l)));

  if (idxPrevios < 0 || idxDespachos < 0) {
    throw new Error(
      "No se reconoció la estructura del PDF (no se encontraron las secciones de previos/despachos). Verifica que sea la plantilla correcta, o usa el Word."
    );
  }

  const headerText = lines.slice(0, idxPrevios).map(lineText).join(" ");
  const headerMap = {};
  const LABELS =
    "Fecha|Coordinador op|Coordinador|Tramitador|Aduana|Hora inicio|Hora cierre|Total previos|Total despachos";
  const labelRegex = new RegExp(
    "(" + LABELS + ")\\s*:?\\s*([^:]*?)(?=(" + LABELS + ")\\s*:|$)",
    "gi"
  );
  let m;
  while ((m = labelRegex.exec(headerText))) {
    const key = HEADER_LABEL_MAP[normLabel(m[1])];
    if (key && !headerMap[key]) headerMap[key] = m[2].trim();
  }

  function extractTable(startIdx, endIdx, fields) {
    const headerLine = lines[startIdx + 1];
    if (!headerLine) return [];
    const headerItems = headerLine.items;
    const bounds = [];
    for (let i = 0; i < headerItems.length - 1; i++)
      bounds.push((headerItems[i].x + headerItems[i + 1].x) / 2);
    const dataLines = lines.slice(startIdx + 2, endIdx);
    const rows = [];
    dataLines.forEach((line) => {
      const cols = Array.from({ length: fields.length }, () => []);
      line.items.forEach((it) => {
        let colIdx = 0;
        for (let i = 0; i < bounds.length; i++) {
          if (it.x >= bounds[i]) colIdx = i + 1;
        }
        if (colIdx >= fields.length) colIdx = fields.length - 1;
        cols[colIdx].push(it.str);
      });
      const row = {};
      fields.forEach((f, i) => (row[f.k] = cols[i].join(" ").trim()));
      if (Object.values(row).some((v) => v)) rows.push(row);
    });
    return rows;
  }

  const previos = extractTable(idxPrevios, idxDespachos, PREVIO_FIELDS);
  const despachosEnd = idxActividades >= 0 ? idxActividades : idxElaboro >= 0 ? idxElaboro : lines.length;
  const despachos = extractTable(idxDespachos, despachosEnd, DESPACHO_FIELDS);

  let otrasActividades = [];
  if (idxActividades >= 0) {
    const end = idxElaboro >= 0 ? idxElaboro : lines.length;
    lines.slice(idxActividades + 1, end).forEach((l) => {
      const t = lineText(l).replace(/^\d+[\.\)]\s*/, "").trim();
      if (t) otrasActividades.push(t);
    });
  }
  if (otrasActividades.length === 0) otrasActividades = ["", "", "", ""];

  let elaboro = "",
    reviso = "",
    autorizo = "";
  if (idxElaboro >= 0) {
    const footerVals = lineText(lines[idxElaboro + 1] || { items: [] })
      .split(/\s{2,}/)
      .map((s) => s.trim())
      .filter(Boolean);
    elaboro = cleanSignaturePlaceholder(footerVals[0] || "");
    reviso = cleanSignaturePlaceholder(footerVals[1] || "");
    autorizo = cleanSignaturePlaceholder(footerVals[2] || "");
  }

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
