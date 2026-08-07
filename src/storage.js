import { db, ref, get, set, remove } from "./firebase.js";

/**
 * Estructura en Firebase:
 * reportes/{fecha}/capturas/[ { ...datos del reporte, sourceFileName, horaCaptura, uploadedBy } ]
 * catalogos/{clientes|almacenes|coordinadores|tramitadores} -> array de strings
 */

export async function loadReports() {
  const snap = await get(ref(db, "reportes"));
  return snap.exists() ? snap.val() : {};
}

export async function saveReportDay(fecha, dayData) {
  await set(ref(db, "reportes/" + fecha), dayData);
}

export async function deleteReportDay(fecha) {
  await remove(ref(db, "reportes/" + fecha));
}

export async function loadCatalogos() {
  const snap = await get(ref(db, "catalogos"));
  return snap.exists() ? snap.val() : null;
}

export async function saveCatalogos(cats) {
  await set(ref(db, "catalogos"), cats);
}

/**
 * Guarda el Excel histórico completo como texto base64 dentro de Realtime Database
 * (no requiere Firebase Storage, que exige plan de pago). Se sobreescribe siempre
 * en la misma ruta, así que el link de descarga (dentro de la app) es siempre el mismo.
 */
export async function saveHistoricoExcel(base64, fechaActualizacion) {
  await set(ref(db, "historico_excel"), { data: base64, actualizado: fechaActualizacion });
}

export async function loadHistoricoExcel() {
  const snap = await get(ref(db, "historico_excel"));
  return snap.exists() ? snap.val() : null;
}
