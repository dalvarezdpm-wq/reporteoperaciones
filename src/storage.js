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
