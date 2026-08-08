import { db, ref, get, set, remove } from "./firebase.js";

/**
 * Estructura en Firebase (todo separado por aduana):
 * aduanas/{aduana}/reportes/{fecha}/capturas/[ {...} ]
 * aduanas/{aduana}/catalogos/{clientes|almacenes|tramitadores|aduanas} -> array de strings
 * aduanas/{aduana}/historico_excel -> { data, actualizado }
 *
 * Datos previos a la separación por aduana (antes de esta versión) vivían en la raíz:
 * reportes/{fecha}, catalogos/..., historico_excel. migrateLegacyDataIfNeeded() los
 * copia una sola vez hacia aduanas/GDL/ la primera vez que se usa esa aduana.
 */

function base(aduana) {
  return "aduanas/" + aduana;
}

export async function loadReports(aduana) {
  const snap = await get(ref(db, base(aduana) + "/reportes"));
  return snap.exists() ? snap.val() : {};
}

export async function saveReportDay(aduana, fecha, dayData) {
  await set(ref(db, base(aduana) + "/reportes/" + fecha), dayData);
}

export async function deleteReportDay(aduana, fecha) {
  await remove(ref(db, base(aduana) + "/reportes/" + fecha));
}

export async function loadCatalogos(aduana) {
  const snap = await get(ref(db, base(aduana) + "/catalogos"));
  return snap.exists() ? snap.val() : null;
}

export async function saveCatalogos(aduana, cats) {
  await set(ref(db, base(aduana) + "/catalogos"), cats);
}

/**
 * Guarda el Excel histórico completo como texto base64 dentro de Realtime Database
 * (no requiere Firebase Storage, que exige plan de pago). Se sobreescribe siempre
 * en la misma ruta, así que el link de descarga (dentro de la app) es siempre el mismo.
 */
export async function saveHistoricoExcel(aduana, base64, fechaActualizacion) {
  await set(ref(db, base(aduana) + "/historico_excel"), { data: base64, actualizado: fechaActualizacion });
}

export async function loadHistoricoExcel(aduana) {
  const snap = await get(ref(db, base(aduana) + "/historico_excel"));
  return snap.exists() ? snap.val() : null;
}

/**
 * Asignaciones: aduanas/{aduana}/asignaciones/{id} -> {
 *   tipo: "previo"|"despacho", guia, cliente, almacen, tramitador,
 *   estatus: "pendiente"|"completada", creadoPor, fechaCreacion,
 *   fechaCompletado, fechaCaptura
 * }
 */
export async function loadAsignaciones(aduana) {
  const snap = await get(ref(db, base(aduana) + "/asignaciones"));
  return snap.exists() ? snap.val() : {};
}

export async function saveAsignacion(aduana, id, data) {
  await set(ref(db, base(aduana) + "/asignaciones/" + id), data);
}

export async function deleteAsignacion(aduana, id) {
  await remove(ref(db, base(aduana) + "/asignaciones/" + id));
}

/**
 * Migración de un solo uso: si "aduanas/{aduana}/reportes" está vacío pero existen
 * datos en la raíz (de antes de separar por aduana), los copia hacia esa aduana.
 * No borra los datos originales de la raíz, solo los copia — así no hay riesgo de pérdida.
 */
export async function migrateLegacyDataIfNeeded(aduana) {
  const already = await get(ref(db, base(aduana) + "/reportes"));
  if (already.exists()) return false; // ya tiene datos propios, no migrar

  const legacyReports = await get(ref(db, "reportes"));
  const legacyCatalogos = await get(ref(db, "catalogos"));
  const legacyHistorico = await get(ref(db, "historico_excel"));

  let migrated = false;
  if (legacyReports.exists()) {
    await set(ref(db, base(aduana) + "/reportes"), legacyReports.val());
    migrated = true;
  }
  if (legacyCatalogos.exists()) {
    await set(ref(db, base(aduana) + "/catalogos"), legacyCatalogos.val());
    migrated = true;
  }
  if (legacyHistorico.exists()) {
    await set(ref(db, base(aduana) + "/historico_excel"), legacyHistorico.val());
    migrated = true;
  }
  return migrated;
}
