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

/** Bandera permanente para la reparación de una sola vez de "Ejecutivo" en renglones
 * viejos (ver repararEjecutivosSiNecesario en main.js) — una vez hecha, nunca se vuelve
 * a intentar, ni aunque después se agreguen más reportes viejos sin ese dato. */
export async function ejecutivosYaReparados(aduana) {
  const snap = await get(ref(db, base(aduana) + "/_ejecutivosReparados"));
  return snap.exists();
}
export async function marcarEjecutivosReparados(aduana) {
  await set(ref(db, base(aduana) + "/_ejecutivosReparados"), true);
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
 * Migración de un solo uso: si esta aduana nunca se ha usado (sin rastro de reportes,
 * catálogos, ni excel histórico) copia los datos previos a la separación por aduana
 * hacia ella. Usa una bandera permanente ("_migrated") para no repetirse NUNCA más,
 * incluso si después se borran todos los reportes de esa aduana (antes este chequeo
 * se basaba solo en "¿hay reportes?", y al borrar el último reporte volvía a
 * resucitar los datos viejos por error — ya corregido).
 */
export async function migrateLegacyDataIfNeeded(aduana) {
  const flagSnap = await get(ref(db, base(aduana) + "/_migrated"));
  if (flagSnap.exists()) return false; // ya se resolvió esto antes, no tocar nunca más

  const [reportesSnap, catalogosSnap, historicoSnap] = await Promise.all([
    get(ref(db, base(aduana) + "/reportes")),
    get(ref(db, base(aduana) + "/catalogos")),
    get(ref(db, base(aduana) + "/historico_excel")),
  ]);
  const yaTieneUso = reportesSnap.exists() || catalogosSnap.exists() || historicoSnap.exists();

  if (yaTieneUso) {
    // Esta aduana ya se ha usado (aunque ahora mismo tenga los reportes vacíos por
    // un borrado) — NO copiar nada encima, solo marcar como resuelto y salir.
    await set(ref(db, base(aduana) + "/_migrated"), true);
    return false;
  }

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
  await set(ref(db, base(aduana) + "/_migrated"), true);
  return migrated;
}
