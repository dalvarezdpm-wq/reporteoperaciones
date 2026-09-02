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

/** Firebase Realtime Database rechaza POR COMPLETO cualquier escritura que tenga un
 * valor "undefined" en cualquier propiedad (sin importar qué tan profundo esté) — y
 * rechaza el objeto entero, no solo esa propiedad puntual. Esta función limpia
 * cualquier "undefined" antes de guardar (un viaje por JSON los quita solos, ya que
 * JSON.stringify omite las propiedades con ese valor), para que un dato viejo o
 * incompleto en algún campo nunca vuelva a tumbar un guardado completo. */
function limpiarUndefined(obj) {
  return JSON.parse(JSON.stringify(obj));
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

/**
 * Respaldos automáticos: aduanas/{aduana}/_respaldos/{fecha} -> copia completa de
 * reportes + catálogos + asignaciones de ese momento. aduanas/{aduana}/_ultimoRespaldo
 * guarda el timestamp del más reciente, para no repetir el respaldo más de una vez al
 * día sin tener que listar todos los que ya existen cada vez.
 */
export async function obtenerUltimoRespaldo(aduana) {
  const snap = await get(ref(db, base(aduana) + "/_ultimoRespaldo"));
  return snap.exists() ? snap.val() : null;
}
export async function guardarRespaldo(aduana, clave, datos) {
  await set(ref(db, base(aduana) + "/_respaldos/" + clave), limpiarUndefined(datos));
  await set(ref(db, base(aduana) + "/_ultimoRespaldo"), Date.now());
}
export async function listarRespaldos(aduana) {
  const snap = await get(ref(db, base(aduana) + "/_respaldos"));
  if (!snap.exists()) return [];
  return Object.keys(snap.val()).sort();
}
export async function descargarRespaldo(aduana, clave) {
  const snap = await get(ref(db, base(aduana) + "/_respaldos/" + clave));
  return snap.exists() ? snap.val() : null;
}
export async function eliminarRespaldo(aduana, clave) {
  await remove(ref(db, base(aduana) + "/_respaldos/" + clave));
}

export async function loadReports(aduana) {
  const snap = await get(ref(db, base(aduana) + "/reportes"));
  if (!snap.exists()) return {};
  const raw = snap.val();
  const out = {};
  Object.entries(raw).forEach(([fecha, day]) => {
    // "capturas" puede venir en dos formatos: arreglo (formato viejo, de cuando se
    // guardaba el día completo de un jalón) u objeto por ID de captura (formato nuevo,
    // de guardar cada hoja por separado — ver saveCapturaUnica). Se acepta cualquiera
    // de los dos al leer, y siempre se entrega como arreglo hacia el resto de la app.
    let capturasArray;
    if (Array.isArray(day.capturas)) {
      capturasArray = day.capturas;
    } else if (day.capturas && typeof day.capturas === "object") {
      capturasArray = Object.values(day.capturas);
    } else {
      capturasArray = [];
    }
    out[fecha] = { ...day, fecha, capturas: capturasArray };
  });
  return out;
}

/** Guarda UNA sola hoja (captura) de un tramitador, en su propio espacio dentro del
 * día — SIN tocar las demás hojas de ese mismo día. Esto es justo lo que evita que dos
 * personas guardando casi al mismo tiempo se borren el trabajo una a la otra: antes,
 * guardar el reporte de una persona reescribía el DÍA COMPLETO con lo que esa persona
 * tuviera en su propia memoria en ese momento — si su copia local no incluía la hoja
 * que alguien más ACABABA de guardar (algo que pasa fácil si llevaba un rato editando
 * sin recargar), esa hoja ajena desaparecía sin que nadie la hubiera tocado. Guardando
 * cada hoja en su propio nodo, un guardado nunca puede afectar los datos de otra hoja
 * que ni siquiera conoce.
 *
 * Los días guardados ANTES de este cambio todavía tienen sus hojas en un arreglo (el
 * formato viejo) — la primera vez que se toca uno de esos días, aquí mismo se convierte
 * a objeto por ID (incluyendo ya la hoja que se está guardando), en una sola escritura,
 * para no dejar mezclados los dos formatos ni crear una hoja duplicada por accidente.
 * Después de esa primera vez, ese día ya queda en el formato nuevo para siempre. */
export async function saveCapturaUnica(aduana, fecha, captura) {
  const rutaDia = base(aduana) + "/reportes/" + fecha;
  const snap = await get(ref(db, rutaDia + "/capturas"));
  if (snap.exists() && Array.isArray(snap.val())) {
    const viejoArray = snap.val();
    const porId = {};
    viejoArray.forEach((c) => {
      if (c && c.id) porId[c.id] = c;
    });
    porId[captura.id] = limpiarUndefined(captura);
    await set(ref(db, rutaDia + "/capturas"), porId);
  } else {
    await set(ref(db, rutaDia + "/capturas/" + captura.id), limpiarUndefined(captura));
  }
}

/** Borra UNA sola hoja de un día, sin tocar las demás — mismo motivo que
 * saveCapturaUnica: nunca se reescribe el día completo a ciegas. Si el día todavía
 * está en el formato viejo (arreglo), se convierte al nuevo (por ID) excluyendo ya la
 * hoja que se está borrando, en la misma escritura. */
export async function deleteCapturaUnica(aduana, fecha, capturaId) {
  const rutaDia = base(aduana) + "/reportes/" + fecha;
  const snap = await get(ref(db, rutaDia + "/capturas"));
  if (snap.exists() && Array.isArray(snap.val())) {
    const viejoArray = snap.val();
    const porId = {};
    viejoArray.forEach((c) => {
      if (c && c.id && c.id !== capturaId) porId[c.id] = c;
    });
    await set(ref(db, rutaDia + "/capturas"), porId);
  } else {
    await remove(ref(db, rutaDia + "/capturas/" + capturaId));
  }
}

export async function saveReportDay(aduana, fecha, dayData) {
  await set(ref(db, base(aduana) + "/reportes/" + fecha), limpiarUndefined(dayData));
}

export async function deleteReportDay(aduana, fecha) {
  await remove(ref(db, base(aduana) + "/reportes/" + fecha));
}

export async function loadCatalogos(aduana) {
  const snap = await get(ref(db, base(aduana) + "/catalogos"));
  return snap.exists() ? snap.val() : null;
}

export async function saveCatalogos(aduana, cats) {
  await set(ref(db, base(aduana) + "/catalogos"), limpiarUndefined(cats));
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
  await set(ref(db, base(aduana) + "/asignaciones/" + id), limpiarUndefined(data));
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
