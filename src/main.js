import "./style.css";
import ExcelJS from "exceljs";
import { PREVIO_FIELDS, DESPACHO_FIELDS, REVALIDADA_FIELDS, PENDIENTE_FIELDS, emptyCaptura } from "./fields.js";
import { parseDocxTemplate, guessFechaFromFileName } from "./parseDocx.js";
import { parsePdfTemplate } from "./parsePdf.js";
import {
  loadReports as fbLoadReports,
  saveReportDay as fbSaveReportDay,
  deleteReportDay as fbDeleteReportDay,
  loadCatalogos as fbLoadCatalogos,
  saveCatalogos as fbSaveCatalogos,
  saveHistoricoExcel,
  loadHistoricoExcel,
  migrateLegacyDataIfNeeded,
  loadAsignaciones as fbLoadAsignaciones,
  saveAsignacion as fbSaveAsignacion,
  deleteAsignacion as fbDeleteAsignacion,
} from "./storage.js";

const USERS = ["Administrador", "Coordinación", "Trámite", "Ejecutivo"];
const TRAMITADORES = ["Monica Ortega", "Luis Arreola", "Mariana Carrillo", "Mayra Romero", "Javier Garcia", "Julio Regalado"];
const ADUANAS = ["GDL", "TOLUCA"];
const ADMIN_PASSWORD = "ow2026";
const APP_VERSION = "4.7.0";

let editableCats = {
  clientes: ["GLXI", "Alkaps", "MTI", "FMI", "IndoUnión", "Foray", "Alpha metal", "BRP", "PMI"],
  almacenes: ["228", "277", "CLA", "WTC"],
  tramitadores: [...TRAMITADORES],
  aduanas: ["GDL"],
};

/**
 * Rellena con arreglos vacíos (o valores por defecto razonables) cualquier clave
 * que falte en un catálogo cargado de Firebase o de la caché local — esto evita que
 * la app truene si el catálogo se guardó con una versión anterior que no tenía
 * todas las categorías actuales (ej. "tramitadores" se agregó después).
 */
function sanitizeCats(cats) {
  // El catálogo de tramitadores de arranque (TRAMITADORES) es el equipo de GDL —
  // solo debe usarse como semilla ahí. Una aduana nueva (ej. Toluca) debe empezar
  // con su catálogo vacío, para que Coordinación dé de alta a SU propio equipo,
  // sin heredar por accidente los nombres de otra aduana.
  const seedTramitadores = state.aduanaActiva === "GDL" ? [...TRAMITADORES] : [];
  const base = { clientes: [], almacenes: [], tramitadores: seedTramitadores, aduanas: [...ADUANAS] };
  const out = { ...base, ...(cats || {}) };
  Object.keys(base).forEach((k) => {
    if (!Array.isArray(out[k])) out[k] = base[k];
  });
  return out;
}

/**
 * Garantiza que una captura (nueva o cargada de un reporte guardado con una versión
 * anterior de la app) siempre tenga todos sus arreglos presentes — evita que funciones
 * como "Traer asignaciones" o "Marcar Listo" fallen en silencio con reportes viejos
 * a los que les falte algún campo (ej. "pendientes", agregado en una versión posterior).
 */
function normalizeCaptura(c) {
  if (!Array.isArray(c.previos)) c.previos = [];
  if (!Array.isArray(c.despachos)) c.despachos = [];
  if (!Array.isArray(c.revalidadas)) c.revalidadas = [];
  if (!Array.isArray(c.pendientes)) c.pendientes = [];
  if (!Array.isArray(c.otrasActividades)) c.otrasActividades = ["", "", "", ""];
  if (!Array.isArray(c.historial)) c.historial = [];
  return c;
}

/**
 * Genera la contraseña automática de cada tramitador: 3 primeras letras del nombre
 * (en minúsculas) + "2026". Si dos nombres coinciden en esas 3 letras, al segundo
 * (y siguientes) se les usa una letra más, para que nunca dos personas compartan
 * la misma contraseña sin que el administrador tenga que hacer nada manual.
 */
function tramitadorPasswordsMap(lista) {
  const map = {};
  const usadas = new Set();
  (lista || []).forEach((nombre) => {
    const letras = (nombre || "").trim().toLowerCase().replace(/[^a-záéíóúñ]/g, "");
    let len = 3;
    let pass = letras.slice(0, len) + "2026";
    while (usadas.has(pass) && len < letras.length) {
      len++;
      pass = letras.slice(0, len) + "2026";
    }
    usadas.add(pass);
    map[nombre] = pass;
  });
  return map;
}
function tramitadorPassword(nombre) {
  return tramitadorPasswordsMap(editableCats.tramitadores)[nombre] || "";
}

let state = {
  user: null,
  userRole: null, // "admin" | "tramite" | "ejecutivo"
  aduanaActiva: null,
  view: "selectAduana",
  reports: {},
  asignaciones: {},
  loading: true,
  currentCaptura: null,
  editingIndex: null,
  editingFecha: null,
  detailFecha: null,
  errorMsg: "",
  processingMsg: "",
  historicoExcelUpdatedAt: null,
  syncingExcel: false,
  installPromptEvent: null,
  installBannerDismissed: false,
  isOnline: navigator.onLine,
  syncingPending: false,
  usingLocalCache: false,
  expandedClientesRef: {},
  expandedHistorial: {},
  asigDraft: { tipo: "previo", guia: "", cliente: "", almacen: "", tramitador: "" },
  notifBannerDismissed: false,
  recentNotifications: [],
  showNotifPanel: false,
  showMisAsignacionesPanel: false,
  kpiDesde: "",
  kpiHasta: "",
  clienteDesde: "",
  clienteHasta: "",
  reportesAnio: "",
  reportesMes: "",
  bulkCatOpen: {},
  bulkCatMsg: "",
  bulkCatMsgKey: "",
  catFilter: {},
  searchQuery: "",
};

const root = document.getElementById("app-root");

function esc(s) {
  return s === undefined || s === null
    ? ""
    : String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
/** Fecha local (no UTC) de hoy menos "diasAtras" días, en formato YYYY-MM-DD.
 * Centralizado aquí para que ningún cálculo de fechas use accidentalmente
 * new Date().toISOString() directamente (eso da la fecha en UTC, que en México
 * ya marca el día siguiente por las noches). */
function dateStrLocal(diasAtras = 0) {
  const d = new Date();
  d.setDate(d.getDate() - diasAtras);
  const local = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}
function todayStr() {
  return dateStrLocal(0);
}
function fmtDateHuman(iso) {
  if (!iso) return "Sin fecha";
  const [y, m, d] = iso.split("-");
  const meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${d} ${meses[parseInt(m, 10) - 1]} ${y}`;
}

// ---------- persistence wrappers ----------
// ---------- caché local (para leer datos sin conexión) ----------
function localCacheKey() {
  return "ow_local_cache_" + (state.aduanaActiva || "sin_aduana");
}
function persistLocalCache() {
  try {
    localStorage.setItem(localCacheKey(), JSON.stringify({ reports: state.reports, catalogos: editableCats }));
  } catch (e) {}
}
function loadLocalCache() {
  try {
    return JSON.parse(localStorage.getItem(localCacheKey()) || "{}");
  } catch (e) {
    return {};
  }
}
/** Superpone cualquier edición pendiente (guardada sin señal) sobre los datos cargados,
 * para que nunca se pierda ni se tape un cambio local todavía no sincronizado. */
function applyPendingOverlay() {
  const map = getPendingMap();
  Object.entries(map).forEach(([fecha, entry]) => {
    if (entry.data === null) delete state.reports[fecha];
    else state.reports[fecha] = entry.data;
  });
}

async function loadReports() {
  if (!navigator.onLine) {
    const cached = loadLocalCache();
    state.reports = cached.reports || {};
    if (cached.catalogos) editableCats = sanitizeCats(cached.catalogos);
    state.usingLocalCache = true;
  } else {
    try {
      await migrateLegacyDataIfNeeded(state.aduanaActiva);
    } catch (e) {}
    try {
      state.reports = await fbLoadReports(state.aduanaActiva);
      state.usingLocalCache = false;
    } catch (e) {
      const cached = loadLocalCache();
      state.reports = cached.reports || {};
      state.usingLocalCache = true;
      state.errorMsg = "No se pudo conectar con Firebase — mostrando la última copia guardada en este dispositivo (puede no ser la más reciente).";
    }
    try {
      const cats = await fbLoadCatalogos(state.aduanaActiva);
      if (cats) editableCats = sanitizeCats(cats);
    } catch (e) {
      const cached = loadLocalCache();
      if (cached.catalogos) editableCats = sanitizeCats(cached.catalogos);
    }
  }
  applyPendingOverlay();
  persistLocalCache();
  try {
    if (navigator.onLine) {
      const hist = await loadHistoricoExcel(state.aduanaActiva);
      if (hist) state.historicoExcelUpdatedAt = hist.actualizado;
    }
  } catch (e) {}
  try {
    if (navigator.onLine) {
      const nuevasAsignaciones = await fbLoadAsignaciones(state.aduanaActiva);
      notificarCambiosAsignaciones(nuevasAsignaciones);
      state.asignaciones = nuevasAsignaciones;
    }
  } catch (e) {}
  state.loading = false;
  render();
  if (!state.historicoExcelUpdatedAt && navigator.onLine && Object.keys(state.reports).length > 0) {
    syncHistoricoExcel();
  }
}

/**
 * Refresca reportes y asignaciones en silencio (sin spinner, sin interrumpir) para
 * que los cambios que haga otra persona (ej. un borrado del coordinador, o una
 * asignación nueva) se reflejen sin que cada quien tenga que cerrar y volver a
 * entrar. Respeta cualquier edición local todavía no sincronizada (applyPendingOverlay)
 * para no perder nada. Las asignaciones se refrescan SIEMPRE (incluso a media
 * captura), porque no interfieren con lo que se está editando; los reportes solo
 * se refrescan si no se está en plena captura, para no interrumpir esa pantalla.
 */
async function backgroundRefresh() {
  if (!navigator.onLine || !state.user || !state.aduanaActiva) return;
  if (state.loading) return;
  if (state.view !== "review") {
    try {
      state.reports = await fbLoadReports(state.aduanaActiva);
      applyPendingOverlay();
      persistLocalCache();
    } catch (e) {
      return; // si falla, se queda con lo que ya tenía, sin avisar (no es crítico)
    }
  }
  try {
    const nuevasAsignaciones = await fbLoadAsignaciones(state.aduanaActiva);
    notificarCambiosAsignaciones(nuevasAsignaciones);
    state.asignaciones = nuevasAsignaciones;
  } catch (e) {}
  render();
}

/**
 * Compara las asignaciones antes/después de un refresco y dispara notificaciones:
 * - Al tramitador, cuando le cae una asignación pendiente nueva.
 * - Al Administrador/Coordinación, cuando una asignación que estaba pendiente se completa.
 */
function seenAsignacionesKey() {
  return "ow_seen_asignaciones_" + (state.aduanaActiva || "x") + "_" + (state.user || "x");
}
function loadSeenAsignaciones() {
  try {
    return JSON.parse(localStorage.getItem(seenAsignacionesKey()) || "null");
  } catch (e) {
    return null;
  }
}
function saveSeenAsignaciones(map) {
  try {
    localStorage.setItem(seenAsignacionesKey(), JSON.stringify(map));
  } catch (e) {}
}

/**
 * Revisa qué cambió en las asignaciones desde la ÚLTIMA vez que este usuario, en este
 * dispositivo, las vio — usando un rastro guardado en localStorage (no solo en memoria),
 * para que funcione también al cerrar sesión y volver a entrar más tarde, no solo
 * mientras la app se queda abierta. La primera vez que se usa (sin rastro guardado
 * todavía) solo establece la línea base, sin notificar nada de golpe.
 */
function notificarCambiosAsignaciones(nuevas) {
  const anteriores = loadSeenAsignaciones();
  const nuevoMapa = {};
  Object.entries(nuevas || {}).forEach(([id, a]) => {
    nuevoMapa[id] = a.estatus;
    if (anteriores === null) return; // primera vez en este dispositivo: solo establecer línea base
    const antes = anteriores[id];
    if (!antes) {
      if (state.userRole === "tramite" && a.estatus === "pendiente" && (a.tramitador || "").trim() === (state.user || "").trim()) {
        dispararNotificacion("📥 Nueva asignación", `Guía ${a.guia} (${a.tipo}) — asignada por ${a.creadoPor || "Coordinación"}`);
      }
    } else if (antes === "pendiente" && a.estatus === "completada" && state.userRole === "admin") {
      dispararNotificacion("✅ Asignación completada", `Guía ${a.guia} — completada por ${a.tramitador}`);
    }
  });
  saveSeenAsignaciones(nuevoMapa);
}

/** Muestra la notificación del navegador (si el permiso está concedido) y siempre
 * la guarda en el historial reciente dentro de la app (funciona aunque el navegador
 * no tenga permiso, o no lo soporte). */
function dispararNotificacion(titulo, cuerpo) {
  state.recentNotifications = [{ titulo, cuerpo, ts: Date.now() }, ...state.recentNotifications].slice(0, 8);
  if (typeof Notification !== "undefined" && Notification.permission === "granted") {
    try {
      new Notification(titulo, { body: cuerpo, icon: "icon-192.png" });
    } catch (e) {}
  }
  render();
}

function notifBanner() {
  if (typeof Notification === "undefined") return "";
  if (Notification.permission !== "default") return "";
  if (state.notifBannerDismissed) return "";
  return `<div class="status-line" style="margin-bottom:16px;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
    <div>🔔 Activa las notificaciones para enterarte al instante de nuevas asignaciones y cambios.</div>
    <div style="display:flex;gap:8px;">
      <button class="btn btn-primary btn-sm" onclick="App.pedirPermisoNotificaciones()">Activar</button>
      <button class="btn btn-ghost btn-sm" onclick="App.dismissNotifBanner()">Ahora no</button>
    </div>
  </div>`;
}

function notifPanelHtml() {
  if (!state.showNotifPanel) return "";
  return `<div class="panel" style="margin-bottom:16px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <div style="font-weight:600;font-size:13.5px;">🔔 Notificaciones recientes</div>
      <button class="btn btn-ghost btn-sm" onclick="App.toggleNotifPanel()">Cerrar</button>
    </div>
    ${
      state.recentNotifications.length === 0
        ? `<div style="color:var(--muted);font-size:12.5px;">Sin notificaciones todavía. Aquí aparecerán las asignaciones nuevas y otros avisos.</div>`
        : state.recentNotifications
            .map(
              (n) => `<div style="padding:8px 0;border-bottom:1px solid var(--line);">
          <div style="font-weight:600;font-size:12.5px;">${esc(n.titulo)}</div>
          <div style="font-size:12px;color:var(--ink);">${esc(n.cuerpo)}</div>
          <div style="font-size:11px;color:var(--muted);">${new Date(n.ts).toLocaleString("es-MX")}</div>
        </div>`
            )
            .join("")
    }
  </div>`;
}

function misAsignacionesPanelHtml() {
  if (!state.showMisAsignacionesPanel) return "";
  const lista = misAsignacionesPendientes();
  return `<div class="panel" style="margin-bottom:16px;background:#FEF6F5;border-color:#F3C9C2;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <div style="font-weight:600;font-size:13.5px;">📥 Tus asignaciones pendientes</div>
      <button class="btn btn-ghost btn-sm" onclick="App.toggleMisAsignacionesPanel()">Cerrar</button>
    </div>
    <div style="color:var(--muted);font-size:12px;margin-bottom:10px;">Solo consulta — para agregarlas a tu reporte, entra a capturar y usa el botón "Traer asignaciones pendientes" en la sección de Pendientes.</div>
    ${
      lista.length === 0
        ? `<div style="color:var(--muted);font-size:12.5px;">No tienes asignaciones pendientes ahora mismo.</div>`
        : lista
            .map(
              (a) => `<div style="padding:8px 0;border-bottom:1px solid var(--line);">
          <div style="font-weight:600;font-size:12.5px;">
            <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span> Guía ${esc(a.guia)}
          </div>
          <div style="font-size:12px;color:var(--ink);">Cliente: ${esc(a.cliente) || "—"} · Almacén: ${esc(a.almacen) || "—"}</div>
          <div style="font-size:11px;color:var(--muted);">Asignado por ${esc(a.creadoPor) || "Coordinación"} el ${
                a.fechaCreacion ? new Date(a.fechaCreacion).toLocaleDateString("es-MX") : "—"
              }</div>
        </div>`
            )
            .join("")
    }
  </div>`;
}

/**
 * Revisa las asignaciones pendientes del tramitador de este reporte; si algún renglón
 * de previos/despachos que acaba de guardar coincide en Guía con una asignación
 * pendiente suya, la marca como completada automáticamente y la enlaza.
 */
async function autoCloseAsignaciones(captura) {
  const tramitador = (captura.tramitador || "").trim();
  if (!tramitador) return;
  const guiasPrevios = new Set((captura.previos || []).map((r) => (r.guia || "").trim()).filter(Boolean));
  const guiasDespachos = new Set((captura.despachos || []).map((r) => (r.guia || "").trim()).filter(Boolean));
  const guiasRevalidadas = new Set((captura.revalidadas || []).map((r) => (r.guia || "").trim()).filter(Boolean));
  const entries = Object.entries(state.asignaciones || {});
  for (const [id, a] of entries) {
    if (a.estatus !== "pendiente") continue;
    if ((a.tramitador || "").trim() !== tramitador) continue;
    const guiaSet = a.tipo === "despacho" ? guiasDespachos : a.tipo === "revalidada" ? guiasRevalidadas : guiasPrevios;
    if (!guiaSet.has((a.guia || "").trim())) continue;
    const updated = { ...a, estatus: "completada", fechaCompletado: new Date().toISOString() };
    state.asignaciones[id] = updated;
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, updated);
    } catch (e) {}
  }
}
/**
 * Regenera el Excel histórico completo y lo guarda como datos (base64) en Realtime
 * Database — no requiere Firebase Storage (que exige plan de pago). Se llama después
 * de guardar, editar o borrar cualquier reporte. Si falla, no interrumpe el guardado
 * normal, solo avisa; siempre queda disponible la descarga manual.
 */
async function syncHistoricoExcel() {
  state.syncingExcel = true;
  render();
  try {
    const days = Object.values(state.reports).sort((a, b) => a.fecha.localeCompare(b.fecha));
    const blob = await buildExcelBlob(days);
    const base64 = await blobToBase64(blob);
    const stamp = new Date().toISOString();
    persistLocalHistoricoExcel(base64, stamp); // respaldo local, funciona aunque falle Firebase
    await saveHistoricoExcel(state.aduanaActiva, base64, stamp);
    state.historicoExcelUpdatedAt = stamp;
  } catch (e) {
    state.errorMsg = "El reporte se guardó, pero no se pudo actualizar el Excel compartido en la nube (" + e.message + "). Se guardó una copia local mientras tanto.";
  }
  state.syncingExcel = false;
  render();
}
function localHistoricoKey() {
  return "ow_local_historico_excel_" + (state.aduanaActiva || "sin_aduana");
}
function persistLocalHistoricoExcel(base64, stamp) {
  try {
    localStorage.setItem(localHistoricoKey(), JSON.stringify({ data: base64, actualizado: stamp }));
  } catch (e) {}
}
function loadLocalHistoricoExcel() {
  try {
    return JSON.parse(localStorage.getItem(localHistoricoKey()) || "null");
  } catch (e) {
    return null;
  }
}
// ---------- cola de sincronización pendiente (guarda localmente si no hay señal) ----------
function pendingKey() {
  return "ow_pending_queue_" + (state.aduanaActiva || "sin_aduana");
}
function getPendingMap() {
  try {
    return JSON.parse(localStorage.getItem(pendingKey()) || "{}");
  } catch (e) {
    return {};
  }
}
function setPendingMap(map) {
  try {
    localStorage.setItem(pendingKey(), JSON.stringify(map));
  } catch (e) {}
}
function markPending(fecha, dayDataOrNullForDelete) {
  const map = getPendingMap();
  map[fecha] = { data: dayDataOrNullForDelete, timestamp: new Date().toISOString() };
  setPendingMap(map);
}
function clearPending(fecha) {
  const map = getPendingMap();
  delete map[fecha];
  setPendingMap(map);
}
function pendingFechas() {
  return Object.keys(getPendingMap());
}
function isPending(fecha) {
  return Object.prototype.hasOwnProperty.call(getPendingMap(), fecha);
}

/**
 * Intenta subir todos los reportes que quedaron pendientes por falta de señal.
 * Se llama automáticamente al detectar que regresó la conexión.
 */
async function processPendingQueue() {
  const map = getPendingMap();
  const fechas = Object.keys(map);
  if (fechas.length === 0) return;
  state.syncingPending = true;
  render();
  let anySucceeded = false;
  for (const fecha of fechas) {
    const entry = map[fecha];
    try {
      if (entry.data === null) {
        await fbDeleteReportDay(state.aduanaActiva, fecha);
      } else {
        await fbSaveReportDay(state.aduanaActiva, fecha, entry.data);
      }
      clearPending(fecha);
      anySucceeded = true;
    } catch (e) {
      // se deja en la cola, se reintentará la próxima vez que regrese la señal
    }
  }
  state.syncingPending = false;
  render();
  if (anySucceeded) syncHistoricoExcel();
}

async function saveReportDay(fecha) {
  if (!navigator.onLine) {
    markPending(fecha, state.reports[fecha]);
    return;
  }
  try {
    await fbSaveReportDay(state.aduanaActiva, fecha, state.reports[fecha]);
    clearPending(fecha);
  } catch (e) {
    markPending(fecha, state.reports[fecha]);
    throw e;
  }
}
async function deleteReportDay(fecha) {
  if (!navigator.onLine) {
    markPending(fecha, null);
    return;
  }
  try {
    await fbDeleteReportDay(state.aduanaActiva, fecha);
    clearPending(fecha);
  } catch (e) {
    markPending(fecha, null);
    throw e;
  }
}
async function saveCatalogos() {
  try {
    await fbSaveCatalogos(state.aduanaActiva, editableCats);
  } catch (e) {}
}
function loadLastUser() {
  try {
    const u = localStorage.getItem("ow_ultimo_usuario");
    const r = localStorage.getItem("ow_ultimo_rol");
    const a = localStorage.getItem("ow_ultima_aduana");
    if (u && r) {
      state.user = u;
      state.userRole = r;
    }
    if (a) state.aduanaActiva = a;
  } catch (e) {}
}
function saveLastUser(u, role) {
  try {
    localStorage.setItem("ow_ultimo_usuario", u);
    localStorage.setItem("ow_ultimo_rol", role);
  } catch (e) {}
}
function saveLastAduana(aduana) {
  try {
    localStorage.setItem("ow_ultima_aduana", aduana);
  } catch (e) {}
}

// ---------- stats ----------
function allCapturas() {
  const out = [];
  Object.values(state.reports).forEach((r) => (r.capturas || []).forEach((c) => out.push({ ...c, fecha: r.fecha })));
  return out;
}

/** Busca por Guía, Referencia o Cliente en todos los renglones (previos, despachos,
 * revalidadas, pendientes) de todos los reportes de la aduana activa. */
function buscarGlobal(query) {
  const q = (query || "").trim().toLowerCase();
  if (!q) return [];
  const results = [];
  allCapturas().forEach((c) => {
    [
      ...(c.previos || []).map((r) => ({ ...r, tabla: "Previo" })),
      ...(c.despachos || []).map((r) => ({ ...r, tabla: "Despacho" })),
      ...(c.revalidadas || []).map((r) => ({ ...r, tabla: "Guía revalidada" })),
      ...(c.pendientes || []).map((r) => ({ ...r, tabla: "Pendiente" })),
    ].forEach((r) => {
      const haystack = [r.guia, r.ref, r.cliente, r.pedimento].filter(Boolean).join(" ").toLowerCase();
      if (haystack.includes(q)) {
        results.push({ ...r, fecha: c.fecha, tramitador: c.tramitador, tabla: r.tabla });
      }
    });
  });
  return results.sort((a, b) => b.fecha.localeCompare(a.fecha));
}

/** Revisa si alguna Guía de este reporte (previos/despachos/revalidadas) ya existe en OTRO
 * reporte ya guardado (misma tabla), para avisar antes de guardar y evitar captura doble. */
function findGuiasDuplicadas(captura) {
  const dup = [];
  const gruposPorTipo = { previo: "previos", despacho: "despachos", revalidada: "revalidadas" };
  const checkTabla = (rows, tipo, numKey) => {
    rows.forEach((row) => {
      const guia = (row.guia || "").trim();
      if (!guia) return;
      if ((row[numKey] || "").trim()) return; // ya etiquetado a propósito como 2do/3er previo, etc. — no avisar
      allCapturas().forEach((c) => {
        if (c.id === captura.id) return; // no compararse contra sí mismo al editar
        const otras = c[gruposPorTipo[tipo]] || [];
        otras.forEach((r) => {
          if ((r.guia || "").trim() === guia) {
            dup.push({ guia, tipo, fecha: c.fecha, tramitador: c.tramitador });
          }
        });
      });
    });
  };
  checkTabla(captura.previos || [], "previo", "numPrevio");
  checkTabla(captura.despachos || [], "despacho", "numDespacho");
  checkTabla(captura.revalidadas || [], "revalidada", "numRevalidada");
  return dup;
}

function totalPrevios() {
  return allCapturas().reduce((s, c) => s + (c.previos ? c.previos.length : 0), 0);
}
function totalDespachos() {
  return allCapturas().reduce((s, c) => s + (c.despachos ? c.despachos.length : 0), 0);
}
function totalRevalidadas() {
  return allCapturas().reduce((s, c) => s + (c.revalidadas ? c.revalidadas.length : 0), 0);
}
function last14Days() {
  const days = [];
  for (let i = 13; i >= 0; i--) {
    days.push(dateStrLocal(i));
  }
  return days.map((iso) => {
    const r = state.reports[iso];
    return { label: iso.slice(8, 10) + "/" + iso.slice(5, 7), value: r ? (r.capturas || []).length : 0 };
  });
}
const MESES_NOMBRES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

/** Años en los que hay al menos un reporte guardado, más recientes primero — para
 * armar el selector "Año" en Reportes por fecha, sin tener que listar todos los
 * días de golpe cuando ya haya varios años de operación acumulados. */
function aniosDisponibles() {
  const set = new Set(Object.keys(state.reports).map((f) => f.slice(0, 4)));
  return [...set].sort((a, b) => b.localeCompare(a));
}

/** Meses (01-12) con al menos un reporte dentro del año dado (o de todos los años,
 * si no se especifica), más recientes primero. */
function mesesDisponibles(anio) {
  const set = new Set(
    Object.keys(state.reports)
      .filter((f) => !anio || f.slice(0, 4) === anio)
      .map((f) => f.slice(5, 7))
  );
  return [...set].sort((a, b) => b.localeCompare(a));
}

function topClientes() {
  const counts = {};
  allCapturas().forEach((c) => {
    [...(c.previos || []), ...(c.despachos || []), ...(c.revalidadas || [])].forEach((r) => {
      const t = (r.cliente || "").trim();
      if (!t) return;
      counts[t] = (counts[t] || 0) + 1;
    });
  });
  return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 5);
}

/** KPIs por tramitador: renglones de previos/despachos/revalidadas/pendientes que dejó
 * cada quien, más cuántos reportes capturó y en cuántos días distintos tuvo actividad. */
function statsByTramitador(desde, hasta) {
  const map = {};
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      const t = (c.tramitador || "").trim() || "Sin asignar";
      if (!map[t]) map[t] = { previos: 0, despachos: 0, revalidadas: 0, pendientes: 0, reportes: 0, dias: new Set() };
      map[t].previos += (c.previos || []).length;
      map[t].despachos += (c.despachos || []).length;
      map[t].revalidadas += (c.revalidadas || []).length;
      map[t].pendientes += (c.pendientes || []).length;
      map[t].reportes += 1;
      map[t].dias.add(c.fecha);
    });
  return Object.entries(map)
    .map(([nombre, s]) => ({
      nombre,
      previos: s.previos,
      despachos: s.despachos,
      revalidadas: s.revalidadas,
      pendientes: s.pendientes,
      total: s.previos + s.despachos + s.revalidadas + s.pendientes,
      reportes: s.reportes,
      dias: s.dias.size,
    }))
    .sort((a, b) => b.total - a.total);
}

/** Detalle día por día y tramitador (fecha × persona), para el rango de fechas dado —
 * así se ve exactamente qué hizo cada quien cada día, no solo el total acumulado. */
function statsDiariasPorTramitador(desde, hasta) {
  const map = {};
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      const t = (c.tramitador || "").trim() || "Sin asignar";
      const key = c.fecha + "||" + t;
      if (!map[key]) map[key] = { fecha: c.fecha, nombre: t, previos: 0, despachos: 0, revalidadas: 0, pendientes: 0 };
      map[key].previos += (c.previos || []).length;
      map[key].despachos += (c.despachos || []).length;
      map[key].revalidadas += (c.revalidadas || []).length;
      map[key].pendientes += (c.pendientes || []).length;
    });
  return Object.values(map)
    .map((r) => ({ ...r, total: r.previos + r.despachos + r.revalidadas + r.pendientes }))
    .sort((a, b) => b.fecha.localeCompare(a.fecha) || a.nombre.localeCompare(b.nombre));
}

/** KPIs por cliente: renglones de previos/despachos/revalidadas capturados para cada
 * cliente en el rango de fechas dado — para ver, por separado, quién es "top cliente"
 * en previos vs. en despachos (no es lo mismo un cliente con muchos previos que uno
 * con muchos despachos). Los pendientes no cuentan aquí porque no llevan Cliente fijo
 * hasta que se resuelven. */
function statsByCliente(desde, hasta) {
  const map = {};
  const add = (rows, c, key) => {
    rows.forEach((r) => {
      const cliente = (r.cliente || "").trim();
      if (!cliente) return;
      if (!map[cliente]) map[cliente] = { previos: 0, despachos: 0, revalidadas: 0, dias: new Set() };
      map[cliente][key] += 1;
      map[cliente].dias.add(c.fecha);
    });
  };
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      add(c.previos || [], c, "previos");
      add(c.despachos || [], c, "despachos");
      add(c.revalidadas || [], c, "revalidadas");
    });
  return Object.entries(map)
    .map(([nombre, s]) => ({
      nombre,
      previos: s.previos,
      despachos: s.despachos,
      revalidadas: s.revalidadas,
      total: s.previos + s.despachos + s.revalidadas,
      dias: s.dias.size,
    }))
    .sort((a, b) => b.total - a.total);
}

/** Detalle día × cliente en el rango dado, ordenado de mayor a menor movimiento —
 * responde directamente "¿qué día hubo más movimiento de qué cliente?". Se limita a
 * los 20 renglones con más actividad para que la tabla siga siendo legible. */
function statsDiariasPorCliente(desde, hasta) {
  const map = {};
  const add = (rows, c, key) => {
    rows.forEach((r) => {
      const cliente = (r.cliente || "").trim();
      if (!cliente) return;
      const k = c.fecha + "||" + cliente;
      if (!map[k]) map[k] = { fecha: c.fecha, nombre: cliente, previos: 0, despachos: 0, revalidadas: 0 };
      map[k][key] += 1;
    });
  };
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      add(c.previos || [], c, "previos");
      add(c.despachos || [], c, "despachos");
      add(c.revalidadas || [], c, "revalidadas");
    });
  return Object.values(map)
    .map((r) => ({ ...r, total: r.previos + r.despachos + r.revalidadas }))
    .sort((a, b) => b.total - a.total || b.fecha.localeCompare(a.fecha))
    .slice(0, 20);
}

function isReferenciaVacia(ref) {
  const v = (ref || "").trim().toLowerCase();
  return v === "" || v === "pendiente";
}

/** Busca si esa Guía ya tiene una Referencia capturada en algún otro renglón (previo,
 * despacho o pendiente) — primero en el MISMO reporte que se está llenando ahorita
 * (aunque todavía no se haya guardado — antes este paso faltaba, y por eso no detectaba
 * una guía repetida dos veces en el mismo día antes de guardar), y si no la encuentra ahí,
 * en el historial ya guardado de otros reportes. La comparación de guía ignora mayúsculas
 * y espacios de más, para que no falle por diferencias de captura entre personas. */
function buscarReferenciaPorGuia(guia, capturaEnEdicion) {
  const g = (guia || "").trim().toLowerCase();
  if (!g) return "";

  if (capturaEnEdicion) {
    const propias = [...(capturaEnEdicion.previos || []), ...(capturaEnEdicion.despachos || []), ...(capturaEnEdicion.pendientes || [])];
    for (const r of propias) {
      if ((r.guia || "").trim().toLowerCase() === g && !isReferenciaVacia(r.ref)) return r.ref;
    }
  }

  let mejor = null;
  allCapturas().forEach((c) => {
    if (capturaEnEdicion && c.id === capturaEnEdicion.id) return; // ya se revisó arriba
    [...(c.previos || []), ...(c.despachos || []), ...(c.pendientes || [])].forEach((r) => {
      if ((r.guia || "").trim().toLowerCase() !== g) return;
      if (isReferenciaVacia(r.ref)) return;
      if (!mejor || (c.fecha || "") > (mejor.fecha || "")) {
        mejor = { ref: r.ref, fecha: c.fecha || "" };
      }
    });
  });
  return mejor ? mejor.ref : "";
}

/** Busca renglones (previos, despachos, revalidadas, pendientes) sin número de Referencia en TODOS
 * los reportes de la aduana activa, agrupados por Guía — para que el Ejecutivo los complete. */
function findFilasSinReferencia() {
  const groups = {};
  allCapturas().forEach((c) => {
    [...(c.previos || []), ...(c.despachos || []), ...(c.revalidadas || []), ...(c.pendientes || [])].forEach((r) => {
      const guia = (r.guia || "").trim();
      if (!guia || !isReferenciaVacia(r.ref)) return;
      if (!groups[guia]) groups[guia] = { guia, cliente: r.cliente || "", almacen: r.almacen || "", count: 0 };
      groups[guia].count += 1;
    });
  });
  return Object.values(groups).sort((a, b) => a.guia.localeCompare(b.guia));
}

/** Igual que findFilasSinReferencia, pero agrupado un nivel arriba por Cliente,
 * para que el Ejecutivo pueda abrir solo el cliente que le interesa. */
function findFilasSinReferenciaPorCliente() {
  const guiaGroups = findFilasSinReferencia();
  const porCliente = {};
  guiaGroups.forEach((g) => {
    const cliente = g.cliente.trim() || "Sin cliente";
    if (!porCliente[cliente]) porCliente[cliente] = [];
    porCliente[cliente].push(g);
  });
  return Object.entries(porCliente)
    .map(([cliente, guias]) => ({ cliente, guias, totalRenglones: guias.reduce((s, g) => s + g.count, 0) }))
    .sort((a, b) => a.cliente.localeCompare(b.cliente));
}

/** Asignaciones pendientes de un tramitador para un tipo (previo/despacho) que
 * todavía NO están como renglón en la captura actual (por Guía) — para el botón
 * "Traer asignaciones pendientes" del formulario de captura. */
function asignacionesPendientesParaFormulario(tramitador) {
  const t = (tramitador || "").trim();
  if (!t) return [];
  const yaEnPendientes = new Set((state.currentCaptura.pendientes || []).map((r) => (r.guia || "").trim()).filter(Boolean));
  return Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .filter((a) => a.estatus === "pendiente" && (a.tramitador || "").trim() === t && !yaEnPendientes.has((a.guia || "").trim()));
}

/** Todas las asignaciones pendientes del tramitador actualmente logueado, sin importar
 * en qué pantalla esté ni si ya inició una captura — para el badge de la barra superior. */
function misAsignacionesPendientes() {
  if (state.userRole !== "tramite") return [];
  const t = (state.user || "").trim();
  return Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .filter((a) => a.estatus === "pendiente" && (a.tramitador || "").trim() === t)
    .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""));
}

/** Convierte un número de ocurrencia a texto tipo "1er previo", "2do previo"... */
function ordinalOcurrencia(n, tipo) {
  const palabra = tipo === "previo" ? "previo" : tipo === "revalidada" ? "revalidación" : "despacho";
  const map = { 1: "1er", 2: "2do", 3: "3er", 4: "4to", 5: "5to", 6: "6to", 7: "7mo", 8: "8vo", 9: "9no", 10: "10mo" };
  return `${map[n] || n + "°"} ${palabra}`;
}
function pillClaseTipo(tipo) {
  return tipo === "previo" ? "pill-navy" : tipo === "revalidada" ? "pill-ambar" : "pill-verde";
}

// ---------- rendering ----------
function render() {
  try {
    if (state.view === "login") root.innerHTML = viewLogin();
    else if (state.view === "selectTramitador") root.innerHTML = viewSelectTramitador();
    else if (state.view === "adminPassword") root.innerHTML = viewAdminPassword();
    else if (state.view === "ejecutivoName") root.innerHTML = viewEjecutivoName();
    else if (state.view === "tramitadorPassword") root.innerHTML = viewTramitadorPassword();
    else if (state.view === "selectAduana") root.innerHTML = viewSelectAduana();
    else {
      root.innerHTML =
        topbar() +
        `<div class="content">${notifPanelHtml()}${misAsignacionesPanelHtml()}${
          state.loading
            ? loadingBlock()
            : state.view === "home"
            ? viewHome()
            : state.view === "capture"
            ? viewCapture()
            : state.view === "review"
            ? viewReview()
            : state.view === "stamp"
            ? viewStamp()
            : state.view === "detail"
            ? viewDetail()
            : state.view === "catalogos"
            ? viewCatalogos()
            : state.view === "kpis"
            ? viewKPIs()
            : state.view === "clientes"
            ? viewClientes()
            : state.view === "asignaciones"
            ? viewAsignaciones()
            : state.view === "completarReferencias"
            ? viewCompletarReferencias()
            : state.view === "buscar"
            ? viewBuscar()
            : ""
        }</div>`;
    }
  } catch (e) {
    // Si algo falla al dibujar la pantalla, mostrar el error en vez de dejar la app en blanco/congelada
    root.innerHTML = `<div style="padding:24px;font-family:monospace;color:#C0453B;background:#FBE9E7;">
      <b>Ocurrió un error al mostrar esta pantalla:</b><br/>${esc(e.message)}<br/><br/>
      <button onclick="App.goHome()" style="padding:8px 14px;cursor:pointer;">Volver al inicio</button>
    </div>`;
    console.error(e);
  }
}
function loadingBlock() {
  return `<div class="status-line" style="max-width:320px;margin:40px auto;"><div class="spinner"></div> Cargando reportes…</div>`;
}
function errorBanner() {
  if (!state.errorMsg) return "";
  return `<div class="status-line status-error" style="margin-bottom:16px;">${esc(state.errorMsg)}</div>`;
}
function isIOS() {
  return /iphone|ipad|ipod/i.test(navigator.userAgent) && !window.MSStream;
}
function isStandalone() {
  return window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone === true;
}
function installBanner() {
  if (isStandalone() || state.installBannerDismissed) return "";
  if (state.installPromptEvent) {
    return `<div class="status-line" style="margin-bottom:16px;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
      <div>📲 Instala esta app en tu celular para acceso rápido, como una app normal.</div>
      <div style="display:flex;gap:8px;">
        <button class="btn btn-primary btn-sm" onclick="App.installApp()">Instalar</button>
        <button class="btn btn-ghost btn-sm" onclick="App.dismissInstallBanner()">Ahora no</button>
      </div>
    </div>`;
  }
  if (isIOS()) {
    return `<div class="status-line" style="margin-bottom:16px;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
      <div>📲 En iPhone: toca <b>Compartir</b> (el ícono de la cajita con flecha ↑) y elige <b>"Agregar a pantalla de inicio"</b> para instalar esta app.</div>
      <button class="btn btn-ghost btn-sm" onclick="App.dismissInstallBanner()">Entendido</button>
    </div>`;
  }
  return "";
}

function topbar() {
  const pendCount = pendingFechas().length;
  const atrasadasCount = Object.values(state.asignaciones || {}).filter(
    (a) => a.estatus === "pendiente" && a.fechaCreacion && Math.floor((Date.now() - new Date(a.fechaCreacion).getTime()) / 86400000) >= 2
  ).length;
  return `
  <div class="topbar">
    <div class="brand">
      <div class="brand-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="brand-text"><div class="t1">Oñate, Willy &amp; Cía.</div><div class="t2">Reporte operativo diario · v${APP_VERSION}</div></div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
      ${
        !state.isOnline
          ? `<span class="pill" style="background:#FBE9E7;color:#C0453B;">🔴 Sin conexión</span>`
          : state.syncingPending
          ? `<span class="pill" style="background:#DCEBF9;color:var(--accent-dark);">🔄 Sincronizando…</span>`
          : pendCount > 0
          ? `<span class="pill" style="background:#FBF3D8;color:#8A6414;" title="Se sincronizará al recuperar señal">⏳ ${pendCount} pendiente${pendCount > 1 ? "s" : ""}</span>`
          : ""
      }
      <div class="user-chip"><span class="user-dot"></span>${esc(state.user)}${state.aduanaActiva ? ` · ${esc(state.aduanaActiva)}` : ""}</div>
      <button class="nav-btn" onclick="App.toggleNotifPanel()" title="Notificaciones recientes">🔔${state.recentNotifications.length > 0 ? ` ${state.recentNotifications.length}` : ""}</button>
      ${
        state.userRole === "tramite" && misAsignacionesPendientes().length > 0
          ? `<button class="nav-btn" style="background:#FBEDEA;color:var(--rojo);" onclick="App.toggleMisAsignacionesPanel()" title="Tus asignaciones pendientes">📥 Asignaciones (${
              misAsignacionesPendientes().length
            })</button>`
          : ""
      }
      ${state.view !== "home" && state.view !== "completarReferencias" ? `<button class="nav-btn" onclick="App.goHome()">Inicio</button>` : ""}
      ${
        state.userRole === "ejecutivo" && state.view === "completarReferencias"
          ? `<button class="nav-btn" onclick="App.verReportes()">📋 Ver reportes (solo lectura)</button>`
          : ""
      }
      ${state.userRole === "ejecutivo" && state.view === "home" ? `<button class="nav-btn" onclick="App.goHome()">🔖 Completar referencias</button>` : ""}
      ${ADUANAS.length > 1 ? `<button class="nav-btn" onclick="App.cambiarAduana()">Cambiar aduana</button>` : ""}
      ${state.userRole === "admin" ? `<button class="nav-btn" onclick="App.goKPIs()">KPIs</button>` : ""}
      ${state.userRole === "admin" ? `<button class="nav-btn" onclick="App.goClientes()">Clientes</button>` : ""}
      ${state.userRole === "admin" ? `<button class="nav-btn" onclick="App.goAsignaciones()">Asignaciones${atrasadasCount > 0 ? ` ⚠️${atrasadasCount}` : ""}</button>` : ""}
      ${state.userRole !== "ejecutivo" ? `<button class="nav-btn" onclick="App.goBuscar()">🔍 Buscar</button>` : ""}
      ${state.userRole !== "ejecutivo" ? `<button class="nav-btn" onclick="App.goCatalogos()">Catálogos</button>` : ""}
      <button class="nav-btn" onclick="App.logout()">Salir</button>
    </div>
  </div>`;
}

function viewLogin() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">Oñate, Willy &amp; Cía.</div>
      <div class="login-sub">Digitalización del reporte operativo diario de previos y despachos</div>
      ${state.aduanaActiva ? `<div class="pill pill-navy" style="margin-bottom:16px;">Aduana: ${esc(state.aduanaActiva)}</div>` : ""}
      <div class="user-pick">
        ${USERS.map(
          (u) =>
            `<button class="user-pick-btn" onclick="App.chooseRole('${esc(u)}')"><span class="user-avatar">${esc(
              u.slice(0, 2).toUpperCase()
            )}</span> ${esc(u)}</button>`
        ).join("")}
      </div>
      ${ADUANAS.length > 1 ? `<button class="btn btn-ghost btn-sm" style="margin-top:16px;" onclick="App.cambiarAduana()">← Cambiar aduana</button>` : ""}
    </div>
  </div>`;
}

function viewSelectTramitador() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">¿Quién eres?</div>
      <div class="login-sub">Elige tu nombre — así quedará marcado automáticamente en cada reporte que captures</div>
      <div class="user-pick">
        ${(editableCats.tramitadores || []).map(
          (n) =>
            `<button class="user-pick-btn" onclick="App.chooseTramitador('${esc(n)}')"><span class="user-avatar">${esc(
              n.split(" ").map((p) => p[0]).join("").slice(0, 2).toUpperCase()
            )}</span> ${esc(n)}</button>`
        ).join("")}
      </div>
      <button class="btn btn-ghost btn-sm" style="margin-top:16px;" onclick="App.backToRoleSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewTramitadorPassword() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">Hola, ${esc(state.pendingTramitador || "")}</div>
      <div class="login-sub">Ingresa tu contraseña personal para continuar</div>
      <div class="field" style="text-align:left;margin-bottom:16px;">
        <input type="password" id="tramitador_password" placeholder="Contraseña" autofocus
          onkeydown="if(event.key==='Enter') App.submitTramitadorPassword()"
          style="width:100%;padding:11px 12px;border:1.3px solid var(--line);border-radius:8px;font-size:15px;"/>
      </div>
      ${state.errorMsg ? `<div class="status-line status-error" style="margin-bottom:16px;">${esc(state.errorMsg)}</div>` : ""}
      <button class="btn btn-primary" style="width:100%;margin-bottom:10px;" onclick="App.submitTramitadorPassword()">Entrar</button>
      <button class="btn btn-ghost btn-sm" onclick="App.backToTramitadorSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewAdminPassword() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">Acceso de ${esc(state.pendingAdminLabel || "Administrador")}</div>
      <div class="login-sub">Ingresa la contraseña para continuar</div>
      <div class="field" style="text-align:left;margin-bottom:16px;">
        <input type="password" id="admin_password" placeholder="Contraseña" autofocus
          onkeydown="if(event.key==='Enter') App.submitAdminPassword()"
          style="width:100%;padding:11px 12px;border:1.3px solid var(--line);border-radius:8px;font-size:15px;"/>
      </div>
      ${state.errorMsg ? `<div class="status-line status-error" style="margin-bottom:16px;">${esc(state.errorMsg)}</div>` : ""}
      <button class="btn btn-primary" style="width:100%;margin-bottom:10px;" onclick="App.submitAdminPassword()">Entrar</button>
      <button class="btn btn-ghost btn-sm" onclick="App.backToRoleSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewEjecutivoName() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">¿Quién eres?</div>
      <div class="login-sub">Escribe tu nombre — así quedará marcado en cada referencia que completes</div>
      <div class="field" style="text-align:left;margin-bottom:16px;">
        <input type="text" id="ejecutivo_nombre" placeholder="Tu nombre" autofocus
          onkeydown="if(event.key==='Enter') App.submitEjecutivoName()"
          style="width:100%;padding:11px 12px;border:1.3px solid var(--line);border-radius:8px;font-size:15px;"/>
      </div>
      ${state.errorMsg ? `<div class="status-line status-error" style="margin-bottom:16px;">${esc(state.errorMsg)}</div>` : ""}
      <button class="btn btn-primary" style="width:100%;margin-bottom:10px;" onclick="App.submitEjecutivoName()">Entrar</button>
      <button class="btn btn-ghost btn-sm" onclick="App.backToRoleSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewSelectAduana() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">¿Con qué aduana vas a trabajar?</div>
      <div class="login-sub">Cada aduana tiene su propio historial, catálogos y Excel — completamente independientes</div>
      <div class="user-pick">
        ${ADUANAS.map(
          (a) => `<button class="user-pick-btn" onclick="App.chooseAduana('${esc(a)}')"><span class="user-avatar">${esc(a.slice(0, 2))}</span> ${esc(a)}</button>`
        ).join("")}
      </div>
    </div>
  </div>`;
}

function viewHome() {
  const days = Object.values(state.reports).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const daysFiltrados = days.filter(
    (d) => (!state.reportesAnio || d.fecha.slice(0, 4) === state.reportesAnio) && (!state.reportesMes || d.fecha.slice(5, 7) === state.reportesMes)
  );
  const bars = last14Days();
  const maxVal = Math.max(1, ...bars.map((b) => b.value));
  const tp = totalPrevios(),
    td = totalDespachos(),
    tr = totalRevalidadas();
  const clienteStats = statsByCliente("", "");
  const topClPrevios = [...clienteStats].filter((s) => s.previos > 0).sort((a, b) => b.previos - a.previos).slice(0, 5);
  const topClDespachos = [...clienteStats].filter((s) => s.despachos > 0).sort((a, b) => b.despachos - a.despachos).slice(0, 5);
  const topClRevalidadas = [...clienteStats].filter((s) => s.revalidadas > 0).sort((a, b) => b.revalidadas - a.revalidadas).slice(0, 5);
  const maxClPrevios = Math.max(1, ...topClPrevios.map((s) => s.previos));
  const maxClDespachos = Math.max(1, ...topClDespachos.map((s) => s.despachos));
  const maxClRevalidadas = Math.max(1, ...topClRevalidadas.map((s) => s.revalidadas));
  const miniRankBlock = (rows, key, max) =>
    rows.length === 0
      ? `<div style="color:var(--muted);font-size:12px;">Sin datos todavía</div>`
      : rows
          .map(
            (s) => `
          <div class="rank-row">
            <div class="rank-name" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
            <div class="rank-bar-bg"><div class="rank-bar" style="width:${(s[key] / max) * 100}%"></div></div>
            <div class="rank-val">${s[key]}</div>
          </div>`
          )
          .join("");

  return `
    ${errorBanner()}
    ${
      state.usingLocalCache
        ? `<div class="status-line status-warn" style="margin-bottom:16px;">📴 Mostrando la última copia guardada en este dispositivo (sin conexión ahora mismo). Puede no incluir cambios hechos por la otra persona desde otro teléfono.</div>`
        : ""
    }
    ${installBanner()}
    ${notifBanner()}
    <div class="stats-grid">
      <div class="stat-card"><div class="stat-num">${days.length}</div><div class="stat-label">Días con reportes</div></div>
      <div class="stat-card"><div class="stat-num">${allCapturas().length}</div><div class="stat-label">Reportes capturados</div></div>
      <div class="stat-card"><div class="stat-num">${tp}</div><div class="stat-label">Total previos</div></div>
      <div class="stat-card"><div class="stat-num">${td}</div><div class="stat-label">Total despachos</div></div>
      <div class="stat-card"><div class="stat-num">${tr}</div><div class="stat-label">Guías revalidadas</div></div>
    </div>

    <div class="chart-row">
      <div class="panel">
        <div class="section-title">Actividad — últimos 14 días</div>
        <div class="section-sub">Reportes capturados por día</div>
        <svg viewBox="0 0 420 130" style="width:100%;height:120px;">
          ${bars
            .map((b, i) => {
              const w = 420 / 14;
              const h = maxVal ? (b.value / maxVal) * 90 : 0;
              const x = i * w + w * 0.2;
              const barW = w * 0.6;
              return `<rect x="${x}" y="${100 - h}" width="${barW}" height="${h}" rx="2" fill="${
                b.value > 0 ? "var(--accent)" : "#e2e7e5"
              }"></rect>
              <text x="${x + barW / 2}" y="115" font-size="7" fill="var(--muted)" text-anchor="middle" font-family="IBM Plex Mono">${b.label}</text>`;
            })
            .join("")}
        </svg>
      </div>
      <div class="panel">
        <div class="section-title">Distribución</div>
        <div class="section-sub">Previos vs. despachos</div>
        ${donut(tp, td)}
      </div>
      <div class="panel">
        <div class="section-title">Top clientes</div>
        <div class="section-sub" style="margin-bottom:10px;">Por número de operaciones — <a href="#" onclick="event.preventDefault();App.goClientes();" style="color:var(--accent);">ver con filtro de fechas →</a></div>
        <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin-bottom:2px;">Previos</div>
        ${miniRankBlock(topClPrevios, "previos", maxClPrevios)}
        <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin:10px 0 2px;">Despachos</div>
        ${miniRankBlock(topClDespachos, "despachos", maxClDespachos)}
        ${
          topClRevalidadas.length > 0
            ? `
        <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin:10px 0 2px;">Revalidadas</div>
        ${miniRankBlock(topClRevalidadas, "revalidadas", maxClRevalidadas)}
        `
            : ""
        }
      </div>
    </div>

    <div class="top-actions">
      <div>
        <div class="section-title" style="margin-bottom:0;">Reportes por fecha</div>
        <div class="section-sub">Cada hoja subida se agrupa automáticamente en el reporte de su fecha</div>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
        ${
          state.userRole === "admin" && state.historicoExcelUpdatedAt
            ? `<a href="${window.location.origin}${window.location.pathname}?historico=1&aduana=${encodeURIComponent(state.aduanaActiva)}" target="_blank" rel="noopener" class="btn btn-ghost">🔗 Link del Excel siempre actualizado</a>`
            : ""
        }
        ${
          state.userRole === "admin"
            ? `<button class="btn btn-ghost" onclick="App.exportExcelAll()" ${days.length === 0 ? "disabled" : ""}>📊 Descargar Excel (histórico completo)</button>`
            : ""
        }
        ${
          state.userRole === "admin"
            ? `<button class="btn btn-ghost" onclick="App.descargarRespaldoJSON()" title="Copia de seguridad completa: reportes, asignaciones y catálogos">💾 Respaldo completo (JSON)</button>`
            : ""
        }
        ${state.userRole !== "ejecutivo" ? `<button class="btn btn-primary" onclick="App.startCapture()">+ Nuevo reporte</button>` : ""}
      </div>
    </div>
    ${
      state.userRole === "admin" && state.historicoExcelUpdatedAt
        ? `<div style="color:var(--muted);font-size:11.5px;margin:-10px 0 14px;">Excel compartido actualizado por última vez: ${new Date(state.historicoExcelUpdatedAt).toLocaleString("es-MX")}</div>`
        : ""
    }
    ${state.userRole === "admin" && state.syncingExcel ? `<div class="status-line" style="max-width:340px;margin-bottom:14px;"><div class="spinner"></div> Actualizando el Excel compartido…</div>` : ""}

    ${
      days.length > 0
        ? `
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px;">
      <div class="field" style="margin-bottom:0;">
        <label>Año</label>
        <select onchange="App.setReportesAnio(this.value)">
          <option value="" ${state.reportesAnio === "" ? "selected" : ""}>Todos</option>
          ${aniosDisponibles()
            .map((a) => `<option value="${a}" ${state.reportesAnio === a ? "selected" : ""}>${a}</option>`)
            .join("")}
        </select>
      </div>
      <div class="field" style="margin-bottom:0;">
        <label>Mes</label>
        <select onchange="App.setReportesMes(this.value)">
          <option value="" ${state.reportesMes === "" ? "selected" : ""}>Todos</option>
          ${mesesDisponibles(state.reportesAnio)
            .map((m) => `<option value="${m}" ${state.reportesMes === m ? "selected" : ""}>${MESES_NOMBRES[parseInt(m, 10) - 1]}</option>`)
            .join("")}
        </select>
      </div>
      ${
        state.reportesAnio || state.reportesMes
          ? `<button class="btn btn-ghost btn-sm" style="margin-top:18px;" onclick="App.setReportesAnio('');App.setReportesMes('');">Quitar filtro</button>`
          : ""
      }
      <div style="margin-top:18px;color:var(--muted);font-size:12.5px;">${daysFiltrados.length} de ${days.length} día(s) con reportes</div>
    </div>
    `
        : ""
    }

    ${
      daysFiltrados.length === 0
        ? days.length === 0
          ? `<div class="empty"><div class="stamp-outline">📋</div><div style="font-weight:600;margin-bottom:4px;">Aún no hay reportes capturados</div><div style="font-size:13px;">Toca "Nuevo reporte" para subir tu primer Word, PDF, o llenarlo manualmente.</div></div>`
          : `<div class="empty"><div class="stamp-outline">📋</div><div style="font-weight:600;margin-bottom:4px;">Sin reportes en ese mes</div><div style="font-size:13px;">Prueba con otro año o mes, o quita el filtro.</div></div>`
        : daysFiltrados
            .map((d) => {
              const p = (d.capturas || []).reduce((s, c) => s + (c.previos ? c.previos.length : 0), 0);
              const de = (d.capturas || []).reduce((s, c) => s + (c.despachos ? c.despachos.length : 0), 0);
              return `
        <div style="display:flex;align-items:stretch;gap:8px;margin-bottom:10px;">
          <div class="day-card" style="flex:1;margin-bottom:0;" onclick="App.openDetail('${d.fecha}')">
            <div class="day-card-head">
              <div><div class="day-date">${fmtDateHuman(d.fecha)}</div><div class="day-meta">${
                (d.capturas || []).length
              } hoja(s) capturada(s)</div></div>
              <div style="display:flex;gap:6px;flex-wrap:wrap;">
                ${p ? `<span class="pill pill-navy">${p} previo${p > 1 ? "s" : ""}</span>` : ""}
                ${de ? `<span class="pill pill-verde">${de} despacho${de > 1 ? "s" : ""}</span>` : ""}
                ${isPending(d.fecha) ? `<span class="pill" style="background:#FBF3D8;color:#8A6414;" title="Todavía no se sube a la nube">⏳ pendiente</span>` : ""}
              </div>
            </div>
          </div>
          ${
            state.userRole === "admin"
              ? `<button class="row-del" style="border:1px solid var(--line);border-radius:var(--radius);background:var(--panel);font-size:18px;" title="Eliminar todo el reporte de este día" onclick="event.stopPropagation(); App.deleteDay('${d.fecha}')">✕</button>`
              : ""
          }
        </div>`;
            })
            .join("")
    }
  `;
}

function donut(a, b) {
  const total = a + b;
  if (total === 0) return `<div style="color:var(--muted);font-size:13px;padding:20px 0;text-align:center;">Sin datos todavía</div>`;
  const pctA = a / total,
    r = 42,
    c = 2 * Math.PI * r;
  const dashA = pctA * c;
  return `
  <div style="display:flex;align-items:center;gap:18px;">
    <svg viewBox="0 0 110 110" width="100" height="100">
      <circle cx="55" cy="55" r="${r}" fill="none" stroke="#E7ECEF" stroke-width="14"/>
      <circle cx="55" cy="55" r="${r}" fill="none" stroke="var(--navy)" stroke-width="14" stroke-dasharray="${dashA} ${
    c - dashA
  }" stroke-dashoffset="${c * 0.25}" stroke-linecap="round"/>
      <circle cx="55" cy="55" r="${r}" fill="none" stroke="var(--verde)" stroke-width="14" stroke-dasharray="${
    c - dashA
  } ${dashA}" stroke-dashoffset="${c * 0.25 - dashA}" stroke-linecap="round"/>
    </svg>
    <div style="font-size:13px;">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px;"><span style="width:10px;height:10px;border-radius:3px;background:var(--navy);display:inline-block;"></span> Previos: <b class="mono">${a}</b></div>
      <div style="display:flex;align-items:center;gap:6px;"><span style="width:10px;height:10px;border-radius:3px;background:var(--verde);display:inline-block;"></span> Despachos: <b class="mono">${b}</b></div>
    </div>
  </div>`;
}

function viewCapture() {
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Nuevo reporte</div>
    <div class="section-sub">Sube el documento ya llenado, o captúralo manualmente si no tienes el archivo a la mano</div>
    <div class="capture-grid">
      <div class="capture-tile" onclick="document.getElementById('docxInput').click()"><div class="icon">📄</div><div class="label">Documento Word (.docx)</div><div class="hint">Lectura exacta por tabla — recomendado</div></div>
      <div class="capture-tile" onclick="document.getElementById('pdfInput').click()"><div class="icon">📑</div><div class="label">Documento PDF</div><div class="hint">Lectura por posición del texto</div></div>
      <div class="capture-tile" onclick="App.startManual()"><div class="icon">✍️</div><div class="label">Llenar manualmente</div><div class="hint">Si no tienes el Word ni el PDF a la mano</div></div>
    </div>
    <input type="file" id="docxInput" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" style="display:none" onchange="App.handleFile(event,'docx')"/>
    <input type="file" id="pdfInput" accept="application/pdf,.pdf" style="display:none" onchange="App.handleFile(event,'pdf')"/>
    ${
      state.processingMsg
        ? `<div class="status-line" style="margin-top:16px;max-width:420px;"><div class="spinner"></div> ${esc(
            state.processingMsg
          )}</div>`
        : ""
    }
    ${state.errorMsg ? `<div class="status-line status-error" style="margin-top:14px;">${esc(state.errorMsg)}</div>` : ""}
  `;
}

function viewReview() {
  const c = state.currentCaptura;
  return `
    <button class="back-link" onclick="App.goHome()">← Cancelar</button>
    <div class="section-title">${state.editingIndex !== null ? "Editar reporte" : "Revisar y corregir"}</div>
    <div class="section-sub">${
      state.editingIndex !== null
        ? "Corrige los datos y confirma para actualizar este reporte"
        : "Verifica que los datos leídos del documento sean correctos antes de guardar"
    }</div>

    <div class="panel" style="margin-bottom:16px;">
      <div class="field-grid" style="grid-template-columns:repeat(3,1fr);">
        <div class="field"><label>Fecha</label><input type="date" id="f_fecha" value="${esc(c.fecha || todayStr())}"/></div>
        <div class="field"><label>Tramitador${state.userRole === "tramite" ? " (automático)" : ""}</label><input id="f_tram" value="${esc(c.tramitador)}" list="dl_tramitadores" ${state.userRole === "tramite" ? "readonly style=\"background:#EAF2FA;color:var(--muted);\"" : ""}/></div>
        <div class="field"><label>Aduana (automático)</label><input id="f_aduana" value="${esc(c.aduana || state.aduanaActiva)}" readonly style="background:#EAF2FA;color:var(--muted);"/></div>
      </div>

      <div class="subhead">1. Control de previos</div>
      <div class="table-wrap">
        <table class="rows-table">
          <thead><tr>${PREVIO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th></th></tr></thead>
          <tbody>${(c.previos || []).map((row, i) => rowHtml(row, i, "previos", PREVIO_FIELDS)).join("")}</tbody>
        </table>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="App.addRow('previos')">+ Agregar renglón de previo</button>

      <div class="subhead">2. Control de despachos</div>
      <div class="table-wrap">
        <table class="rows-table">
          <thead><tr>${DESPACHO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th></th></tr></thead>
          <tbody>${(c.despachos || []).map((row, i) => rowHtml(row, i, "despachos", DESPACHO_FIELDS)).join("")}</tbody>
        </table>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="App.addRow('despachos')">+ Agregar renglón de despacho</button>

      <div class="subhead">3. Control de guías revalidadas</div>
      <div class="table-wrap">
        <table class="rows-table">
          <thead><tr>${REVALIDADA_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th></th></tr></thead>
          <tbody>${(c.revalidadas || []).map((row, i) => rowHtml(row, i, "revalidadas", REVALIDADA_FIELDS)).join("")}</tbody>
        </table>
      </div>
      <button class="btn btn-ghost btn-sm" onclick="App.addRow('revalidadas')">+ Agregar renglón de guía revalidada</button>

      <div class="subhead">4. Pendientes</div>
      <div class="section-sub" style="margin-bottom:10px;">Operaciones que la persona de trámite deja pendientes ese día. Si tienen Tipo (previo/despacho/revalidada), aparece el botón "✅ Listo" para subirlas directo al cuadro correspondiente cuando se terminen.</div>
      <div class="table-wrap">
        <table class="rows-table">
          <thead><tr>${PENDIENTE_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th></th></tr></thead>
          <tbody>${(c.pendientes || []).map((row, i) => rowHtml(row, i, "pendientes", PENDIENTE_FIELDS)).join("")}</tbody>
        </table>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-ghost btn-sm" onclick="App.addRow('pendientes')">+ Agregar pendiente</button>
        ${
          asignacionesPendientesParaFormulario(c.tramitador).length > 0
            ? `<button class="btn btn-primary btn-sm" onclick="App.traerAsignaciones()">📥 Traer asignaciones pendientes (${
                asignacionesPendientesParaFormulario(c.tramitador).length
              })</button>`
            : ""
        }
      </div>

      <div class="subhead">5. Otras actividades</div>
      <div class="actividades-list" id="actividadesList">
        ${(c.otrasActividades || [])
          .map(
            (a, i) => `
          <div class="actividad-row"><span class="actividad-num">${i + 1}.</span><input value="${esc(
              a
            )}" oninput="App.updateActividad(${i},this.value)"/>
            <button class="row-del" onclick="App.deleteActividad(${i})">✕</button></div>
        `
          )
          .join("")}
      </div>
      <button class="btn btn-ghost btn-sm" onclick="App.addActividad()">+ Agregar actividad</button>

      <div class="review-actions">
        <button class="btn btn-ghost" onclick="App.goHome()">Cancelar</button>
        <button class="btn btn-primary" onclick="App.confirmSave()">${
          state.editingIndex !== null ? "Guardar cambios" : "Confirmar y guardar"
        }</button>
      </div>
    </div>
    ${datalists()}
  `;
}

function rowHtml(row, i, group, fields) {
  return `<tr>
    ${fields
      .map((f) => {
        if (f.select) {
          return `<td><select onchange="App.updateRowSelect('${group}',${i},'${f.k}',this.value)">${f.select
            .map((opt) => `<option value="${esc(opt)}" ${row[f.k] === opt ? "selected" : ""}>${opt ? esc(opt[0].toUpperCase() + opt.slice(1)) : "—"}</option>`)
            .join("")}</select></td>`;
        }
        const dl = f.cat ? ` list="dl_${f.cat}"` : "";
        const blur = f.k === "guia" ? ` onblur="App.autocompletarRefPorGuia('${group}',${i})"` : "";
        return `<td><input${dl} value="${esc(row[f.k])}" oninput="App.updateRow('${group}',${i},'${f.k}',this.value)"${blur}/></td>`;
      })
      .join("")}
    <td style="white-space:nowrap;">
      ${
        group === "pendientes" && (row.tipo === "previo" || row.tipo === "despacho" || row.tipo === "revalidada")
          ? `<button class="btn btn-primary btn-sm" style="margin-right:4px;" onclick="App.marcarPendienteListo(${i})">✅ Listo</button>`
          : ""
      }
      <button class="row-del" onclick="App.deleteRow('${group}',${i})" title="Eliminar renglón">✕</button>
    </td>
  </tr>`;
}

function datalists() {
  const merged = { ...editableCats, resultados: ["Verde", "Rojo", "Desaduanamiento libre", "Reconocimiento aduanero"] };
  return Object.entries(merged)
    .map((entry) => `<datalist id="dl_${entry[0]}">${entry[1].map((v) => `<option value="${esc(v)}">`).join("")}</datalist>`)
    .join("");
}

function viewStamp() {
  const wasQueuedOffline = state.currentCaptura && isPending(state.currentCaptura.fecha) && !state.isOnline;
  return `
  <div class="stamp-screen">
    <div class="stamp"><div class="s1">CONFIRMADO</div><div class="s2 mono">${fmtDateHuman(
      state.currentCaptura ? state.currentCaptura.fecha : todayStr()
    )}</div></div>
    <div style="font-weight:600;font-size:16px;margin-bottom:4px;">Reporte guardado</div>
    ${
      wasQueuedOffline
        ? `<div class="status-line status-warn" style="max-width:360px;margin:0 auto 16px;">⏳ Sin conexión ahora mismo — se guardó en este teléfono y se subirá solo en cuanto regrese la señal.</div>`
        : `<div style="color:var(--muted);font-size:13px;margin-bottom:20px;">Se agregó al reporte del día correspondiente</div>`
    }
    <div style="display:flex;gap:10px;">
      <button class="btn btn-ghost" onclick="App.goHome()">Ver reportes</button>
      <button class="btn btn-primary" onclick="App.startCapture()">Capturar otra hoja</button>
    </div>
  </div>`;
}

function viewDetail() {
  const d = state.reports[state.detailFecha];
  if (!d) return `<button class="back-link" onclick="App.goHome()">← Inicio</button><div class="empty">Reporte no encontrado.</div>`;
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="top-actions">
      <div><div class="section-title" style="margin-bottom:0;">${fmtDateHuman(d.fecha)}</div><div class="section-sub">${
    (d.capturas || []).length
  } hoja(s) capturada(s) este día</div></div>
      <div style="display:flex;gap:8px;">
        <button class="btn btn-ghost btn-sm" onclick="App.exportCsv('${d.fecha}')">Exportar CSV</button>
        ${state.userRole === "admin" ? `<button class="btn btn-primary btn-sm" onclick="App.exportExcel('${d.fecha}')">Descargar Excel</button>` : ""}
      </div>
    </div>
    ${(d.capturas || [])
      .map(
        (c, ci) => `
      <div class="captura-item">
        <div class="captura-head">
          <div class="meta-line">
            <span class="pill ${c.sourceType === "pdf" ? "pill-ambar" : c.sourceType === "manual" ? "pill-verde" : "pill-navy"}">${
          c.sourceType === "pdf" ? "📑 PDF (lectura por posición)" : c.sourceType === "manual" ? "✍️ Captura manual" : "📄 Word (lectura exacta)"
        }</span>
            ${c.sourceFileName ? ` <span class="mono" style="font-size:11px;">${esc(c.sourceFileName)}</span>` : ""}<br/>
            <b>Cargado:</b> ${c.horaCaptura ? new Date(c.horaCaptura).toLocaleString("es-MX") : "—"}<br/>
            <b>Tramitador:</b> ${esc(c.tramitador) || "—"} &nbsp;·&nbsp; <b>Aduana:</b> ${esc(c.aduana) || "—"} &nbsp;·&nbsp; Capturado por ${esc(
          c.uploadedBy
        )}
          </div>
          <div style="display:flex;gap:6px;">
            ${state.userRole !== "ejecutivo" ? `<button class="btn btn-ghost btn-sm" onclick="App.editCaptura('${d.fecha}',${ci})">Editar</button>` : ""}
            ${state.userRole === "admin" ? `<button class="row-del" onclick="App.deleteCaptura('${d.fecha}',${ci})" title="Eliminar hoja">✕</button>` : ""}
          </div>
        </div>
        <div class="mini-title">1. Control de previos (${(c.previos || []).length})</div>
        <div style="overflow-x:auto;">
        <table class="mini-table">
          <thead><tr>${PREVIO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
          <tbody>${
            (c.previos || []).map((r) => `<tr>${PREVIO_FIELDS.map((f) => `<td>${esc(r[f.k])}</td>`).join("")}</tr>`).join("") ||
            `<tr><td colspan="${PREVIO_FIELDS.length}" style="color:var(--muted);font-family:'IBM Plex Sans';">Sin renglones</td></tr>`
          }</tbody>
        </table>
        </div>
        <div class="mini-title">2. Control de despachos (${(c.despachos || []).length})</div>
        <div style="overflow-x:auto;">
        <table class="mini-table">
          <thead><tr>${DESPACHO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
          <tbody>${
            (c.despachos || []).map((r) => `<tr>${DESPACHO_FIELDS.map((f) => `<td>${esc(r[f.k])}</td>`).join("")}</tr>`).join("") ||
            `<tr><td colspan="${DESPACHO_FIELDS.length}" style="color:var(--muted);font-family:'IBM Plex Sans';">Sin renglones</td></tr>`
          }</tbody>
        </table>
        </div>
        ${
          (c.revalidadas || []).length
            ? `
          <div class="mini-title">3. Control de guías revalidadas (${(c.revalidadas || []).length})</div>
          <div style="overflow-x:auto;">
          <table class="mini-table">
            <thead><tr>${REVALIDADA_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
            <tbody>${(c.revalidadas || []).map((r) => `<tr>${REVALIDADA_FIELDS.map((f) => `<td>${esc(r[f.k])}</td>`).join("")}</tr>`).join("")}</tbody>
          </table>
          </div>
        `
            : ""
        }
        ${
          (c.pendientes || []).length
            ? `
          <div class="mini-title">4. Pendientes (${(c.pendientes || []).length})</div>
          <div style="overflow-x:auto;">
          <table class="mini-table">
            <thead><tr>${PENDIENTE_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
            <tbody>${(c.pendientes || []).map((r) => `<tr>${PENDIENTE_FIELDS.map((f) => `<td>${esc(r[f.k])}</td>`).join("")}</tr>`).join("")}</tbody>
          </table>
          </div>
        `
            : ""
        }
        ${
          (c.otrasActividades || []).filter((a) => a && a.trim()).length
            ? `
          <div class="mini-title">5. Otras actividades</div>
          <div style="font-size:12.5px;color:var(--ink);">${(c.otrasActividades || [])
            .filter((a) => a && a.trim())
            .map((a, i) => `${i + 1}. ${esc(a)}`)
            .join("<br/>")}</div>
        `
            : ""
        }
        ${
          (c.historial || []).length
            ? `
          <div style="margin-top:8px;">
            <button class="btn btn-ghost btn-sm" onclick="App.toggleHistorial('${d.fecha}',${ci})">
              🕒 Historial de cambios (${(c.historial || []).length}) ${state.expandedHistorial[`${d.fecha}_${ci}`] ? "▲" : "▼"}
            </button>
            ${
              state.expandedHistorial[`${d.fecha}_${ci}`]
                ? `<div style="font-size:11.5px;color:var(--muted);margin-top:6px;">${(c.historial || [])
                    .map((h) => `${h.accion === "creado" ? "🆕 Creado" : "✏️ Editado"} por ${esc(h.por)} — ${new Date(h.fecha).toLocaleString("es-MX")}`)
                    .join("<br/>")}</div>`
                : ""
            }
          </div>
        `
            : ""
        }
      </div>
    `
      )
      .join("")}
  `;
}

function viewCatalogos() {
  editableCats = sanitizeCats(editableCats);
  const categorias = ["clientes", "almacenes", "aduanas"];
  if (state.userRole === "admin") categorias.splice(2, 0, "tramitadores");
  const pwMap = tramitadorPasswordsMap(editableCats.tramitadores);
  const UMBRAL_BUSCADOR = 25; // a partir de cuántos registros aparece el buscador en vez de mostrar todo de un jalón
  const LIMITE_VISIBLE = 60; // tope de chips dibujados a la vez, aunque el filtro deje más resultados
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Catálogos</div>
    <div class="section-sub">Nombres sugeridos automáticamente al capturar (clientes, almacenes, tramitadores)</div>
    ${categorias
      .map((k) => {
        const total = editableCats[k].length;
        const necesitaBuscador = total > UMBRAL_BUSCADOR;
        const filtro = (state.catFilter[k] || "").trim().toLowerCase();
        let visibles = editableCats[k].map((v, i) => ({ v, i }));
        if (necesitaBuscador) {
          visibles = filtro ? visibles.filter((x) => x.v.toLowerCase().includes(filtro)) : [];
        }
        const recortado = visibles.length > LIMITE_VISIBLE;
        const visiblesFinal = recortado ? visibles.slice(0, LIMITE_VISIBLE) : visibles;
        return `
      <div class="panel" style="margin-bottom:14px;">
        <div style="font-weight:600;font-size:13.5px;margin-bottom:10px;text-transform:capitalize;">${k} ${
          total > 0 ? `<span style="color:var(--muted);font-weight:400;">(${total})</span>` : ""
        }</div>
        ${
          k === "tramitadores"
            ? `<div style="color:var(--muted);font-size:11.5px;margin-bottom:8px;">Contraseña automática: 3 primeras letras del nombre + 2026 (más letras si se repite con otra persona). Compártela con cada quien.</div>`
            : ""
        }
        ${
          necesitaBuscador
            ? `<input type="text" id="catsearch_${k}" placeholder="Buscar entre ${total} registros…" value="${esc(
                state.catFilter[k] || ""
              )}" oninput="App.setCatFilter('${k}', this.value)" style="width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:var(--radius);font-size:13px;margin-bottom:8px;"/>`
            : ""
        }
        <div class="chip-list">${
          necesitaBuscador && !filtro
            ? `<span style="color:var(--muted);font-size:12.5px;">Escribe arriba para buscar entre tus ${total} registros — con tantos, mostrarlos todos de un jalón haría la pantalla muy lenta.</span>`
            : visiblesFinal
                .map(
                  ({ v, i }) =>
                    `<span class="chip">${esc(v)}${
                      k === "tramitadores" ? ` <span style="font-family:'IBM Plex Mono',monospace;color:var(--muted);font-size:11px;">(${esc(pwMap[v] || "")})</span>` : ""
                    }<button onclick="App.removeCat('${k}',${i})">✕</button></span>`
                )
                .join("") || `<span style="color:var(--muted);font-size:12.5px;">${necesitaBuscador ? "Sin resultados para esa búsqueda" : "Sin registros"}</span>`
        }</div>
        ${
          recortado
            ? `<div style="font-size:11.5px;color:var(--muted);margin-top:4px;">Mostrando ${LIMITE_VISIBLE} de ${visibles.length} resultados — afina la búsqueda para ver el resto.</div>`
            : ""
        }
        <div class="cat-add">
          <input type="text" id="new_${k}" placeholder="Agregar…" onkeydown="if(event.key==='Enter') App.addCat('${k}')"/>
          <button class="btn btn-ghost btn-sm" onclick="App.addCat('${k}')">Agregar</button>
        </div>
        <button class="btn btn-ghost btn-sm" style="margin-top:8px;" onclick="App.toggleBulkCat('${k}')">${
          state.bulkCatOpen[k] ? "▲ Ocultar carga masiva" : "+ Agregar varios de un jalón"
        }</button>
        ${
          state.bulkCatOpen[k]
            ? `
          <div style="margin-top:8px;">
            <div style="font-size:11.5px;color:var(--muted);margin-bottom:6px;">Pega aquí varios nombres, uno por línea (o separados por coma) — se agregan todos juntos, sin repetir los que ya existan.</div>
            <textarea id="bulk_${k}" rows="6" placeholder="Cliente 1&#10;Cliente 2&#10;Cliente 3" style="width:100%;font-family:inherit;font-size:13px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;"></textarea>
            <button class="btn btn-primary btn-sm" style="margin-top:6px;" onclick="App.addCatBulk('${k}')">Agregar todos</button>
          </div>
        `
            : ""
        }
        ${
          state.bulkCatMsg && state.bulkCatMsgKey === k
            ? `<div class="status-line" style="margin-top:8px;">${esc(state.bulkCatMsg)}</div>`
            : ""
        }
      </div>`;
      })
      .join("")}
  `;
}

function viewKPIs() {
  const desde = state.kpiDesde;
  const hasta = state.kpiHasta;
  const stats = statsByTramitador(desde, hasta);
  const detalle = statsDiariasPorTramitador(desde, hasta);
  const maxTotal = Math.max(1, ...stats.map((s) => s.total));
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">KPIs por tramitador</div>
    <div class="section-sub">Previos, despachos, guías revalidadas y pendientes capturados por cada persona — aduana ${esc(state.aduanaActiva || "")}</div>

    <div class="panel" style="margin-bottom:16px;">
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end;">
        <div class="field" style="margin-bottom:0;"><label>Desde</label><input type="date" value="${esc(desde)}" onchange="App.setKpiRango('desde', this.value)"/></div>
        <div class="field" style="margin-bottom:0;"><label>Hasta</label><input type="date" value="${esc(hasta)}" onchange="App.setKpiRango('hasta', this.value)"/></div>
        <button class="btn btn-ghost btn-sm" onclick="App.setKpiPreset('hoy')">Hoy</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setKpiPreset('semana')">Últimos 7 días</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setKpiPreset('mes')">Este mes</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setKpiPreset('todo')">Todo</button>
      </div>
    </div>

    ${
      stats.length === 0
        ? `<div class="empty"><div class="stamp-outline">📊</div><div style="font-weight:600;margin-bottom:4px;">Sin datos en este rango</div><div style="font-size:13px;">Prueba con "Todo" o cambia las fechas.</div></div>`
        : `
    <div class="panel" style="margin-bottom:16px;overflow-x:auto;">
      <table class="mini-table" style="min-width:620px;">
        <thead><tr><th>Tramitador</th><th>Previos</th><th>Despachos</th><th>Revalidadas</th><th>Pendientes</th><th>Total</th><th>Reportes</th><th>Días activos</th></tr></thead>
        <tbody>
        ${stats
          .map(
            (s) => `<tr>
          <td style="font-family:'IBM Plex Sans',sans-serif;font-weight:600;">${esc(s.nombre)}</td>
          <td>${s.previos}</td><td>${s.despachos}</td><td>${s.revalidadas}</td><td>${s.pendientes}</td>
          <td style="font-weight:700;">${s.total}</td>
          <td>${s.reportes}</td><td>${s.dias}</td>
        </tr>`
          )
          .join("")}
        </tbody>
      </table>
    </div>
    <div class="panel" style="margin-bottom:16px;">
      <div class="section-title" style="font-size:15px;margin-bottom:12px;">Comparativo — total de operaciones (previos + despachos + revalidadas + pendientes)</div>
      ${stats
        .map(
          (s) => `
        <div class="rank-row">
          <div class="rank-name" style="width:140px;" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
          <div class="rank-bar-bg"><div class="rank-bar" style="width:${(s.total / maxTotal) * 100}%"></div></div>
          <div class="rank-val" style="width:36px;">${s.total}</div>
        </div>
      `
        )
        .join("")}
    </div>
    <div class="panel" style="overflow-x:auto;">
      <div class="section-title" style="font-size:15px;margin-bottom:12px;">Detalle día por día</div>
      <table class="mini-table" style="min-width:620px;">
        <thead><tr><th>Fecha</th><th>Tramitador</th><th>Previos</th><th>Despachos</th><th>Revalidadas</th><th>Pendientes</th><th>Total</th></tr></thead>
        <tbody>
        ${detalle
          .map(
            (r) => `<tr>
          <td>${esc(r.fecha)}</td>
          <td style="font-family:'IBM Plex Sans',sans-serif;font-weight:600;">${esc(r.nombre)}</td>
          <td>${r.previos}</td><td>${r.despachos}</td><td>${r.revalidadas}</td><td>${r.pendientes}</td>
          <td style="font-weight:700;">${r.total}</td>
        </tr>`
          )
          .join("")}
        </tbody>
      </table>
    </div>
    `
    }
  `;
}

function viewClientes() {
  const desde = state.clienteDesde;
  const hasta = state.clienteHasta;
  const stats = statsByCliente(desde, hasta);
  const detalle = statsDiariasPorCliente(desde, hasta);
  const topPrevios = [...stats].filter((s) => s.previos > 0).sort((a, b) => b.previos - a.previos).slice(0, 8);
  const topDespachos = [...stats].filter((s) => s.despachos > 0).sort((a, b) => b.despachos - a.despachos).slice(0, 8);
  const topRevalidadas = [...stats].filter((s) => s.revalidadas > 0).sort((a, b) => b.revalidadas - a.revalidadas).slice(0, 8);
  const maxPrevios = Math.max(1, ...topPrevios.map((s) => s.previos));
  const maxDespachos = Math.max(1, ...topDespachos.map((s) => s.despachos));
  const maxRevalidadas = Math.max(1, ...topRevalidadas.map((s) => s.revalidadas));
  const rankBlock = (titulo, rows, key, max) =>
    rows.length === 0
      ? `<div style="color:var(--muted);font-size:12.5px;">Sin datos en este rango</div>`
      : rows
          .map(
            (s) => `
        <div class="rank-row">
          <div class="rank-name" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
          <div class="rank-bar-bg"><div class="rank-bar" style="width:${(s[key] / max) * 100}%"></div></div>
          <div class="rank-val">${s[key]}</div>
        </div>`
          )
          .join("");
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Dashboard de clientes</div>
    <div class="section-sub">Previos, despachos y guías revalidadas por cliente — aduana ${esc(state.aduanaActiva || "")}</div>

    <div class="panel" style="margin-bottom:16px;">
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end;">
        <div class="field" style="margin-bottom:0;"><label>Desde</label><input type="date" value="${esc(desde)}" onchange="App.setClienteRango('desde', this.value)"/></div>
        <div class="field" style="margin-bottom:0;"><label>Hasta</label><input type="date" value="${esc(hasta)}" onchange="App.setClienteRango('hasta', this.value)"/></div>
        <button class="btn btn-ghost btn-sm" onclick="App.setClientePreset('hoy')">Hoy</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setClientePreset('semana')">Últimos 7 días</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setClientePreset('mes')">Este mes</button>
        <button class="btn btn-ghost btn-sm" onclick="App.setClientePreset('todo')">Todo</button>
      </div>
    </div>

    ${
      stats.length === 0
        ? `<div class="empty"><div class="stamp-outline">🏢</div><div style="font-weight:600;margin-bottom:4px;">Sin datos en este rango</div><div style="font-size:13px;">Prueba con "Todo" o cambia las fechas.</div></div>`
        : `
    <div class="dashboard-grid" style="margin-bottom:16px;">
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Top clientes — Previos</div>
        <div class="section-sub">Por número de renglones de previos</div>
        ${rankBlock("Previos", topPrevios, "previos", maxPrevios)}
      </div>
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Top clientes — Despachos</div>
        <div class="section-sub">Por número de renglones de despachos</div>
        ${rankBlock("Despachos", topDespachos, "despachos", maxDespachos)}
      </div>
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Top clientes — Guías revalidadas</div>
        <div class="section-sub">Por número de renglones de revalidaciones</div>
        ${rankBlock("Revalidadas", topRevalidadas, "revalidadas", maxRevalidadas)}
      </div>
    </div>

    <div class="panel" style="margin-bottom:16px;overflow-x:auto;">
      <div class="section-title" style="font-size:15px;margin-bottom:12px;">Todos los clientes en este rango</div>
      <table class="mini-table" style="min-width:560px;">
        <thead><tr><th>Cliente</th><th>Previos</th><th>Despachos</th><th>Revalidadas</th><th>Total</th><th>Días con movimiento</th></tr></thead>
        <tbody>
        ${stats
          .map(
            (s) => `<tr>
          <td style="font-family:'IBM Plex Sans',sans-serif;font-weight:600;">${esc(s.nombre)}</td>
          <td>${s.previos}</td><td>${s.despachos}</td><td>${s.revalidadas}</td>
          <td style="font-weight:700;">${s.total}</td>
          <td>${s.dias}</td>
        </tr>`
          )
          .join("")}
        </tbody>
      </table>
    </div>

    <div class="panel" style="overflow-x:auto;">
      <div class="section-title" style="font-size:15px;margin-bottom:4px;">Días con más movimiento</div>
      <div class="section-sub" style="margin-bottom:12px;">Los 20 renglones día + cliente con más actividad en este rango — así se ve directo qué día se movió más un cliente</div>
      <table class="mini-table" style="min-width:560px;">
        <thead><tr><th>Fecha</th><th>Cliente</th><th>Previos</th><th>Despachos</th><th>Revalidadas</th><th>Total</th></tr></thead>
        <tbody>
        ${detalle
          .map(
            (r) => `<tr>
          <td>${esc(r.fecha)}</td>
          <td style="font-family:'IBM Plex Sans',sans-serif;font-weight:600;">${esc(r.nombre)}</td>
          <td>${r.previos}</td><td>${r.despachos}</td><td>${r.revalidadas}</td>
          <td style="font-weight:700;">${r.total}</td>
        </tr>`
          )
          .join("")}
        </tbody>
      </table>
    </div>
    `
    }
  `;
}

function viewAsignaciones() {
  const list = Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .sort((a, b) => (b.fechaCreacion || "").localeCompare(a.fechaCreacion || ""));
  const pendientes = list.filter((a) => a.estatus === "pendiente");
  const completadas = list.filter((a) => a.estatus === "completada");
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Asignaciones</div>
    <div class="section-sub">Asigna previos, despachos o revalidaciones de guía a cada tramitador por número de Guía. Se cierran solas en cuanto ese tramitador guarda un reporte con esa misma Guía.</div>
    ${errorBanner()}

    <div class="panel" style="margin-bottom:20px;">
      <div class="field-grid" style="grid-template-columns:repeat(5,1fr);">
        <div class="field"><label>Tipo</label>
          <select id="asig_tipo" onchange="App.updateAsigDraft('tipo',this.value)">
            <option value="previo" ${state.asigDraft.tipo === "previo" ? "selected" : ""}>Previo</option>
            <option value="despacho" ${state.asigDraft.tipo === "despacho" ? "selected" : ""}>Despacho</option>
            <option value="revalidada" ${state.asigDraft.tipo === "revalidada" ? "selected" : ""}>Guía revalidada</option>
          </select>
        </div>
        <div class="field"><label>Guía</label><input id="asig_guia" placeholder="Número de guía" value="${esc(state.asigDraft.guia)}" oninput="App.updateAsigDraft('guia',this.value)"/></div>
        <div class="field"><label>Cliente</label><input id="asig_cliente" list="dl_clientes" value="${esc(state.asigDraft.cliente)}" oninput="App.updateAsigDraft('cliente',this.value)"/></div>
        <div class="field"><label>Almacén</label><input id="asig_almacen" list="dl_almacenes" value="${esc(state.asigDraft.almacen)}" oninput="App.updateAsigDraft('almacen',this.value)"/></div>
        <div class="field"><label>Tramitador</label>
          <select id="asig_tramitador" onchange="App.updateAsigDraft('tramitador',this.value)">
            ${(editableCats.tramitadores || []).map((t) => `<option value="${esc(t)}" ${state.asigDraft.tramitador === t ? "selected" : ""}>${esc(t)}</option>`).join("")}
          </select>
        </div>
      </div>
      <button class="btn btn-primary btn-sm" style="margin-top:10px;" onclick="App.crearAsignacion()">+ Crear asignación</button>
      ${datalists()}
    </div>

    <div class="subhead">Pendientes (${pendientes.length})</div>
    ${
      pendientes.length === 0
        ? `<div style="color:var(--muted);font-size:13px;margin-bottom:16px;">Sin asignaciones pendientes.</div>`
        : [...pendientes]
            .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""))
            .map((a) => {
              const dias = a.fechaCreacion ? Math.floor((Date.now() - new Date(a.fechaCreacion).getTime()) / 86400000) : 0;
              const atrasada = dias >= 2;
              return `
      <div class="captura-item" style="background:${atrasada ? "#FCDDD6" : "#FEF6F5"};border-color:${atrasada ? "#E8A79A" : "#F3C9C2"};">
        <div class="captura-head">
          <div class="meta-line">
            ${atrasada ? `<span class="pill" style="background:#C0453B;color:#fff;">⚠️ Atrasada (${dias} días)</span>` : `<span class="pill" style="background:#F8D9D4;color:var(--rojo);">⏳ pendiente</span>`}
            <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
            &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"} &nbsp;·&nbsp; <b>Almacén:</b> ${esc(a.almacen) || "—"}<br/>
            <b>Asignado a:</b> ${esc(a.tramitador)} &nbsp;·&nbsp; Creado por ${esc(a.creadoPor)} el ${a.fechaCreacion ? new Date(a.fechaCreacion).toLocaleDateString("es-MX") : "—"}
          </div>
          <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar asignación">✕</button>
        </div>
      </div>`;
            })
            .join("")
    }

    <div class="subhead">Completadas (${completadas.length})</div>
    ${
      completadas.length === 0
        ? `<div style="color:var(--muted);font-size:13px;">Sin asignaciones completadas todavía.</div>`
        : completadas
            .map(
              (a) => `
      <div class="captura-item" style="background:#F1F7F5;border-color:#BFE0CC;">
        <div class="captura-head">
          <div class="meta-line">
            <span class="pill" style="background:#DCF3E3;color:#2E7D4F;">✅ completada</span>
            <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
            &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"}<br/>
            <b>Completado por:</b> ${esc(a.tramitador)} el ${a.fechaCompletado ? new Date(a.fechaCompletado).toLocaleString("es-MX") : "—"}
          </div>
          <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar del historial">✕</button>
        </div>
      </div>`
            )
            .join("")
    }
  `;
}

function viewCompletarReferencias() {
  const porCliente = findFilasSinReferenciaPorCliente();
  const totalPendientes = porCliente.reduce((s, c) => s + c.totalRenglones, 0);
  return `
    <div class="section-title">Completar referencias</div>
    <div class="section-sub">Renglones capturados sin número de Referencia, agrupados por cliente — toca un cliente para ver sus guías pendientes.</div>
    ${errorBanner()}
    ${state.processingMsg ? `<div class="status-line" style="margin-bottom:16px;"><div class="spinner"></div> ${esc(state.processingMsg)}</div>` : ""}
    ${
      porCliente.length === 0
        ? `<div class="empty"><div class="stamp-outline">✅</div><div style="font-weight:600;">No hay referencias pendientes por completar</div></div>`
        : `
      <div class="stats-grid" style="grid-template-columns:1fr 1fr;margin-bottom:18px;">
        <div class="stat-card"><div class="stat-num">${porCliente.length}</div><div class="stat-label">Clientes con pendientes</div></div>
        <div class="stat-card"><div class="stat-num">${totalPendientes}</div><div class="stat-label">Renglones sin referencia</div></div>
      </div>
      ` +
          porCliente
            .map((grupo) => {
              const isOpen = !!state.expandedClientesRef[grupo.cliente];
              return `
        <div class="panel" style="margin-bottom:10px;padding:0;overflow:hidden;">
          <div style="padding:14px 16px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;background:#FDEDEB;" onclick="App.toggleClienteRef('${esc(
                grupo.cliente
              )}')">
            <div style="font-weight:600;">${esc(grupo.cliente)}</div>
            <div style="display:flex;align-items:center;gap:10px;">
              <span class="pill" style="background:#F8D9D4;color:var(--rojo);">⏳ ${grupo.totalRenglones} pendiente${grupo.totalRenglones > 1 ? "s" : ""}</span>
              <span style="font-size:13px;color:var(--muted);">${isOpen ? "▲" : "▼"}</span>
            </div>
          </div>
          ${
            isOpen
              ? `<div style="padding:14px 16px;border-top:1px solid var(--line);">
              ${grupo.guias
                .map(
                  (g) => `
                <div style="border:1px solid var(--line);border-radius:8px;padding:12px;margin-bottom:10px;background:#FEF6F5;">
                  <div class="meta-line" style="margin-bottom:10px;">
                    <b>Guía:</b> ${esc(g.guia)} &nbsp;·&nbsp; <b>Almacén:</b> ${esc(g.almacen) || "—"} &nbsp;·&nbsp; ${g.count} renglón(es) sin referencia
                  </div>
                  <div style="display:flex;gap:8px;">
                    <input id="ref_${esc(g.guia)}" placeholder="Número de referencia correcto" style="flex:1;padding:9px 10px;border:1.3px solid var(--line);border-radius:7px;font-size:13.5px;"/>
                    <button class="btn btn-primary btn-sm" onclick="App.completarReferencia('${esc(g.guia)}')">Actualizar</button>
                  </div>
                </div>`
                )
                .join("")}
            </div>`
              : ""
          }
        </div>`;
            })
            .join("")
    }
  `;
}

function viewBuscar() {
  const results = buscarGlobal(state.searchQuery);
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Buscar</div>
    <div class="section-sub">Busca por Guía, Referencia, Pedimento o Cliente en todo el historial de la aduana</div>
    <div class="field" style="margin-bottom:16px;">
      <input type="text" value="${esc(state.searchQuery)}" placeholder="Escribe una guía, referencia o cliente…" autofocus
        oninput="App.updateSearchQuery(this.value)"
        style="width:100%;padding:12px 14px;border:1.3px solid var(--line);border-radius:8px;font-size:15px;"/>
    </div>
    ${
      !state.searchQuery.trim()
        ? `<div style="color:var(--muted);font-size:13px;">Escribe algo para empezar a buscar.</div>`
        : results.length === 0
        ? `<div class="empty"><div class="stamp-outline">🔍</div><div style="font-weight:600;">Sin resultados para "${esc(state.searchQuery)}"</div></div>`
        : `
      <div style="color:var(--muted);font-size:12.5px;margin-bottom:10px;">${results.length} resultado(s)</div>
      <div class="table-wrap">
        <table class="mini-table">
          <thead><tr><th>Fecha</th><th>Tabla</th><th>Ref.</th><th>Guía</th><th>Pedimento</th><th>Cliente</th><th>Tramitador</th><th></th></tr></thead>
          <tbody>
            ${results
              .map(
                (r) => `<tr>
              <td>${esc(r.fecha)}</td>
              <td>${esc(r.tabla)}</td>
              <td>${esc(r.ref)}</td>
              <td>${esc(r.guia)}</td>
              <td>${esc(r.pedimento || "")}</td>
              <td>${esc(r.cliente)}</td>
              <td>${esc(r.tramitador)}</td>
              <td><button class="btn btn-ghost btn-sm" onclick="App.openDetail('${esc(r.fecha)}')">Ver reporte</button></td>
            </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>
      `
    }
  `;
}

// ---------- app logic ----------
const App = {
  async chooseAduana(aduana) {
    state.aduanaActiva = aduana;
    saveLastAduana(aduana);
    state.view = "login";
    state.errorMsg = "";
    render();
    // Cargar catálogos (tramitadores incluidos) YA, para que la lista y las
    // contraseñas de la pantalla de login estén actualizadas con lo último guardado.
    try {
      if (navigator.onLine) {
        const cats = await fbLoadCatalogos(aduana);
        if (cats) editableCats = sanitizeCats(cats);
      } else {
        const cached = loadLocalCache();
        if (cached.catalogos) editableCats = sanitizeCats(cached.catalogos);
      }
    } catch (e) {
      const cached = loadLocalCache();
      if (cached.catalogos) editableCats = sanitizeCats(cached.catalogos);
    }
    render();
  },
  cambiarAduana() {
    state.aduanaActiva = null;
    state.view = "selectAduana";
    state.errorMsg = "";
    state.reportesAnio = "";
    state.reportesMes = "";
    render();
  },
  chooseRole(role) {
    if (role === "Administrador" || role === "Coordinación") {
      state.pendingAdminLabel = role;
      state.view = "adminPassword";
      state.errorMsg = "";
      render();
    } else if (role === "Trámite") {
      state.view = "selectTramitador";
      render();
    } else if (role === "Ejecutivo") {
      state.view = "ejecutivoName";
      state.errorMsg = "";
      render();
    }
  },
  submitAdminPassword() {
    const input = document.getElementById("admin_password");
    const val = input ? input.value : "";
    if (val === ADMIN_PASSWORD) {
      state.errorMsg = "";
      App.enterApp(state.pendingAdminLabel || "Administrador", "admin");
    } else {
      state.errorMsg = "Contraseña incorrecta.";
      render();
      const el = document.getElementById("admin_password");
      if (el) {
        el.value = "";
        el.focus();
      }
    }
  },
  submitEjecutivoName() {
    const input = document.getElementById("ejecutivo_nombre");
    const val = input ? input.value.trim() : "";
    if (!val) {
      state.errorMsg = "Escribe tu nombre para continuar.";
      render();
      return;
    }
    state.errorMsg = "";
    App.enterApp(val, "ejecutivo");
  },
  chooseTramitador(nombre) {
    state.pendingTramitador = nombre;
    state.view = "tramitadorPassword";
    state.errorMsg = "";
    render();
  },
  submitTramitadorPassword() {
    const input = document.getElementById("tramitador_password");
    const val = input ? input.value : "";
    const correcta = tramitadorPassword(state.pendingTramitador);
    if (val === correcta) {
      state.errorMsg = "";
      App.enterApp(state.pendingTramitador, "tramite");
    } else {
      state.errorMsg = "Contraseña incorrecta.";
      render();
      const el = document.getElementById("tramitador_password");
      if (el) {
        el.value = "";
        el.focus();
      }
    }
  },
  backToTramitadorSelect() {
    state.view = "selectTramitador";
    state.errorMsg = "";
    render();
  },
  backToRoleSelect() {
    state.view = "login";
    state.errorMsg = "";
    render();
  },
  enterApp(nombre, role) {
    state.user = nombre;
    state.userRole = role;
    saveLastUser(nombre, role);
    state.view = role === "ejecutivo" ? "completarReferencias" : "home";
    state.loading = true;
    render();
    loadReports();
  },
  logout() {
    state.user = null;
    state.userRole = null;
    try {
      localStorage.removeItem("ow_ultimo_usuario");
      localStorage.removeItem("ow_ultimo_rol");
    } catch (e) {}
    state.view = state.aduanaActiva ? "login" : "selectAduana";
    render();
  },
  goHome() {
    state.view = state.userRole === "ejecutivo" ? "completarReferencias" : "home";
    state.currentCaptura = null;
    state.editingIndex = null;
    state.editingFecha = null;
    state.errorMsg = "";
    state.processingMsg = "";
    render();
    backgroundRefresh();
  },
  verReportes() {
    state.view = "home";
    state.errorMsg = "";
    render();
    backgroundRefresh();
  },
  goCatalogos() {
    state.view = "catalogos";
    render();
  },
  goBuscar() {
    state.view = "buscar";
    render();
    const el = document.querySelector('#app-root input[type="text"]');
    if (el) el.focus();
  },
  updateSearchQuery(val) {
    state.searchQuery = val;
    render();
    const el = document.querySelector('#app-root input[type="text"]');
    if (el) {
      el.focus();
      el.setSelectionRange(el.value.length, el.value.length);
    }
  },
  pedirPermisoNotificaciones() {
    if (typeof Notification === "undefined") return;
    Notification.requestPermission().then(() => render());
  },
  dismissNotifBanner() {
    state.notifBannerDismissed = true;
    render();
  },
  toggleNotifPanel() {
    state.showNotifPanel = !state.showNotifPanel;
    render();
  },
  toggleMisAsignacionesPanel() {
    state.showMisAsignacionesPanel = !state.showMisAsignacionesPanel;
    render();
  },
  goKPIs() {
    if (state.userRole !== "admin") return;
    state.view = "kpis";
    render();
  },
  setKpiRango(cual, val) {
    if (cual === "desde") state.kpiDesde = val;
    else state.kpiHasta = val;
    render();
  },
  setKpiPreset(preset) {
    const hoy = todayStr();
    if (preset === "hoy") {
      state.kpiDesde = hoy;
      state.kpiHasta = hoy;
    } else if (preset === "semana") {
      state.kpiDesde = dateStrLocal(6);
      state.kpiHasta = hoy;
    } else if (preset === "mes") {
      state.kpiDesde = hoy.slice(0, 8) + "01";
      state.kpiHasta = hoy;
    } else {
      state.kpiDesde = "";
      state.kpiHasta = "";
    }
    render();
  },
  goClientes() {
    if (state.userRole !== "admin") return;
    state.view = "clientes";
    render();
  },
  setClienteRango(cual, val) {
    if (cual === "desde") state.clienteDesde = val;
    else state.clienteHasta = val;
    render();
  },
  setClientePreset(preset) {
    const hoy = todayStr();
    if (preset === "hoy") {
      state.clienteDesde = hoy;
      state.clienteHasta = hoy;
    } else if (preset === "semana") {
      state.clienteDesde = dateStrLocal(6);
      state.clienteHasta = hoy;
    } else if (preset === "mes") {
      state.clienteDesde = hoy.slice(0, 8) + "01";
      state.clienteHasta = hoy;
    } else {
      state.clienteDesde = "";
      state.clienteHasta = "";
    }
    render();
  },
  setReportesAnio(val) {
    state.reportesAnio = val;
    state.reportesMes = ""; // al cambiar de año, no dejar seleccionado un mes que quizá no aplica
    render();
  },
  setReportesMes(val) {
    state.reportesMes = val;
    render();
  },
  goAsignaciones() {
    if (state.userRole !== "admin") return;
    state.view = "asignaciones";
    state.errorMsg = "";
    state.asigDraft = { tipo: "previo", guia: "", cliente: "", almacen: "", tramitador: (editableCats.tramitadores || [])[0] || "" };
    render();
  },
  updateAsigDraft(field, value) {
    state.asigDraft[field] = value;
    // Sin render() aquí a propósito: así no se interrumpe mientras se está escribiendo.
    // El valor ya queda guardado en memoria, así que aunque llegue la actualización
    // automática de cada 25 segundos, el formulario no se borra.
  },
  async crearAsignacion() {
    const tipo = state.asigDraft.tipo;
    const guia = (state.asigDraft.guia || "").trim();
    const cliente = (state.asigDraft.cliente || "").trim();
    const almacen = (state.asigDraft.almacen || "").trim();
    const tramitador = state.asigDraft.tramitador;
    if (!guia) {
      state.errorMsg = "Escribe el número de guía para crear la asignación.";
      render();
      return;
    }
    const id = "a_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
    const data = {
      tipo,
      guia,
      cliente,
      almacen,
      tramitador,
      estatus: "pendiente",
      creadoPor: state.user,
      fechaCreacion: new Date().toISOString(),
      fechaCompletado: null,
    };
    state.asignaciones[id] = data;
    state.asigDraft = { tipo: "previo", guia: "", cliente: "", almacen: "", tramitador };
    render();
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, data);
    } catch (e) {
      state.errorMsg = "No se pudo guardar la asignación en la nube (" + e.message + "), pero quedó visible aquí por ahora.";
      render();
    }
  },
  async eliminarAsignacion(id) {
    if (!confirm("¿Eliminar esta asignación?")) return;
    delete state.asignaciones[id];
    render();
    try {
      await fbDeleteAsignacion(state.aduanaActiva, id);
    } catch (e) {}
  },
  toggleClienteRef(cliente) {
    state.expandedClientesRef[cliente] = !state.expandedClientesRef[cliente];
    render();
  },
  toggleHistorial(fecha, ci) {
    const key = `${fecha}_${ci}`;
    state.expandedHistorial[key] = !state.expandedHistorial[key];
    render();
  },
  async completarReferencia(guia) {
    const input = document.getElementById("ref_" + guia);
    const nuevaRef = input ? input.value.trim() : "";
    if (!nuevaRef) {
      state.errorMsg = "Escribe el número de referencia antes de actualizar.";
      render();
      return;
    }
    state.processingMsg = "Actualizando referencia en todos los reportes con esa guía…";
    render();
    const fechasAfectadas = new Set();
    Object.values(state.reports).forEach((day) => {
      (day.capturas || []).forEach((c) => {
        [...(c.previos || []), ...(c.despachos || []), ...(c.revalidadas || []), ...(c.pendientes || [])].forEach((r) => {
          if ((r.guia || "").trim() === guia) {
            r.ref = nuevaRef;
            fechasAfectadas.add(c.fecha);
          }
        });
      });
    });
    persistLocalCache();
    for (const fecha of fechasAfectadas) {
      try {
        await saveReportDay(fecha);
      } catch (e) {}
    }
    state.processingMsg = "";
    render();
    if (navigator.onLine) syncHistoricoExcel();
  },
  async installApp() {
    if (!state.installPromptEvent) return;
    state.installPromptEvent.prompt();
    await state.installPromptEvent.userChoice;
    state.installPromptEvent = null;
    render();
  },
  dismissInstallBanner() {
    state.installBannerDismissed = true;
    render();
  },
  openDetail(fecha) {
    state.view = "detail";
    state.detailFecha = fecha;
    state.errorMsg = "";
    render();
  },
  startCapture() {
    if (state.userRole === "ejecutivo") return; // solo lectura
    state.view = "capture";
    state.errorMsg = "";
    state.processingMsg = "";
    render();
  },
  startManual() {
    const cap = emptyCaptura();
    cap.uploadedBy = state.user;
    cap.horaCaptura = new Date().toISOString();
    cap.sourceType = "manual";
    cap.fecha = todayStr();
    cap.tramitador = state.userRole === "tramite" ? state.user : "";
    cap.aduana = state.aduanaActiva;
    state.currentCaptura = normalizeCaptura(cap);
    state.editingIndex = null;
    state.editingFecha = null;
    state.view = "review";
    render();
  },

  async addCat(k) {
    const el = document.getElementById("new_" + k);
    const v = el.value.trim();
    if (!v) return;
    if (!editableCats[k].includes(v)) editableCats[k].push(v);
    await saveCatalogos();
    render();
  },
  async removeCat(k, i) {
    editableCats[k].splice(i, 1);
    await saveCatalogos();
    render();
  },
  toggleBulkCat(k) {
    state.bulkCatOpen[k] = !state.bulkCatOpen[k];
    state.bulkCatMsg = "";
    render();
  },
  setCatFilter(k, val) {
    state.catFilter[k] = val;
    render();
    const el = document.getElementById("catsearch_" + k);
    if (el) {
      el.focus();
      el.setSelectionRange(el.value.length, el.value.length);
    }
  },
  /** Da de alta muchos nombres de un jalón (uno por línea o separados por coma) —
   * pensado para cuando se está armando el catálogo de una aduana nueva desde cero
   * (ej. Toluca) y sería muy lento agregarlos uno por uno. No repite los que ya existan. */
  async addCatBulk(k) {
    const el = document.getElementById("bulk_" + k);
    if (!el) return;
    const nombres = (el.value || "")
      .split(/[\n,]/)
      .map((s) => s.trim())
      .filter(Boolean);
    if (nombres.length === 0) return;
    let agregados = 0;
    let repetidos = 0;
    nombres.forEach((n) => {
      if (editableCats[k].includes(n)) {
        repetidos++;
      } else {
        editableCats[k].push(n);
        agregados++;
      }
    });
    state.bulkCatOpen[k] = false;
    state.bulkCatMsgKey = k;
    state.bulkCatMsg =
      `✅ Se agregaron ${agregados} registro${agregados === 1 ? "" : "s"}` + (repetidos > 0 ? ` (${repetidos} ya existían y no se repitieron).` : ".");
    try {
      await saveCatalogos();
    } catch (e) {
      state.bulkCatMsg = "No se pudo guardar en la nube (" + e.message + "), pero quedó visible aquí por ahora.";
    }
    render();
  },

  async handleFile(evt, kind) {
    const file = evt.target.files[0];
    if (!file) return;
    state.errorMsg = "";
    const name = file.name || "";
    try {
      if (kind === "docx") {
        state.processingMsg = "Leyendo documento Word…";
        render();
        const parsed = await parseDocxTemplate(file);
        App.applyParsedData(parsed, name, "docx");
      } else {
        state.processingMsg = "Leyendo documento PDF…";
        render();
        const parsed = await parsePdfTemplate(file);
        App.applyParsedData(parsed, name, "pdf");
      }
      state.processingMsg = "";
    } catch (e) {
      state.processingMsg = "";
      state.errorMsg = "No se pudo leer el archivo (" + (e && e.message ? e.message : "error desconocido") + ").";
      render();
    }
  },

  applyParsedData(parsed, fileName, sourceType) {
    const cap = emptyCaptura();
    cap.uploadedBy = state.user;
    cap.horaCaptura = new Date().toISOString();
    cap.sourceType = sourceType;
    cap.sourceFileName = fileName;
    const guessed = guessFechaFromFileName(fileName);
    cap.fecha = parsed.fecha && /^\d{4}-\d{2}-\d{2}$/.test(parsed.fecha) ? parsed.fecha : guessed || todayStr();
    cap.tramitador = state.userRole === "tramite" ? state.user : parsed.tramitador || "";
    cap.aduana = state.aduanaActiva;
    cap.previos = Array.isArray(parsed.previos) ? parsed.previos : [];
    cap.despachos = Array.isArray(parsed.despachos) ? parsed.despachos : [];
    cap.pendientes = Array.isArray(parsed.pendientes) ? parsed.pendientes : [];
    cap.otrasActividades =
      Array.isArray(parsed.otrasActividades) && parsed.otrasActividades.length ? parsed.otrasActividades : ["", "", "", ""];
    state.currentCaptura = normalizeCaptura(cap);
    state.editingIndex = null;
    state.editingFecha = null;
    state.view = "review";
    render();
  },

  traerAsignaciones() {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      const pendientes = asignacionesPendientesParaFormulario(c.tramitador);
      if (pendientes.length === 0) return;
      pendientes.forEach((a) => {
        const row = {};
        PENDIENTE_FIELDS.forEach((f) => (row[f.k] = ""));
        row.guia = a.guia || "";
        row.cliente = a.cliente || "";
        row.almacen = a.almacen || "";
        row.tipo = a.tipo === "previo" || a.tipo === "despacho" || a.tipo === "revalidada" ? a.tipo : "";
        c.pendientes.push(row);
      });
      render();
    } catch (e) {
      state.errorMsg = "No se pudieron traer las asignaciones (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  /** Sube un renglón de Pendientes directo a Previos o Despachos (según su Tipo) y
   * lo quita de Pendientes. Si la Guía ya existe en ese cuadro (2do previo, etc.),
   * lo etiqueta automáticamente en la columna "N° previo/despacho" en vez de avisar
   * como duplicado, porque puede pasar legítimamente (varios previos por guía). */
  marcarPendienteListo(i) {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      const pend = c.pendientes[i];
      if (!pend || (pend.tipo !== "previo" && pend.tipo !== "despacho" && pend.tipo !== "revalidada")) return;
      const tipo = pend.tipo;
      const grupo = tipo === "previo" ? "previos" : tipo === "revalidada" ? "revalidadas" : "despachos";
      const fields = tipo === "previo" ? PREVIO_FIELDS : tipo === "revalidada" ? REVALIDADA_FIELDS : DESPACHO_FIELDS;
      const guia = (pend.guia || "").trim();

      let count = (c[grupo] || []).filter((r) => (r.guia || "").trim() === guia).length;
      if (guia) {
        allCapturas().forEach((cap) => {
          if (cap.id === c.id) return; // evitar contar dos veces la misma hoja que se está editando
          const filas = cap[grupo] || [];
          filas.forEach((r) => {
            if ((r.guia || "").trim() === guia) count++;
          });
        });
      }

      const row = {};
      fields.forEach((f) => (row[f.k] = ""));
      row.guia = pend.guia || "";
      row.cliente = pend.cliente || "";
      row.almacen = pend.almacen || "";
      row.ref = pend.ref || "";
      if (tipo === "despacho") row.pedimento = pend.pedimento || "";
      const numKey = tipo === "previo" ? "numPrevio" : tipo === "revalidada" ? "numRevalidada" : "numDespacho";
      row[numKey] = guia ? ordinalOcurrencia(count + 1, tipo) : "";

      c[grupo].push(row);
      c.pendientes.splice(i, 1);
      render();
    } catch (e) {
      state.errorMsg = "No se pudo mover el pendiente (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  addRow(group) {
    const fields = group === "previos" ? PREVIO_FIELDS : group === "pendientes" ? PENDIENTE_FIELDS : group === "revalidadas" ? REVALIDADA_FIELDS : DESPACHO_FIELDS;
    const empty = {};
    fields.forEach((f) => (empty[f.k] = ""));
    state.currentCaptura[group].push(empty);
    render();
  },
  updateRow(group, i, key, val) {
    state.currentCaptura[group][i][key] = val;
  },
  /** Al salir del campo Guía (no en cada tecla, para no interrumpir al escribir): si esa
   * misma guía ya tiene una Referencia capturada en un reporte anterior, la rellena sola —
   * así no hay que volver a buscarla o teclearla cuando se repite una guía (ej. 2do previo). */
  autocompletarRefPorGuia(group, i) {
    try {
      if (!("ref" in (state.currentCaptura[group][i] || {}))) return;
      const row = state.currentCaptura[group][i];
      const guia = (row.guia || "").trim();
      if (!guia) return;
      if (!isReferenciaVacia(row.ref)) return; // ya tiene una referencia escrita a mano, no se la pisamos
      const encontrada = buscarReferenciaPorGuia(guia, state.currentCaptura);
      if (encontrada) {
        row.ref = encontrada;
        render();
      }
    } catch (e) {}
  },
  /** Igual que updateRow, pero para selectores (Tipo, Dificultad, etc.) — estos SÍ
   * necesitan refrescar la pantalla de inmediato (para que aparezca el botón "Listo"
   * o cambie alguna otra parte visible), a diferencia de los campos de texto normales
   * donde refrescar en cada tecla interrumpiría al escribir. */
  updateRowSelect(group, i, key, val) {
    state.currentCaptura[group][i][key] = val;
    render();
  },
  deleteRow(group, i) {
    state.currentCaptura[group].splice(i, 1);
    render();
  },
  addActividad() {
    state.currentCaptura.otrasActividades.push("");
    render();
  },
  updateActividad(i, val) {
    state.currentCaptura.otrasActividades[i] = val;
  },
  deleteActividad(i) {
    state.currentCaptura.otrasActividades.splice(i, 1);
    render();
  },

  async confirmSave() {
    try {
      const c = state.currentCaptura;
      if (!c) throw new Error("No hay un reporte en edición.");
      const fFecha = document.getElementById("f_fecha");
      const fTram = document.getElementById("f_tram");
      const fAduana = document.getElementById("f_aduana");
      c.fecha = (fFecha && fFecha.value) || todayStr();
      c.tramitador = fTram ? fTram.value : c.tramitador;
      c.aduana = fAduana ? fAduana.value : c.aduana;

      const duplicadas = findGuiasDuplicadas(c);
      if (duplicadas.length > 0) {
        const lista = duplicadas
          .map((d) => `• Guía ${d.guia} (${d.tipo}) ya está en el reporte del ${d.fecha} (${d.tramitador || "sin tramitador"})`)
          .join("\n");
        const continuar = confirm(
          "⚠️ Esta(s) guía(s) ya aparecen en otro reporte:\n\n" + lista + "\n\n¿Quieres guardar de todas formas?"
        );
        if (!continuar) return;
      }

      if (!state.reports[c.fecha]) state.reports[c.fecha] = { fecha: c.fecha, capturas: [] };

      const wasEditing = state.editingIndex !== null && state.editingFecha;
      if (!Array.isArray(c.historial)) c.historial = [];
      c.historial.push({
        accion: wasEditing ? "editado" : "creado",
        por: state.user,
        fecha: new Date().toISOString(),
      });

      if (wasEditing && state.editingFecha === c.fecha) {
        // Misma fecha: reemplazar el renglón en su lugar, sin borrar el día ni volver
        // a agregarlo aparte (evita el bug de "borro el día vacío y luego intento
        // agregarle algo, que ya no existe").
        const day = state.reports[c.fecha];
        if (day.capturas && day.capturas[state.editingIndex]) {
          day.capturas[state.editingIndex] = c;
        } else {
          day.capturas.push(c);
        }
      } else {
        if (wasEditing) {
          // La fecha sí cambió: quitar la versión anterior de donde estaba
          const oldDay = state.reports[state.editingFecha];
          if (oldDay && oldDay.capturas && oldDay.capturas[state.editingIndex]) {
            oldDay.capturas.splice(state.editingIndex, 1);
            if (oldDay.capturas.length === 0) delete state.reports[state.editingFecha];
          }
        }
        state.reports[c.fecha].capturas.push(c);
      }
      persistLocalCache();

      // Si el día viejo cambió (edición que cambió de fecha), guardar ambos días en Firebase
      if (wasEditing && state.editingFecha !== c.fecha && state.reports[state.editingFecha]) {
        await saveReportDay(state.editingFecha);
      } else if (wasEditing && state.editingFecha !== c.fecha && !state.reports[state.editingFecha]) {
        await deleteReportDay(state.editingFecha);
      }
      await saveReportDay(c.fecha);
      if (navigator.onLine) await autoCloseAsignaciones(c);

      state.editingIndex = null;
      state.editingFecha = null;
      state.view = "stamp";
      render();
      if (navigator.onLine) syncHistoricoExcel();
    } catch (e) {
      state.errorMsg = "No se pudo guardar el reporte (" + e.message + "). Intenta de nuevo.";
      state.view = "capture";
      render();
    }
  },

  editCaptura(fecha, idx) {
    if (state.userRole === "ejecutivo") return; // solo lectura
    try {
      const day = state.reports[fecha];
      if (!day || !day.capturas || !day.capturas[idx]) {
        state.errorMsg = "No se encontró ese reporte para editar (puede que ya haya sido eliminado o modificado).";
        render();
        return;
      }
      const c = day.capturas[idx];
      state.currentCaptura = normalizeCaptura(JSON.parse(JSON.stringify(c)));
      state.editingFecha = fecha;
      state.editingIndex = idx;
      state.errorMsg = "";
      state.view = "review";
      render();
    } catch (e) {
      state.errorMsg = "No se pudo abrir este reporte para editar (" + e.message + ").";
      render();
    }
  },

  async deleteCaptura(fecha, idx) {
    if (state.userRole !== "admin") return;
    if (!confirm("¿Eliminar esta hoja del reporte?")) return;
    state.reports[fecha].capturas.splice(idx, 1);
    if (state.reports[fecha].capturas.length === 0) {
      await deleteReportDay(fecha);
      delete state.reports[fecha];
      state.view = "home";
    } else {
      await saveReportDay(fecha);
    }
    persistLocalCache();
    render();
    if (navigator.onLine) syncHistoricoExcel();
  },

  async deleteDay(fecha) {
    if (state.userRole !== "admin") return;
    if (!confirm(`¿Eliminar TODO el reporte de ${fmtDateHuman(fecha)}? Esto borra todas las hojas capturadas ese día.`)) return;
    try {
      await deleteReportDay(fecha);
      delete state.reports[fecha];
      persistLocalCache();
      render();
      if (navigator.onLine) syncHistoricoExcel();
    } catch (e) {
      state.errorMsg = "No se pudo eliminar (" + e.message + ").";
      render();
    }
  },

  exportCsv(fecha) {
    const d = state.reports[fecha];
    const lines = [];
    lines.push(["Fecha", "Tramitador", "Aduana"].join(","));
    (d.capturas || []).forEach((c) =>
      lines.push([c.fecha, c.tramitador, c.aduana].map((v) => `"${String(v || "").replace(/"/g, '""')}"`).join(","))
    );
    lines.push("");
    lines.push("--- PREVIOS ---");
    lines.push(PREVIO_FIELDS.map((f) => f.label).join(","));
    (d.capturas || []).forEach((c) =>
      (c.previos || []).forEach((r) => lines.push(PREVIO_FIELDS.map((f) => `"${String(r[f.k] || "").replace(/"/g, '""')}"`).join(",")))
    );
    lines.push("");
    lines.push("--- DESPACHOS ---");
    lines.push(DESPACHO_FIELDS.map((f) => f.label).join(","));
    (d.capturas || []).forEach((c) =>
      (c.despachos || []).forEach((r) => lines.push(DESPACHO_FIELDS.map((f) => `"${String(r[f.k] || "").replace(/"/g, '""')}"`).join(",")))
    );
    lines.push("");
    lines.push("--- GUÍAS REVALIDADAS ---");
    lines.push(REVALIDADA_FIELDS.map((f) => f.label).join(","));
    (d.capturas || []).forEach((c) =>
      (c.revalidadas || []).forEach((r) => lines.push(REVALIDADA_FIELDS.map((f) => `"${String(r[f.k] || "").replace(/"/g, '""')}"`).join(",")))
    );
    lines.push("");
    lines.push("--- PENDIENTES ---");
    lines.push(PENDIENTE_FIELDS.map((f) => f.label).join(","));
    (d.capturas || []).forEach((c) =>
      (c.pendientes || []).forEach((r) => lines.push(PENDIENTE_FIELDS.map((f) => `"${String(r[f.k] || "").replace(/"/g, '""')}"`).join(",")))
    );
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte_${fecha}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  },

  async exportExcel(fecha) {
    if (state.userRole !== "admin") return;
    const d = state.reports[fecha];
    await buildAndDownloadExcel([d], `reporte_${fecha}.xlsx`);
  },

  async exportExcelAll() {
    if (state.userRole !== "admin") return;
    const days = Object.values(state.reports).sort((a, b) => a.fecha.localeCompare(b.fecha));
    if (days.length === 0) return;
    const stamp = todayStr();
    await buildAndDownloadExcel(days, `reporte_historico_${stamp}.xlsx`);
  },
  descargarRespaldoJSON() {
    if (state.userRole !== "admin") return;
    const respaldo = {
      aduana: state.aduanaActiva,
      fechaExportacion: new Date().toISOString(),
      reportes: state.reports,
      asignaciones: state.asignaciones,
      catalogos: editableCats,
    };
    const blob = new Blob([JSON.stringify(respaldo, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `respaldo_${state.aduanaActiva}_${todayStr()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  },
};
window.App = App;

const OW_NAVY = "FF0A2E52";
const OW_ACCENT = "FF1D63A6";
const OW_BG_CLARO = "FFEAF2FA";
const OW_MUTED = "FF5B7690";

let _owLogoBuffer = null;
/** Descarga el logo de OW una sola vez por sesión y lo deja listo para insertarse en
 * cualquier hoja de Excel. Si por alguna razón no se puede cargar (ej. sin internet en
 * ese instante), el Excel se genera igual, nomás sin el logo — nunca bloquea la descarga. */
async function cargarLogoOW() {
  if (_owLogoBuffer) return _owLogoBuffer;
  try {
    const res = await fetch("logo-ow.jpg");
    if (!res.ok) return null;
    _owLogoBuffer = await res.arrayBuffer();
    return _owLogoBuffer;
  } catch (e) {
    return null;
  }
}

/** Arma una hoja con la identidad de OW: logo arriba, título con los colores de marca,
 * y los datos en forma de Tabla de Excel real (con flechitas de filtro y franjas de
 * color alternadas), con el encabezado forzado al azul marino exacto de OW. */
function agregarHojaOW(wb, logoImageId, nombreHoja, columnas, filas, subtitulo) {
  const ws = wb.addWorksheet(nombreHoja.slice(0, 31), {
    views: [{ state: "frozen", ySplit: 5, showGridLines: false }],
  });

  if (logoImageId !== null) {
    ws.addImage(logoImageId, { tl: { col: 0.15, row: 0.15 }, ext: { width: 150, height: 52 } });
  }
  ws.getCell("C1").value = "Oñate, Willy y Cía., S.C.";
  ws.getCell("C1").font = { bold: true, size: 15, color: { argb: OW_NAVY } };
  ws.getCell("C2").value = subtitulo || nombreHoja;
  ws.getCell("C2").font = { bold: true, size: 12, color: { argb: OW_ACCENT } };
  ws.getCell("C3").value = `Generado ${new Date().toLocaleString("es-MX")} — ${filas.length} registro${filas.length === 1 ? "" : "s"}`;
  ws.getCell("C3").font = { italic: true, size: 9, color: { argb: OW_MUTED } };
  ws.getRow(1).height = 22;

  const filaEncabezado = 5;
  if (filas.length === 0) {
    ws.getCell(filaEncabezado, 1).value = "Sin registros";
    ws.getCell(filaEncabezado, 1).font = { italic: true, color: { argb: OW_MUTED } };
    return ws;
  }

  ws.addTable({
    name: "Tabla_" + nombreHoja.replace(/[^a-zA-Z0-9_]/g, "_") + "_" + Math.random().toString(36).slice(2, 6),
    ref: `A${filaEncabezado}`,
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: columnas.map((c) => ({ name: c, filterButton: true })),
    rows: filas,
  });

  // Se fuerzan los colores exactos de marca OW en el encabezado, encima del tema
  // genérico que trae la Tabla de Excel por defecto.
  const header = ws.getRow(filaEncabezado);
  header.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: OW_NAVY } };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle" };
  });
  header.height = 20;

  columnas.forEach((label, idx) => {
    let maxLen = String(label || "").length;
    filas.forEach((r) => {
      const v = r[idx];
      if (v != null) maxLen = Math.max(maxLen, String(v).length);
    });
    ws.getColumn(idx + 1).width = Math.min(Math.max(maxLen + 3, 11), 44);
  });

  return ws;
}

async function buildExcelWorkbook(days) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Oñate, Willy y Cía., S.C.";
  wb.created = new Date();
  const logoBuf = await cargarLogoOW();
  const logoImageId = logoBuf !== null ? wb.addImage({ buffer: logoBuf, extension: "jpeg" }) : null;

  const resumenCols = ["Fecha", "Tramitador", "Aduana", "Total previos", "Total despachos", "Total revalidadas", "Archivo origen", "Capturado por"];
  const resumenRows = [];
  days.forEach((d) =>
    (d.capturas || []).forEach((c) =>
      resumenRows.push([
        c.fecha,
        c.tramitador,
        c.aduana,
        (c.previos || []).length,
        (c.despachos || []).length,
        (c.revalidadas || []).length,
        c.sourceFileName || "(captura manual)",
        c.uploadedBy || "",
      ])
    )
  );
  agregarHojaOW(wb, logoImageId, "Resumen", resumenCols, resumenRows, "Reporte operativo — Resumen");

  const previosCols = ["Fecha", "Tramitador", ...PREVIO_FIELDS.map((f) => f.label)];
  const previosRows = [];
  days.forEach((d) =>
    (d.capturas || []).forEach((c) =>
      (c.previos || []).forEach((r) => previosRows.push([c.fecha, c.tramitador, ...PREVIO_FIELDS.map((f) => r[f.k] || "")]))
    )
  );
  agregarHojaOW(wb, logoImageId, "Previos", previosCols, previosRows, "Control de previos");

  const despachosCols = ["Fecha", "Tramitador", ...DESPACHO_FIELDS.map((f) => f.label)];
  const despachosRows = [];
  days.forEach((d) =>
    (d.capturas || []).forEach((c) =>
      (c.despachos || []).forEach((r) => despachosRows.push([c.fecha, c.tramitador, ...DESPACHO_FIELDS.map((f) => r[f.k] || "")]))
    )
  );
  agregarHojaOW(wb, logoImageId, "Despachos", despachosCols, despachosRows, "Control de despachos");

  const revalidadasCols = ["Fecha", "Tramitador", ...REVALIDADA_FIELDS.map((f) => f.label)];
  const revalidadasRows = [];
  days.forEach((d) =>
    (d.capturas || []).forEach((c) =>
      (c.revalidadas || []).forEach((r) => revalidadasRows.push([c.fecha, c.tramitador, ...REVALIDADA_FIELDS.map((f) => r[f.k] || "")]))
    )
  );
  agregarHojaOW(wb, logoImageId, "Revalidadas", revalidadasCols, revalidadasRows, "Control de guías revalidadas");

  const pendientesCols = ["Fecha", "Tramitador", "Dejado por", ...PENDIENTE_FIELDS.map((f) => f.label)];
  const pendientesRows = [];
  days.forEach((d) =>
    (d.capturas || []).forEach((c) =>
      (c.pendientes || []).forEach((r) => pendientesRows.push([c.fecha, c.tramitador, c.uploadedBy || "", ...PENDIENTE_FIELDS.map((f) => r[f.k] || "")]))
    )
  );
  agregarHojaOW(wb, logoImageId, "Pendientes", pendientesCols, pendientesRows, "Pendientes");

  return wb;
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(String(reader.result).split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
function base64ToBlob(base64, mime) {
  const byteChars = atob(base64);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
  return new Blob([new Uint8Array(byteNumbers)], { type: mime });
}

async function buildExcelBlob(days) {
  const wb = await buildExcelWorkbook(days);
  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

async function buildAndDownloadExcel(days, filename) {
  const wb = await buildExcelWorkbook(days);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ---------- link especial: descarga automática del Excel compartido, sin login ----------
async function handleHistoricoRoute(aduana) {
  state.aduanaActiva = aduana;
  root.innerHTML = `<div style="padding:60px 20px;text-align:center;font-family:'IBM Plex Sans',sans-serif;">
    <div class="spinner" style="margin:0 auto 14px;"></div>
    <div>Cargando el Excel más reciente de ${esc(aduana)}…</div>
  </div>`;
  try {
    let hist = null;
    let fromLocalCache = false;
    if (navigator.onLine) {
      try {
        hist = await loadHistoricoExcel(aduana);
      } catch (e) {
        hist = null;
      }
    }
    if (!hist || !hist.data) {
      hist = loadLocalHistoricoExcel();
      fromLocalCache = true;
    }
    if (!hist || !hist.data) {
      root.innerHTML = `<div style="padding:60px 20px;text-align:center;font-family:'IBM Plex Sans',sans-serif;">
        Todavía no se ha generado ningún Excel compartido en este dispositivo. Entra a la app y guarda al menos un reporte primero.<br/><br/>
        <a href="${window.location.origin}${window.location.pathname}" style="color:#2B6E63;font-weight:600;">Ir a la app</a>
      </div>`;
      return;
    }
    const blob = base64ToBlob(hist.data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte_historico_${aduana}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    root.innerHTML = `<div style="padding:60px 20px;text-align:center;font-family:'IBM Plex Sans',sans-serif;">
      <div style="font-size:40px;margin-bottom:10px;">✅</div>
      <div style="font-weight:600;margin-bottom:6px;">Descarga iniciada</div>
      ${
        fromLocalCache
          ? `<div style="color:#8A6414;font-size:12.5px;margin-bottom:10px;">📴 Sin conexión: esta es la última copia guardada en este dispositivo, puede no ser la más reciente.</div>`
          : ""
      }
      <div style="color:#5B6B72;font-size:13px;margin-bottom:16px;">Actualizado por última vez: ${
        hist.actualizado ? new Date(hist.actualizado).toLocaleString("es-MX") : "—"
      }</div>
      <div style="color:#5B6B72;font-size:12px;margin-bottom:16px;">Guarda este link para volver a descargar la versión más reciente cuando quieras.</div>
      <a href="${window.location.origin}${window.location.pathname}" style="color:#2B6E63;font-weight:600;">Ir a la app</a>
    </div>`;
  } catch (e) {
    root.innerHTML = `<div style="padding:60px 20px;text-align:center;font-family:'IBM Plex Sans',sans-serif;color:#C0453B;">
      No se pudo cargar el Excel compartido (${esc(e.message)}).<br/><br/>
      <a href="${window.location.origin}${window.location.pathname}" style="color:#2B6E63;font-weight:600;">Ir a la app</a>
    </div>`;
  }
}

// ---------- instalación como app (PWA) ----------
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  state.installPromptEvent = e;
  render();
});

// ---------- detección de conexión ----------
window.addEventListener("online", () => {
  state.isOnline = true;
  render();
  processPendingQueue().then(() => {
    if (state.user && !state.loading) loadReports();
  });
});
window.addEventListener("offline", () => {
  state.isOnline = false;
  render();
});
// Red de seguridad: si algo truena en un botón/acción (ej. error de programación no
// previsto), que se vea en pantalla en vez de "no pasar nada" silenciosamente.
window.addEventListener("error", (e) => {
  if (state.view === "login" || state.view === "selectAduana") return; // no molestar en pantallas de entrada
  state.errorMsg = "Ocurrió un error inesperado (" + (e.message || "desconocido") + "). Si vuelve a pasar, avísale a soporte.";
  render();
});
// respaldo: revisa cada 20s por si el navegador no dispara el evento "online" de forma confiable
setInterval(() => {
  if (navigator.onLine && pendingFechas().length > 0 && !state.syncingPending) {
    processPendingQueue();
  }
}, 20000);
// refresco silencioso cada 25s, para que borrados/cambios de otra persona se vean
// sin tener que cerrar y volver a entrar
setInterval(() => {
  backgroundRefresh();
}, 25000);

// ---------- init ----------
const urlParams = new URLSearchParams(window.location.search);
if (urlParams.get("historico")) {
  handleHistoricoRoute(urlParams.get("aduana") || "GDL");
} else {
  loadLastUser();
  if (state.user && state.userRole && state.aduanaActiva) {
    state.view = state.userRole === "ejecutivo" ? "completarReferencias" : "home";
    render();
    state.loading = true;
    render();
    loadReports();
    if (navigator.onLine && pendingFechas().length > 0) {
      processPendingQueue();
    }
  } else if (state.aduanaActiva) {
    // ya se eligió aduana antes, pero falta iniciar sesión con un rol
    state.view = "login";
    render();
  } else {
    render();
  }
}
