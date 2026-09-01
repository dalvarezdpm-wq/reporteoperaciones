import "./style.css";
import ExcelJS from "exceljs";
import { PREVIO_FIELDS, DESPACHO_FIELDS, REVALIDADA_FIELDS, PENDIENTE_FIELDS, emptyCaptura } from "./fields.js";
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
  ejecutivosYaReparados,
  marcarEjecutivosReparados,
  obtenerUltimoRespaldo,
  guardarRespaldo,
  listarRespaldos,
  descargarRespaldo,
  eliminarRespaldo,
} from "./storage.js";

// (roles de login ahora se arman dinámicamente en viewLogin, ver ahí)
const TRAMITADORES = ["Monica Ortega", "Luis Arreola", "Mariana Carrillo", "Mayra Romero", "Javier Garcia", "Julio Regalado"];
const EJECUTIVOS = ["Alberto Pichardo", "Fernanda Narez", "Claudia Barrera", "Fernanda Ramirez", "Ulises Bautista", "Erendira Calderon"];
const ADUANAS = ["GDL", "TOLUCA", "AIFA", "AICM", "QUERETARO"];
const ADMIN_PASSWORDS = { GDL: "ow2026", TOLUCA: "OWTOLUCA", AIFA: "OWAIFA", AICM: "OWAICM", QUERETARO: "OWQUERETARO" };
/** Algunas aduanas trabajan con un único almacén fijo (ej. Querétaro, por manifiesto) —
 * si la aduana activa aparece aquí, el campo Almacén se autocompleta y queda bloqueado
 * en todas las capturas y en el formulario de asignaciones, para que nadie lo cambie
 * por error. */
const ALMACEN_FIJO = { QUERETARO: "210 TERMINAL" };
function almacenFijo() {
  return ALMACEN_FIJO[state.aduanaActiva] || null;
}
/** Todas las aduanas dejan que cada tramitador se autoasigne las guías que va a
 * trabajar de la "bolsa" de asignaciones sin tramitador (ej. Coordinación asigna
 * masivo sin decidir quién hace cada una, y cada tramitador entra y agarra las suyas). */
const AUTOASIGNACION_ADUANAS = [...ADUANAS];
function permiteAutoasignacion() {
  return AUTOASIGNACION_ADUANAS.includes(state.aduanaActiva);
}
const APP_VERSION = "5.48.0";

let editableCats = {
  clientes: ["GLXI", "Alkaps", "MTI", "FMI", "IndoUnión", "Foray", "Alpha metal", "BRP", "PMI"],
  almacenes: ["228", "277", "CLA", "WTC"],
  tramitadores: [...TRAMITADORES],
  ejecutivos: [...EJECUTIVOS],
};

/**
 * Rellena con arreglos vacíos (o valores por defecto razonables) cualquier clave
 * que falte en un catálogo cargado de Firebase o de la caché local — esto evita que
 * la app truene si el catálogo se guardó con una versión anterior que no tenía
 * todas las categorías actuales (ej. "tramitadores" se agregó después).
 */
function sanitizeCats(cats) {
  // El catálogo de tramitadores/ejecutivos de arranque es el equipo de GDL —
  // solo debe usarse como semilla ahí. Una aduana nueva (ej. Toluca) debe empezar
  // con su catálogo vacío, para que Coordinación dé de alta a SU propio equipo,
  // sin heredar por accidente los nombres de otra aduana.
  const esGDL = state.aduanaActiva === "GDL";
  const fijo = almacenFijo();
  const base = {
    clientes: [],
    almacenes: fijo ? [fijo] : [],
    tramitadores: esGDL ? [...TRAMITADORES] : [],
    ejecutivos: esGDL ? [...EJECUTIVOS] : [],
    coordinadores: [],
  };
  const out = { ...base, ...(cats || {}) };
  Object.keys(base).forEach((k) => {
    if (!Array.isArray(out[k])) out[k] = base[k];
  });
  if (fijo && !out.almacenes.includes(fijo)) out.almacenes.push(fijo);
  return out;
}

/**
 * Garantiza que una captura (nueva o cargada de un reporte guardado con una versión
 * anterior de la app) siempre tenga todos sus arreglos presentes — evita que funciones
 * como "Traer asignaciones" o "Marcar Listo" fallen en silencio con reportes viejos
 * a los que les falte algún campo (ej. "pendientes", agregado en una versión posterior).
 */
/** Reduce el tamaño de una foto (ancho máximo + calidad JPEG) antes de guardarla —
 * una foto de cámara sin comprimir puede pesar varios MB, y eso sí infla rápido la
 * base de datos (a diferencia del resto de los datos de la app, que son puro texto).
 * Regresa un data URL (base64) ya comprimido, listo para guardarse en el renglón.
 *
 * Usa URL.createObjectURL en vez de FileReader.readAsDataURL a propósito: leer el
 * archivo completo como texto base64 primero (antes de siquiera empezar a comprimir)
 * puede usar mucha memoria de golpe con una foto de cámara sin comprimir (varios MB se
 * vuelven ~33% más grandes como texto) — en celulares con poca memoria libre, eso podía
 * tumbar la pestaña/app a medio proceso, antes de llegar siquiera a comprimir nada.
 * createObjectURL es una referencia ligera al archivo, sin copiarlo entero a memoria. */
function comprimirImagenArchivo(file, maxAncho = 900, calidad = 0.55) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const limpiar = () => {
      try {
        URL.revokeObjectURL(url);
      } catch (e) {}
    };
    const img = new Image();
    img.onload = () => {
      try {
        let w = img.width,
          h = img.height;
        if (w > maxAncho) {
          h = Math.round(h * (maxAncho / w));
          w = maxAncho;
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, w, h);
        const dataUrl = canvas.toDataURL("image/jpeg", calidad);
        limpiar();
        resolve(dataUrl);
      } catch (err) {
        limpiar();
        reject(err);
      }
    };
    img.onerror = () => {
      limpiar();
      reject(new Error("No se pudo leer la imagen"));
    };
    img.src = url;
  });
}

function normalizeCaptura(c) {
  if (!Array.isArray(c.previos)) c.previos = [];
  if (!Array.isArray(c.despachos)) c.despachos = [];
  if (!Array.isArray(c.revalidadas)) c.revalidadas = [];
  if (!Array.isArray(c.pendientes)) c.pendientes = [];
  if (!Array.isArray(c.otrasActividades)) c.otrasActividades = ["", "", "", ""];
  if (!Array.isArray(c.historial)) c.historial = [];
  // Si esta captura es de una aduana con almacén fijo (ej. Querétaro), se fuerza ese
  // valor en todos sus renglones — así el Excel, el CSV y el detalle del reporte siempre
  // muestran el almacén correcto, no solo el formulario de captura.
  const fijo = ALMACEN_FIJO[c.aduana];
  if (fijo) {
    ["previos", "despachos", "revalidadas", "pendientes"].forEach((grupo) => {
      (c[grupo] || []).forEach((r) => {
        if ("almacen" in r) r.almacen = fijo;
      });
    });
  }
  return c;
}

/** Contraseña automática de tramitadores: 3 letras iniciales + "2026" para GDL (como
 * ya se venía usando, para no invalidar las que ya se compartieron), y 4 letras iniciales
 * + "2026" para cualquier otra aduana (ej. Toluca) — a propósito distinto de GDL. */
function tramitadorPasswordsMap(lista, aduana) {
  const map = {};
  const usadas = new Set();
  const lenInicial = aduana === "GDL" ? 3 : 4;
  (lista || []).forEach((nombre) => {
    const letras = (nombre || "").trim().toLowerCase().replace(/[^a-záéíóúñ]/g, "");
    let len = lenInicial;
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
  return tramitadorPasswordsMap(editableCats.tramitadores, state.aduanaActiva)[nombre] || "";
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
  asigDraft: { tipo: "previo", ref: "", guia: "", cliente: "", almacen: "", sector: "", ejecutivo: "", pedimento: "", tramitador: "" },
  asigMasivaOpen: false,
  asigMasivaDraft: { tipo: "previo", sector: "", tramitador: "", textoGuias: "", textoClientes: "", textoRefs: "", textoEjecutivos: "", textoPedimentos: "" },
  asigMasivaMsg: "",
  asigEditId: null,
  notifBannerDismissed: false,
  recentNotifications: [],
  respaldosList: null,
  homologarEjecutivosMsg: "",
  homologarClientesMsg: "",
  soloSinRastro: false,
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
/** La mayoría de las aduanas cierran operaciones a las 8:00 PM, pero Querétaro trabaja
 * hasta medianoche — un valor de 24 nunca se activa antes de las 00:00 reales, así que
 * en la práctica el día cambia justo a medianoche para ellos, como cualquier calendario
 * normal. */
const HORA_CIERRE_ADUANA = { QUERETARO: 24 };
function horaCierreAduana() {
  return HORA_CIERRE_ADUANA[state.aduanaActiva] !== undefined ? HORA_CIERRE_ADUANA[state.aduanaActiva] : 20;
}
function todayStr() {
  // La aduana cierra operaciones a una hora fija (ver horaCierreAduana) — de ahí en
  // adelante, para la app ya empezó el "día operativo" siguiente (aunque el reloj real
  // todavía no llegue a medianoche). Esto hace que, por ejemplo, los pendientes sin
  // resolver a esa hora ya se puedan trasladar a "mañana" desde el cierre real de
  // operaciones, no hasta las 00:00 (salvo en aduanas cuyo cierre YA es medianoche).
  const HORA_CIERRE = horaCierreAduana();
  const ahora = new Date();
  return ahora.getHours() >= HORA_CIERRE ? dateStrLocal(-1) : dateStrLocal(0);
}
/** Igual que todayStr(), pero para una fecha/hora cualquiera (no "ahora") — a qué "día
 * operativo" pertenece ese momento, con el mismo corte que use la aduana activa. Sirve
 * para saber si algo se creó "hoy" (en el sentido operativo) o si ya cruzó el cierre del
 * día sin resolverse — ej. una asignación sin tramitador creada a las 9pm ya cuenta como
 * de "mañana", igual que pasaría con un pendiente sin resolver a esa hora. */
function diaOperativoDe(fechaISO) {
  if (!fechaISO) return "";
  const HORA_CIERRE = horaCierreAduana();
  const d = new Date(fechaISO);
  const diasAtras = d.getHours() >= HORA_CIERRE ? -1 : 0;
  const ajustada = new Date(d);
  ajustada.setDate(ajustada.getDate() - diasAtras);
  const local = new Date(ajustada.getTime() - ajustada.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
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
  } catch (e) {
    // Lo más probable, ahora que hay fotos, es que se llenó el espacio local del
    // navegador — se avisa una sola vez por sesión para no ser repetitivo, ya que
    // Firebase (si hay señal) sigue funcionando normal; lo que se pierde es solo el
    // respaldo para el modo sin conexión de este dispositivo en particular.
    if (!state._avisoCacheLocalMostrado) {
      state._avisoCacheLocalMostrado = true;
      state.errorMsg = "No se pudo guardar la copia local de respaldo en este dispositivo (probablemente se llenó su espacio de almacenamiento) — el modo sin conexión podría mostrar datos desactualizados. Esto no afecta lo que ya está guardado en la nube.";
      render();
    }
  }
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

/** Repara datos viejos, UNA SOLA VEZ por aduana: antes de que existiera la columna
 * "Ejecutivo" en Previos/Despachos/Pendientes, ese dato se guardaba en la Asignación
 * pero se perdía al traerla al reporte. Esta reparación busca, para cada guía sin
 * ejecutivo en algún renglón, si existe una asignación con esa misma guía que sí tenga
 * ejecutivo guardado, y lo copia. Solo rellena huecos — nunca pisa uno que ya esté
 * puesto. Usa una bandera permanente para no volver a correr jamás, ni con reportes
 * nuevos que lleguen a faltarles el dato por otra razón — es solo el arreglo del
 * hueco que dejó esta columna al agregarse tarde, no algo que deba repetirse solo. */
async function repararEjecutivosSiNecesario() {
  try {
    if (await ejecutivosYaReparados(state.aduanaActiva)) return;
    const norm = (s) => (s || "").trim().toLowerCase();
    const mapaEjecutivos = {};
    Object.values(state.asignaciones || {})
      .filter((a) => (a.guia || "").trim() && (a.ejecutivo || "").trim())
      .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || "")) // ascendente: la más reciente pisa al final
      .forEach((a) => {
        mapaEjecutivos[norm(a.guia)] = a.ejecutivo;
      });

    const diasParaGuardar = new Set();
    if (Object.keys(mapaEjecutivos).length > 0) {
      Object.values(state.reports || {}).forEach((day) => {
        (day.capturas || []).forEach((c) => {
          normalizeCaptura(c);
          ["previos", "despachos", "pendientes"].forEach((grupo) => {
            (c[grupo] || []).forEach((r) => {
              if (!("ejecutivo" in r)) return; // ej. revalidadas no tiene esta columna
              if ((r.ejecutivo || "").trim()) return; // ya tiene uno puesto, no se toca
              const key = norm(r.guia);
              if (key && mapaEjecutivos[key]) {
                r.ejecutivo = mapaEjecutivos[key];
                diasParaGuardar.add(day.fecha);
              }
            });
          });
        });
      });
    }

    for (const fecha of diasParaGuardar) {
      try {
        await saveReportDay(fecha);
      } catch (e) {}
    }
    if (diasParaGuardar.size > 0) render(); // refleja el cambio en pantalla si algo estaba abierto
    await marcarEjecutivosReparados(state.aduanaActiva);
  } catch (e) {} // no crítico — si falla, simplemente se reintenta la próxima vez que cargue (no se marcó la bandera)
}

/** Respaldo automático completo (reportes + catálogos + asignaciones), como máximo una
 * vez al día por aduana — para tener siempre un punto de recuperación reciente sin
 * depender de que alguien se acuerde de descargar el Excel histórico a mano. Se guarda
 * en una rama aparte de Firebase (no se mezcla con los datos "en vivo"), y se conservan
 * solo los últimos 14 respaldos por aduana para no crecer sin límite — los más viejos
 * se borran solos al pasarse de ese número. */
async function respaldoAutomaticoSiNecesario() {
  try {
    const ultimo = await obtenerUltimoRespaldo(state.aduanaActiva);
    const HORAS_MINIMAS = 20; // margen bajo 24h por si la hora exacta de entrada varía día a día
    if (ultimo && (Date.now() - ultimo) / 3600000 < HORAS_MINIMAS) return;

    const datos = {
      reportes: state.reports,
      catalogos: editableCats,
      asignaciones: state.asignaciones,
      guardadoEl: new Date().toISOString(),
    };
    const clave = todayStr() + "_" + Date.now();
    await guardarRespaldo(state.aduanaActiva, clave, datos);

    const lista = await listarRespaldos(state.aduanaActiva);
    const LIMITE = 14;
    if (lista.length > LIMITE) {
      const sobrantes = lista.slice(0, lista.length - LIMITE); // ya vienen ordenados por clave (fecha)
      for (const key of sobrantes) {
        try {
          await eliminarRespaldo(state.aduanaActiva, key);
        } catch (e) {}
      }
    }
  } catch (e) {} // no crítico — se reintenta el próximo día
}

async function loadReports() {
  if (!navigator.onLine) {
    const cached = loadLocalCache();
    state.reports = cached.reports || {};
    if (cached.catalogos) editableCats = sanitizeCats(cached.catalogos);
    state.usingLocalCache = true;
  } else {
    // Si quedó algo pendiente de subir de una sesión anterior (por falta de señal, o
    // porque falló el guardado por cualquier otro motivo), se intenta subir ANTES de
    // traer datos frescos — así, si ya hay señal ahora, no se vuelve a tapar lo nuevo
    // del servidor con una copia local vieja que ya no hacía falta conservar. Sin esto,
    // una copia pendiente que nunca se llegó a subir podía revivir sola cada vez que la
    // app cargaba, borrando en la pantalla cualquier cambio que otra persona hubiera
    // hecho después en ese mismo día.
    if (pendingFechas().length > 0) {
      try {
        await processPendingQueue();
      } catch (e) {}
    }
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
  // Reparación de datos viejos de una sola vez (ver función) — solo la dispara Admin,
  // y solo con conexión, para no repetir trabajo ni hacerlo desde una sesión de trámite.
  if (state.userRole === "admin" && navigator.onLine) {
    repararEjecutivosSiNecesario();
    respaldoAutomaticoSiNecesario();
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
  // Igual que en loadReports: si algo se quedó pendiente de subir, se intenta aquí
  // también — esto pasa cada 25 segundos mientras haya señal, así que una copia
  // pendiente atorada tiene muchas más oportunidades de subirse sola, en vez de
  // quedarse tapando lo nuevo del servidor indefinidamente hasta que alguien note
  // el problema.
  if (pendingFechas().length > 0) {
    try {
      await processPendingQueue();
    } catch (e) {}
  }
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
      if (state.userRole === "tramite" && a.estatus === "pendiente" && mismoNombre(a.tramitador, state.user)) {
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
          <div style="font-size:12px;color:var(--ink);">Cliente: ${esc(a.cliente) || "—"} · Almacén: ${esc(a.almacen) || "—"}${a.ref ? ` · Ref.: ${esc(a.ref)}` : ""}${
                a.pedimento ? ` · Pedimento: ${esc(a.pedimento)}` : ""
              }${a.sector ? ` · Sector: ${esc(a.sector)}` : ""}${a.ejecutivo ? ` · Ejecutivo: ${esc(a.ejecutivo)}` : ""}</div>
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
    if (!mismoNombre(a.tramitador, tramitador)) continue;
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
    return true;
  } catch (e) {
    return false;
  }
}
function markPending(fecha, dayDataOrNullForDelete) {
  const map = getPendingMap();
  map[fecha] = { data: dayDataOrNullForDelete, timestamp: new Date().toISOString() };
  const ok = setPendingMap(map);
  if (!ok) {
    // Caso serio: no se pudo ni siquiera guardar la señal de "esto quedó pendiente de
    // subir" — sin esto, el cambio se perdería sin dejar ningún rastro para recuperarlo
    // después. Se avisa fuerte y de inmediato, no en silencio como el resto de los
    // catches de este archivo.
    state.errorMsg =
      "⚠️ No se pudo guardar tu cambio para reintentar después (el dispositivo se quedó sin espacio de almacenamiento) — sin señal, este cambio específico podría perderse. Intenta guardar de nuevo en cuanto tengas internet.";
    render();
  }
}
function clearPending(fecha) {
  const map = getPendingMap();
  delete map[fecha];
  setPendingMap(map);
}
function pendingFechas() {
  return Object.keys(getPendingMap());
}
/** De todo lo que sigue sin subirse, cuánto lleva atorado MÁS de lo normal — con los
 * reintentos automáticos (al cargar la app y cada 25 segundos con señal), algo debería
 * subirse solo en cuestión de segundos si de verdad hay internet. Si algo lleva más de
 * 10 minutos sin lograrlo, ya no es una simple demora — hay que avisar claro, en vez de
 * dejarlo revivir en silencio cada vez que la app carga. */
function pendientesAtoradosCount() {
  const LIMITE_MINUTOS = 10;
  const ahora = Date.now();
  return Object.values(getPendingMap()).filter((entry) => {
    const t = entry.timestamp ? new Date(entry.timestamp).getTime() : 0;
    return t && (ahora - t) / 60000 >= LIMITE_MINUTOS;
  }).length;
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

/** Si se está editando un reporte que YA existe guardado, guarda el cambio de inmediato
 * (foto, solicitud de eliminación, etc.) sin esperar al botón final "Guardar cambios" —
 * buscando la hoja por su ID único, NUNCA por su posición en el arreglo. Esto importa de
 * verdad cuando varias personas trabajan en la app al mismo tiempo: mientras alguien edita
 * su hoja, otra persona puede guardar la suya del mismo día, cambiando el orden — buscar
 * por ID evita que ese cambio de orden termine actualizando o borrando la hoja de alguien
 * más por accidente. Si la hoja ya no está ahí (se movió o se borró mientras tanto), no
 * hace nada — mejor eso a arriesgarse a escribir en el lugar equivocado. */
async function guardarCambioInmediatoSiExiste() {
  if (!state.editingFecha || !state.currentCaptura || !state.currentCaptura.id) return;
  const day = state.reports[state.editingFecha];
  if (!day || !Array.isArray(day.capturas)) return;
  const idx = day.capturas.findIndex((c) => c.id === state.currentCaptura.id);
  if (idx === -1) return;
  day.capturas[idx] = state.currentCaptura;
  persistLocalCache();
  try {
    await saveReportDay(state.editingFecha);
  } catch (e) {}
}

/** Recorre los días ANTERIORES al de hoy buscando pendientes de este tramitador que
 * sigan sin resolver (sin campo "bloqueado" todavía) — los junta para llevárselos al día
 * de hoy, y congela ("bloqueado": true) el renglón original en su día viejo, para que
 * quede fijo ahí como respaldo histórico: si alguien borra por error la copia de hoy, el
 * renglón de ese día viejo sigue existiendo tal cual. No distingue si el pendiente vino de
 * una asignación o se escribió a mano — aplica por igual a cualquiera.
 *
 * Incluye una reconciliación de seguridad: si un traslado anterior se quedó a medias (el
 * pendiente viejo se marcó "bloqueado" pero, por lo que sea — se cerró la app a medio
 * camino, falló el guardado de ese día — la copia nunca llegó a guardarse en ningún día
 * posterior), se destraba solo para que se vuelva a intentar, en vez de quedar perdido
 * para siempre sin que nadie lo note. */
async function migrarPendientesAlDiaDeHoy(tramitador) {
  const hoy = todayStr();
  const todosLosDias = Object.values(state.reports).sort((a, b) => a.fecha.localeCompare(b.fecha));
  const diasModificados = new Set();

  // --- Reconciliación: destraba pendientes "bloqueados" que nunca llegaron a tener
  // una copia real en ningún día posterior. ---
  todosLosDias.forEach((day) => {
    if (day.fecha >= hoy) return;
    (day.capturas || []).forEach((c) => {
      if (!mismoNombre(c.tramitador, tramitador)) return;
      (c.pendientes || []).forEach((p) => {
        if (!p.bloqueado) return;
        if (p.descartado) return; // se borró a propósito — nunca se vuelve a traer de regreso
        const guia = (p.guia || "").trim();
        if (!guia) return; // sin guía no hay forma confiable de rastrear la copia, se deja como está
        const origen = p.origenFecha || day.fecha;
        const yaLlego = todosLosDias.some(
          (otroDia) =>
            otroDia.fecha > day.fecha &&
            (otroDia.capturas || []).some(
              (c2) =>
                mismoNombre(c2.tramitador, tramitador) &&
                (c2.pendientes || []).some((p2) => p2.origenFecha === origen && (p2.guia || "").trim() === guia)
            )
        );
        if (!yaLlego) {
          p.bloqueado = false; // se destraba: se vuelve a intentar el traslado más abajo
          diasModificados.add(day.fecha);
        }
      });
    });
  });

  // --- Traslado normal de lo que sigue pendiente y sin trasladar ---
  const heredados = [];
  todosLosDias
    .filter((d) => d.fecha < hoy)
    .forEach((day) => {
      (day.capturas || []).forEach((c) => {
        if (!mismoNombre(c.tramitador, tramitador)) return;
        (c.pendientes || []).forEach((p) => {
          if (p.bloqueado) return; // ya se había migrado antes, no se vuelve a mover
          p.bloqueado = true;
          if (!p.origenFecha) p.origenFecha = day.fecha;
          const copia = { ...p };
          delete copia.bloqueado;
          heredados.push(copia);
          diasModificados.add(day.fecha);
        });
      });
    });

  for (const fecha of diasModificados) {
    try {
      await saveReportDay(fecha);
    } catch (e) {}
  }
  return heredados;
}

/** Junta TODOS los pendientes abiertos (sin "bloqueado") de todos los tramitadores y
 * todos los días — para la pantalla de "Pendientes" que ve Admin/Coordinación, donde se
 * puede reasignar cualquiera a otra persona, venga o no de una asignación formal. */
function todosPendientesAbiertos() {
  const list = [];
  Object.values(state.reports).forEach((day) => {
    (day.capturas || []).forEach((c) => {
      (c.pendientes || []).forEach((p) => {
        if (p.bloqueado) return;
        list.push({ fecha: day.fecha, tramitador: c.tramitador, capturaId: c.id, row: p });
      });
    });
  });
  return list.sort((a, b) => b.fecha.localeCompare(a.fecha));
}

/** Cuenta (sin modificar nada todavía) cuántos pendientes de días anteriores de este
 * tramitador siguen sin trasladar — para mostrar el número en el botón "Traer pendientes"
 * antes de que la persona le dé clic. */
function pendientesSinTrasladarCount(tramitador) {
  const hoy = todayStr();
  const t = (tramitador || "").trim();
  if (!t) return 0;
  let n = 0;
  Object.values(state.reports).forEach((day) => {
    if (day.fecha >= hoy) return;
    (day.capturas || []).forEach((c) => {
      if (!mismoNombre(c.tramitador, t)) return;
      (c.pendientes || []).forEach((p) => {
        if (!p.bloqueado) n++;
      });
    });
  });
  return n;
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
  } catch (e) {
    state.errorMsg = "No se pudo guardar el catálogo en la nube (" + e.message + ") — el cambio se ve aquí, pero vuelve a intentarlo con internet para que no se pierda ni se quede desactualizado en otros dispositivos.";
    render();
  }
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
      const key = normalizarNombre(t);
      if (!counts[key]) counts[key] = { nombre: nombreCanonicoCliente(t), n: 0 };
      counts[key].n += 1;
    });
  });
  return Object.values(counts)
    .sort((a, b) => b.n - a.n)
    .slice(0, 5)
    .map((c) => [c.nombre, c.n]);
}

/** Escalas de puntos por Dificultad — cada tipo de operación tiene su propia escala
 * (los previos pesan más que los despachos). Un renglón sin dificultad asignada (el
 * campo es opcional) cuenta como A, el mínimo de su escala, para no descartar el
 * trabajo hecho, pero sin premiarlo como si fuera complejo. */
const PUNTOS_PREVIO = { A: 0.06, B: 0.1, C: 0.14, D: 0.18 };
const PUNTOS_DESPACHO = { A: 0.02, B: 0.05, C: 0.08, D: 0.11 };
/** Un "2do previo" (numPrevio) resta en vez de sumar — indica que el previo tuvo que
 * repetirse, así que penaliza la productividad en vez de contarla normal, sin importar
 * la dificultad que se le haya puesto. */
const PUNTOS_SEGUNDO_PREVIO = -0.06;
const PUNTOS_REVALIDADA = 0.02;
const PUNTOS_OTRAS_ACTIVIDADES = 0.02;
/** Convierte cualquier valor crudo del campo Dificultad a una letra válida (A/B/C/D) —
 * en blanco o cualquier cosa rara cuenta como A, el mínimo. */
function normalizarDificultad(valor) {
  const v = (valor || "").trim().toUpperCase();
  return ["A", "B", "C", "D"].includes(v) ? v : "A";
}

/** Productividad ponderada por dificultad: a diferencia del conteo simple de operaciones,
 * aquí cada tipo de trabajo suma según qué tan difícil fue y qué tan pesado es ese tipo
 * de operación — así se distingue a alguien con pocas operaciones pero muy complejas de
 * alguien con muchas operaciones sencillas. Un "2do previo" resta en vez de sumar (ver
 * PUNTOS_SEGUNDO_PREVIO). Guías revalidadas y Otras actividades suman un valor fijo cada
 * una, sin escala de dificultad. También lleva el conteo individual de cuántos A, B, C y
 * D hizo cada quien (en Previos y Despachos), para ver el desglose y no solo el total. */
function statsProductividadPonderada(desde, hasta) {
  const map = {};
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      const t = (c.tramitador || "").trim();
      if (!t) return;
      const key = normalizarNombre(t);
      if (!map[key])
        map[key] = {
          nombre: nombreCanonico(t),
          puntos: 0,
          operaciones: 0,
          conteo: { A: 0, B: 0, C: 0, D: 0 },
          revalidadas: 0,
          otrasActividades: 0,
          segundosPrevios: 0,
        };
      (c.previos || []).forEach((r) => {
        const dif = normalizarDificultad(r.tipoPrevio);
        map[key].conteo[dif] += 1;
        if (r.numPrevio === "2do previo") {
          map[key].puntos += PUNTOS_SEGUNDO_PREVIO;
          map[key].segundosPrevios += 1;
        } else {
          map[key].puntos += PUNTOS_PREVIO[dif];
        }
        map[key].operaciones += 1;
      });
      (c.despachos || []).forEach((r) => {
        const dif = normalizarDificultad(r.dificultad);
        map[key].conteo[dif] += 1;
        map[key].puntos += PUNTOS_DESPACHO[dif];
        map[key].operaciones += 1;
      });
      (c.revalidadas || []).forEach(() => {
        map[key].puntos += PUNTOS_REVALIDADA;
        map[key].revalidadas += 1;
        map[key].operaciones += 1;
      });
      (c.otrasActividades || []).forEach((texto) => {
        if (!(texto || "").trim()) return;
        map[key].puntos += PUNTOS_OTRAS_ACTIVIDADES;
        map[key].otrasActividades += 1;
        map[key].operaciones += 1;
      });
    });
  return Object.values(map)
    .map((s) => ({
      nombre: s.nombre,
      puntos: s.puntos,
      operaciones: s.operaciones,
      promedio: s.operaciones ? s.puntos / s.operaciones : 0,
      conteo: s.conteo,
      revalidadas: s.revalidadas,
      otrasActividades: s.otrasActividades,
      segundosPrevios: s.segundosPrevios,
    }))
    .sort((a, b) => b.puntos - a.puntos);
}

/** KPIs por tramitador: renglones de previos/despachos/revalidadas/pendientes que dejó
 * cada quien, más cuántos reportes capturó y en cuántos días distintos tuvo actividad. */
function statsByTramitador(desde, hasta) {
  const map = {};
  allCapturas()
    .filter((c) => (!desde || c.fecha >= desde) && (!hasta || c.fecha <= hasta))
    .forEach((c) => {
      const raw = (c.tramitador || "").trim();
      const key = raw ? normalizarNombre(raw) : "sin_asignar";
      if (!map[key]) map[key] = { nombre: raw ? nombreCanonico(raw) : "Sin asignar", previos: 0, despachos: 0, revalidadas: 0, pendientes: 0, reportes: 0, dias: new Set() };
      map[key].previos += (c.previos || []).length;
      map[key].despachos += (c.despachos || []).length;
      map[key].revalidadas += (c.revalidadas || []).length;
      map[key].pendientes += (c.pendientes || []).length;
      map[key].reportes += 1;
      map[key].dias.add(c.fecha);
    });
  return Object.values(map)
    .map((s) => ({
      nombre: s.nombre,
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
      const raw = (c.tramitador || "").trim();
      const nombre = raw ? nombreCanonico(raw) : "Sin asignar";
      const key = c.fecha + "||" + (raw ? normalizarNombre(raw) : "sin_asignar");
      if (!map[key]) map[key] = { fecha: c.fecha, nombre, previos: 0, despachos: 0, revalidadas: 0, pendientes: 0 };
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
      const raw = (r.cliente || "").trim();
      if (!raw) return;
      const cliente = normalizarNombre(raw);
      if (!map[cliente]) map[cliente] = { nombre: nombreCanonicoCliente(raw), previos: 0, despachos: 0, revalidadas: 0, dias: new Set() };
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
  return Object.values(map)
    .map((s) => ({
      nombre: s.nombre,
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
      const raw = (r.cliente || "").trim();
      if (!raw) return;
      const cliente = normalizarNombre(raw);
      const k = c.fecha + "||" + cliente;
      if (!map[k]) map[k] = { fecha: c.fecha, nombre: nombreCanonicoCliente(raw), previos: 0, despachos: 0, revalidadas: 0 };
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
 * los reportes de la aduana activa, agrupados por Guía — para que el Ejecutivo los complete.
 * También guarda la fecha MÁS ANTIGUA en la que esa guía apareció sin referencia, para poder
 * mostrar cuánto tiempo lleva pendiente. */
function findFilasSinReferencia() {
  const groups = {};
  allCapturas().forEach((c) => {
    // Las guías revalidadas NO tienen campo de Referencia (solo Guía/Cliente/Almacén) —
    // se excluyen a propósito para que no aparezcan siempre como "pendientes de
    // referencia" sin serlo de verdad.
    [...(c.previos || []), ...(c.despachos || []), ...(c.pendientes || [])].forEach((r) => {
      const guia = (r.guia || "").trim();
      if (!guia || !isReferenciaVacia(r.ref)) return;
      if (!groups[guia])
        groups[guia] = { guia, cliente: r.cliente || "", almacen: r.almacen || "", ejecutivo: r.ejecutivo || "", count: 0, desde: c.fecha || "" };
      groups[guia].count += 1;
      if (!groups[guia].ejecutivo && r.ejecutivo) groups[guia].ejecutivo = r.ejecutivo; // por si el primer renglón encontrado no lo traía pero otro sí
      if (c.fecha && (!groups[guia].desde || c.fecha < groups[guia].desde)) groups[guia].desde = c.fecha;
    });
  });
  return Object.values(groups).sort((a, b) => a.guia.localeCompare(b.guia));
}

/** Días completos transcurridos entre dos fechas en formato YYYY-MM-DD (fin - inicio). */
function diasEntreFechas(fechaInicio, fechaFin) {
  if (!fechaInicio || !fechaFin) return 0;
  const d1 = new Date(fechaInicio + "T00:00:00");
  const d2 = new Date(fechaFin + "T00:00:00");
  return Math.max(0, Math.round((d2 - d1) / 86400000));
}

/** Igual que findFilasSinReferencia, pero agrupado un nivel arriba por Cliente,
 * para que el Ejecutivo pueda abrir solo el cliente que le interesa. */
function findFilasSinReferenciaPorCliente() {
  const guiaGroups = findFilasSinReferencia();
  const porCliente = {};
  guiaGroups.forEach((g) => {
    const raw = g.cliente.trim() || "Sin cliente";
    const key = raw === "Sin cliente" ? "Sin cliente" : normalizarNombre(raw);
    if (!porCliente[key]) porCliente[key] = { nombre: raw === "Sin cliente" ? raw : nombreCanonicoCliente(raw), guias: [] };
    porCliente[key].guias.push(g);
  });
  return Object.values(porCliente)
    .map(({ nombre, guias }) => ({
      cliente: nombre,
      guias: guias.slice().sort((a, b) => (a.desde || "").localeCompare(b.desde || "")),
      totalRenglones: guias.reduce((s, g) => s + g.count, 0),
    }))
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
    .filter((a) => a.estatus === "pendiente" && mismoNombre(a.tramitador, t) && !yaEnPendientes.has((a.guia || "").trim()));
}

/** Todas las asignaciones pendientes del tramitador actualmente logueado, sin importar
 * en qué pantalla esté ni si ya inició una captura — para el badge de la barra superior. */
function misAsignacionesPendientes() {
  if (state.userRole !== "tramite") return [];
  const t = (state.user || "").trim();
  return Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .filter((a) => a.estatus === "pendiente" && mismoNombre(a.tramitador, t))
    .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""));
}

/** La "bolsa" de guías que Coordinación asignó sin decidir todavía quién las va a
 * trabajar — en aduanas con autoasignación habilitada (ver permiteAutoasignacion),
 * cualquier tramitador puede entrar aquí y agarrar las que él mismo va a hacer. */
function guiasEnBolsa() {
  return Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .filter((a) => a.estatus === "pendiente" && !(a.tramitador || "").trim())
    .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""));
}

/** Versión "canónica" de un nombre de persona — sin mayúsculas/minúsculas ni espacios de
 * más — para poder agrupar o comparar aunque el nombre se haya escrito con variaciones
 * mínimas en distintos lugares (ej. "Luis Arreola" vs "luis arreola"). */
function normalizarNombre(s) {
  return (s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

/** Junta, en un solo Set, la combinación tramitador+tipo+guía de TODOS los renglones
 * capturados en cualquier reporte — para poder revisar rápido (sin recorrer todo de
 * nuevo por cada asignación) si una asignación "completada" todavía tiene rastro real
 * en algún reporte, o si ese renglón se borró/movió después de que el cierre automático
 * la marcó — algo que el cierre automático nunca revisa hacia atrás por su cuenta. */
function indiceGuiasEnReportes() {
  const set = new Set();
  allCapturas().forEach((c) => {
    const t = normalizarNombre(c.tramitador);
    if (!t) return;
    (c.previos || []).forEach((r) => {
      if ((r.guia || "").trim()) set.add(t + "||previo||" + normalizarNombre(r.guia));
    });
    (c.despachos || []).forEach((r) => {
      if ((r.guia || "").trim()) set.add(t + "||despacho||" + normalizarNombre(r.guia));
    });
    (c.revalidadas || []).forEach((r) => {
      if ((r.guia || "").trim()) set.add(t + "||revalidada||" + normalizarNombre(r.guia));
    });
  });
  return set;
}
function asignacionSinRastro(a, indice) {
  return !indice.has(normalizarNombre(a.tramitador) + "||" + a.tipo + "||" + normalizarNombre(a.guia));
}

/** Da el nombre "oficial" para mostrar en pantalla: si hay alguien en el catálogo actual
 * de tramitadores que coincide (ignorando mayúsculas/espacios), se usa esa forma exacta;
 * si no, se usa el nombre tal cual venía. Así, aunque los datos históricos tengan pequeñas
 * variaciones de captura, en pantalla siempre se ve un solo nombre consistente. */
function nombreCanonico(raw) {
  const t = (raw || "").trim();
  if (!t) return "Sin asignar";
  const enCatalogo = (editableCats.tramitadores || []).find((n) => mismoNombre(n, t));
  return enCatalogo || t;
}

/** Igual que nombreCanonico, pero comparando contra el catálogo de Clientes en vez del
 * de Tramitadores — para que "stanley" y "STANLEY" se vean como un solo cliente en las
 * estadísticas, aunque los datos históricos tengan la variación. */
function nombreCanonicoCliente(raw) {
  const t = (raw || "").trim();
  if (!t) return "";
  const enCatalogo = (editableCats.clientes || []).find((n) => mismoNombre(n, t));
  return enCatalogo || t;
}

/** "Título Case": primera letra de cada palabra en mayúscula, el resto en minúscula —
 * la forma que se usa como nombre "bonito" al homologar duplicados (ej. "claudia
 * barrera" y "CLAUDIA BARRERA" se homologan a "Claudia Barrera"). */
function tituloCase(s) {
  return (s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .split(" ")
    .map((p) => (p ? p[0].toUpperCase() + p.slice(1) : p))
    .join(" ");
}

/** Unifica de verdad los nombres de Ejecutivos que quedaron duplicados solo por
 * mayúsculas/espacios distintos (a diferencia de la alerta de "posibles duplicados",
 * que solo avisa) — deja UNA sola entrada por persona en el catálogo (en Título Case),
 * y reescribe cualquier renglón ya guardado (Previos, Despachos, Observaciones,
 * Asignaciones) que tuviera alguna de las variantes viejas, para que todo apunte al
 * mismo nombre de una vez. No toca a nadie que no tenga duplicados. */
async function homologarDuplicadosEjecutivos() {
  const grupos = {};
  (editableCats.ejecutivos || []).forEach((v) => {
    const key = normalizarNombre(v);
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(v);
  });
  const gruposConDuplicado = Object.values(grupos).filter((g) => g.length > 1);
  if (gruposConDuplicado.length === 0) return { fusionados: 0, renglonesActualizados: 0 };

  // Arma el catálogo nuevo: una sola entrada por persona (en Título Case si tenía
  // duplicados; se respeta tal cual si no los tenía), y el mapa de qué nombre viejo
  // apunta a cuál nombre nuevo.
  const mapaViejoACanonico = {};
  const nuevosEjecutivos = [];
  const yaAgregados = new Set();
  (editableCats.ejecutivos || []).forEach((v) => {
    const key = normalizarNombre(v);
    if (yaAgregados.has(key)) return;
    const grupo = grupos[key];
    const canonico = grupo.length > 1 ? tituloCase(grupo[0]) : v;
    mapaViejoACanonico[key] = canonico;
    nuevosEjecutivos.push(canonico);
    yaAgregados.add(key);
  });
  editableCats.ejecutivos = nuevosEjecutivos;
  await saveCatalogos();

  // Reescribe renglones guardados: "ejecutivo" en Previos/Despachos, y "observadoPor"
  // en Previos/Despachos/Pendientes (donde el Ejecutivo deja observaciones).
  let renglonesActualizados = 0;
  const diasParaGuardar = new Set();
  Object.values(state.reports || {}).forEach((day) => {
    (day.capturas || []).forEach((c) => {
      ["previos", "despachos"].forEach((grupo) => {
        (c[grupo] || []).forEach((r) => {
          const key = normalizarNombre(r.ejecutivo);
          if (key && mapaViejoACanonico[key] && r.ejecutivo !== mapaViejoACanonico[key]) {
            r.ejecutivo = mapaViejoACanonico[key];
            renglonesActualizados++;
            diasParaGuardar.add(day.fecha);
          }
        });
      });
      ["previos", "despachos", "pendientes"].forEach((grupo) => {
        (c[grupo] || []).forEach((r) => {
          const key = normalizarNombre(r.observadoPor);
          if (key && mapaViejoACanonico[key] && r.observadoPor !== mapaViejoACanonico[key]) {
            r.observadoPor = mapaViejoACanonico[key];
            renglonesActualizados++;
            diasParaGuardar.add(day.fecha);
          }
        });
      });
    });
  });
  for (const fecha of diasParaGuardar) {
    try {
      await saveReportDay(fecha);
    } catch (e) {}
  }

  // Reescribe asignaciones que tengan Ejecutivo puesto.
  for (const [id, a] of Object.entries(state.asignaciones || {})) {
    const key = normalizarNombre(a.ejecutivo);
    if (key && mapaViejoACanonico[key] && a.ejecutivo !== mapaViejoACanonico[key]) {
      a.ejecutivo = mapaViejoACanonico[key];
      renglonesActualizados++;
      try {
        await fbSaveAsignacion(state.aduanaActiva, id, a);
      } catch (e) {}
    }
  }

  return { fusionados: gruposConDuplicado.length, renglonesActualizados };
}

/** Junta, agrupadas por nombre normalizado, todas las variantes de un mismo cliente
 * que existan de verdad — tanto en el catálogo como en cualquier renglón ya capturado
 * (Previos/Despachos/Revalidadas/Pendientes) y en Asignaciones. A diferencia del
 * catálogo de Tramitadores/Ejecutivos, el campo Cliente es texto libre — un cliente
 * puede estar escrito con variaciones en los reportes sin nunca haberse agregado al
 * catálogo, así que solo revisar el catálogo se quedaría corto. */
function gruposDuplicadosClientes() {
  const grupos = {};
  const registrar = (v) => {
    const t = (v || "").trim();
    if (!t) return;
    const key = normalizarNombre(t);
    if (!grupos[key]) grupos[key] = [];
    if (!grupos[key].includes(t)) grupos[key].push(t);
  };
  (editableCats.clientes || []).forEach(registrar);
  Object.values(state.reports || {}).forEach((day) => {
    (day.capturas || []).forEach((c) => {
      ["previos", "despachos", "revalidadas", "pendientes"].forEach((grupo) => {
        (c[grupo] || []).forEach((r) => registrar(r.cliente));
      });
    });
  });
  Object.values(state.asignaciones || {}).forEach((a) => registrar(a.cliente));
  return Object.values(grupos).filter((g) => g.length > 1);
}

/** Unifica de verdad los nombres de Clientes que quedaron duplicados solo por
 * mayúsculas/espacios distintos (mismo criterio que ya existe para Ejecutivos) —
 * deja UNA sola entrada por cliente en el catálogo, y reescribe cualquier renglón ya
 * guardado (Previos, Despachos, Revalidadas, Pendientes) y Asignaciones que tuvieran
 * alguna de las variantes viejas, para que todo apunte al mismo nombre de una vez. Si
 * alguna variante ya coincidía con una entrada del catálogo, se usa esa forma como la
 * "oficial"; si no, se usa Título Case de la primera variante encontrada. */
async function homologarDuplicadosClientes() {
  const gruposConDuplicado = gruposDuplicadosClientes();
  if (gruposConDuplicado.length === 0) return { fusionados: 0, renglonesActualizados: 0 };

  const mapaViejoACanonico = {};
  gruposConDuplicado.forEach((variantes) => {
    const key = normalizarNombre(variantes[0]);
    const enCatalogo = (editableCats.clientes || []).find((n) => normalizarNombre(n) === key);
    mapaViejoACanonico[key] = enCatalogo || tituloCase(variantes[0]);
  });

  // Reescribe el catálogo: una sola entrada por cliente.
  const nuevosClientes = [];
  const yaAgregados = new Set();
  (editableCats.clientes || []).forEach((v) => {
    const key = normalizarNombre(v);
    if (yaAgregados.has(key)) return;
    nuevosClientes.push(mapaViejoACanonico[key] || v);
    yaAgregados.add(key);
  });
  // Por si alguna variante duplicada nunca había estado en el catálogo, se agrega su
  // forma canónica también — así el catálogo queda completo, no solo "limpio".
  gruposConDuplicado.forEach((variantes) => {
    const key = normalizarNombre(variantes[0]);
    if (!yaAgregados.has(key)) {
      nuevosClientes.push(mapaViejoACanonico[key]);
      yaAgregados.add(key);
    }
  });
  editableCats.clientes = nuevosClientes;
  await saveCatalogos();

  // Reescribe renglones guardados en cualquier reporte.
  let renglonesActualizados = 0;
  const diasParaGuardar = new Set();
  Object.values(state.reports || {}).forEach((day) => {
    (day.capturas || []).forEach((c) => {
      ["previos", "despachos", "revalidadas", "pendientes"].forEach((grupo) => {
        (c[grupo] || []).forEach((r) => {
          const key = normalizarNombre(r.cliente);
          if (key && mapaViejoACanonico[key] && r.cliente !== mapaViejoACanonico[key]) {
            r.cliente = mapaViejoACanonico[key];
            renglonesActualizados++;
            diasParaGuardar.add(day.fecha);
          }
        });
      });
    });
  });
  for (const fecha of diasParaGuardar) {
    try {
      await saveReportDay(fecha);
    } catch (e) {}
  }

  // Reescribe asignaciones que tengan Cliente puesto.
  for (const [id, a] of Object.entries(state.asignaciones || {})) {
    const key = normalizarNombre(a.cliente);
    if (key && mapaViejoACanonico[key] && a.cliente !== mapaViejoACanonico[key]) {
      a.cliente = mapaViejoACanonico[key];
      renglonesActualizados++;
      try {
        await fbSaveAsignacion(state.aduanaActiva, id, a);
      } catch (e) {}
    }
  }

  return { fusionados: gruposConDuplicado.length, renglonesActualizados };
}

/** Compara dos nombres de personas de forma tolerante — sin importar mayúsculas/minúsculas
 * ni espacios de más entre palabras. Evita que una asignación se "pierda" solo porque el
 * nombre quedó escrito con una letra en mayúscula distinta o un espacio extra en algún lado
 * (ej. al crearla desde Asignaciones vs. el nombre exacto con el que inició sesión). */
function mismoNombre(a, b) {
  const na = normalizarNombre(a);
  return na !== "" && na === normalizarNombre(b);
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
    else if (state.view === "selectCoordinadorLogin") root.innerHTML = viewSelectCoordinadorLogin();
    else if (state.view === "coordinadorPassword") root.innerHTML = viewCoordinadorPassword();
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
            : state.view === "pendientesGlobal"
            ? viewPendientesGlobal()
            : state.view === "guiasDisponibles"
            ? viewGuiasDisponibles()
            : state.view === "respaldos"
            ? viewRespaldos()
            : state.view === "solicitudesEliminacion"
            ? viewSolicitudesEliminacion()
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
  const hoyOperativo = todayStr();
  const sinTramitadorCount = Object.values(state.asignaciones || {}).filter(
    (a) => a.estatus === "pendiente" && !(a.tramitador || "").trim() && diaOperativoDe(a.fechaCreacion) !== hoyOperativo
  ).length;
  const pendientesGlobalCount = todosPendientesAbiertos().length + sinTramitadorCount;
  const misAsigCount = misAsignacionesPendientes().length;
  const solicitudesCount = state.userRole === "admin" ? todasLasSolicitudesEliminacion().length : 0;
  return `
  <div class="topbar">
    <div class="brand">
      <div class="brand-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="brand-text"><div class="t1">Oñate Reporte</div><div class="t2">Reporte operativo diario · v${APP_VERSION}</div></div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
      ${
        !state.isOnline
          ? `<span class="pill" style="background:#FBE9E7;color:#C0453B;">🔴 Sin conexión</span>`
          : state.syncingPending
          ? `<span class="pill" style="background:#DCEBF9;color:var(--accent-dark);">🔄 Sincronizando…</span>`
          : pendientesAtoradosCount() > 0
          ? `<span class="pill" style="background:#FBE9E7;color:#C0453B;cursor:pointer;" onclick="App.reintentarPendientesAhora()" title="Llevan más de 10 minutos sin poder subirse a pesar de tener señal — toca para reintentar a mano">⚠️ ${pendientesAtoradosCount()} sin subir hace rato — reintentar</span>`
          : pendCount > 0
          ? `<span class="pill" style="background:#FBF3D8;color:#8A6414;" title="Se sincronizará al recuperar señal">⏳ ${pendCount} pendiente${pendCount > 1 ? "s" : ""}</span>`
          : ""
      }
      <div class="user-chip"><span class="user-dot"></span>${esc(state.user)}${state.aduanaActiva ? ` · ${esc(state.aduanaActiva)}` : ""}</div>
      <button class="nav-btn" onclick="App.toggleNotifPanel()" title="Notificaciones recientes">🔔${state.recentNotifications.length > 0 ? ` ${state.recentNotifications.length}` : ""}</button>
      ${
        state.userRole === "tramite" && misAsigCount > 0
          ? `<button class="nav-btn" style="background:#FBEDEA;color:var(--rojo);" onclick="App.toggleMisAsignacionesPanel()" title="Tus asignaciones pendientes">📥 Asignaciones (${misAsigCount})</button>`
          : ""
      }
      ${
        state.userRole === "tramite" && permiteAutoasignacion()
          ? `<button class="nav-btn" onclick="App.goGuiasDisponibles()" title="Guías que Coordinación asignó sin decidir quién las trabaja — agarra las tuyas">🎯 Guías disponibles${
              guiasEnBolsa().length > 0 ? ` (${guiasEnBolsa().length})` : ""
            }</button>`
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
      ${state.userRole === "admin" ? `<button class="nav-btn" onclick="App.goPendientesGlobal()">Pendientes${pendientesGlobalCount > 0 ? ` (${pendientesGlobalCount})` : ""}</button>` : ""}
      ${state.userRole === "admin" ? `<button class="nav-btn" onclick="App.goRespaldos()">🗄️ Respaldos</button>` : ""}
      ${
        solicitudesCount > 0
          ? `<button class="nav-btn" style="background:#FBEDD8;color:var(--ambar);" onclick="App.goSolicitudesEliminacion()">🗑️ Solicitudes (${solicitudesCount})</button>`
          : ""
      }
      ${state.userRole !== "ejecutivo" ? `<button class="nav-btn" onclick="App.goBuscar()">🔍 Buscar</button>` : ""}
      ${state.userRole !== "ejecutivo" ? `<button class="nav-btn" onclick="App.goCatalogos()">Catálogos</button>` : ""}
      <button class="nav-btn" onclick="App.logout()">Salir</button>
    </div>
  </div>`;
}

function viewLogin() {
  const roles = ["Administrador"];
  if ((editableCats.coordinadores || []).length > 0) roles.push("Coordinador");
  roles.push("Trámite", "Ejecutivo");
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">Oñate Reporte</div>
      <div class="login-sub">Digitalización del reporte operativo diario de previos y despachos</div>
      ${state.aduanaActiva ? `<div class="pill pill-navy" style="margin-bottom:16px;">Aduana: ${esc(state.aduanaActiva)}</div>` : ""}
      <div class="user-pick">
        ${roles
          .map(
            (u) =>
              `<button class="user-pick-btn" onclick="App.chooseRole('${esc(u)}')"><span class="user-avatar">${esc(
                u.slice(0, 2).toUpperCase()
              )}</span> ${esc(u)}</button>`
          )
          .join("")}
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

function viewSelectCoordinadorLogin() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">¿Quién eres?</div>
      <div class="login-sub">Elige tu nombre — así quedará marcado en cada cosa que hagas como Coordinador</div>
      <div class="user-pick">
        ${(editableCats.coordinadores || []).map(
          (n) =>
            `<button class="user-pick-btn" onclick="App.chooseCoordinadorLogin('${esc(n)}')"><span class="user-avatar">${esc(
              n.split(" ").map((p) => p[0]).join("").slice(0, 2).toUpperCase()
            )}</span> ${esc(n)}</button>`
        ).join("")}
      </div>
      <button class="btn btn-ghost btn-sm" style="margin-top:16px;" onclick="App.backToRoleSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewCoordinadorPassword() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">Hola, ${esc(state.pendingCoordinador || "")}</div>
      <div class="login-sub">Ingresa tu contraseña personal para continuar</div>
      <div class="field" style="text-align:left;margin-bottom:16px;">
        <input type="password" id="coordinador_password" placeholder="Contraseña" autofocus
          onkeydown="if(event.key==='Enter') App.submitCoordinadorPassword()"
          style="width:100%;padding:11px 12px;border:1.3px solid var(--line);border-radius:8px;font-size:15px;"/>
      </div>
      ${state.errorMsg ? `<div class="status-line status-error" style="margin-bottom:16px;">${esc(state.errorMsg)}</div>` : ""}
      <button class="btn btn-primary" style="width:100%;margin-bottom:10px;" onclick="App.submitCoordinadorPassword()">Entrar</button>
      <button class="btn btn-ghost btn-sm" onclick="App.backToCoordinadorSelect()">← Volver</button>
    </div>
  </div>`;
}

function viewEjecutivoName() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark"><img src="logo-ow.jpg" alt="OW" /></div>
      <div class="login-title display">¿Quién eres?</div>
      <div class="login-sub">Elige tu nombre — así quedará marcado en cada observación y referencia que dejes</div>
      ${
        (editableCats.ejecutivos || []).length === 0
          ? `<div class="status-line status-error" style="margin-bottom:16px;">Todavía no hay ejecutivos dados de alta en esta aduana. Pídele a Coordinación que los agregue en Catálogos.</div>`
          : `<div class="user-pick">
        ${(editableCats.ejecutivos || []).map(
          (n) =>
            `<button class="user-pick-btn" onclick="App.submitEjecutivoName('${esc(n)}')"><span class="user-avatar">${esc(
              n.split(" ").map((p) => p[0]).join("").slice(0, 2).toUpperCase()
            )}</span> ${esc(n)}</button>`
        ).join("")}
      </div>`
      }
      ${state.errorMsg ? `<div class="status-line status-error" style="margin-top:16px;">${esc(state.errorMsg)}</div>` : ""}
      <button class="btn btn-ghost btn-sm" style="margin-top:16px;" onclick="App.backToRoleSelect()">← Volver</button>
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

    <div class="chart-row-2col">
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
    </div>

    <div class="panel" style="margin-bottom:20px;">
      <div class="section-title">Top clientes</div>
      <div class="section-sub" style="margin-bottom:10px;">Por número de operaciones — <a href="#" onclick="event.preventDefault();App.goClientes();" style="color:var(--accent);">ver con filtro de fechas →</a></div>
      <div class="dashboard-grid">
        <div>
          <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin-bottom:2px;">Previos</div>
          ${miniRankBlock(topClPrevios, "previos", maxClPrevios)}
        </div>
        <div>
          <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin-bottom:2px;">Despachos</div>
          ${miniRankBlock(topClDespachos, "despachos", maxClDespachos)}
        </div>
        ${
          topClRevalidadas.length > 0
            ? `
        <div>
          <div style="font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.02em;margin-bottom:2px;">Revalidadas</div>
          ${miniRankBlock(topClRevalidadas, "revalidadas", maxClRevalidadas)}
        </div>
        `
            : ""
        }
      </div>
    </div>

    ${
      state.userRole === "admin"
        ? (() => {
            const statsSemana = statsByTramitador(dateStrLocal(6), todayStr()).slice(0, 6);
            const maxSemana = Math.max(1, ...statsSemana.map((s) => s.total));
            const statsPonderada = statsProductividadPonderada(dateStrLocal(6), todayStr())
              .filter((s) => s.operaciones > 0)
              .slice(0, 6);
            const maxPonderada = Math.max(1, ...statsPonderada.map((s) => s.puntos));
            const atrasadosPorTramitador = (() => {
              const hoy = todayStr();
              const day = state.reports[hoy];
              const map = {};
              if (day && day.capturas) {
                day.capturas.forEach((c) => {
                  const raw = (c.tramitador || "").trim();
                  if (!raw) return;
                  const key = normalizarNombre(raw);
                  const n = (c.pendientes || []).filter((p) => p.origenFecha && !p.bloqueado).length;
                  if (n > 0) map[key] = { nombre: nombreCanonico(raw), n: (map[key] ? map[key].n : 0) + n };
                });
              }
              return Object.values(map)
                .map((x) => [x.nombre, x.n])
                .sort((a, b) => b[1] - a[1]);
            })();
            return `
    <div class="dashboard-grid" style="margin-bottom:20px;">
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Productividad — últimos 7 días</div>
        <div class="section-sub" style="margin-bottom:10px;">Total de operaciones por tramitador — <a href="#" onclick="event.preventDefault();App.goKPIs();" style="color:var(--accent);">ver KPIs completos →</a></div>
        ${
          statsSemana.length === 0
            ? `<div style="color:var(--muted);font-size:12px;">Sin datos todavía</div>`
            : statsSemana
                .map(
                  (s) => `
          <div class="rank-row">
            <div class="rank-name" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
            <div class="rank-bar-bg"><div class="rank-bar" style="width:${(s.total / maxSemana) * 100}%"></div></div>
            <div class="rank-val">${s.total}</div>
          </div>`
                )
                .join("")
        }
      </div>
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Pendientes atrasados por tramitador</div>
        <div class="section-sub" style="margin-bottom:10px;">Pendientes de hoy que vienen arrastrados de días anteriores sin resolver</div>
        ${
          atrasadosPorTramitador.length === 0
            ? `<div style="color:var(--muted);font-size:12px;">Nadie tiene pendientes atrasados 🎉</div>`
            : atrasadosPorTramitador
                .map(
                  ([nombre, n]) => `
          <div class="rank-row">
            <div class="rank-name" title="${esc(nombre)}">${esc(nombre)}</div>
            <div class="rank-bar-bg"><div class="rank-bar" style="width:${(n / atrasadosPorTramitador[0][1]) * 100}%;background:var(--ambar);"></div></div>
            <div class="rank-val">${n}</div>
          </div>`
                )
                .join("")
        }
      </div>
      <div class="panel">
        <div class="section-title" style="font-size:15px;">Productividad ponderada por dificultad</div>
        <div class="section-sub" style="margin-bottom:10px;">Últimos 7 días — Previos (A 0.06 a D 0.18), Despachos (A 0.02 a D 0.11), Revalidadas (0.02) y Otras actividades (0.02). Un "2do previo" resta -0.06. <a href="#" onclick="event.preventDefault();App.goKPIs();" style="color:var(--accent);">Ver detalle diario y filtrar fechas →</a></div>
        ${
          statsPonderada.length === 0
            ? `<div style="color:var(--muted);font-size:12px;">Sin datos todavía</div>`
            : statsPonderada
                .map(
                  (s) => `
          <div class="rank-row">
            <div class="rank-name" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
            <div class="rank-bar-bg"><div class="rank-bar" style="width:${Math.max(0, (s.puntos / maxPonderada) * 100)}%;background:${
                    s.puntos < 0 ? "var(--rojo)" : "var(--navy)"
                  };"></div></div>
            <div class="rank-val" title="${s.operaciones} operación(es), promedio ${s.promedio.toFixed(3)} pts">${s.puntos.toFixed(2)}</div>
          </div>
          <div style="font-size:10.5px;color:var(--muted);margin:-4px 0 8px 0;padding-left:2px;">A:${s.conteo.A} · B:${s.conteo.B} · C:${s.conteo.C} · D:${
                    s.conteo.D
                  } · Revalidadas:${s.revalidadas} · Otras act.:${s.otrasActividades}${s.segundosPrevios > 0 ? ` · <span style="color:var(--rojo);">2dos previos:${s.segundosPrevios}</span>` : ""}</div>`
                )
                .join("")
        }
      </div>
    </div>
    `;
          })()
        : ""
    }

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
    ${state.processingMsg ? `<div class="status-line" style="margin-bottom:16px;max-width:340px;"><div class="spinner"></div> ${esc(state.processingMsg)}</div>` : ""}
    ${errorBanner()}

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
          <thead><tr>${REVALIDADA_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th>Foto</th><th></th></tr></thead>
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
      ${(() => {
        // Detecta lotes (renglones que llegaron de la misma carga masiva de asignaciones) —
        // si hay 2 o más pendientes de DESPACHO del mismo lote sin resolver, se ofrece un
        // botón para marcarlos "Listo" a todos juntos de un jalón. Solo aplica a Despacho
        // (normalmente una sola persona con todo el consolidado) — Previo se deja siempre
        // individual, porque ahí sí pueden repartirse entre varios tramitadores distintos.
        // Cualquier guía se puede sacar del lote a mano con su ✕ normal antes de usar el botón.
        const lotes = {};
        (c.pendientes || []).forEach((row, i) => {
          if (row.bloqueado || !row._loteId) return;
          if (row.tipo !== "despacho") return;
          if (!lotes[row._loteId]) lotes[row._loteId] = [];
          lotes[row._loteId].push(i);
        });
        const gruposConDosOMas = Object.entries(lotes).filter(([, idxs]) => idxs.length > 1);
        if (gruposConDosOMas.length === 0) return "";
        return `<div style="margin-bottom:10px;display:flex;flex-direction:column;gap:6px;">
          ${gruposConDosOMas
            .map(([loteId, idxs]) => {
              const cliente = (c.pendientes[idxs[0]].cliente || "").trim();
              return `<button class="btn btn-primary btn-sm" style="align-self:flex-start;" onclick="App.marcarLoteListo('${loteId}')">✅ Marcar TODO este lote como Listo (${idxs.length} guías${
                cliente ? ` — ${esc(cliente)}` : ""
              })</button>`;
            })
            .join("")}
        </div>`;
      })()}
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-ghost btn-sm" onclick="App.addRow('pendientes')">+ Agregar pendiente</button>
        ${
          asignacionesPendientesParaFormulario(c.tramitador).length > 0
            ? `<button class="btn btn-primary btn-sm" onclick="App.traerAsignaciones()">📥 Traer asignaciones pendientes (${
                asignacionesPendientesParaFormulario(c.tramitador).length
              })</button>`
            : ""
        }
        ${
          c.fecha === todayStr() && pendientesSinTrasladarCount(c.tramitador) > 0
            ? `<button class="btn btn-primary btn-sm" onclick="App.traerPendientesAnteriores()">📥 Traer pendientes de días anteriores (${pendientesSinTrasladarCount(
                c.tramitador
              )})</button>`
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
  // Un pendiente "bloqueado" es el respaldo histórico de un día anterior que ya se
  // trasladó al día siguiente — queda fijo, de solo lectura, para que nunca se pierda
  // el rastro de que ese pendiente existió, aunque alguien borre por error la copia viva.
  if (group === "pendientes" && row.bloqueado) {
    return `<tr style="opacity:.65;background:#F4F6F5;">
      ${fields.map((f) => `<td>${esc(row[f.k])}</td>`).join("")}
      <td style="white-space:nowrap;font-size:11px;color:var(--muted);">
        🔒 Trasladado al día siguiente
        ${
          row._solicitudEliminacion
            ? state.userRole === "admin"
              ? ` <span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:9px;">🗑️ ${esc(row._solicitadoPor)}</span>
                  <button onclick="App.deleteRow('${group}',${i})" title="Autorizar y borrar" style="border:none;background:none;color:#2E7D4F;cursor:pointer;">✅</button>
                  <button onclick="App.rechazarEliminacionRenglon('${group}',${i})" title="Rechazar" style="border:none;background:none;color:var(--rojo);cursor:pointer;">❌</button>`
              : ` <span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:9px;">🗑️ Esperando</span>
                  <button onclick="App.deleteRow('${group}',${i})" title="Cancelar tu solicitud" style="border:none;background:none;color:var(--muted);cursor:pointer;">Cancelar</button>`
            : `<button class="row-del" onclick="App.deleteRow('${group}',${i})" title="${
                state.userRole === "admin" ? "Borrar este respaldo histórico (con confirmación)" : "Solicitar eliminación (la debe autorizar Admin)"
              }">✕</button>`
        }
      </td>
    </tr>`;
  }
  return `<tr>
    ${
      group === "pendientes" && row.origenFecha
        ? `<td colspan="${fields.length + 1}" style="padding:2px 6px 0;border:none;font-size:10.5px;color:var(--ambar);font-weight:600;">⚠️ Pendiente desde ${fmtDateHuman(row.origenFecha)} — no se resolvió ese día</td></tr><tr>`
        : ""
    }
    ${fields
      .map((f) => {
        if (f.k === "almacen" && almacenFijo()) {
          return `<td style="padding:6px 6px;font-size:12px;font-family:'IBM Plex Mono',monospace;color:var(--ink);white-space:nowrap;">${esc(
            almacenFijo()
          )} <span title="Fijo para esta aduana" style="font-size:10px;">🔒</span></td>`;
        }
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
    ${
      group === "revalidadas"
        ? `<td style="white-space:nowrap;text-align:center;">
        ${
          row.foto
            ? `<img src="${row.foto}" onclick="App.verFotoRevalidada('${group}',${i})" title="Ver foto" style="width:34px;height:34px;object-fit:cover;border-radius:6px;cursor:pointer;border:1px solid var(--line);vertical-align:middle;${
                row._solicitudQuitarFoto ? "opacity:.5;" : ""
              }"/>
               ${
                 row._solicitudQuitarFoto
                   ? state.userRole === "admin"
                     ? `<br/><span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:9px;">🗑️ ${esc(row._solicitadoPorFoto)}</span>
                        <button onclick="App.quitarFotoRevalidada('${group}',${i})" title="Autorizar" style="border:none;background:none;color:#2E7D4F;cursor:pointer;font-size:13px;vertical-align:middle;">✅</button>
                        <button onclick="App.rechazarQuitarFotoRevalidada('${group}',${i})" title="Rechazar" style="border:none;background:none;color:var(--rojo);cursor:pointer;font-size:13px;vertical-align:middle;">❌</button>`
                     : `<br/><span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:9px;">🗑️ Esperando</span>
                        <button onclick="App.quitarFotoRevalidada('${group}',${i})" title="Cancelar solicitud" style="border:none;background:none;color:var(--muted);cursor:pointer;font-size:11px;vertical-align:middle;">Cancelar</button>`
                   : `<button onclick="App.quitarFotoRevalidada('${group}',${i})" title="Quitar foto" style="border:none;background:none;color:var(--rojo);cursor:pointer;font-size:13px;vertical-align:middle;margin-left:2px;">✕</button>`
               }`
            : `<label title="Tomar o subir foto de la guía" style="display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border:1px dashed var(--line);border-radius:6px;cursor:pointer;font-size:15px;">
                 📷
                 <input type="file" accept="image/*" capture="environment" style="display:none" onchange="App.capturarFotoRevalidada('${group}',${i},this)"/>
               </label>`
        }
      </td>`
        : ""
    }
    <td style="white-space:nowrap;">
      ${
        group === "pendientes" && (row.tipo === "previo" || row.tipo === "despacho" || row.tipo === "revalidada")
          ? `<button class="btn btn-primary btn-sm" style="margin-right:4px;" onclick="App.marcarPendienteListo(${i})">✅ Listo</button>`
          : ""
      }
      ${
        row._solicitudEliminacion
          ? state.userRole === "admin"
            ? `<span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:10px;" title="Solicitado por ${esc(row._solicitadoPor)}">🗑️ ${esc(
                row._solicitadoPor
              )}</span>
             <button class="btn btn-primary btn-sm" style="margin-left:4px;" onclick="App.deleteRow('${group}',${i})" title="Autorizar y borrar">✅ Autorizar</button>
             <button class="btn btn-ghost btn-sm" onclick="App.rechazarEliminacionRenglon('${group}',${i})" title="Rechazar la solicitud">❌</button>`
            : `<span class="pill" style="background:#FBEDD8;color:var(--ambar);font-size:10px;">🗑️ Esperando autorización</span>
             <button class="btn btn-ghost btn-sm" onclick="App.deleteRow('${group}',${i})" title="Cancelar tu solicitud">Cancelar</button>`
          : `<button class="row-del" onclick="App.deleteRow('${group}',${i})" title="${state.userRole === "admin" ? "Eliminar renglón" : "Solicitar eliminación (la debe autorizar Admin)"}">✕</button>`
      }
    </td>
  </tr>`;
}

function datalists() {
  const merged = { ...editableCats, resultados: ["Desaduanamiento libre", "Reconocimiento aduanero"] };
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

/** Dibuja la celda de "Observaciones" para Previos, Despachos o Pendientes: editable
 * (con el nombre del Ejecutivo que la escribió, guardado automáticamente) solo si el
 * usuario actual es Ejecutivo — para cualquier otro rol se ve como texto fijo, con el
 * mismo nombre de autor visible debajo para que quede claro quién la dejó. */
function celdaObservaciones(r, grupo, fecha, capturaId, ri) {
  const autor = r.observadoPor ? `<div style="font-size:10px;color:var(--muted);margin-top:2px;">— ${esc(r.observadoPor)}</div>` : "";
  if (state.userRole === "ejecutivo") {
    return `<td><input value="${esc(r.observaciones)}" placeholder="Agregar observación…" style="width:100%;min-width:140px;padding:5px 6px;border:1px solid var(--line);border-radius:6px;font-size:12.5px;" onblur="App.updateObservacion('${grupo}','${fecha}','${esc(
      capturaId
    )}','${esc((r.guia || "").trim())}',${ri},this.value)"/>${autor}</td>`;
  }
  return `<td>${esc(r.observaciones)}${autor}</td>`;
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
            ${state.userRole !== "ejecutivo" ? `<button class="btn btn-ghost btn-sm" onclick="App.editCaptura('${d.fecha}','${esc(c.id)}')">Editar</button>` : ""}
            ${state.userRole === "admin" ? `<button class="row-del" onclick="App.deleteCaptura('${d.fecha}','${esc(c.id)}')" title="Eliminar hoja">✕</button>` : ""}
          </div>
        </div>
        <div class="mini-title">1. Control de previos (${(c.previos || []).length})</div>
        <div style="overflow-x:auto;">
        <table class="mini-table">
          <thead><tr>${PREVIO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
          <tbody>${
            (c.previos || [])
              .map(
                (r, ri) =>
                  `<tr>${PREVIO_FIELDS.map((f) =>
                    f.k === "observaciones" ? celdaObservaciones(r, "previos", d.fecha, c.id, ri) : `<td>${esc(r[f.k])}</td>`
                  ).join("")}</tr>`
              )
              .join("") ||
            `<tr><td colspan="${PREVIO_FIELDS.length}" style="color:var(--muted);font-family:'IBM Plex Sans';">Sin renglones</td></tr>`
          }</tbody>
        </table>
        </div>
        <div class="mini-title">2. Control de despachos (${(c.despachos || []).length})</div>
        <div style="overflow-x:auto;">
        <table class="mini-table">
          <thead><tr>${DESPACHO_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}</tr></thead>
          <tbody>${
            (c.despachos || [])
              .map(
                (r, ri) =>
                  `<tr>${DESPACHO_FIELDS.map((f) =>
                    f.k === "observaciones" ? celdaObservaciones(r, "despachos", d.fecha, c.id, ri) : `<td>${esc(r[f.k])}</td>`
                  ).join("")}</tr>`
              )
              .join("") ||
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
            <thead><tr>${REVALIDADA_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th>Foto</th></tr></thead>
            <tbody>${(c.revalidadas || [])
              .map(
                (r, ri) =>
                  `<tr>${REVALIDADA_FIELDS.map((f) => `<td>${esc(r[f.k])}</td>`).join("")}<td>${
                    r.foto
                      ? `<img src="${r.foto}" onclick="App.verFotoRevalidadaDetalle('${d.fecha}','${esc(c.id)}','${esc((r.guia || "").trim())}')" title="Ver foto" style="width:30px;height:30px;object-fit:cover;border-radius:5px;cursor:pointer;border:1px solid var(--line);"/>`
                      : "—"
                  }</td></tr>`
              )
              .join("")}</tbody>
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
            <thead><tr>${PENDIENTE_FIELDS.map((f) => `<th>${f.label}</th>`).join("")}<th>Estado</th></tr></thead>
            <tbody>${(c.pendientes || [])
              .map(
                (r, ri) =>
                  `<tr${r.bloqueado ? ` style="opacity:.6;"` : ""}>${PENDIENTE_FIELDS.map((f) =>
                    f.k === "observaciones" ? celdaObservaciones(r, "pendientes", d.fecha, c.id, ri) : `<td>${esc(r[f.k])}</td>`
                  ).join("")}<td style="font-size:10px;">${
                    r.bloqueado
                      ? "🔒 Trasladado al día siguiente"
                      : r.origenFecha
                      ? `⚠️ Pendiente desde ${fmtDateHuman(r.origenFecha)}`
                      : ""
                  }</td></tr>`
              )
              .join("")}</tbody>
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
  const categorias = ["clientes", "almacenes"];
  if (state.userRole === "admin") categorias.splice(2, 0, "tramitadores", "ejecutivos", "coordinadores");
  const pwMap = tramitadorPasswordsMap(editableCats.tramitadores, state.aduanaActiva);
  const pwMapCoordinadores = tramitadorPasswordsMap(editableCats.coordinadores, state.aduanaActiva);
  const UMBRAL_BUSCADOR = 25; // a partir de cuántos registros aparece el buscador en vez de mostrar todo de un jalón
  const LIMITE_VISIBLE = 60; // tope de chips dibujados a la vez, aunque el filtro deje más resultados
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Catálogos</div>
    <div class="section-sub">Nombres sugeridos automáticamente al capturar (clientes, almacenes, tramitadores, ejecutivos)</div>
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
        const posiblesDuplicados =
          k === "clientes"
            ? gruposDuplicadosClientes()
            : (() => {
                const vistos = {};
                const grupos = [];
                editableCats[k].forEach((v) => {
                  const key = normalizarNombre(v);
                  if (!vistos[key]) vistos[key] = [];
                  vistos[key].push(v);
                });
                Object.values(vistos).forEach((grupo) => {
                  if (grupo.length > 1) grupos.push(grupo);
                });
                return grupos;
              })();
        return `
      <div class="panel" style="margin-bottom:14px;">
        <div style="font-weight:600;font-size:13.5px;margin-bottom:10px;text-transform:capitalize;">${k} ${
          total > 0 ? `<span style="color:var(--muted);font-weight:400;">(${total})</span>` : ""
        }</div>
        ${
          posiblesDuplicados.length > 0
            ? `<div class="status-line status-warn" style="margin-bottom:10px;font-size:12px;">⚠️ Posibles duplicados (mismo nombre, escrito distinto${
                k === "clientes" ? " — se revisó en reportes y asignaciones, no solo el catálogo" : ""
              }): ${posiblesDuplicados.map((g) => g.map((n) => `"${esc(n)}"`).join(" / ")).join(" — ")}. ${
                k === "ejecutivos" || k === "clientes"
                  ? `Toca el botón de abajo para unificarlos solos — deja un solo nombre y corrige también los reportes${
                      k === "ejecutivos" ? "/observaciones" : ""
                    } y asignaciones ya guardadas que tuvieran cualquiera de las variantes.`
                  : `Revisa cuál dejar y borra el resto con el ✕ de abajo — sus registros históricos no se mueven solos al hacerlo.`
              }</div>
              ${
                k === "ejecutivos"
                  ? `<button class="btn btn-primary btn-sm" style="margin-bottom:10px;" onclick="App.homologarEjecutivos()">🔀 Unificar duplicados de Ejecutivos</button>`
                  : k === "clientes"
                  ? `<button class="btn btn-primary btn-sm" style="margin-bottom:10px;" onclick="App.homologarClientes()">🔀 Unificar duplicados de Clientes</button>`
                  : ""
              }`
            : ""
        }
        ${state.homologarEjecutivosMsg && k === "ejecutivos" ? `<div class="status-line" style="margin-bottom:10px;font-size:12px;">${esc(state.homologarEjecutivosMsg)}</div>` : ""}
        ${state.homologarClientesMsg && k === "clientes" ? `<div class="status-line" style="margin-bottom:10px;font-size:12px;">${esc(state.homologarClientesMsg)}</div>` : ""}
        ${
          k === "tramitadores"
            ? `<div style="color:var(--muted);font-size:11.5px;margin-bottom:8px;">Contraseña automática: ${
                state.aduanaActiva === "GDL" ? "3" : "4"
              } primeras letras del nombre + 2026 (más letras si se repite con otra persona). Compártela con cada quien.</div>`
            : k === "ejecutivos"
            ? `<div style="color:var(--muted);font-size:11.5px;margin-bottom:8px;">Sin contraseña — solo eligen su nombre de esta lista al entrar como Ejecutivo.</div>`
            : k === "coordinadores"
            ? `<div style="color:var(--muted);font-size:11.5px;margin-bottom:8px;">Coordinador es un rol aparte de Administrador — cada quien entra con su nombre y su PROPIA contraseña (mismo criterio que Trámite: iniciales del nombre + 2026), y tiene los mismos permisos que Administrador dentro de la app. Si esta lista queda vacía, el botón "Coordinador" no aparece en la pantalla de entrada.</div>`
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
                      k === "tramitadores"
                        ? ` <span style="font-family:'IBM Plex Mono',monospace;color:var(--muted);font-size:11px;">(${esc(pwMap[v] || "")})</span>`
                        : k === "coordinadores"
                        ? ` <span style="font-family:'IBM Plex Mono',monospace;color:var(--muted);font-size:11px;">(${esc(pwMapCoordinadores[v] || "")})</span>`
                        : ""
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
  const ponderada = statsProductividadPonderada(desde, hasta).filter((s) => s.operaciones > 0);
  const maxPonderada = Math.max(1, ...ponderada.map((s) => s.puntos));
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
        <tfoot>
          <tr style="font-weight:700;border-top:2px solid var(--navy);">
            <td>TOTAL</td>
            <td>${stats.reduce((s, r) => s + r.previos, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.despachos, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.revalidadas, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.pendientes, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.total, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.reportes, 0)}</td>
            <td>—</td>
          </tr>
        </tfoot>
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
    <div class="panel" style="margin-bottom:16px;overflow-x:auto;">
      <div class="section-title" style="font-size:15px;margin-bottom:4px;">Productividad ponderada por dificultad</div>
      <div class="section-sub" style="margin-bottom:12px;">Previos (A 0.06 a D 0.18) y Despachos (A 0.02 a D 0.11) según dificultad — Revalidadas y Otras actividades suman 0.02 fijo cada una. Un "2do previo" resta -0.06 en vez de sumar.</div>
      ${
        ponderada.length === 0
          ? `<div style="color:var(--muted);font-size:12.5px;">Sin datos en este rango</div>`
          : `
      <table class="mini-table" style="min-width:760px;">
        <thead><tr><th>Tramitador</th><th>A</th><th>B</th><th>C</th><th>D</th><th>Revalidadas</th><th>Otras act.</th><th>2dos previos</th><th>Puntos totales</th><th>Operaciones</th><th>Promedio</th></tr></thead>
        <tbody>
        ${ponderada
          .map(
            (s) => `<tr>
          <td style="font-family:'IBM Plex Sans',sans-serif;font-weight:600;">${esc(s.nombre)}</td>
          <td>${s.conteo.A}</td>
          <td>${s.conteo.B}</td>
          <td>${s.conteo.C}</td>
          <td>${s.conteo.D}</td>
          <td>${s.revalidadas}</td>
          <td>${s.otrasActividades}</td>
          <td style="${s.segundosPrevios > 0 ? "color:var(--rojo);font-weight:700;" : ""}">${s.segundosPrevios}</td>
          <td style="font-weight:700;${s.puntos < 0 ? "color:var(--rojo);" : ""}">${s.puntos.toFixed(2)}</td>
          <td>${s.operaciones}</td>
          <td>${s.promedio.toFixed(3)}</td>
        </tr>`
          )
          .join("")}
        </tbody>
      </table>
      <div style="margin-top:12px;">
        ${ponderada
          .map(
            (s) => `
        <div class="rank-row">
          <div class="rank-name" style="width:140px;" title="${esc(s.nombre)}">${esc(s.nombre)}</div>
          <div class="rank-bar-bg"><div class="rank-bar" style="width:${Math.max(0, (s.puntos / maxPonderada) * 100)}%;background:${
              s.puntos < 0 ? "var(--rojo)" : "var(--navy)"
            };"></div></div>
          <div class="rank-val" style="width:50px;">${s.puntos.toFixed(2)}</div>
        </div>
      `
          )
          .join("")}
      </div>
      `
      }
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
        <tfoot>
          <tr style="font-weight:700;border-top:2px solid var(--navy);">
            <td colspan="2">TOTAL</td>
            <td>${detalle.reduce((s, r) => s + r.previos, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.despachos, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.revalidadas, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.pendientes, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.total, 0)}</td>
          </tr>
        </tfoot>
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
        <tfoot>
          <tr style="font-weight:700;border-top:2px solid var(--navy);">
            <td>TOTAL</td>
            <td>${stats.reduce((s, r) => s + r.previos, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.despachos, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.revalidadas, 0)}</td>
            <td>${stats.reduce((s, r) => s + r.total, 0)}</td>
            <td>—</td>
          </tr>
        </tfoot>
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
        <tfoot>
          <tr style="font-weight:700;border-top:2px solid var(--navy);">
            <td colspan="2">TOTAL (de estos ${detalle.length} renglones)</td>
            <td>${detalle.reduce((s, r) => s + r.previos, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.despachos, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.revalidadas, 0)}</td>
            <td>${detalle.reduce((s, r) => s + r.total, 0)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
    `
    }
  `;
}

/** Pantalla exclusiva de Admin/Coordinación: TODOS los pendientes abiertos de todos los
 * tramitadores y todos los días, con opción de reasignar cualquiera a otra persona —
 * venga o no de una asignación formal (a diferencia de "Asignaciones", que solo cubre lo
 * que Coordinación asignó a propósito). */
/** Pantalla para Trámite en aduanas con autoasignación habilitada (ver
 * permiteAutoasignacion): la "bolsa" de guías que Coordinación asignó masivamente sin
 * decidir quién las va a trabajar. Cada tramitador agarra las suyas — SIEMPRE se
 * asignan a sí mismo (el nombre con el que entró a la app), nunca a otra persona. */
function viewGuiasDisponibles() {
  const bolsa = guiasEnBolsa();
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Guías disponibles</div>
    <div class="section-sub">Coordinación las asignó sin decidir todavía quién las va a trabajar — toca "Agarrar esta guía" en las que tú vayas a hacer. Se van a asignar a tu nombre (${esc(
      state.user
    )}).</div>
    ${errorBanner()}
    ${
      bolsa.length === 0
        ? `<div class="empty"><div class="stamp-outline">🎉</div><div style="font-weight:600;">No hay guías disponibles en la bolsa ahorita</div></div>`
        : bolsa
            .map(
              (a) => `
      <div class="panel" style="margin-bottom:12px;">
        <div class="meta-line">
          <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
          &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"} &nbsp;·&nbsp; <b>Almacén:</b> ${esc(a.almacen) || "—"}
          ${a.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(a.ref)}` : ""}
          ${a.pedimento ? ` &nbsp;·&nbsp; <b>Pedimento:</b> ${esc(a.pedimento)}` : ""}
          ${a.sector ? ` &nbsp;·&nbsp; <b>Sector:</b> ${esc(a.sector)}` : ""}
        </div>
        <button class="btn btn-primary btn-sm" style="margin-top:10px;" onclick="App.asignarmeGuia('${a.id}')">🎯 Agarrar esta guía</button>
      </div>`
            )
            .join("")
    }
  `;
}

/** Pantalla exclusiva de Admin: lista de respaldos automáticos guardados (uno por
 * día, hasta los últimos 14), con opción de descargar cualquiera como archivo JSON —
 * el punto de recuperación manual en caso de que algo salga mal con los datos "en vivo". */
/** Junta, de todos los reportes guardados, cualquier renglón marcado con una solicitud
 * de eliminación (de renglón completo o solo de foto) pendiente de que Admin la revise —
 * para verlas todas juntas en un solo lugar, sin tener que abrir reporte por reporte. */
function todasLasSolicitudesEliminacion() {
  const list = [];
  Object.values(state.reports || {}).forEach((day) => {
    (day.capturas || []).forEach((c) => {
      ["previos", "despachos", "revalidadas", "pendientes"].forEach((grupo) => {
        (c[grupo] || []).forEach((row) => {
          if (row._solicitudEliminacion) {
            list.push({ fecha: day.fecha, capturaId: c.id, grupo, row, tipo: "renglon", tramitador: c.tramitador });
          }
          if (row._solicitudQuitarFoto) {
            list.push({ fecha: day.fecha, capturaId: c.id, grupo, row, tipo: "foto", tramitador: c.tramitador });
          }
        });
      });
    });
  });
  return list.sort((a, b) =>
    (a.row._fechaSolicitud || a.row._fechaSolicitudFoto || "").localeCompare(b.row._fechaSolicitud || b.row._fechaSolicitudFoto || "")
  );
}

/** Pantalla exclusiva de Admin: todas las solicitudes de eliminación pendientes de
 * cualquier reporte, con botones para autorizar (borra de verdad) o rechazar (se queda
 * tal cual) cada una desde aquí mismo, sin necesidad de abrir cada reporte por separado. */
function viewSolicitudesEliminacion() {
  const lista = todasLasSolicitudesEliminacion();
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Solicitudes de eliminación</div>
    <div class="section-sub">Trámite y Ejecutivo no pueden borrar directo — piden autorización, y aparece aquí hasta que Admin la resuelva.</div>
    ${errorBanner()}
    ${
      lista.length === 0
        ? `<div class="empty"><div class="stamp-outline">🎉</div><div style="font-weight:600;">No hay solicitudes pendientes</div></div>`
        : lista
            .map((item) => {
              const solicitadoPor = item.tipo === "foto" ? item.row._solicitadoPorFoto : item.row._solicitadoPor;
              const fechaSolicitud = item.tipo === "foto" ? item.row._fechaSolicitudFoto : item.row._fechaSolicitud;
              return `
      <div class="panel" style="margin-bottom:12px;">
        <div class="meta-line">
          <b>Reporte del:</b> ${fmtDateHuman(item.fecha)} &nbsp;·&nbsp; <b>Tramitador:</b> ${esc(item.tramitador) || "—"} &nbsp;·&nbsp;
          <b>Sección:</b> ${esc(item.grupo)} &nbsp;·&nbsp; <b>Solicitud:</b> ${item.tipo === "foto" ? "Quitar foto" : "Borrar renglón completo"}
          <br/>
          ${item.row.guia ? `<b>Guía:</b> ${esc(item.row.guia)}` : ""} ${item.row.cliente ? ` &nbsp;·&nbsp; <b>Cliente:</b> ${esc(item.row.cliente)}` : ""}
          <br/>
          <span style="color:var(--ambar);">Solicitado por ${esc(solicitadoPor) || "—"} el ${fechaSolicitud ? new Date(fechaSolicitud).toLocaleString("es-MX") : "—"}</span>
        </div>
        <div style="display:flex;gap:8px;margin-top:10px;">
          <button class="btn btn-primary btn-sm" onclick="App.autorizarEliminacionDesdeGlobal('${item.fecha}','${esc(item.capturaId)}','${item.grupo}','${item.tipo}','${esc(
                fechaSolicitud || ""
              )}')">✅ Autorizar</button>
          <button class="btn btn-ghost btn-sm" onclick="App.rechazarEliminacionDesdeGlobal('${item.fecha}','${esc(item.capturaId)}','${item.grupo}','${item.tipo}','${esc(
                fechaSolicitud || ""
              )}')">❌ Rechazar</button>
        </div>
      </div>`;
            })
            .join("")
    }
  `;
}

function viewRespaldos() {
  const lista = state.respaldosList;
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Respaldos automáticos</div>
    <div class="section-sub">Copia completa de reportes, catálogos y asignaciones de esta aduana — se guarda sola, como máximo una vez al día. Se conservan los últimos 14; los más viejos se van borrando solos.</div>
    ${errorBanner()}
    ${state.processingMsg ? `<div class="status-line" style="margin-bottom:16px;"><div class="spinner"></div> ${esc(state.processingMsg)}</div>` : ""}
    ${
      lista === null
        ? `<div class="status-line"><div class="spinner"></div> Cargando…</div>`
        : lista.length === 0
        ? `<div class="empty"><div class="stamp-outline">🗄️</div><div style="font-weight:600;">Todavía no hay respaldos — se crea el primero automáticamente la próxima vez que entres como Admin.</div></div>`
        : lista
            .slice()
            .reverse()
            .map((clave) => {
              const fechaLegible = clave.split("_")[0];
              return `
      <div class="panel" style="margin-bottom:10px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
        <div style="font-weight:600;">${fmtDateHuman(fechaLegible)}</div>
        <button class="btn btn-ghost btn-sm" onclick="App.descargarRespaldoClave('${clave}')">⬇️ Descargar</button>
      </div>`;
            })
            .join("")
    }
  `;
}

function viewPendientesGlobal() {
  const abiertos = todosPendientesAbiertos();
  const hoy = todayStr();
  const sinTramitador = Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .filter((a) => a.estatus === "pendiente" && !(a.tramitador || "").trim() && diaOperativoDe(a.fechaCreacion) !== hoy)
    .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""));
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Pendientes de todos los tramitadores</div>
    <div class="section-sub">Todo lo que sigue pendiente sin resolver, sin importar el día ni de dónde salió — reasigna aquí si alguien más lo va a terminar</div>
    ${errorBanner()}
    ${
      sinTramitador.length > 0
        ? `
    <div class="panel" style="margin-bottom:20px;border-color:#EAD199;background:#FFFBF2;">
      <div class="subhead" style="margin-top:0;">⚠️ Asignaciones sin tramitador (${sinTramitador.length})</div>
      <div class="section-sub" style="margin-bottom:12px;">Cruzaron el cierre de operaciones (8:00 PM) sin que se les pusiera tramitador — elige a alguien para cada una, y aparecerá sola en el reporte de esa persona en cuanto entre a capturar.</div>
      ${sinTramitador
        .map((a) => {
          const dias = a.fechaCreacion ? Math.floor((Date.now() - new Date(a.fechaCreacion).getTime()) / 86400000) : 0;
          return `
      <div class="captura-item" style="background:#FBF3E3;border-color:#EAD199;margin-bottom:10px;">
        <div class="meta-line">
          <span class="pill" style="background:#FBEDD8;color:var(--ambar);">⚠️ Sin asignar desde hace ${dias} día${dias === 1 ? "" : "s"}</span>
          <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
          &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"}
          ${a.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(a.ref)}` : ""}
          ${a.pedimento ? ` &nbsp;·&nbsp; <b>Pedimento:</b> ${esc(a.pedimento)}` : ""}
          ${a.sector ? ` &nbsp;·&nbsp; <b>Sector:</b> ${esc(a.sector)}` : ""}
        </div>
        <div style="display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap;">
          <select id="asignar_rapido_${a.id}" style="flex:1;min-width:160px;">
            <option value="">— Elige un tramitador —</option>
            ${(editableCats.tramitadores || []).map((t) => `<option value="${esc(t)}">${esc(t)}</option>`).join("")}
          </select>
          <button class="btn btn-primary btn-sm" onclick="App.asignarTramitadorRapido('${a.id}',document.getElementById('asignar_rapido_${a.id}').value)">Asignar</button>
          <button class="btn btn-ghost btn-sm" onclick="App.editarAsignacion('${a.id}')" title="Corregir cualquier dato (guía, cliente, referencia, etc.)">✏️ Editar</button>
          <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar asignación">✕</button>
        </div>
      </div>`;
        })
        .join("")}
    </div>
    `
        : ""
    }
    <div class="subhead" style="margin-top:0;">Pendientes en reportes (${abiertos.length})</div>
    ${
      abiertos.length === 0
        ? `<div class="empty"><div class="stamp-outline">🎉</div><div style="font-weight:600;">No hay pendientes abiertos</div></div>`
        : abiertos
            .map((item) => {
              const p = item.row;
              const key = `${item.fecha}_${item.capturaId}_${(p.guia || "").trim()}`.replace(/[^a-zA-Z0-9_]/g, "_");
              return `
      <div class="panel" style="margin-bottom:12px;">
        <div class="meta-line">
          <b>Fecha:</b> ${fmtDateHuman(item.fecha)} ${item.fecha === hoy ? "(hoy)" : ""} &nbsp;·&nbsp;
          <b>Tramitador actual:</b> ${esc(item.tramitador) || "—"} &nbsp;·&nbsp;
          <b>Guía:</b> ${esc(p.guia) || "—"} &nbsp;·&nbsp;
          <b>Cliente:</b> ${esc(p.cliente) || "—"}
          ${p.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(p.ref)}` : ""}
          ${p.tipo ? ` &nbsp;·&nbsp; <b>Tipo:</b> ${esc(p.tipo)}` : ""}
          <br/>
          ${
            p.origenFecha
              ? `<span style="color:var(--ambar);font-weight:600;">⚠️ Arrastrado desde ${fmtDateHuman(p.origenFecha)}</span>`
              : `<span style="color:var(--muted);">Sin arrastrar de días anteriores</span>`
          }
        </div>
        <div style="display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap;">
          <select id="reasignar_${key}" style="flex:1;min-width:160px;">
            ${(editableCats.tramitadores || [])
              .map((t) => `<option value="${esc(t)}" ${mismoNombre(t, item.tramitador) ? "selected" : ""}>${esc(t)}</option>`)
              .join("")}
          </select>
          <button class="btn btn-primary btn-sm" onclick="App.reasignarPendienteGlobal('${item.fecha}','${esc(item.capturaId)}','${esc(
            (p.guia || "").trim()
          )}',document.getElementById('reasignar_${key}').value)">↪ Reasignar a esta persona</button>
        </div>
      </div>`;
            })
            .join("")
    }
  `;
}

function viewAsignaciones() {
  const hoy = todayStr();
  const list = Object.entries(state.asignaciones || {})
    .map(([id, a]) => ({ id, ...a }))
    .sort((a, b) => (b.fechaCreacion || "").localeCompare(a.fechaCreacion || ""));
  const pendientesTodas = list.filter((a) => a.estatus === "pendiente");
  // Sin tramitador Y todavía dentro del día operativo de hoy (antes de las 8pm) — las que
  // ya cruzaron el cierre de operaciones sin asignarse se van solas a la pantalla de
  // Pendientes, no se quedan aquí.
  const sinTramitador = pendientesTodas.filter((a) => !(a.tramitador || "").trim() && diaOperativoDe(a.fechaCreacion) === hoy);
  const pendientes = pendientesTodas.filter((a) => (a.tramitador || "").trim());
  const completadas = list.filter((a) => a.estatus === "completada");
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Asignaciones</div>
    <div class="section-sub">Asigna previos, despachos o revalidaciones de guía a cada tramitador por número de Guía. Se cierran solas en cuanto ese tramitador guarda un reporte con esa misma Guía.</div>
    ${errorBanner()}
    ${state.processingMsg ? `<div class="status-line" style="margin-bottom:16px;"><div class="spinner"></div> ${esc(state.processingMsg)}</div>` : ""}

    <div class="panel" style="margin-bottom:20px;">
      <div class="field-grid">
        <div class="field"><label>Tipo</label>
          <select id="asig_tipo" onchange="App.updateAsigDraft('tipo',this.value)">
            <option value="previo" ${state.asigDraft.tipo === "previo" ? "selected" : ""}>Previo</option>
            <option value="despacho" ${state.asigDraft.tipo === "despacho" ? "selected" : ""}>Despacho</option>
            <option value="revalidada" ${state.asigDraft.tipo === "revalidada" ? "selected" : ""}>Guía revalidada</option>
          </select>
        </div>
        <div class="field"><label>Referencia</label><input id="asig_ref" placeholder="Número de referencia" value="${esc(state.asigDraft.ref)}" oninput="App.updateAsigDraft('ref',this.value)"/></div>
        <div class="field"><label>Guía</label><input id="asig_guia" placeholder="Número de guía" value="${esc(state.asigDraft.guia)}" oninput="App.updateAsigDraft('guia',this.value)"/></div>
        <div class="field"><label>Cliente</label><input id="asig_cliente" list="dl_clientes" value="${esc(state.asigDraft.cliente)}" oninput="App.updateAsigDraft('cliente',this.value)"/></div>
        <div class="field"><label>Almacén</label>${
          almacenFijo()
            ? `<div style="padding:9px 10px;border:1.3px solid var(--line);border-radius:7px;color:var(--muted);background:#F4F6F5;font-size:13.5px;font-family:'IBM Plex Sans',sans-serif;">${esc(
                almacenFijo()
              )} 🔒</div>`
            : `<input id="asig_almacen" list="dl_almacenes" value="${esc(state.asigDraft.almacen)}" oninput="App.updateAsigDraft('almacen',this.value)"/>`
        }</div>
        <div class="field"><label>Sector</label>
          <select id="asig_sector" onchange="App.updateAsigDraft('sector',this.value)">
            ${["", "QUIMICO", "PERECEDERO", "METALURGICO", "TEXTIL", "AGRICULTURA", "MANUFACTURA", "SALUD", "DIGITAL"]
              .map((s) => `<option value="${s}" ${state.asigDraft.sector === s ? "selected" : ""}>${s || "—"}</option>`)
              .join("")}
          </select>
        </div>
        <div class="field"><label>Ejecutivo</label>
          <select id="asig_ejecutivo" onchange="App.updateAsigDraft('ejecutivo',this.value)">
            <option value="" ${state.asigDraft.ejecutivo === "" ? "selected" : ""}>— Sin asignar —</option>
            ${(editableCats.ejecutivos || []).map((e) => `<option value="${esc(e)}" ${state.asigDraft.ejecutivo === e ? "selected" : ""}>${esc(e)}</option>`).join("")}
          </select>
        </div>
        <div class="field"><label>Pedimento (para Despacho)</label><input id="asig_pedimento" placeholder="Número de pedimento" value="${esc(state.asigDraft.pedimento)}" oninput="App.updateAsigDraft('pedimento',this.value)"/></div>
        <div class="field"><label>Tramitador</label>
          <select id="asig_tramitador" onchange="App.updateAsigDraft('tramitador',this.value)">
            <option value="" ${state.asigDraft.tramitador === "" ? "selected" : ""}>— Sin asignar aún —</option>
            ${(editableCats.tramitadores || []).map((t) => `<option value="${esc(t)}" ${state.asigDraft.tramitador === t ? "selected" : ""}>${esc(t)}</option>`).join("")}
          </select>
        </div>
      </div>
      <button class="btn btn-primary btn-sm" style="margin-top:10px;" onclick="App.crearAsignacion()">${
        state.asigEditId ? "💾 Guardar cambios" : "+ Crear asignación"
      }</button>
      ${state.asigEditId ? `<button class="btn btn-ghost btn-sm" style="margin-top:10px;margin-left:8px;" onclick="App.cancelarEdicionAsignacion()">Cancelar edición</button>` : ""}
      ${datalists()}
    </div>

    <div class="panel" style="margin-bottom:20px;">
      <button class="btn btn-ghost btn-sm" onclick="App.toggleAsigMasiva()">${
        state.asigMasivaOpen ? "▲ Ocultar carga masiva" : "📋 Carga masiva de asignaciones (por manifiesto)"
      }</button>
      ${
        state.asigMasivaOpen
          ? `
      <div style="margin-top:14px;">
        <div class="section-sub" style="margin-bottom:10px;">
          Ve a tu manifiesto o Excel, selecciona la columna completa de <b>Guías</b> y pégala en el primer cuadro. Luego selecciona la columna completa de <b>Clientes</b> y pégala en el segundo, y así con Referencia y Ejecutivo.
          Deben quedar en el mismo orden — el renglón 1 de cada cuadro se junta con el renglón 1 de los demás. Cliente, Referencia y Ejecutivo son opcionales (puedes dejar esos cuadros vacíos). El mismo Tipo, Sector y Tramitador que elijas abajo se aplica a todas.
        </div>
        <div class="field-grid" style="margin-bottom:10px;">
          <div class="field"><label>Tipo (para todas)</label>
            <select id="asig_masiva_tipo" onchange="App.updateAsigMasivaDraft('tipo',this.value)">
              <option value="previo" ${state.asigMasivaDraft.tipo === "previo" ? "selected" : ""}>Previo</option>
              <option value="despacho" ${state.asigMasivaDraft.tipo === "despacho" ? "selected" : ""}>Despacho</option>
              <option value="revalidada" ${state.asigMasivaDraft.tipo === "revalidada" ? "selected" : ""}>Guía revalidada</option>
            </select>
          </div>
          <div class="field"><label>Sector (para todas, opcional)</label>
            <select id="asig_masiva_sector" onchange="App.updateAsigMasivaDraft('sector',this.value)">
              ${["", "QUIMICO", "PERECEDERO", "METALURGICO", "TEXTIL", "AGRICULTURA", "MANUFACTURA", "SALUD", "DIGITAL"]
                .map((s) => `<option value="${s}" ${state.asigMasivaDraft.sector === s ? "selected" : ""}>${s || "—"}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>Tramitador (para todas, opcional)</label>
            <select id="asig_masiva_tramitador" onchange="App.updateAsigMasivaDraft('tramitador',this.value)">
              <option value="" ${state.asigMasivaDraft.tramitador === "" ? "selected" : ""}>— Sin asignar aún —</option>
              ${(editableCats.tramitadores || [])
                .map((t) => `<option value="${esc(t)}" ${state.asigMasivaDraft.tramitador === t ? "selected" : ""}>${esc(t)}</option>`)
                .join("")}
            </select>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1.6fr;gap:10px;margin-bottom:10px;">
          <div class="field" style="margin-bottom:0;">
            <label>Guías <span style="color:var(--rojo);">*</span></label>
            <textarea id="asig_masiva_guias" rows="10" placeholder="875782762327&#10;861234567890&#10;745609590220" oninput="App.updateAsigMasivaTexto('textoGuias',this.value)" style="width:100%;font-family:'IBM Plex Mono',monospace;font-size:12.5px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;">${esc(
              state.asigMasivaDraft.textoGuias
            )}</textarea>
          </div>
          <div class="field" style="margin-bottom:0;">
            <label>Clientes (opcional — nombre completo)</label>
            <textarea id="asig_masiva_clientes" rows="10" placeholder="IMPRO INDUSTRIAL DE MEXICO&#10;BRP MEXICANA MANUFACTURING&#10;CPQ TECHNOLOGIES" oninput="App.updateAsigMasivaTexto('textoClientes',this.value)" style="width:100%;font-family:'IBM Plex Mono',monospace;font-size:12.5px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;">${esc(
              state.asigMasivaDraft.textoClientes
            )}</textarea>
          </div>
        </div>
        <div class="field-grid" style="grid-template-columns:repeat(3,1fr);">
          <div class="field">
            <label>Referencia (opcional)</label>
            <textarea id="asig_masiva_refs" rows="6" placeholder="GLSI262480" oninput="App.updateAsigMasivaTexto('textoRefs',this.value)" style="width:100%;font-family:'IBM Plex Mono',monospace;font-size:12.5px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;">${esc(
              state.asigMasivaDraft.textoRefs
            )}</textarea>
          </div>
          <div class="field">
            <label>Pedimento (opcional — Despacho)</label>
            <textarea id="asig_masiva_pedimentos" rows="6" placeholder="24 44 1234 5678901" oninput="App.updateAsigMasivaTexto('textoPedimentos',this.value)" style="width:100%;font-family:'IBM Plex Mono',monospace;font-size:12.5px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;">${esc(
              state.asigMasivaDraft.textoPedimentos
            )}</textarea>
          </div>
          <div class="field">
            <label>Ejecutivo (opcional)</label>
            <textarea id="asig_masiva_ejecutivos" rows="6" placeholder="Alberto Pichardo" oninput="App.updateAsigMasivaTexto('textoEjecutivos',this.value)" style="width:100%;font-family:'IBM Plex Mono',monospace;font-size:12.5px;padding:8px;border:1px solid var(--line);border-radius:var(--radius);resize:vertical;">${esc(
              state.asigMasivaDraft.textoEjecutivos
            )}</textarea>
          </div>
        </div>
        <button class="btn btn-primary btn-sm" style="margin-top:8px;" onclick="App.crearAsignacionesMasivas()">Crear todas las asignaciones</button>
        ${state.asigMasivaMsg ? `<div class="status-line" style="margin-top:8px;">${esc(state.asigMasivaMsg)}</div>` : ""}
      </div>
      `
          : ""
      }
    </div>
    ${
      sinTramitador.length > 0
        ? `
    <div class="panel" style="margin-bottom:20px;border-color:#EAD199;background:#FFFBF2;">
      <div class="subhead" style="margin-top:0;">⚠️ Sin tramitador asignado — hoy (${sinTramitador.length})</div>
      <div class="section-sub" style="margin-bottom:12px;">Se crearon hoy sin saber todavía quién las va a trabajar. Si no se les pone tramitador antes del cierre de operaciones (8:00 PM), se mueven solas a la pantalla de Pendientes.</div>
      ${sinTramitador
        .map(
          (a) => `
      <div class="captura-item" style="background:#fff;margin-bottom:10px;">
        <div class="meta-line">
          <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
          &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"}
          ${a.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(a.ref)}` : ""}
          ${a.pedimento ? ` &nbsp;·&nbsp; <b>Pedimento:</b> ${esc(a.pedimento)}` : ""}
          ${a.sector ? ` &nbsp;·&nbsp; <b>Sector:</b> ${esc(a.sector)}` : ""}
        </div>
        <div style="display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap;">
          <select id="asignar_rapido_${a.id}" style="flex:1;min-width:160px;">
            <option value="">— Elige un tramitador —</option>
            ${(editableCats.tramitadores || []).map((t) => `<option value="${esc(t)}">${esc(t)}</option>`).join("")}
          </select>
          <button class="btn btn-primary btn-sm" onclick="App.asignarTramitadorRapido('${a.id}',document.getElementById('asignar_rapido_${a.id}').value)">Asignar</button>
          <button class="btn btn-ghost btn-sm" onclick="App.editarAsignacion('${a.id}')" title="Corregir cualquier dato (guía, cliente, referencia, etc.)">✏️ Editar</button>
          <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar asignación">✕</button>
        </div>
      </div>`
        )
        .join("")}
    </div>
    `
        : ""
    }
    <div class="subhead">Pendientes (${pendientes.length})</div>
    ${
      pendientes.length === 0
        ? `<div style="color:var(--muted);font-size:13px;margin-bottom:16px;">Sin asignaciones pendientes.</div>`
        : [...pendientes]
            .sort((a, b) => (a.fechaCreacion || "").localeCompare(b.fechaCreacion || ""))
            .map((a) => {
              const dias = a.fechaCreacion ? Math.floor((Date.now() - new Date(a.fechaCreacion).getTime()) / 86400000) : 0;
              const esDeHoy = (a.fechaCreacion || "").slice(0, 10) === todayStr();
              // Rojo = se creó hoy. Amarillo = se quedó pendiente de un día anterior (alguien no
              // la resolvió) — así se distingue de un vistazo lo nuevo de lo que sigue arrastrándose.
              return `
      <div class="captura-item" style="background:${esDeHoy ? "#FEF6F5" : "#FBF3E3"};border-color:${esDeHoy ? "#F3C9C2" : "#EAD199"};">
        <div class="captura-head">
          <div class="meta-line">
            ${
              esDeHoy
                ? `<span class="pill" style="background:#F8D9D4;color:var(--rojo);">⏳ Pendiente hoy</span>`
                : `<span class="pill" style="background:#FBEDD8;color:var(--ambar);">⚠️ Pendiente desde hace ${dias} día${dias === 1 ? "" : "s"}</span>`
            }
            <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
            &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"} &nbsp;·&nbsp; <b>Almacén:</b> ${esc(a.almacen) || "—"}${
              a.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(a.ref)}` : ""
            }${a.pedimento ? ` &nbsp;·&nbsp; <b>Pedimento:</b> ${esc(a.pedimento)}` : ""}${a.sector ? ` &nbsp;·&nbsp; <b>Sector:</b> ${esc(a.sector)}` : ""}${
              a.ejecutivo ? ` &nbsp;·&nbsp; <b>Ejecutivo:</b> ${esc(a.ejecutivo)}` : ""
            }<br/>
            <b>Asignado a:</b> ${esc(a.tramitador)} &nbsp;·&nbsp; Creado por ${esc(a.creadoPor)} el ${a.fechaCreacion ? new Date(a.fechaCreacion).toLocaleDateString("es-MX") : "—"}
          </div>
          <div style="display:flex;gap:4px;flex-shrink:0;">
            <button class="btn btn-ghost btn-sm" onclick="App.editarAsignacion('${a.id}')" title="Corregir cualquier dato: guía, cliente, referencia, tramitador, etc.">✏️ Editar</button>
            <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar asignación">✕</button>
          </div>
        </div>
      </div>`;
            })
            .join("")
    }

    <div class="subhead">Completadas (${completadas.length})</div>
    ${
      completadas.length === 0
        ? `<div style="color:var(--muted);font-size:13px;">Sin asignaciones completadas todavía.</div>`
        : (() => {
            const indice = indiceGuiasEnReportes();
            const conRastroInfo = completadas.map((a) => ({ a, sinRastro: asignacionSinRastro(a, indice) }));
            const totalSinRastro = conRastroInfo.filter((x) => x.sinRastro).length;
            const mostrar = state.soloSinRastro ? conRastroInfo.filter((x) => x.sinRastro) : conRastroInfo;
            return `
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:14px;">
        ${
          totalSinRastro > 0
            ? `<button class="btn btn-ghost btn-sm" onclick="App.toggleSoloSinRastro()">${
                state.soloSinRastro ? "Ver todas las completadas" : `⚠️ Ver solo las ${totalSinRastro} sin rastro`
              }</button>
             <button class="btn btn-primary btn-sm" onclick="App.regresarTodasSinRastroAPendiente()">↩️ Regresar TODAS (${totalSinRastro}) a pendiente</button>`
            : ""
        }
      </div>
      ${mostrar
        .map(({ a, sinRastro }) => {
          return `
      <div class="captura-item" style="background:${sinRastro ? "#FFF8ED" : "#F1F7F5"};border-color:${sinRastro ? "#EAD199" : "#BFE0CC"};">
        <div class="captura-head">
          <div class="meta-line">
            <span class="pill" style="background:#DCF3E3;color:#2E7D4F;">✅ completada</span>
            <span class="pill ${pillClaseTipo(a.tipo)}">${esc(a.tipo)}</span>
            &nbsp;<b>Guía:</b> ${esc(a.guia)} &nbsp;·&nbsp; <b>Cliente:</b> ${esc(a.cliente) || "—"}${
            a.ref ? ` &nbsp;·&nbsp; <b>Ref.:</b> ${esc(a.ref)}` : ""
          }${a.pedimento ? ` &nbsp;·&nbsp; <b>Pedimento:</b> ${esc(a.pedimento)}` : ""}${a.sector ? ` &nbsp;·&nbsp; <b>Sector:</b> ${esc(a.sector)}` : ""}${
            a.ejecutivo ? ` &nbsp;·&nbsp; <b>Ejecutivo:</b> ${esc(a.ejecutivo)}` : ""
          }<br/>
            <b>Completado por:</b> ${esc(a.tramitador)} el ${a.fechaCompletado ? new Date(a.fechaCompletado).toLocaleString("es-MX") : "—"}
            ${
              sinRastro
                ? `<br/><span style="color:var(--ambar);font-weight:600;">⚠️ No se encuentra esta guía en ningún reporte de ${esc(
                    a.tramitador
                  )} ahorita — puede que se haya borrado o corregido después de completarse.</span>`
                : ""
            }
          </div>
          <div style="display:flex;gap:6px;align-items:center;">
            ${sinRastro ? `<button class="btn btn-ghost btn-sm" onclick="App.regresarAPendiente('${a.id}')" title="Regresarla a pendiente para reasignarla o corregirla">↩️ Regresar a pendiente</button>` : ""}
            <button class="row-del" onclick="App.eliminarAsignacion('${a.id}')" title="Eliminar del historial">✕</button>
          </div>
        </div>
      </div>`;
        })
        .join("")}`;
          })()
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
                .map((g) => {
                  const dias = diasEntreFechas(g.desde, todayStr());
                  const colorDias = dias >= 5 ? "var(--rojo)" : dias >= 2 ? "var(--ambar)" : "var(--muted)";
                  return `
                <div style="border:1px solid var(--line);border-radius:8px;padding:12px;margin-bottom:10px;background:#FEF6F5;">
                  <div class="meta-line" style="margin-bottom:10px;">
                    <b>Guía:</b> ${esc(g.guia)} &nbsp;·&nbsp; <b>Almacén:</b> ${esc(g.almacen) || "—"} &nbsp;·&nbsp; ${g.count} renglón(es) sin referencia
                    ${
                      g.ejecutivo
                        ? `<br/><span style="background:#DCEBF9;color:var(--accent-dark);font-weight:700;padding:2px 8px;border-radius:5px;font-size:12px;display:inline-block;margin-top:4px;">👤 Le toca a: ${esc(
                            g.ejecutivo
                          )}</span>`
                        : `<br/><span style="color:var(--muted);font-size:12px;">Sin ejecutivo asignado</span>`
                    }
                    <br/>
                    <span style="color:${colorDias};font-weight:600;">🕒 Pendiente desde ${fmtDateHuman(g.desde)} — ${dias} día${dias === 1 ? "" : "s"}</span>
                  </div>
                  <div style="display:flex;gap:8px;">
                    <input id="ref_${esc(g.guia)}" placeholder="Número de referencia correcto" style="flex:1;padding:9px 10px;border:1.3px solid var(--line);border-radius:7px;font-size:13.5px;"/>
                    <button class="btn btn-primary btn-sm" onclick="App.completarReferencia('${esc(g.guia)}')">Actualizar</button>
                  </div>
                </div>`;
                })
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
    if (role === "Administrador") {
      state.pendingAdminLabel = role;
      state.view = "adminPassword";
      state.errorMsg = "";
      render();
    } else if (role === "Coordinador") {
      state.view = "selectCoordinadorLogin";
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
    if (val === (ADMIN_PASSWORDS[state.aduanaActiva] || ADMIN_PASSWORDS.GDL)) {
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
  /** Coordinador es un rol aparte de Administrador (contraseña compartida): cada
   * coordinador elige su nombre y entra con su PROPIA contraseña individual (mismo
   * criterio que Trámite — iniciales del nombre + 2026). Adentro de la app tiene
   * exactamente los mismos permisos que Administrador (a propósito — así se decidió),
   * incluyendo poder borrar directo sin pedir autorización — nada más queda su propio
   * nombre marcado en vez de un genérico "Administrador" en todo lo que haga. */
  chooseCoordinadorLogin(nombre) {
    state.pendingCoordinador = nombre;
    state.view = "coordinadorPassword";
    state.errorMsg = "";
    render();
  },
  submitCoordinadorPassword() {
    const input = document.getElementById("coordinador_password");
    const val = input ? input.value : "";
    const correcta = tramitadorPasswordsMap(editableCats.coordinadores, state.aduanaActiva)[state.pendingCoordinador] || "";
    if (val === correcta) {
      state.errorMsg = "";
      App.enterApp(state.pendingCoordinador, "admin");
    } else {
      state.errorMsg = "Contraseña incorrecta.";
      render();
      const el = document.getElementById("coordinador_password");
      if (el) {
        el.value = "";
        el.focus();
      }
    }
  },
  backToCoordinadorSelect() {
    state.view = "selectCoordinadorLogin";
    state.errorMsg = "";
    render();
  },
  submitEjecutivoName(nombre) {
    if (!nombre) return;
    state.errorMsg = "";
    App.enterApp(nombre, "ejecutivo");
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
  /** Botón manual del aviso "⚠️ sin subir hace rato" — fuerza un intento inmediato,
   * sin esperar al siguiente ciclo de 25 segundos ni a que el navegador dispare su
   * propio evento de "regresó la señal" (que a veces no se dispara bien en celulares). */
  async reintentarPendientesAhora() {
    if (!navigator.onLine) {
      state.errorMsg = "Todavía no hay señal — espera a tener internet para reintentar.";
      render();
      return;
    }
    await processPendingQueue();
    try {
      state.reports = await fbLoadReports(state.aduanaActiva);
      applyPendingOverlay();
      persistLocalCache();
    } catch (e) {}
    render();
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
  async homologarEjecutivos() {
    if (
      !confirm(
        "Esto va a unificar en el catálogo cualquier nombre de Ejecutivo que esté duplicado solo por mayúsculas o espacios (dejando una sola entrada por persona), y va a corregir cualquier reporte, observación o asignación ya guardada que tuviera alguna de las variantes viejas. ¿Continuar?"
      )
    )
      return;
    state.homologarEjecutivosMsg = "Unificando…";
    render();
    try {
      const resultado = await homologarDuplicadosEjecutivos();
      state.homologarEjecutivosMsg =
        resultado.fusionados > 0
          ? `✅ Se unificaron ${resultado.fusionados} nombre(s) duplicado(s), y se corrigieron ${resultado.renglonesActualizados} renglón(es)/registro(s) que tenían alguna variante vieja.`
          : "No se encontró ningún duplicado para unificar.";
    } catch (e) {
      state.homologarEjecutivosMsg = "No se pudo completar la unificación (" + e.message + "). Vuelve a intentarlo.";
    }
    render();
  },
  async homologarClientes() {
    if (
      !confirm(
        "Esto va a unificar cualquier nombre de Cliente que esté duplicado solo por mayúsculas o espacios (dejando una sola entrada), y va a corregir cualquier renglón de Previos/Despachos/Revalidadas/Pendientes y Asignaciones ya guardado que tuviera alguna de las variantes viejas. ¿Continuar?"
      )
    )
      return;
    state.homologarClientesMsg = "Unificando…";
    render();
    try {
      const resultado = await homologarDuplicadosClientes();
      state.homologarClientesMsg =
        resultado.fusionados > 0
          ? `✅ Se unificaron ${resultado.fusionados} nombre(s) duplicado(s), y se corrigieron ${resultado.renglonesActualizados} renglón(es)/registro(s) que tenían alguna variante vieja.`
          : "No se encontró ningún duplicado para unificar.";
    } catch (e) {
      state.homologarClientesMsg = "No se pudo completar la unificación (" + e.message + "). Vuelve a intentarlo.";
    }
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
    state.asigEditId = null;
    state.asigDraft = {
      tipo: "previo",
      ref: "",
      guia: "",
      cliente: "",
      almacen: almacenFijo() || "",
      sector: "",
      ejecutivo: "",
      pedimento: "",
      tramitador: "",
    };
    state.asigMasivaOpen = false;
    state.asigMasivaMsg = "";
    state.asigMasivaDraft = {
      tipo: "previo",
      sector: "",
      tramitador: "",
      textoGuias: "",
      textoClientes: "",
      textoRefs: "",
      textoEjecutivos: "",
      textoPedimentos: "",
    };
    render();
  },
  goPendientesGlobal() {
    if (state.userRole !== "admin") return;
    state.view = "pendientesGlobal";
    state.errorMsg = "";
    render();
  },
  async goRespaldos() {
    if (state.userRole !== "admin") return;
    state.view = "respaldos";
    state.respaldosList = null; // null = todavía cargando
    state.errorMsg = "";
    render();
    try {
      state.respaldosList = await listarRespaldos(state.aduanaActiva);
    } catch (e) {
      state.errorMsg = "No se pudo cargar la lista de respaldos (" + e.message + ").";
      state.respaldosList = [];
    }
    render();
  },
  goSolicitudesEliminacion() {
    if (state.userRole !== "admin") return;
    state.view = "solicitudesEliminacion";
    state.errorMsg = "";
    render();
  },
  /** Admin autoriza una solicitud desde la pantalla centralizada — actúa directo sobre
   * el reporte guardado (sin necesitar tenerlo abierto para editar) y lo guarda.
   * Busca la captura por su ID único (no por posición en la lista — esa posición pudo
   * haber cambiado desde que se dibujó la pantalla, ej. porque alguien más guardó su
   * propio reporte de ese mismo día mientras tanto) y el renglón exacto por su "huella"
   * (la fecha/hora exacta en que se pidió esa solicitud puntual, prácticamente única) —
   * así nunca se autoriza/borra el renglón equivocado por casualidad. */
  async autorizarEliminacionDesdeGlobal(fecha, capturaId, grupo, tipo, marca) {
    const day = state.reports[fecha];
    const c = day && (day.capturas || []).find((cap) => cap.id === capturaId);
    if (!c) {
      state.errorMsg = "No se encontró esa hoja (puede que ya se haya movido o guardado distinto). Refresca e inténtalo de nuevo.";
      render();
      return;
    }
    const idx = (c[grupo] || []).findIndex((r) => (tipo === "foto" ? r._fechaSolicitudFoto === marca : r._fechaSolicitud === marca));
    if (idx === -1) {
      state.errorMsg = "Esa solicitud ya no existe (puede que alguien más ya la haya resuelto). Refresca la pantalla.";
      render();
      return;
    }
    const row = c[grupo][idx];
    if (tipo === "foto") {
      delete row.foto;
      delete row._solicitudQuitarFoto;
      delete row._solicitadoPorFoto;
      delete row._fechaSolicitudFoto;
    } else {
      c[grupo].splice(idx, 1);
    }
    render();
    try {
      await saveReportDay(fecha);
    } catch (e) {
      state.errorMsg = "Se autorizó aquí, pero no se pudo guardar en la nube (" + e.message + "). Vuelve a intentarlo con internet.";
      render();
    }
  },
  /** Admin rechaza una solicitud desde la pantalla centralizada — el renglón/foto se
   * queda tal cual, solo se le quita la marca de "pendiente de autorizar". Mismo criterio
   * de búsqueda segura (ID de captura + huella exacta) que autorizarEliminacionDesdeGlobal. */
  async rechazarEliminacionDesdeGlobal(fecha, capturaId, grupo, tipo, marca) {
    const day = state.reports[fecha];
    const c = day && (day.capturas || []).find((cap) => cap.id === capturaId);
    if (!c) {
      state.errorMsg = "No se encontró esa hoja (puede que ya se haya movido o guardado distinto). Refresca e inténtalo de nuevo.";
      render();
      return;
    }
    const row = (c[grupo] || []).find((r) => (tipo === "foto" ? r._fechaSolicitudFoto === marca : r._fechaSolicitud === marca));
    if (!row) {
      state.errorMsg = "Esa solicitud ya no existe (puede que alguien más ya la haya resuelto). Refresca la pantalla.";
      render();
      return;
    }
    if (tipo === "foto") {
      delete row._solicitudQuitarFoto;
      delete row._solicitadoPorFoto;
      delete row._fechaSolicitudFoto;
    } else {
      delete row._solicitudEliminacion;
      delete row._solicitadoPor;
      delete row._fechaSolicitud;
    }
    render();
    try {
      await saveReportDay(fecha);
    } catch (e) {
      state.errorMsg = "Se rechazó aquí, pero no se pudo guardar en la nube (" + e.message + "). Vuelve a intentarlo con internet.";
      render();
    }
  },
  /** Descarga un respaldo puntual como archivo JSON — sirve como punto de recuperación
   * manual: si algo saliera mal con los datos "en vivo", este archivo tiene una copia
   * completa de reportes, catálogos y asignaciones tal como estaban ese día. */
  async descargarRespaldoClave(clave) {
    try {
      state.processingMsg = "Preparando descarga…";
      render();
      const datos = await descargarRespaldo(state.aduanaActiva, clave);
      state.processingMsg = "";
      if (!datos) {
        state.errorMsg = "No se encontró ese respaldo (puede que ya se haya borrado por antigüedad).";
        render();
        return;
      }
      const blob = new Blob([JSON.stringify(datos, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `respaldo_${state.aduanaActiva}_${clave}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 3000);
      render();
    } catch (e) {
      state.processingMsg = "";
      state.errorMsg = "No se pudo descargar el respaldo (" + e.message + ").";
      render();
    }
  },
  /** Mueve un pendiente puntual (venga o no de una asignación formal) de la hoja donde
   * estaba a la hoja de HOY del tramitador elegido — se crea esa hoja si todavía no
   * existe. Deja marcado desde cuándo viene arrastrando (origenFecha), igual que el
   * traslado automático normal, para que se siga viendo el aviso de "arrastrado desde…".
   * Busca la captura por su ID único y el renglón por su guía — no por posición, que
   * puede cambiar entre que se dibuja la pantalla y que se le da clic a "Reasignar". */
  async reasignarPendienteGlobal(fecha, capturaId, guia, nuevoTramitador) {
    try {
      const day = state.reports[fecha];
      const captura = day && (day.capturas || []).find((c) => c.id === capturaId);
      if (!captura) return;
      const idx = (captura.pendientes || []).findIndex((p) => (p.guia || "").trim() === guia);
      if (idx === -1) return;
      const pend = captura.pendientes[idx];
      if (!nuevoTramitador || mismoNombre(nuevoTramitador, captura.tramitador)) return; // sin cambio real

      captura.pendientes.splice(idx, 1);
      const copia = { ...pend };
      if (!copia.origenFecha) copia.origenFecha = fecha;
      delete copia.bloqueado;

      const hoy = todayStr();
      if (!state.reports[hoy]) state.reports[hoy] = { fecha: hoy, capturas: [] };
      let destino = state.reports[hoy].capturas.find((c) => mismoNombre(c.tramitador, nuevoTramitador));
      if (!destino) {
        destino = emptyCaptura();
        destino.uploadedBy = state.user;
        destino.horaCaptura = new Date().toISOString();
        destino.sourceType = "manual";
        destino.fecha = hoy;
        destino.tramitador = nuevoTramitador;
        destino.aduana = state.aduanaActiva;
        normalizeCaptura(destino);
        state.reports[hoy].capturas.push(destino);
      }
      normalizeCaptura(destino);
      destino.pendientes.push(copia);

      render();
      let fallo = [];
      try {
        await saveReportDay(fecha);
      } catch (e) {
        fallo.push(fmtDateHuman(fecha));
      }
      if (hoy !== fecha) {
        try {
          await saveReportDay(hoy);
        } catch (e) {
          fallo.push(fmtDateHuman(hoy));
        }
      }
      if (fallo.length > 0) {
        state.errorMsg = `Se movió el pendiente aquí, pero no se pudo guardar en la nube el reporte del ${fallo.join(
          " y del "
        )} — revisa tu conexión y vuelve a intentarlo, para que no quede a medias entre los dos días.`;
        render();
      }
    } catch (e) {
      state.errorMsg = "No se pudo reasignar el pendiente (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  updateAsigDraft(field, value) {
    state.asigDraft[field] = value;
    // Sin render() aquí a propósito: así no se interrumpe mientras se está escribiendo.
    // El valor ya queda guardado en memoria, así que aunque llegue la actualización
    // automática de cada 25 segundos, el formulario no se borra.
  },
  /** Carga una asignación pendiente ya existente en el formulario de arriba, para poder
   * reasignarla a otro tramitador (o corregir cualquier otro dato) SIN borrarla y crear
   * una nueva — así conserva su historial (quién la creó, cuándo) en vez de perderlo. */
  /** Le pone tramitador a una asignación que se creó "sin asignar aún" (típico de una
   * carga masiva hecha antes de saber quién la va a trabajar) — sin tener que abrir el
   * formulario de editar/reasignar de arriba, un asignado rápido directo en la lista. */
  async asignarTramitadorRapido(id, nuevoTramitador) {
    const a = state.asignaciones[id];
    if (!a || !nuevoTramitador) {
      if (a && !nuevoTramitador) {
        state.errorMsg = "Elige un tramitador de la lista antes de asignar.";
        render();
      }
      return;
    }
    a.tramitador = nuevoTramitador;
    render();
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, a);
    } catch (e) {
      state.errorMsg = "No se pudo guardar el tramitador (" + e.message + "), pero quedó visible aquí por ahora.";
      render();
    }
  },
  goGuiasDisponibles() {
    if (state.userRole !== "tramite" || !permiteAutoasignacion()) return;
    state.view = "guiasDisponibles";
    state.errorMsg = "";
    render();
  },
  /** Un tramitador agarra una guía de la "bolsa" (sin tramitador asignado) para sí
   * mismo — SIEMPRE con el nombre con el que entró a la app (state.user), nunca puede
   * elegir a otra persona. Es la versión de autoservicio de asignarTramitadorRapido. */
  async asignarmeGuia(id) {
    if (state.userRole !== "tramite" || !permiteAutoasignacion()) return;
    const a = state.asignaciones[id];
    if (!a) return;
    if ((a.tramitador || "").trim()) {
      state.errorMsg = "Esta guía ya se la agarró alguien más justo antes que tú.";
      render();
      return;
    }
    a.tramitador = state.user;
    render();
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, a);
    } catch (e) {
      state.errorMsg = "No se pudo guardar (" + e.message + "), pero quedó visible aquí por ahora — vuelve a intentarlo con internet.";
      render();
    }
  },
  editarAsignacion(id) {
    const a = state.asignaciones[id];
    if (!a) return;
    state.asigEditId = id;
    state.asigDraft = {
      tipo: a.tipo || "previo",
      ref: a.ref || "",
      guia: a.guia || "",
      cliente: a.cliente || "",
      almacen: almacenFijo() || a.almacen || "",
      sector: a.sector || "",
      ejecutivo: a.ejecutivo || "",
      pedimento: a.pedimento || "",
      tramitador: a.tramitador || "",
    };
    state.errorMsg = "";
    state.view = "asignaciones"; // por si se llamó desde otra pantalla (ej. Pendientes), lleva al formulario
    render();
    const el = document.getElementById("asig_guia");
    if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
  },
  cancelarEdicionAsignacion() {
    state.asigEditId = null;
    state.asigDraft = {
      tipo: "previo",
      ref: "",
      guia: "",
      cliente: "",
      almacen: almacenFijo() || "",
      sector: "",
      ejecutivo: "",
      pedimento: "",
      tramitador: "",
    };
    render();
  },
  async crearAsignacion() {
    const tipo = state.asigDraft.tipo;
    const ref = (state.asigDraft.ref || "").trim();
    const guia = (state.asigDraft.guia || "").trim();
    const cliente = (state.asigDraft.cliente || "").trim();
    const almacen = almacenFijo() || (state.asigDraft.almacen || "").trim();
    const sector = state.asigDraft.sector || "";
    const ejecutivo = state.asigDraft.ejecutivo || "";
    const pedimento = (state.asigDraft.pedimento || "").trim();
    const tramitador = state.asigDraft.tramitador;
    if (!guia) {
      state.errorMsg = "Escribe el número de guía para crear la asignación.";
      render();
      return;
    }
    // Si se abrió con "✏️ Reasignar", se actualiza esa misma asignación en su lugar —
    // conserva su id, quién la creó originalmente y cuándo, en vez de crear una aparte.
    const id = state.asigEditId || "a_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
    const existente = state.asigEditId ? state.asignaciones[state.asigEditId] : null;
    const data = {
      tipo,
      ref,
      guia,
      cliente,
      almacen,
      sector,
      ejecutivo,
      pedimento,
      tramitador,
      estatus: existente && existente.estatus ? existente.estatus : "pendiente",
      creadoPor: existente && existente.creadoPor ? existente.creadoPor : state.user,
      fechaCreacion: existente && existente.fechaCreacion ? existente.fechaCreacion : new Date().toISOString(),
      fechaCompletado: (existente && existente.fechaCompletado) || null,
    };
    state.asignaciones[id] = data;
    state.asigEditId = null;
    state.asigDraft = { tipo: "previo", ref: "", guia: "", cliente: "", almacen: almacenFijo() || "", sector: "", ejecutivo: "", pedimento: "", tramitador };
    render();
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, data);
    } catch (e) {
      state.errorMsg = "No se pudo guardar la asignación en la nube (" + e.message + "), pero quedó visible aquí por ahora.";
      render();
    }
  },
  toggleAsigMasiva() {
    state.asigMasivaOpen = !state.asigMasivaOpen;
    state.asigMasivaMsg = "";
    render();
  },
  updateAsigMasivaDraft(field, value) {
    state.asigMasivaDraft[field] = value;
  },
  /** Guarda lo que se va escribiendo/pegando en cada uno de los 4 cuadros de carga masiva,
   * al instante y sin volver a dibujar la pantalla — así, aunque llegue la actualización
   * automática de cada 25 segundos, lo que ya se pegó no se borra. */
  updateAsigMasivaTexto(campo, value) {
    state.asigMasivaDraft[campo] = value;
  },
  /** Crea muchas asignaciones de un jalón, una por línea del cuadro de texto — pensado
   * para Querétaro, que recibe manifiestos de 5 a 100 guías de un jalón. Cada línea puede
   * traer Guía sola, o Guía + Cliente + Referencia separados por tabulador (al pegar de
   * Excel) o por coma. Tipo, Sector y Tramitador se comparten para todas las líneas. */
  /** Junta las 4 columnas pegadas por separado (Guías, Clientes, Referencia, Ejecutivo)
   * por POSICIÓN — el renglón 1 de cada cuadro se junta con el renglón 1 de los demás,
   * el renglón 2 con el renglón 2, etc. Así se puede pegar columna por columna directo
   * desde un manifiesto o Excel, sin tener que acomodar todo en un solo renglón por guía. */
  async crearAsignacionesMasivas() {
    const partir = (texto) =>
      (texto || "")
        .split("\n")
        .map((l) => l.trim())
        .filter(Boolean);
    const guias = partir(state.asigMasivaDraft.textoGuias);
    const clientes = partir(state.asigMasivaDraft.textoClientes);
    const refs = partir(state.asigMasivaDraft.textoRefs);
    const ejecutivos = partir(state.asigMasivaDraft.textoEjecutivos);
    const pedimentos = partir(state.asigMasivaDraft.textoPedimentos);
    if (guias.length === 0) {
      state.asigMasivaMsg = "Pega al menos una guía en el primer cuadro para crear asignaciones.";
      render();
      return;
    }
    const tipo = state.asigMasivaDraft.tipo;
    const sector = state.asigMasivaDraft.sector || "";
    const tramitador = state.asigMasivaDraft.tramitador;
    const almacen = almacenFijo() || "";
    // Todas las asignaciones creadas en esta misma carga masiva comparten un "loteId" —
    // así, más adelante en Pendientes, se pueden marcar "Listo" todas juntas de un jalón
    // en vez de una por una, sin perder la posibilidad de sacar alguna del lote a mano.
    const loteId = "lote_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6);
    let creadas = 0;
    let fallidas = 0;
    for (let i = 0; i < guias.length; i++) {
      const guia = guias[i];
      const cliente = clientes[i] || "";
      const ref = refs[i] || "";
      const ejecutivo = ejecutivos[i] || "";
      const pedimento = pedimentos[i] || "";
      const id = "a_" + Date.now() + "_" + i + "_" + Math.random().toString(36).slice(2, 6);
      const data = {
        tipo,
        ref,
        guia,
        cliente,
        almacen,
        sector,
        ejecutivo,
        pedimento,
        tramitador,
        loteId,
        estatus: "pendiente",
        creadoPor: state.user,
        fechaCreacion: new Date().toISOString(),
        fechaCompletado: null,
      };
      state.asignaciones[id] = data;
      creadas++;
      try {
        await fbSaveAsignacion(state.aduanaActiva, id, data);
      } catch (e) {
        fallidas++;
      }
    }
    state.asigMasivaDraft.textoGuias = "";
    state.asigMasivaDraft.textoClientes = "";
    state.asigMasivaDraft.textoRefs = "";
    state.asigMasivaDraft.textoEjecutivos = "";
    state.asigMasivaDraft.textoPedimentos = "";
    const avisosOrden = [];
    if (clientes.length && clientes.length !== guias.length) avisosOrden.push(`Clientes tenía ${clientes.length} renglón(es) en vez de ${guias.length}`);
    if (refs.length && refs.length !== guias.length) avisosOrden.push(`Referencia tenía ${refs.length} renglón(es) en vez de ${guias.length}`);
    if (pedimentos.length && pedimentos.length !== guias.length) avisosOrden.push(`Pedimento tenía ${pedimentos.length} renglón(es) en vez de ${guias.length}`);
    if (ejecutivos.length && ejecutivos.length !== guias.length) avisosOrden.push(`Ejecutivo tenía ${ejecutivos.length} renglón(es) en vez de ${guias.length}`);
    let msg = `✅ Se crearon ${creadas} asignación${creadas === 1 ? "" : "es"}.`;
    if (avisosOrden.length) msg += ` ⚠️ Revisa el orden: ${avisosOrden.join("; ")} — los renglones de más se ignoraron, los que faltaron quedaron en blanco.`;
    if (fallidas > 0)
      msg += ` ⚠️ ${fallidas} de ${creadas} no se pudieron subir a la nube todavía (se ven aquí, pero revisa tu conexión y vuelve a intentar más tarde para que no falten en otros dispositivos).`;
    state.asigMasivaMsg = msg;
    render();
  },
  /** Regresa una asignación "completada" a "pendiente" — para cuando el cierre
   * automático la marcó, pero después esa guía se borró o se corrigió en el reporte
   * (el cierre automático nunca revisa hacia atrás por su cuenta). No borra nada del
   * reporte, solo corrige el estatus de la asignación para poder reasignarla o
   * dejarla visible como pendiente de nuevo. */
  async regresarAPendiente(id) {
    const a = state.asignaciones[id];
    if (!a) return;
    if (!confirm(`¿Regresar la guía ${a.guia} a "pendiente"? Se le quita la marca de completada, pero no se toca nada del reporte.`)) return;
    const updated = { ...a, estatus: "pendiente", fechaCompletado: null };
    state.asignaciones[id] = updated;
    render();
    try {
      await fbSaveAsignacion(state.aduanaActiva, id, updated);
    } catch (e) {
      state.errorMsg = "Se regresó a pendiente aquí, pero no se pudo guardar en la nube (" + e.message + "). Vuelve a intentarlo con internet.";
      render();
    }
  },
  toggleSoloSinRastro() {
    state.soloSinRastro = !state.soloSinRastro;
    render();
  },
  /** Regresa a "pendiente" TODAS las asignaciones "completadas" que no tienen rastro
   * real en ningún reporte, de un jalón — para cuando hay muchas y no tiene caso hacerlo
   * una por una. No toca nada de los reportes, solo corrige el estatus de cada
   * asignación, igual que hacer "Regresar a pendiente" individual muchas veces seguidas. */
  async regresarTodasSinRastroAPendiente() {
    const indice = indiceGuiasEnReportes();
    const completadas = Object.entries(state.asignaciones || {})
      .map(([id, a]) => ({ id, ...a }))
      .filter((a) => a.estatus === "completada");
    const sinRastro = completadas.filter((a) => asignacionSinRastro(a, indice));
    if (sinRastro.length === 0) return;
    if (
      !confirm(
        `Esto va a regresar ${sinRastro.length} asignación(es) "completada" a "pendiente" (las que no tienen rastro en ningún reporte ahorita), para que se puedan revisar o reasignar. No se toca nada de los reportes ya guardados. ¿Continuar?`
      )
    )
      return;
    state.processingMsg = `Regresando ${sinRastro.length} a pendiente…`;
    render();
    let ok = 0;
    let fallidas = 0;
    for (const a of sinRastro) {
      const updated = { ...a, estatus: "pendiente", fechaCompletado: null };
      delete updated.id;
      state.asignaciones[a.id] = updated;
      try {
        await fbSaveAsignacion(state.aduanaActiva, a.id, updated);
        ok++;
      } catch (e) {
        fallidas++;
      }
    }
    state.processingMsg = "";
    state.errorMsg =
      fallidas > 0
        ? `Se regresaron ${ok} a pendiente. ${fallidas} no se pudieron guardar en la nube — revisa tu conexión y vuelve a intentar más tarde.`
        : "";
    render();
  },
  async eliminarAsignacion(id) {
    if (!confirm("¿Eliminar esta asignación?")) return;
    delete state.asignaciones[id];
    render();
    try {
      await fbDeleteAsignacion(state.aduanaActiva, id);
    } catch (e) {
      state.errorMsg = "Se quitó de aquí, pero no se pudo borrar en la nube (" + e.message + ") — podría reaparecer más tarde. Vuelve a intentarlo con internet.";
      render();
    }
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
    App.startManual();
  },
  /** Crea una hoja nueva — salvo que el usuario sea Trámite y YA tenga una hoja propia
   * capturada hoy: en ese caso la abre para seguir editándola, en vez de crear otra y
   * duplicar sus datos. Coordinación/Admin sí pueden crear varias hojas manuales (no
   * quedan ligadas a un solo tramitador). Antes de abrir/crear la hoja de hoy, si es
   * Trámite se revisan sus pendientes de días anteriores sin resolver y se trasladan
   * solos al día de hoy (ver migrarPendientesAlDiaDeHoy). */
  async startManual() {
    const hoy = todayStr();
    let heredados = [];
    if (state.userRole === "tramite") {
      heredados = await migrarPendientesAlDiaDeHoy(state.user);
      const day = state.reports[hoy];
      const idxExistente = day && day.capturas ? day.capturas.findIndex((c) => mismoNombre(c.tramitador, state.user)) : -1;
      if (idxExistente !== -1) {
        const captura = day.capturas[idxExistente];
        let cambios = false;
        heredados.forEach((p) => {
          const yaExiste = (captura.pendientes || []).some(
            (x) => x.origenFecha === p.origenFecha && (x.guia || "").trim() === (p.guia || "").trim() && (x.guia || "").trim() !== ""
          );
          if (!yaExiste) {
            captura.pendientes.push(p);
            cambios = true;
          }
        });
        if (cambios) {
          try {
            await saveReportDay(hoy);
          } catch (e) {}
        }
        App.editCaptura(hoy, captura.id);
        return;
      }
    }
    const cap = emptyCaptura();
    cap.uploadedBy = state.user;
    cap.horaCaptura = new Date().toISOString();
    cap.sourceType = "manual";
    cap.fecha = hoy;
    cap.tramitador = state.userRole === "tramite" ? state.user : "";
    cap.aduana = state.aduanaActiva;
    if (heredados.length) cap.pendientes = heredados;
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
    if (!editableCats[k].some((existente) => mismoNombre(existente, v))) editableCats[k].push(v);
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
      if (editableCats[k].some((existente) => mismoNombre(existente, n))) {
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

  traerAsignaciones() {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      const pendientes = asignacionesPendientesParaFormulario(c.tramitador);
      if (pendientes.length === 0) return;
      pendientes.forEach((a) => {
        const row = {};
        PENDIENTE_FIELDS.forEach((f) => (row[f.k] = ""));
        row.ref = a.ref || "";
        row.guia = a.guia || "";
        row.cliente = a.cliente || "";
        row.almacen = a.almacen || "";
        row.ejecutivo = a.ejecutivo || "";
        row.pedimento = a.pedimento || "";
        row.tipo = a.tipo === "previo" || a.tipo === "despacho" || a.tipo === "revalidada" ? a.tipo : "";
        row._sector = a.sector || ""; // no es columna visible en Pendientes, viaja oculto hasta que se marque "Listo"
        row._loteId = a.loteId || ""; // idem — identifica si viene de una carga masiva, para poder marcar "Listo" todo el lote junto
        c.pendientes.push(row);
      });
      render();
    } catch (e) {
      state.errorMsg = "No se pudieron traer las asignaciones (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  /** Trae a mano los pendientes de días anteriores de este tramitador que sigan sin
   * resolver — igual que el traslado automático, pero disparado por la propia persona
   * en cualquier momento, sin depender de que el traslado automático se haya activado
   * justo al crear el reporte del día. Congela el renglón original en su día viejo,
   * igual que el traslado automático. */
  async traerPendientesAnteriores() {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      if (!c.tramitador) return;
      const heredados = await migrarPendientesAlDiaDeHoy(c.tramitador);
      let agregados = 0;
      heredados.forEach((p) => {
        const yaExiste = (c.pendientes || []).some(
          (x) => x.origenFecha === p.origenFecha && (x.guia || "").trim() === (p.guia || "").trim() && (x.guia || "").trim() !== ""
        );
        if (!yaExiste) {
          c.pendientes.push(p);
          agregados++;
        }
      });
      render();
      if (agregados > 0 && c.fecha) {
        try {
          await saveReportDay(c.fecha);
        } catch (e) {}
      }
    } catch (e) {
      state.errorMsg = "No se pudieron traer los pendientes de días anteriores (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  /** Sube un renglón de Pendientes directo a Previos o Despachos (según su Tipo) y
   * lo quita de Pendientes. Si la Guía ya existe en ese cuadro (2do previo, etc.),
   * lo etiqueta automáticamente en la columna "N° previo/despacho" en vez de avisar
   * como duplicado, porque puede pasar legítimamente (varios previos por guía). */
  /** Marca "Listo" a TODOS los pendientes de un mismo lote (los que llegaron juntos de una
   * misma carga masiva de asignaciones) de un solo jalón — sin tener que marcar casillas a
   * mano. Cualquier guía se puede sacar del lote antes con su ✕ normal si necesita tratarse
   * aparte; ya no se contará en el botón del lote una vez borrada. */
  marcarLoteListo(loteId) {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      const indices = [];
      (c.pendientes || []).forEach((row, i) => {
        if (row.bloqueado || row._loteId !== loteId) return;
        if (row.tipo !== "despacho") return; // los lotes solo aplican a Despacho — Previo se queda individual
        indices.push(i);
      });
      indices.sort((a, b) => b - a).forEach((i) => App.marcarPendienteListo(i));
      render();
    } catch (e) {
      state.errorMsg = "No se pudo marcar el lote como Listo (" + e.message + "). Avísale a soporte con este mensaje.";
      render();
    }
  },
  marcarPendienteListo(i) {
    try {
      const c = normalizeCaptura(state.currentCaptura);
      const pend = c.pendientes[i];
      if (!pend || pend.bloqueado) return; // respaldo histórico fijo, no se puede accionar desde ahí
      if (pend.tipo !== "previo" && pend.tipo !== "despacho" && pend.tipo !== "revalidada") return;
      const tipo = pend.tipo;
      const grupo = tipo === "previo" ? "previos" : tipo === "revalidada" ? "revalidadas" : "despachos";
      const fields = tipo === "previo" ? PREVIO_FIELDS : tipo === "revalidada" ? REVALIDADA_FIELDS : DESPACHO_FIELDS;
      const guia = (pend.guia || "").trim();

      let count = (c[grupo] || []).filter((r) => (r.guia || "").trim() === guia).length;
      // Ojo: solo cuenta las veces que ESTE MISMO tramitador ya hizo esta guía — no un
      // conteo global de todos los reportes. Si se cuenta contra todo mundo, una guía
      // reasignada legítimamente a otra persona (algo normal en Coordinación) le saldría
      // "2do previo" a alguien que en realidad la está haciendo por primera vez — y con
      // el sistema de puntos, un "2do previo" resta en vez de sumar. No es justo castigar
      // así un trabajo que sí es nuevo para esta persona.
      if (guia) {
        allCapturas().forEach((cap) => {
          if (cap.id === c.id) return; // evitar contar dos veces la misma hoja que se está editando
          if (!mismoNombre(cap.tramitador, c.tramitador)) return;
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
      if (tipo === "previo" || tipo === "despacho") row.ejecutivo = pend.ejecutivo || "";
      if (tipo === "previo" && pend._sector) row.sector = pend._sector;
      if (tipo === "despacho") row.pedimento = pend.pedimento || "";
      if ((tipo === "previo" || tipo === "despacho") && pend.observaciones) {
        row.observaciones = pend.observaciones;
        row.observadoPor = pend.observadoPor || "";
      }
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
    const fijo = almacenFijo();
    if (fijo && "almacen" in empty) empty.almacen = fijo;
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
  /** Permite al Ejecutivo escribir SOLO en la columna "Observaciones" de Previos/Despachos
   * desde la vista de detalle (de solo lectura para todo lo demás) — sin pasar por el
   * formulario completo de edición, al que el Ejecutivo no tiene acceso. Guarda al salir
   * del campo (no en cada tecla), igual que el resto de campos de texto de la app. Busca
   * la captura por su ID único y el renglón por su guía — no por posición, que puede
   * cambiar entre que se dibuja la pantalla y que se sale del campo. */
  async updateObservacion(grupo, fecha, capturaId, guia, value) {
    if (state.userRole !== "ejecutivo") return;
    try {
      const day = state.reports[fecha];
      const c = day && (day.capturas || []).find((cap) => cap.id === capturaId);
      const row = c && (c[grupo] || []).find((r) => (r.guia || "").trim() === guia);
      if (!row) return;
      if (row.observaciones === value) return; // sin cambios, no gastar un guardado de más
      row.observaciones = value;
      row.observadoPor = value.trim() ? state.user : "";
      await saveReportDay(fecha);
      render();
    } catch (e) {
      state.errorMsg = "No se pudo guardar la observación (" + e.message + "). Intenta de nuevo.";
      render();
    }
  },
  /** Igual que updateRow, pero para selectores (Tipo, Dificultad, etc.) — estos SÍ
   * necesitan refrescar la pantalla de inmediato (para que aparezca el botón "Listo"
   * o cambie alguna otra parte visible), a diferencia de los campos de texto normales
   * donde refrescar en cada tecla interrumpiría al escribir. */
  updateRowSelect(group, i, key, val) {
    state.currentCaptura[group][i][key] = val;
    render();
  },
  /** Borra un renglón. Para Pendientes tiene cuidado especial: el respaldo histórico
   * fijo (bloqueado) ahora SÍ se puede borrar, pero pide confirmación aparte por ser
   * el respaldo. Y si se borra la copia viva de un pendiente que vino de un traslado
   * automático, se marca el original histórico como "descartado" — si no se hiciera
   * esto, la reconciliación automática (ver migrarPendientesAlDiaDeHoy) pensaría que
   * la copia nunca llegó a existir y la volvería a traer sola al día siguiente. */
  /** Admin rechaza una solicitud de eliminación — el renglón se queda tal cual, sin la
   * marca de "pendiente de borrar". Para APROBARLA, Admin simplemente usa el botón de
   * borrar normal (App.deleteRow) — como Admin sí puede borrar directo, no hace falta
   * un método aparte para eso. */
  async rechazarEliminacionRenglon(group, i) {
    const row = state.currentCaptura[group] && state.currentCaptura[group][i];
    if (!row) return;
    delete row._solicitudEliminacion;
    delete row._solicitadoPor;
    delete row._fechaSolicitud;
    render();
    await guardarCambioInmediatoSiExiste();
  },
  async deleteRow(group, i) {
    const row = state.currentCaptura[group] && state.currentCaptura[group][i];
    if (!row) return;

    // Solo Admin puede borrar directo. Trámite/Ejecutivo mandan una "solicitud de
    // eliminación" en vez de borrar — el renglón se queda visible, marcado, hasta que
    // Admin la autorice o la rechace. Así siempre hay control de quién puede de verdad
    // hacer desaparecer algo de un reporte.
    if (state.userRole !== "admin") {
      if (row._solicitudEliminacion) {
        // ya la había pedido — puede cancelarla ella misma
        delete row._solicitudEliminacion;
        delete row._solicitadoPor;
        delete row._fechaSolicitud;
      } else {
        row._solicitudEliminacion = true;
        row._solicitadoPor = state.user;
        row._fechaSolicitud = new Date().toISOString();
      }
      render();
      await guardarCambioInmediatoSiExiste();
      return;
    }

    if (group === "pendientes" && row.bloqueado) {
      if (!confirm("Este es el respaldo histórico fijo de un pendiente que ya se trasladó a otro día. ¿Seguro que lo quieres borrar de todos modos? No se puede deshacer.")) return;
    }
    const debeMarcarOriginalComoDescartado = group === "pendientes" && !row.bloqueado && row.origenFecha;
    const tramitadorActual = state.currentCaptura.tramitador;
    const guiaRow = row.guia;
    const origenFechaRow = row.origenFecha;
    state.currentCaptura[group].splice(i, 1);
    render();
    if (debeMarcarOriginalComoDescartado) {
      const origenDay = state.reports[origenFechaRow];
      if (origenDay) {
        let cambiado = false;
        (origenDay.capturas || []).forEach((c) => {
          if (!mismoNombre(c.tramitador, tramitadorActual)) return;
          (c.pendientes || []).forEach((p) => {
            if (p.bloqueado && p.origenFecha === origenFechaRow && normalizarNombre(p.guia) === normalizarNombre(guiaRow)) {
              p.descartado = true;
              cambiado = true;
            }
          });
        });
        if (cambiado) {
          try {
            await saveReportDay(origenFechaRow);
          } catch (e) {}
        }
      }
    }
  },
  /** Toma o sube una foto para un renglón de Guía revalidada — se comprime antes de
   * guardarse (se reduce tamaño y calidad) para no inflar la base de datos con fotos
   * de tamaño completo, que pesan mucho más que cualquier otro dato de la app. */
  async capturarFotoRevalidada(group, i, inputEl) {
    const file = inputEl && inputEl.files && inputEl.files[0];
    if (!file) return;
    const LIMITE_MB = 20;
    if (file.size > LIMITE_MB * 1024 * 1024) {
      state.errorMsg = `Esa foto pesa demasiado (más de ${LIMITE_MB} MB) — intenta con una foto más chica, o baja la calidad de la cámara del celular.`;
      if (inputEl) inputEl.value = "";
      render();
      return;
    }
    try {
      state.processingMsg = "Procesando foto…";
      render();
      const dataUrl = await comprimirImagenArchivo(file);
      const row = state.currentCaptura[group] && state.currentCaptura[group][i];
      if (row) row.foto = dataUrl;
      state.processingMsg = "";
      render();
      // Si se está editando un reporte que YA existe guardado, la foto se guarda de
      // inmediato en la nube — así, aunque la persona se salga sin darle al botón
      // final de "Guardar cambios" (se cierra la app, se le va la señal, etc.), la
      // foto no se pierde. En una hoja NUEVA (todavía sin guardar nada), la foto se
      // queda en memoria como el resto de los campos hasta el guardado final normal.
      // Se busca la hoja por su ID único, no por posición — importante cuando varias
      // personas trabajan en la app al mismo tiempo.
      if (state.editingFecha && state.currentCaptura.id) {
        const day = state.reports[state.editingFecha];
        const idx = day && Array.isArray(day.capturas) ? day.capturas.findIndex((c) => c.id === state.currentCaptura.id) : -1;
        if (idx !== -1) {
          day.capturas[idx] = state.currentCaptura;
          persistLocalCache();
          try {
            await saveReportDay(state.editingFecha);
          } catch (e2) {
            state.errorMsg = "La foto se ve aquí, pero no se pudo guardar en la nube todavía (" + e2.message + "). Vuelve a intentarlo con internet, o guarda el reporte completo con el botón de abajo.";
            render();
          }
        }
      }
    } catch (e) {
      state.processingMsg = "";
      state.errorMsg = "No se pudo procesar la foto (" + (e && e.message ? e.message : "error desconocido") + ").";
      render();
    }
  },
  /** Admin rechaza una solicitud de quitar foto — la foto se queda tal cual. */
  async rechazarQuitarFotoRevalidada(group, i) {
    const row = state.currentCaptura[group] && state.currentCaptura[group][i];
    if (!row) return;
    delete row._solicitudQuitarFoto;
    delete row._solicitadoPorFoto;
    delete row._fechaSolicitudFoto;
    render();
    await guardarCambioInmediatoSiExiste();
  },
  async quitarFotoRevalidada(group, i) {
    const row = state.currentCaptura[group] && state.currentCaptura[group][i];
    if (!row) return;

    if (state.userRole !== "admin") {
      if (row._solicitudQuitarFoto) {
        delete row._solicitudQuitarFoto;
        delete row._solicitadoPorFoto;
        delete row._fechaSolicitudFoto;
      } else {
        if (!confirm("Esto le manda a Admin una solicitud para quitar esta foto — no se quita directo. ¿Continuar?")) return;
        row._solicitudQuitarFoto = true;
        row._solicitadoPorFoto = state.user;
        row._fechaSolicitudFoto = new Date().toISOString();
      }
      render();
      await guardarCambioInmediatoSiExiste();
      return;
    }

    if (!confirm("¿Quitar esta foto?")) return;
    delete row.foto;
    delete row._solicitudQuitarFoto;
    delete row._solicitadoPorFoto;
    delete row._fechaSolicitudFoto;
    render();
    await guardarCambioInmediatoSiExiste();
  },
  /** Abre la foto de un renglón que se está capturando/editando ahorita mismo (todavía
   * en memoria, puede no estar guardado). */
  verFotoRevalidada(group, i) {
    const row = state.currentCaptura[group] && state.currentCaptura[group][i];
    if (!row || !row.foto) return;
    const w = window.open("", "_blank");
    if (w) w.document.write(`<title>Foto</title><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh;"><img src="${row.foto}" style="max-width:100%;max-height:100vh;"/></body>`);
  },
  /** Igual, pero para una foto ya guardada dentro de un reporte histórico (pantalla de
   * detalle de solo lectura) — se lee de state.reports en vez de la captura en edición. */
  verFotoRevalidadaDetalle(fecha, capturaId, guia) {
    const day = state.reports[fecha];
    const c = day && (day.capturas || []).find((cap) => cap.id === capturaId);
    const r = c && (c.revalidadas || []).find((row) => (row.guia || "").trim() === guia);
    if (!r || !r.foto) return;
    const w = window.open("", "_blank");
    if (w) w.document.write(`<title>Foto</title><body style="margin:0;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh;"><img src="${r.foto}" style="max-width:100%;max-height:100vh;"/></body>`);
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
        // Se busca por el ID único de la captura, NO por su posición en el arreglo —
        // si algo más se agregó o se quitó de ese mismo día mientras esta hoja estaba
        // abierta (otro tramitador guardando su propio reporte, por ejemplo), la
        // POSICIÓN puede haber cambiado desde que se abrió para editar. Guardar por
        // posición en ese caso sobrescribiría la captura equivocada — con el ID
        // siempre se encuentra la correcta, sin importar en qué lugar del arreglo
        // haya quedado.
        const day = state.reports[c.fecha];
        const idxReal = (day.capturas || []).findIndex((cap) => cap.id === c.id);
        if (idxReal !== -1) {
          day.capturas[idxReal] = c;
        } else {
          day.capturas.push(c);
        }
      } else {
        if (wasEditing) {
          // La fecha sí cambió: quitar la versión anterior de donde estaba (también
          // por ID, mismo motivo)
          const oldDay = state.reports[state.editingFecha];
          if (oldDay && oldDay.capturas) {
            const idxViejo = oldDay.capturas.findIndex((cap) => cap.id === c.id);
            if (idxViejo !== -1) {
              oldDay.capturas.splice(idxViejo, 1);
              if (oldDay.capturas.length === 0) delete state.reports[state.editingFecha];
            }
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

  /** Recibe el ID único de la captura (no su posición en el arreglo) — la posición se
   * busca aquí mismo, en el momento del clic, para no arriesgarse a que haya cambiado
   * desde que se dibujó la lista (ej. otro tramitador guardando su propio reporte ese
   * mismo día, entre que se vio la lista y se le dio clic a "Editar"). */
  editCaptura(fecha, capturaId) {
    if (state.userRole === "ejecutivo") return; // solo lectura
    try {
      const day = state.reports[fecha];
      const idx = day && day.capturas ? day.capturas.findIndex((c) => c.id === capturaId) : -1;
      if (!day || idx === -1) {
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

  async deleteCaptura(fecha, capturaId) {
    if (state.userRole !== "admin") return;
    if (!confirm("¿Eliminar esta hoja del reporte?")) return;
    const day = state.reports[fecha];
    const idx = day && day.capturas ? day.capturas.findIndex((c) => c.id === capturaId) : -1;
    if (!day || idx === -1) {
      state.errorMsg = "No se encontró esa hoja (puede que ya se haya movido o eliminado).";
      render();
      return;
    }
    day.capturas.splice(idx, 1);
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
