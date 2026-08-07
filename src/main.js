import "./style.css";
import * as XLSX from "xlsx";
import { PREVIO_FIELDS, DESPACHO_FIELDS, emptyCaptura } from "./fields.js";
import { parseDocxTemplate, guessFechaFromFileName } from "./parseDocx.js";
import { parsePdfTemplate } from "./parsePdf.js";
import {
  loadReports as fbLoadReports,
  saveReportDay as fbSaveReportDay,
  deleteReportDay as fbDeleteReportDay,
  loadCatalogos as fbLoadCatalogos,
  saveCatalogos as fbSaveCatalogos,
} from "./storage.js";

const USERS = ["Davis", "Compañero(a)"];

let editableCats = {
  clientes: ["GLXI", "Alkaps", "MTI", "FMI", "IndoUnión", "Foray", "Alpha metal", "BRP", "PMI"],
  almacenes: ["228", "277", "CLA", "WTC"],
  tramitadores: ["Luis Arreola"],
};

let state = {
  user: null,
  view: "login",
  reports: {},
  loading: true,
  currentCaptura: null,
  editingIndex: null,
  editingFecha: null,
  detailFecha: null,
  errorMsg: "",
  processingMsg: "",
};

const root = document.getElementById("app-root");

function esc(s) {
  return s === undefined || s === null
    ? ""
    : String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
function fmtDateHuman(iso) {
  if (!iso) return "Sin fecha";
  const [y, m, d] = iso.split("-");
  const meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${d} ${meses[parseInt(m, 10) - 1]} ${y}`;
}

// ---------- persistence wrappers ----------
async function loadReports() {
  try {
    state.reports = await fbLoadReports();
  } catch (e) {
    state.reports = {};
    state.errorMsg = "No se pudo conectar con Firebase (" + e.message + "). Verifica tu configuración y reglas de Realtime Database.";
  }
  try {
    const cats = await fbLoadCatalogos();
    if (cats) editableCats = cats;
  } catch (e) {}
  state.loading = false;
  render();
}
async function saveReportDay(fecha) {
  await fbSaveReportDay(fecha, state.reports[fecha]);
}
async function deleteReportDay(fecha) {
  await fbDeleteReportDay(fecha);
}
async function saveCatalogos() {
  try {
    await fbSaveCatalogos(editableCats);
  } catch (e) {}
}
function loadLastUser() {
  try {
    const u = localStorage.getItem("ow_ultimo_usuario");
    if (u) state.user = u;
  } catch (e) {}
}
function saveLastUser(u) {
  try {
    localStorage.setItem("ow_ultimo_usuario", u);
  } catch (e) {}
}

// ---------- stats ----------
function allCapturas() {
  const out = [];
  Object.values(state.reports).forEach((r) => (r.capturas || []).forEach((c) => out.push({ ...c, fecha: r.fecha })));
  return out;
}
function totalPrevios() {
  return allCapturas().reduce((s, c) => s + (c.previos ? c.previos.length : 0), 0);
}
function totalDespachos() {
  return allCapturas().reduce((s, c) => s + (c.despachos ? c.despachos.length : 0), 0);
}
function last14Days() {
  const days = [];
  for (let i = 13; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    days.push(d.toISOString().slice(0, 10));
  }
  return days.map((iso) => {
    const r = state.reports[iso];
    return { label: iso.slice(8, 10) + "/" + iso.slice(5, 7), value: r ? (r.capturas || []).length : 0 };
  });
}
function topClientes() {
  const counts = {};
  allCapturas().forEach((c) => {
    [...(c.previos || []), ...(c.despachos || [])].forEach((r) => {
      const t = (r.cliente || "").trim();
      if (!t) return;
      counts[t] = (counts[t] || 0) + 1;
    });
  });
  return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 5);
}

// ---------- rendering ----------
function render() {
  try {
    if (state.view === "login") root.innerHTML = viewLogin();
    else {
      root.innerHTML =
        topbar() +
        `<div class="content">${
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

function topbar() {
  return `
  <div class="topbar">
    <div class="brand">
      <div class="brand-mark">OW</div>
      <div class="brand-text"><div class="t1">Oñate, Willy &amp; Cía.</div><div class="t2">Reporte operativo diario</div></div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
      <div class="user-chip"><span class="user-dot"></span>${esc(state.user)}</div>
      ${state.view !== "home" ? `<button class="nav-btn" onclick="App.goHome()">Inicio</button>` : ""}
      <button class="nav-btn" onclick="App.goCatalogos()">Catálogos</button>
      <button class="nav-btn" onclick="App.logout()">Salir</button>
    </div>
  </div>`;
}

function viewLogin() {
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-mark">OW</div>
      <div class="login-title display">Oñate, Willy &amp; Cía.</div>
      <div class="login-sub">Digitalización del reporte operativo diario de previos y despachos</div>
      <div class="user-pick">
        ${USERS.map(
          (u) =>
            `<button class="user-pick-btn" onclick="App.login('${esc(u)}')"><span class="user-avatar">${esc(
              u.slice(0, 2).toUpperCase()
            )}</span> ${esc(u)}</button>`
        ).join("")}
      </div>
    </div>
  </div>`;
}

function viewHome() {
  const days = Object.values(state.reports).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const bars = last14Days();
  const maxVal = Math.max(1, ...bars.map((b) => b.value));
  const tp = totalPrevios(),
    td = totalDespachos();
  const ranks = topClientes();
  const maxRank = Math.max(1, ...ranks.map((r) => r[1]));

  return `
    ${errorBanner()}
    <div class="stats-grid">
      <div class="stat-card"><div class="stat-num">${days.length}</div><div class="stat-label">Días con reportes</div></div>
      <div class="stat-card"><div class="stat-num">${allCapturas().length}</div><div class="stat-label">Reportes capturados</div></div>
      <div class="stat-card"><div class="stat-num">${tp}</div><div class="stat-label">Total previos</div></div>
      <div class="stat-card"><div class="stat-num">${td}</div><div class="stat-label">Total despachos</div></div>
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
        <div class="section-sub">Por número de operaciones</div>
        ${
          ranks.length === 0
            ? `<div style="color:var(--muted);font-size:12.5px;">Sin datos todavía</div>`
            : ranks
                .map(
                  ([name, val]) => `
          <div class="rank-row">
            <div class="rank-name" title="${esc(name)}">${esc(name)}</div>
            <div class="rank-bar-bg"><div class="rank-bar" style="width:${(val / maxRank) * 100}%"></div></div>
            <div class="rank-val">${val}</div>
          </div>`
                )
                .join("")
        }
      </div>
    </div>

    <div class="top-actions">
      <div>
        <div class="section-title" style="margin-bottom:0;">Reportes por fecha</div>
        <div class="section-sub">Cada hoja subida se agrupa automáticamente en el reporte de su fecha</div>
      </div>
      <button class="btn btn-primary" onclick="App.startCapture()">+ Nuevo reporte</button>
    </div>

    ${
      days.length === 0
        ? `<div class="empty"><div class="stamp-outline">📋</div><div style="font-weight:600;margin-bottom:4px;">Aún no hay reportes capturados</div><div style="font-size:13px;">Toca "Nuevo reporte" para subir tu primer Word, PDF, o llenarlo manualmente.</div></div>`
        : days
            .map((d) => {
              const p = (d.capturas || []).reduce((s, c) => s + (c.previos ? c.previos.length : 0), 0);
              const de = (d.capturas || []).reduce((s, c) => s + (c.despachos ? c.despachos.length : 0), 0);
              return `
        <div class="day-card" onclick="App.openDetail('${d.fecha}')">
          <div class="day-card-head">
            <div><div class="day-date">${fmtDateHuman(d.fecha)}</div><div class="day-meta">${
                (d.capturas || []).length
              } hoja(s) capturada(s)</div></div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;">
              ${p ? `<span class="pill pill-navy">${p} previo${p > 1 ? "s" : ""}</span>` : ""}
              ${de ? `<span class="pill pill-verde">${de} despacho${de > 1 ? "s" : ""}</span>` : ""}
            </div>
          </div>
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
        <div class="field"><label>Tramitador</label><input id="f_tram" value="${esc(c.tramitador)}" list="dl_tramitadores"/></div>
        <div class="field"><label>Aduana</label><input id="f_aduana" value="${esc(c.aduana)}"/></div>
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

      <div class="subhead">3. Otras actividades</div>
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
        const dl = f.cat ? ` list="dl_${f.cat}"` : "";
        return `<td><input${dl} value="${esc(row[f.k])}" oninput="App.updateRow('${group}',${i},'${f.k}',this.value)"/></td>`;
      })
      .join("")}
    <td><button class="row-del" onclick="App.deleteRow('${group}',${i})" title="Eliminar renglón">✕</button></td>
  </tr>`;
}

function datalists() {
  const merged = { ...editableCats, resultados: ["Verde", "Rojo", "Desaduanamiento libre", "Reconocimiento aduanero"] };
  return Object.entries(merged)
    .map((entry) => `<datalist id="dl_${entry[0]}">${entry[1].map((v) => `<option value="${esc(v)}">`).join("")}</datalist>`)
    .join("");
}

function viewStamp() {
  return `
  <div class="stamp-screen">
    <div class="stamp"><div class="s1">CONFIRMADO</div><div class="s2 mono">${fmtDateHuman(
      state.currentCaptura ? state.currentCaptura.fecha : todayStr()
    )}</div></div>
    <div style="font-weight:600;font-size:16px;margin-bottom:4px;">Reporte guardado</div>
    <div style="color:var(--muted);font-size:13px;margin-bottom:20px;">Se agregó al reporte del día correspondiente</div>
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
        <button class="btn btn-primary btn-sm" onclick="App.exportExcel('${d.fecha}')">Descargar Excel</button>
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
            <button class="btn btn-ghost btn-sm" onclick="App.editCaptura('${d.fecha}',${ci})">Editar</button>
            <button class="row-del" onclick="App.deleteCaptura('${d.fecha}',${ci})" title="Eliminar hoja">✕</button>
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
          (c.otrasActividades || []).filter((a) => a && a.trim()).length
            ? `
          <div class="mini-title">3. Otras actividades</div>
          <div style="font-size:12.5px;color:var(--ink);">${(c.otrasActividades || [])
            .filter((a) => a && a.trim())
            .map((a, i) => `${i + 1}. ${esc(a)}`)
            .join("<br/>")}</div>
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
  return `
    <button class="back-link" onclick="App.goHome()">← Inicio</button>
    <div class="section-title">Catálogos</div>
    <div class="section-sub">Nombres sugeridos automáticamente al capturar (clientes, almacenes, tramitadores)</div>
    ${["clientes", "almacenes", "tramitadores"]
      .map(
        (k) => `
      <div class="panel" style="margin-bottom:14px;">
        <div style="font-weight:600;font-size:13.5px;margin-bottom:10px;text-transform:capitalize;">${k}</div>
        <div class="chip-list">${
          editableCats[k].map((v, i) => `<span class="chip">${esc(v)}<button onclick="App.removeCat('${k}',${i})">✕</button></span>`).join("") ||
          `<span style="color:var(--muted);font-size:12.5px;">Sin registros</span>`
        }</div>
        <div class="cat-add">
          <input type="text" id="new_${k}" placeholder="Agregar…" onkeydown="if(event.key==='Enter') App.addCat('${k}')"/>
          <button class="btn btn-ghost btn-sm" onclick="App.addCat('${k}')">Agregar</button>
        </div>
      </div>`
      )
      .join("")}
  `;
}

// ---------- app logic ----------
const App = {
  login(u) {
    state.user = u;
    saveLastUser(u);
    state.view = "home";
    state.loading = true;
    render();
    loadReports();
  },
  logout() {
    state.user = null;
    state.view = "login";
    render();
  },
  goHome() {
    state.view = "home";
    state.currentCaptura = null;
    state.editingIndex = null;
    state.editingFecha = null;
    state.errorMsg = "";
    state.processingMsg = "";
    render();
  },
  goCatalogos() {
    state.view = "catalogos";
    render();
  },
  startCapture() {
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
    state.currentCaptura = cap;
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
    cap.tramitador = parsed.tramitador || "";
    cap.aduana = parsed.aduana || "";
    cap.previos = Array.isArray(parsed.previos) ? parsed.previos : [];
    cap.despachos = Array.isArray(parsed.despachos) ? parsed.despachos : [];
    cap.otrasActividades =
      Array.isArray(parsed.otrasActividades) && parsed.otrasActividades.length ? parsed.otrasActividades : ["", "", "", ""];
    state.currentCaptura = cap;
    state.editingIndex = null;
    state.editingFecha = null;
    state.view = "review";
    render();
  },

  addRow(group) {
    const fields = group === "previos" ? PREVIO_FIELDS : DESPACHO_FIELDS;
    const empty = {};
    fields.forEach((f) => (empty[f.k] = ""));
    state.currentCaptura[group].push(empty);
    render();
  },
  updateRow(group, i, key, val) {
    state.currentCaptura[group][i][key] = val;
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

      if (!state.reports[c.fecha]) state.reports[c.fecha] = { fecha: c.fecha, capturas: [] };

      const wasEditing = state.editingIndex !== null && state.editingFecha;
      if (wasEditing) {
        // Quitar la versión anterior de donde estaba (pudo cambiar de fecha al editar)
        const oldDay = state.reports[state.editingFecha];
        if (oldDay && oldDay.capturas && oldDay.capturas[state.editingIndex]) {
          oldDay.capturas.splice(state.editingIndex, 1);
          if (oldDay.capturas.length === 0) delete state.reports[state.editingFecha];
        }
      }
      state.reports[c.fecha].capturas.push(c);

      // Si el día viejo cambió (edición que cambió de fecha), guardar ambos días en Firebase
      if (wasEditing && state.editingFecha !== c.fecha && state.reports[state.editingFecha]) {
        await saveReportDay(state.editingFecha);
      } else if (wasEditing && state.editingFecha !== c.fecha && !state.reports[state.editingFecha]) {
        await deleteReportDay(state.editingFecha);
      }
      await saveReportDay(c.fecha);

      state.editingIndex = null;
      state.editingFecha = null;
      state.view = "stamp";
      render();
    } catch (e) {
      state.errorMsg = "No se pudo guardar el reporte (" + e.message + "). Intenta de nuevo.";
      state.view = "capture";
      render();
    }
  },

  editCaptura(fecha, idx) {
    try {
      const day = state.reports[fecha];
      if (!day || !day.capturas || !day.capturas[idx]) {
        state.errorMsg = "No se encontró ese reporte para editar (puede que ya haya sido eliminado o modificado).";
        render();
        return;
      }
      const c = day.capturas[idx];
      state.currentCaptura = JSON.parse(JSON.stringify(c));
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
    if (!confirm("¿Eliminar esta hoja del reporte?")) return;
    state.reports[fecha].capturas.splice(idx, 1);
    if (state.reports[fecha].capturas.length === 0) {
      await deleteReportDay(fecha);
      delete state.reports[fecha];
      state.view = "home";
    } else {
      await saveReportDay(fecha);
    }
    render();
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
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte_${fecha}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  },

  exportExcel(fecha) {
    const d = state.reports[fecha];
    const wb = XLSX.utils.book_new();

    const resumenRows = (d.capturas || []).map((c) => ({
      Fecha: c.fecha,
      Tramitador: c.tramitador,
      Aduana: c.aduana,
      "Total previos": (c.previos || []).length,
      "Total despachos": (c.despachos || []).length,
      "Archivo origen": c.sourceFileName || "(captura manual)",
      "Capturado por": c.uploadedBy || "",
    }));
    const wsResumen = XLSX.utils.json_to_sheet(resumenRows);
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");

    const previosRows = [];
    (d.capturas || []).forEach((c) =>
      (c.previos || []).forEach((r) => {
        const row = { Fecha: c.fecha, Tramitador: c.tramitador };
        PREVIO_FIELDS.forEach((f) => (row[f.label] = r[f.k] || ""));
        previosRows.push(row);
      })
    );
    const wsPrevios = XLSX.utils.json_to_sheet(previosRows);
    XLSX.utils.book_append_sheet(wb, wsPrevios, "Previos");

    const despachosRows = [];
    (d.capturas || []).forEach((c) =>
      (c.despachos || []).forEach((r) => {
        const row = { Fecha: c.fecha, Tramitador: c.tramitador };
        DESPACHO_FIELDS.forEach((f) => (row[f.label] = r[f.k] || ""));
        despachosRows.push(row);
      })
    );
    const wsDespachos = XLSX.utils.json_to_sheet(despachosRows);
    XLSX.utils.book_append_sheet(wb, wsDespachos, "Despachos");

    XLSX.writeFile(wb, `reporte_${fecha}.xlsx`);
  },
};
window.App = App;

// ---------- init ----------
loadLastUser();
render();
if (state.user) {
  state.loading = true;
  render();
  loadReports();
}
