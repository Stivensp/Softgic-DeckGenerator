// Softgic Deck Generator — editor de CONTENIDO web (sin diseno/posiciones).
// Construye un formulario dinamico por slide a partir de /api/schema (espejo
// de los modelos pydantic en src/models/slides.py), parte del ultimo
// input_template.yaml, y envia los datos a /api/render para generar el .pptx.
//
// Un solo slide esta "enfocado" (state.focusedIdx) a la vez — la lista de
// slides es navegacion/seleccion, el canvas y el panel de propiedades
// muestran y editan ese slide. El canvas es una aproximacion HTML/CSS
// (posicion/color/texto reales de config/generated/layout.yaml +
// config/theme.yaml vivo desde el formulario) — NO es un render real del
// .pptx ni permite seleccionar elementos individuales; eso depende del motor
// de render+fidelidad del backend OOXML (ver memoria del proyecto), todavia
// no construido.

const state = {
  slideTypes: [],
  schema: {},
  limits: {},
  slides: [],
  focusedIdx: 0,
  canvasLayoutCache: {},
  searchQuery: "",
  zoom: 1.0,
};

const CANVAS_W_IN = 13.33;
const CANVAS_H_IN = 7.5;
const CANVAS_BASE_PX = 640; // debe coincidir con --canvas width en style.css (.canvas-slide)

// ───────────────────────────── iconos ──────────────────────────────────────

const ICONS = {
  chevron: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>',
  up: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"></polyline></svg>',
  down: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>',
  trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"></path><path d="M10 11v6"></path><path d="M14 11v6"></path></svg>',
  plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>',
  copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>',
  image: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>',
};

const SLIDE_META = {
  cover: { label: "Portada", short: "CO", color: "var(--t-cover)" },
  section_divider: { label: "Sección", short: "SD", color: "var(--t-section_divider)" },
  content_two_col: { label: "2 Columnas", short: "2C", color: "var(--t-content_two_col)" },
  content_one_col: { label: "1 Columna", short: "1C", color: "var(--t-content_one_col)" },
  pricing_table: { label: "Precios", short: "$", color: "var(--t-pricing_table)" },
  profile_card: { label: "Perfil", short: "PF", color: "var(--t-profile_card)" },
  stat_callout: { label: "Estadística", short: "%", color: "var(--t-stat_callout)" },
  closing: { label: "Cierre", short: "CL", color: "var(--t-closing)" },
  grafico: { label: "Gráfico", short: "GR", color: "var(--t-grafico)" },
};

function meta(type) {
  return SLIDE_META[type] || { label: type, short: "?", color: "#64748b" };
}

// ───────────────────────────── utilidades ────────────────────────────────

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function getPath(obj, path) {
  return path.split(".").reduce((o, k) => (o == null ? undefined : o[k]), obj);
}

function setPath(obj, path, value) {
  const keys = path.split(".");
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    const k = keys[i];
    if (cur[k] === undefined || cur[k] === null) {
      cur[k] = /^\d+$/.test(keys[i + 1]) ? [] : {};
    }
    cur = cur[k];
  }
  cur[keys[keys.length - 1]] = value;
}

// ───────────────────────────── toasts ──────────────────────────────────────

function showToast(message, type = "ok", duration = 4500) {
  const container = document.getElementById("toastContainer");
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${message}</span>`;
  container.appendChild(el);
  if (duration) {
    setTimeout(() => {
      el.classList.add("leaving");
      setTimeout(() => el.remove(), 200);
    }, duration);
  }
  return el;
}

// ───────────────────────────── modal de confirmacion ──────────────────────

function showConfirm({ title, body, confirmText = "Confirmar", cancelText = "Cancelar", danger = false }) {
  return new Promise((resolve) => {
    const root = document.getElementById("modalRoot");
    root.innerHTML = `
      <div class="modal-card">
        <p class="modal-title">${escapeHtml(title)}</p>
        <p class="modal-body">${escapeHtml(body)}</p>
        <div class="modal-actions">
          <button class="btn btn-ghost" id="modalCancel">${escapeHtml(cancelText)}</button>
          <button class="btn btn-primary" id="modalConfirm" ${danger ? 'style="background:var(--c-danger);border-color:var(--c-danger);color:#fff"' : ""}>${escapeHtml(confirmText)}</button>
        </div>
      </div>`;
    root.hidden = false;
    const close = (result) => { root.hidden = true; root.innerHTML = ""; resolve(result); };
    document.getElementById("modalCancel").addEventListener("click", () => close(false));
    document.getElementById("modalConfirm").addEventListener("click", () => close(true));
    root.addEventListener("click", (e) => { if (e.target === root) close(false); }, { once: true });
  });
}

// ───────────────────────────── defaults por tipo ──────────────────────────

function defaultSlideForType(type) {
  const slide = { type };
  (state.schema[type] || []).forEach((f) => {
    if (f.type === "list") slide[f.name] = [];
    else if (f.type === "bullet_column") slide[f.name] = { title: "", bullets: [] };
    else if (f.type === "pricing_rows") slide[f.name] = [];
    else if (f.type === "chart_series") slide[f.name] = [];
    else if (f.type === "number") slide[f.name] = null;
    else if (f.type === "select") slide[f.name] = (f.options && f.options[0]) || "";
    else slide[f.name] = "";
  });
  return slide;
}

// ───────────────────────────── render de campos ───────────────────────────

function charMeter(field, val) {
  if (!field.max_chars_key || !state.limits[field.max_chars_key]) return "";
  const max = state.limits[field.max_chars_key];
  const len = (val || "").length;
  const pct = Math.min(100, (len / max) * 100);
  const over = len > max;
  return `<div class="char-meter">
    <div class="char-meter-bar"><div class="char-meter-fill${over ? " over" : ""}" data-max="${max}" style="width:${pct}%"></div></div>
    <span class="char-meter-text${over ? " over" : ""}">${len}/${max}</span>
  </div>`;
}

function reqMark(field) {
  return field.required ? '<span class="f-required">*</span>' : "";
}

function textRow(field, slideIdx, path, val) {
  return `<div class="f-row">
    <label>${field.label} ${reqMark(field)}</label>
    <input type="text" data-required="${!!field.required}" data-slide-idx="${slideIdx}" data-path="${path}" value="${escapeHtml(val)}">
    ${charMeter(field, val)}
  </div>`;
}

function numberRow(field, slideIdx, path, val) {
  return `<div class="f-row">
    <label>${field.label} ${reqMark(field)}</label>
    <input type="number" step="any" data-required="${!!field.required}" data-slide-idx="${slideIdx}" data-path="${path}" value="${val ?? ""}">
  </div>`;
}

function selectRow(field, slideIdx, path, val) {
  const opts = (field.options || []).map((o) => `<option value="${o}" ${o === val ? "selected" : ""}>${o}</option>`).join("");
  return `<div class="f-row">
    <label>${field.label}</label>
    <select data-slide-idx="${slideIdx}" data-path="${path}">${opts}</select>
  </div>`;
}

function listRow(label, required, slideIdx, path, items) {
  let html = `<div class="f-row"><label>${label} ${required ? '<span class="f-required">*</span>' : ""}</label><div class="list-editor">`;
  items.forEach((item, i) => {
    html += `<div class="list-item">
      <input type="text" data-slide-idx="${slideIdx}" data-path="${path}.${i}" value="${escapeHtml(item)}">
      <button type="button" class="icon-btn danger" data-action="remove-row" data-slide-idx="${slideIdx}" data-path="${path}" data-row="${i}" title="Quitar">${ICONS.trash}</button>
    </div>`;
  });
  html += `<button type="button" class="btn btn-sm" data-action="add-list-item" data-slide-idx="${slideIdx}" data-path="${path}">${ICONS.plus} Agregar</button>`;
  html += "</div></div>";
  return html;
}

function bulletColumnRow(field, slideIdx, path, col) {
  col = col || { title: "", bullets: [] };
  let html = `<div class="f-row bullet-column"><label>${field.label} ${reqMark(field)}</label>`;
  html += `<input type="text" placeholder="Titulo de columna" data-slide-idx="${slideIdx}" data-path="${path}.title" value="${escapeHtml(col.title || "")}">`;
  html += listRow("Bullets", true, slideIdx, `${path}.bullets`, col.bullets || []);
  html += "</div>";
  return html;
}

function pricingRowsField(field, slideIdx, path, rows) {
  rows = rows || [];
  let computed = 0;
  rows.forEach((r) => { computed += Number(r.total) || 0; });

  let html = `<div class="f-row"><label>${field.label} ${reqMark(field)}</label>`;
  html += `<div class="pricing-rows-head"><span>Descripcion</span><span>Cant.</span><span>Unidad</span><span>P.Unit</span><span>Total</span><span></span></div>`;
  html += `<div class="pricing-rows">`;
  rows.forEach((r, i) => {
    html += `<div class="pricing-row">
      <input type="text" placeholder="Descripcion" data-slide-idx="${slideIdx}" data-path="${path}.${i}.description" value="${escapeHtml(r.description || "")}">
      <input type="number" placeholder="1" data-slide-idx="${slideIdx}" data-path="${path}.${i}.quantity" value="${r.quantity ?? ""}">
      <input type="text" placeholder="hrs" data-slide-idx="${slideIdx}" data-path="${path}.${i}.unit" value="${escapeHtml(r.unit || "")}">
      <input type="number" step="0.01" placeholder="0.00" data-slide-idx="${slideIdx}" data-path="${path}.${i}.unit_price" value="${r.unit_price ?? ""}">
      <input type="number" step="0.01" placeholder="0.00" data-slide-idx="${slideIdx}" data-path="${path}.${i}.total" value="${r.total ?? ""}">
      <button type="button" class="icon-btn danger" data-action="remove-row" data-slide-idx="${slideIdx}" data-path="${path}" data-row="${i}" title="Quitar fila">${ICONS.trash}</button>
    </div>`;
  });
  html += `</div>`;
  html += `<button type="button" class="btn btn-sm" data-action="add-pricing-row" data-slide-idx="${slideIdx}" data-path="${path}">${ICONS.plus} Agregar fila</button>`;
  html += `<div class="pricing-rows-foot" data-role="pricing-foot" data-slide-idx="${slideIdx}">Suma de filas: <strong>${computed.toFixed(2)}</strong></div>`;
  html += "</div>";
  return html;
}

function chartSeriesField(field, slideIdx, path, series) {
  series = series || [];
  let html = `<div class="f-row"><label>${field.label} ${reqMark(field)}</label><div class="chart-series">`;
  series.forEach((s, i) => {
    const valuesStr = (s.values || []).join(", ");
    html += `<div class="series-row">
      <input type="text" placeholder="Nombre de la serie" data-slide-idx="${slideIdx}" data-path="${path}.${i}.name" value="${escapeHtml(s.name || "")}">
      <input type="text" placeholder="Valores separados por coma (10, 20, 30)" data-is-csv-numbers="1" data-slide-idx="${slideIdx}" data-path="${path}.${i}.values" value="${escapeHtml(valuesStr)}">
      <button type="button" class="icon-btn danger" data-action="remove-row" data-slide-idx="${slideIdx}" data-path="${path}" data-row="${i}" title="Quitar serie">${ICONS.trash}</button>
    </div>`;
  });
  html += `<button type="button" class="btn btn-sm" data-action="add-series-row" data-slide-idx="${slideIdx}" data-path="${path}">${ICONS.plus} Agregar serie</button>`;
  html += "</div></div>";
  return html;
}

function renderField(field, slide, slideIdx) {
  const val = slide[field.name];
  switch (field.type) {
    case "text": return textRow(field, slideIdx, field.name, val ?? "");
    case "number": return numberRow(field, slideIdx, field.name, val);
    case "select": return selectRow(field, slideIdx, field.name, val);
    case "list": return listRow(field.label, field.required, slideIdx, field.name, val || []);
    case "bullet_column": return bulletColumnRow(field, slideIdx, field.name, val);
    case "pricing_rows": return pricingRowsField(field, slideIdx, field.name, val);
    case "chart_series": return chartSeriesField(field, slideIdx, field.name, val);
    default: return "";
  }
}

// ───────────────────────────── panel de propiedades (1 slide enfocado) ────

// Campos presentes en los datos del slide que NO estan en el esquema estatico
// (ej. una forma personalizada nombrada "promo_code" en el editor visual,
// que apply_visual_config.py si incluye en input_template.yaml pero que
// FIELD_SCHEMAS no conoce de antemano). Sin esto, esos valores se cargarian
// en memoria pero quedarian invisibles/inaccesibles en el formulario.
function extraFieldNames(slide, schema) {
  const known = new Set(["type", ...schema.map((f) => f.name)]);
  return Object.keys(slide).filter((k) => !known.has(k));
}

function extraFieldRow(slide, slideIdx, key) {
  const val = slide[key];
  if (val !== null && typeof val === "object") {
    return `<div class="f-row f-row-dynamic">
      <label>${escapeHtml(key)} <span class="badge-dynamic">campo dinamico (estructura compleja)</span></label>
      <input type="text" value="${escapeHtml(JSON.stringify(val))}" disabled>
    </div>`;
  }
  return `<div class="f-row f-row-dynamic">
    <label>${escapeHtml(key)} <span class="badge-dynamic">campo dinamico</span></label>
    <input type="text" data-slide-idx="${slideIdx}" data-path="${escapeHtml(key)}" value="${escapeHtml(val ?? "")}">
  </div>`;
}

function renderPropertiesPanel() {
  const root = document.getElementById("slidesList");
  if (!state.slides.length) {
    root.innerHTML = "";
    return;
  }
  const idx = state.focusedIdx;
  const slide = state.slides[idx];
  const schema = state.schema[slide.type] || [];
  const m = meta(slide.type);
  const typeOpts = state.slideTypes.map((t) => `<option value="${t}" ${t === slide.type ? "selected" : ""}>${meta(t).label}</option>`).join("");
  const extraKeys = extraFieldNames(slide, schema);
  const visibleFields = schema.filter((field) => !field.hidden);
  const hiddenCount = schema.length - visibleFields.length;

  let bodyHtml = visibleFields.map((field) => renderField(field, slide, idx)).join("")
    + extraKeys.map((key) => extraFieldRow(slide, idx, key)).join("");
  if (!visibleFields.length && !extraKeys.length) {
    bodyHtml = `<p class="no-visible-fields">
      Ningun campo de este slide esta visible en tu editor visual actual
      (los ${hiddenCount} campo(s) de ${escapeHtml(m.label)} estan ocultos/reemplazados por tu propio diseño).
      Abre el editor visual para agregar o mostrar elementos de este tipo.
    </p>`;
  }

  root.innerHTML = `
    <div class="properties-header">
      <span class="properties-type-badge" style="background:${m.color}">${m.short}</span>
      <div class="properties-title">
        <span class="properties-num">Slide ${idx + 1} de ${state.slides.length}</span>
        <span class="properties-type-name">${escapeHtml(m.label)}</span>
      </div>
      <select data-role="slide-type-select" data-slide-idx="${idx}">${typeOpts}</select>
    </div>
    <div class="properties-body">${bodyHtml}</div>`;
}

// ───────────────────────────── lista de slides (panel izquierdo) ──────────

function slideMatchesSearch(slide, query) {
  if (!query) return true;
  const haystack = (meta(slide.type).label + " " + JSON.stringify(slide)).toLowerCase();
  return haystack.includes(query.toLowerCase());
}

function renderSlideListPanel() {
  const panel = document.getElementById("slideListPanel");
  panel.innerHTML = state.slides.map((s, i) => {
    if (!slideMatchesSearch(s, state.searchQuery)) return "";
    const m = meta(s.type);
    const active = i === state.focusedIdx;
    return `<div class="slide-list-item${active ? " active" : ""}" data-list-idx="${i}">
      <span class="slide-list-num">${i + 1}</span>
      <span class="slide-list-dot" style="background:${m.color}"></span>
      <span class="slide-list-label">${escapeHtml(m.label)}</span>
      <div class="slide-list-actions">
        <button type="button" class="icon-btn" data-action="move-up" data-slide-idx="${i}" title="Subir">${ICONS.up}</button>
        <button type="button" class="icon-btn" data-action="move-down" data-slide-idx="${i}" title="Bajar">${ICONS.down}</button>
        <button type="button" class="icon-btn accent" data-action="duplicate-slide" data-slide-idx="${i}" title="Duplicar">${ICONS.copy}</button>
        <button type="button" class="icon-btn danger" data-action="delete-slide" data-slide-idx="${i}" title="Eliminar">${ICONS.trash}</button>
      </div>
    </div>`;
  }).join("");

  panel.querySelectorAll("[data-list-idx]").forEach((item) => {
    item.addEventListener("click", (e) => {
      if (e.target.closest("[data-action]")) return; // los botones de accion manejan su propio click
      state.focusedIdx = parseInt(item.dataset.listIdx, 10);
      renderSlides();
    });
  });
}

function updateDeckProgress() {
  const el = document.getElementById("deckProgress");
  const n = state.slides.length;
  el.textContent = n ? `${n} slide${n === 1 ? "" : "s"}` : "vacío";
}

// ───────────────────────────── canvas (vista previa en vivo) ──────────────

async function getCanvasLayout(type) {
  if (!(type in state.canvasLayoutCache)) {
    try {
      const res = await fetch(`/api/canvas-layout/${encodeURIComponent(type)}`);
      state.canvasLayoutCache[type] = res.ok ? await res.json() : {};
    } catch {
      state.canvasLayoutCache[type] = {};
    }
  }
  return state.canvasLayoutCache[type];
}

function canvasElementHtml(el, slide) {
  const leftPct = (el.left / CANVAS_W_IN) * 100;
  const topPct = (el.top / CANVAS_H_IN) * 100;
  const widthPct = (el.width / CANVAS_W_IN) * 100;
  const heightPct = (el.height / CANVAS_H_IN) * 100;
  const posStyle = `left:${leftPct}%;top:${topPct}%;width:${widthPct}%;height:${heightPct}%;`;

  const nameForGuess = (el.name || "").toLowerCase();
  if (el.kind === "image" || /photo|logo/.test(nameForGuess)) {
    return `<div class="canvas-el canvas-el--image" style="${posStyle}">${ICONS.image}</div>`;
  }

  const looksDecorative = !el.kind && !("font_size" in el) && !(el.fields && el.fields.length) && !el.text;
  if (looksDecorative) {
    const bg = el.fill ? `#${el.fill}` : "rgba(0,0,0,0.08)";
    return `<div class="canvas-el canvas-el--box" style="${posStyle}background:${bg}"></div>`;
  }

  let text = el.text || "";
  let isPlaceholder = !text;
  if (el.fields && el.fields.length) {
    const liveVal = slide[el.fields[0]];
    if (typeof liveVal === "string" && liveVal.trim()) {
      text = liveVal;
      isPlaceholder = false;
    } else if (!text) {
      text = `(${el.fields[0]})`;
    }
  }

  const fontPx = el.font_size ? Math.max(5, (el.font_size * (CANVAS_BASE_PX / CANVAS_W_IN)) / 72) : 10;
  const color = el.color ? `#${el.color}` : "#1a1a1a";
  const bg = el.fill ? `background:#${el.fill};padding:2px 4px;` : "";
  const fontWeight = el.bold ? "font-weight:700;" : "";
  const cls = `canvas-el canvas-el--text${isPlaceholder ? " canvas-el--placeholder" : ""}`;
  return `<div class="${cls}" style="${posStyle}font-size:${fontPx}px;color:${color};${fontWeight}${bg}">${escapeHtml(text)}</div>`;
}

async function renderCanvasPreview() {
  const viewport = document.getElementById("canvasViewport");
  if (!state.slides.length) {
    viewport.innerHTML = `<p class="canvas-empty">Agrega un slide para ver la vista previa.</p>`;
    return;
  }
  const slide = state.slides[state.focusedIdx];
  const layout = await getCanvasLayout(slide.type);
  const names = Object.keys(layout);

  const slideDiv = document.createElement("div");
  slideDiv.className = "canvas-slide";
  slideDiv.style.transform = `scale(${state.zoom})`;
  if (!names.length) {
    slideDiv.innerHTML = `<p class="canvas-empty" style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;">
      Este tipo de slide no tiene elementos visibles en tu editor visual actual.
    </p>`;
  } else {
    slideDiv.innerHTML = names.map((name) => canvasElementHtml({ ...layout[name], name }, slide)).join("");
  }

  // Si el slide enfocado no cambio (solo el contenido), reemplazar in-place
  // evita un parpadeo visible en cada tecla.
  viewport.innerHTML = "";
  viewport.appendChild(slideDiv);
}

function setZoom(value) {
  state.zoom = Math.min(2, Math.max(0.4, value));
  document.getElementById("zoomLabel").textContent = `${Math.round(state.zoom * 100)}%`;
  const slideEl = document.querySelector(".canvas-slide");
  if (slideEl) slideEl.style.transform = `scale(${state.zoom})`;
}

// ───────────────────────────── orquestador de render ───────────────────────

function renderSlides() {
  if (state.focusedIdx > state.slides.length - 1) state.focusedIdx = Math.max(0, state.slides.length - 1);

  if (!state.slides.length) {
    document.getElementById("slideListPanel").innerHTML = "";
    document.getElementById("slidesList").innerHTML = `<div class="empty-state">
      <h3>Aun no hay slides</h3>
      <p>Carga un archivo existente arriba o agrega tu primer slide para empezar.</p>
      <button class="btn btn-primary" id="emptyAddSlideBtn">+ Agregar slide</button>
    </div>`;
    const btn = document.getElementById("emptyAddSlideBtn");
    if (btn) btn.addEventListener("click", addSlide);
    renderCanvasPreview();
    updateDeckProgress();
    return;
  }

  renderSlideListPanel();
  renderPropertiesPanel();
  renderCanvasPreview();
  updateDeckProgress();
}

function addSlide() {
  state.slides.push(defaultSlideForType(state.slideTypes[0]));
  state.focusedIdx = state.slides.length - 1;
  renderSlides();
}

function duplicateSlide(idx) {
  const clone = JSON.parse(JSON.stringify(state.slides[idx]));
  state.slides.splice(idx + 1, 0, clone);
  state.focusedIdx = idx + 1;
  renderSlides();
}

// ───────────────────────────── eventos delegados ──────────────────────────

function initDelegatedEvents() {
  const propertiesContainer = document.getElementById("slidesList");

  // Edicion de campos simples: NO se re-renderiza el formulario (evita perder
  // el foco mientras se escribe) — solo se actualiza el estado en memoria,
  // los indicadores visuales (contador de caracteres, suma de precios), y el
  // canvas (que es un arbol DOM separado, sin riesgo de robar el foco).
  propertiesContainer.addEventListener("input", (e) => {
    const target = e.target;
    if (!target.dataset.path) return;
    const slideIdx = parseInt(target.dataset.slideIdx, 10);
    const slide = state.slides[slideIdx];
    let value = target.value;

    if (target.dataset.isCsvNumbers) {
      value = value.split(",").map((s) => parseFloat(s.trim())).filter((n) => !Number.isNaN(n));
    } else if (target.type === "number") {
      value = value === "" ? null : parseFloat(value);
    }
    setPath(slide, target.dataset.path, value);

    if (target.dataset.required === "true") {
      target.classList.toggle("invalid", !String(value || "").trim());
    }

    const fill = target.closest(".f-row")?.querySelector(".char-meter-fill");
    const text = target.closest(".f-row")?.querySelector(".char-meter-text");
    if (fill && text) {
      const max = parseInt(fill.dataset.max, 10);
      const len = target.value.length;
      const over = len > max;
      fill.style.width = `${Math.min(100, (len / max) * 100)}%`;
      fill.classList.toggle("over", over);
      text.textContent = `${len}/${max}`;
      text.classList.toggle("over", over);
    }

    const pricingRows = target.closest(".pricing-rows");
    if (pricingRows) {
      const foot = pricingRows.parentElement.querySelector('[data-role="pricing-foot"]');
      const rows = getPath(slide, target.dataset.path.split(".").slice(0, -2).join(".")) || [];
      const computed = rows.reduce((sum, r) => sum + (Number(r.total) || 0), 0);
      if (foot) foot.innerHTML = `Suma de filas: <strong>${computed.toFixed(2)}</strong>`;
    }

    renderCanvasPreview();
  });

  // Cambios estructurales (agregar/quitar filas, cambiar tipo) que si
  // requieren reconstruir el panel de propiedades.
  propertiesContainer.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-action]");
    if (!btn) return;
    const action = btn.dataset.action;
    const slideIdx = parseInt(btn.dataset.slideIdx, 10);
    const path = btn.dataset.path;
    const slide = state.slides[slideIdx];

    if (action === "add-list-item") {
      const arr = getPath(slide, path) || [];
      arr.push("");
      setPath(slide, path, arr);
      renderSlides();
    } else if (action === "add-pricing-row") {
      const arr = getPath(slide, path) || [];
      arr.push({ description: "", quantity: 1, unit: "", unit_price: 0, total: 0 });
      setPath(slide, path, arr);
      renderSlides();
    } else if (action === "add-series-row") {
      const arr = getPath(slide, path) || [];
      arr.push({ name: "", values: [] });
      setPath(slide, path, arr);
      renderSlides();
    } else if (action === "remove-row") {
      const row = parseInt(btn.dataset.row, 10);
      const arr = getPath(slide, path) || [];
      arr.splice(row, 1);
      setPath(slide, path, arr);
      renderSlides();
    }
  });

  propertiesContainer.addEventListener("change", (e) => {
    if (e.target.dataset.role !== "slide-type-select") return;
    const slideIdx = parseInt(e.target.dataset.slideIdx, 10);
    const newType = e.target.value;
    if (newType === state.slides[slideIdx].type) return;
    showConfirm({
      title: "Cambiar tipo de slide",
      body: "Esto reinicia todos los campos de este slide a un estado vacio. ¿Continuar?",
      confirmText: "Cambiar tipo",
      danger: true,
    }).then((confirmed) => {
      if (confirmed) {
        state.slides[slideIdx] = defaultSlideForType(newType);
        renderSlides();
      } else {
        e.target.value = state.slides[slideIdx].type;
      }
    });
  });

  // Acciones de la lista de slides (panel izquierdo): mover, duplicar, eliminar.
  document.getElementById("slideListPanel").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-action]");
    if (!btn) return;
    const action = btn.dataset.action;
    const slideIdx = parseInt(btn.dataset.slideIdx, 10);
    const slide = state.slides[slideIdx];

    if (action === "duplicate-slide") {
      duplicateSlide(slideIdx);
    } else if (action === "delete-slide") {
      showConfirm({
        title: "Eliminar slide",
        body: `¿Eliminar el slide ${slideIdx + 1} (${meta(slide.type).label}) de este deck?`,
        confirmText: "Eliminar",
        danger: true,
      }).then((confirmed) => {
        if (!confirmed) return;
        state.slides.splice(slideIdx, 1);
        if (state.focusedIdx >= slideIdx) state.focusedIdx = Math.max(0, state.focusedIdx - 1);
        renderSlides();
      });
    } else if (action === "move-up") {
      if (slideIdx === 0) return;
      [state.slides[slideIdx - 1], state.slides[slideIdx]] = [state.slides[slideIdx], state.slides[slideIdx - 1]];
      if (state.focusedIdx === slideIdx) state.focusedIdx -= 1;
      else if (state.focusedIdx === slideIdx - 1) state.focusedIdx += 1;
      renderSlides();
    } else if (action === "move-down") {
      if (slideIdx === state.slides.length - 1) return;
      [state.slides[slideIdx + 1], state.slides[slideIdx]] = [state.slides[slideIdx], state.slides[slideIdx + 1]];
      if (state.focusedIdx === slideIdx) state.focusedIdx += 1;
      else if (state.focusedIdx === slideIdx + 1) state.focusedIdx -= 1;
      renderSlides();
    }
  });

  document.getElementById("searchSlides").addEventListener("input", (e) => {
    state.searchQuery = e.target.value;
    renderSlideListPanel();
  });

  document.getElementById("zoomInBtn").addEventListener("click", () => setZoom(state.zoom + 0.1));
  document.getElementById("zoomOutBtn").addEventListener("click", () => setZoom(state.zoom - 0.1));
}

// ───────────────────────────── carga / guardado de archivos ──────────────

async function loadDeckFilesList() {
  const res = await fetch("/api/deck-files");
  const files = await res.json();
  const sel = document.getElementById("deckFileSelect");
  const current = sel.value;
  sel.innerHTML = '<option value="">-- elegir archivo --</option>'
    + files.map((f) => {
      const label = f.startsWith("generated/") ? `⚙ Plantilla auto-generada (${f})` : f;
      return `<option value="${f}">${label}</option>`;
    }).join("");
  if (current && files.includes(current)) sel.value = current;
  return files;
}

async function loadDeckFile(filename) {
  if (!filename) return;
  const res = await fetch(`/api/deck-data?file=${encodeURIComponent(filename)}`);
  const data = await res.json();
  if (!data.slides) {
    showToast("Error: " + escapeHtml(data.error || "no encontrado"), "err");
    return;
  }

  state.slides = data.slides;
  state.focusedIdx = 0;

  // "generated/input_template.yaml" es auto-generada por 'Aplicar cambios del
  // editor visual' y se sobreescribe en cada corrida — NUNCA debe guardarse
  // de vuelta con ese mismo nombre, o se crea un examples/input_template.yaml
  // confuso (mismo nombre, contenido distinto, y la plantilla real nunca se
  // actualiza). Forzamos al usuario a elegir un nombre real para su deck.
  const filenameInput = document.getElementById("deckFilenameInput");
  const isGenerated = filename.startsWith("generated/");
  filenameInput.value = isGenerated ? "" : filename.split("/").pop();
  filenameInput.placeholder = isGenerated ? "ej: propuesta_cliente_acme.yaml" : "mi_deck.yaml";

  renderSlides();

  if (isGenerated) {
    showToast(
      `Cargada la plantilla auto-generada (${data.slides.length} slide(s) de muestra). ` +
      `Es solo punto de partida — escribe un nombre nuevo en "Guardar como" antes de guardar tus datos reales.`,
      "ok", 8000,
    );
  } else {
    showToast(`Cargado <strong>${escapeHtml(filename)}</strong> — ${data.slides.length} slide(s)`, "ok");
  }
}

async function saveDeckData() {
  const filename = document.getElementById("deckFilenameInput").value.trim();
  if (!filename) {
    showToast("Escribe un nombre de archivo primero.", "err");
    return;
  }
  try {
    const res = await fetch("/api/deck-data", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename, slides: state.slides }),
    });
    const data = await res.json();
    if (data.ok) {
      showToast(`Guardado como <strong>examples/${escapeHtml(data.filename)}</strong>`, "ok");
      await loadDeckFilesList();
    } else {
      showToast("Error: " + escapeHtml(data.error), "err");
    }
  } catch (err) {
    showToast("Error de red: " + escapeHtml(err.message), "err");
  }
}

async function renderDeck() {
  const filenameInput = document.getElementById("deckFilenameInput").value.trim() || "preview";
  const outputName = filenameInput.replace(/\.ya?ml$/i, "");
  const btn = document.getElementById("renderDeckBtn");
  const label = btn.querySelector(".btn-label");
  const originalLabel = label.textContent;

  btn.disabled = true;
  label.innerHTML = '<span class="spinner"></span> Generando...';

  try {
    const res = await fetch("/api/render", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ slides: state.slides, output_name: outputName }),
    });
    const data = await res.json();
    if (data.ok) {
      showToast(`Deck generado: <a href="${escapeHtml(data.download_url)}">descargar .pptx</a>`, "ok", 9000);
    } else {
      showToast("Error: " + escapeHtml(data.error), "err", 9000);
    }
  } catch (err) {
    showToast("Error de red: " + escapeHtml(err.message), "err");
  } finally {
    btn.disabled = false;
    label.textContent = originalLabel;
  }
}

// ───────────────────────────── init ────────────────────────────────────────

async function init() {
  const [typesRes, schemaRes, limitsRes] = await Promise.all([
    fetch("/api/slide-types"),
    fetch("/api/schema"),
    fetch("/api/limits"),
  ]);
  state.slideTypes = await typesRes.json();
  state.schema = await schemaRes.json();
  state.limits = await limitsRes.json();

  initDelegatedEvents();

  const sel = document.getElementById("deckFileSelect");
  sel.addEventListener("change", (e) => loadDeckFile(e.target.value));

  const files = await loadDeckFilesList();
  const latest = files.find((f) => f.startsWith("generated/"));
  if (latest) {
    sel.value = latest;
    await loadDeckFile(latest);
  } else if (files.length) {
    sel.value = files[0];
    await loadDeckFile(files[0]);
  } else {
    state.slides = [];
    renderSlides();
  }

  document.getElementById("newDeckBtn").addEventListener("click", () => {
    showConfirm({
      title: "Nuevo deck",
      body: "Esto reemplaza el contenido actual del formulario (sin guardar). ¿Continuar?",
    }).then((confirmed) => {
      if (!confirmed) return;
      state.slides = [defaultSlideForType(state.slideTypes[0])];
      state.focusedIdx = 0;
      document.getElementById("deckFilenameInput").value = "";
      document.getElementById("deckFileSelect").value = "";
      renderSlides();
    });
  });
  document.getElementById("addSlideBtn").addEventListener("click", addSlide);
  document.getElementById("saveDeckBtn").addEventListener("click", saveDeckData);
  document.getElementById("renderDeckBtn").addEventListener("click", renderDeck);
}

init();
