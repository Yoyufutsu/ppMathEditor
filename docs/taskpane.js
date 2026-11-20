/* ppMath 数式エディタ — PowerPoint 作業ウィンドウ */
"use strict";

// ---- 定数 -------------------------------------------------------------------
const TAG = {
  LATEX: "PPMATH_LATEX",
  SIZE: "PPMATH_SIZE",
  COLOR: "PPMATH_COLOR",
  DISPLAY: "PPMATH_DISPLAY",
  H: "PPMATH_H", // 挿入時の高さ(pt)。手動で拡大縮小された倍率を更新時に保つため
};
const EM_PER_EX = 0.442; // MathJax TeX フォントの x-height
const SHAPE_NAME_PREFIX = "ppMath: ";

// ---- 状態 -------------------------------------------------------------------
const state = {
  latex: "",
  inOffice: false,
  apiOk: false,
  busy: false,
  dirty: false,       // 編集対象を読み込んでから変更したか
  editing: null,      // { slideId, shapeId }
  rendered: null,     // 最新のレンダリング結果
  selSeq: 0,
};

const $ = (id) => document.getElementById(id);
const els = {};

// ---- 起動 -------------------------------------------------------------------
window.addEventListener("DOMContentLoaded", async () => {
  for (const id of ["src", "preview", "error", "fontSize", "color", "display", "insertBtn", "updateBtn",
    "newBtn", "status", "eqList", "refreshBtn", "macros", "modeBadge", "modeText", "autoUpdate"]) {
    els[id] = $(id);
  }
  loadPrefs();

  await MathJax.startup.promise;
  bindUi();
  render();

  const info = await officeReady();
  state.inOffice = info && info.host === Office.HostType.PowerPoint;
  if (!state.inOffice) {
    setStatus("PowerPoint の外で開いています（プレビューのみ）", false);
  } else {
    state.apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
    if (!state.apiOk) {
      setStatus("この PowerPoint では編集機能 (PowerPointApi 1.5) が使えません。Microsoft 365 版を更新してください。", true);
    }
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    onSelectionChanged();
    refreshList();
  }
  updateButtons();
});

function officeReady() {
  if (typeof Office === "undefined") return Promise.resolve(null);
  return Promise.race([
    Office.onReady(),
    new Promise((r) => setTimeout(() => r(null), 4000)),
  ]);
}

// ---- 設定の保存 --------------------------------------------------------------
function loadPrefs() {
  try {
    const p = JSON.parse(localStorage.getItem("ppmath.prefs") || "{}");
    if (p.fontSize) els.fontSize.value = p.fontSize;
    if (p.color) els.color.value = p.color;
    if (typeof p.display === "boolean") els.display.checked = p.display;
    if (typeof p.autoUpdate === "boolean") els.autoUpdate.checked = p.autoUpdate;
    els.macros.value = localStorage.getItem("ppmath.macros") || "";
  } catch (_) { /* storage unavailable */ }
}
function savePrefs() {
  try {
    localStorage.setItem("ppmath.prefs", JSON.stringify({
      fontSize: Number(els.fontSize.value), color: els.color.value, display: els.display.checked,
      autoUpdate: els.autoUpdate.checked,
    }));
    localStorage.setItem("ppmath.macros", els.macros.value);
  } catch (_) { /* storage unavailable */ }
}

// ---- 入力欄 ------------------------------------------------------------------
function setLatex(latex) {
  state.latex = latex;
  els.src.value = latex;
  render();
}

function onContentChanged() {
  if (state.editing) state.dirty = true;
  render();
  scheduleAuto();
}

// 編集中の数式は、入力が止まって少ししたら自動でスライドに反映する
const AUTO_DELAY_MS = 600;
let autoTimer = null;
function scheduleAuto() {
  clearTimeout(autoTimer);
  if (!state.editing || !els.autoUpdate.checked) return;
  autoTimer = setTimeout(() => {
    if (!state.editing || !state.dirty) return;
    if (state.busy) {
      state.autoPending = true; // 実行中の更新が終わったら再実行
      return;
    }
    commit({ auto: true });
  }, AUTO_DELAY_MS);
}

function contentKey() {
  return JSON.stringify([state.latex, currentOptions(), els.macros.value]);
}

function bindUi() {
  els.src.addEventListener("input", () => {
    state.latex = els.src.value;
    onContentChanged();
  });

  for (const el of [els.fontSize, els.color, els.display]) {
    el.addEventListener("input", () => { savePrefs(); onContentChanged(); });
  }
  els.macros.addEventListener("input", () => { savePrefs(); onContentChanged(); });
  els.autoUpdate.addEventListener("change", () => { savePrefs(); updateButtons(); scheduleAuto(); });

  els.insertBtn.addEventListener("click", () => commit());
  els.updateBtn.addEventListener("click", () => commit());
  els.newBtn.addEventListener("click", () => { exitEdit(); setLatex(""); els.src.focus(); });
  els.refreshBtn.addEventListener("click", (e) => { e.preventDefault(); refreshList(); });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      e.stopPropagation();
      commit();
    }
  }, true);
}

// ---- レンダリング --------------------------------------------------------------
function texWithMacros(latex) {
  const macros = els.macros.value.trim();
  return macros ? macros + "\n" + latex : latex;
}

/** LaTeX を SVG にして、PowerPoint へ入れる SVG 文字列と寸法(pt)を返す */
function renderEquation(latex, { fontSize, color, display }) {
  if (!latex.trim()) return { empty: true };
  const node = MathJax.tex2svg(texWithMacros(latex), { display });
  const svg = node.querySelector("svg");
  const err = svg.querySelector('[data-mml-node="merror"]');
  if (err) {
    return { error: err.getAttribute("data-mjx-error") || err.getAttribute("title") || "LaTeX の構文エラー" };
  }
  const wEx = parseFloat(svg.getAttribute("width"));
  const hEx = parseFloat(svg.getAttribute("height"));
  const widthPt = wEx * EM_PER_EX * fontSize;
  const heightPt = hEx * EM_PER_EX * fontSize;

  const out = svg.cloneNode(true);
  for (const a of ["style", "role", "focusable", "aria-hidden"]) out.removeAttribute(a);
  out.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  out.setAttribute("width", (widthPt * 4 / 3).toFixed(2)); // px
  out.setAttribute("height", (heightPt * 4 / 3).toFixed(2));
  const svgText = new XMLSerializer().serializeToString(out).replace(/currentColor/g, color);

  return { svg, svgText, widthPt, heightPt };
}

function currentOptions() {
  return {
    fontSize: Math.max(1, Number(els.fontSize.value) || 24),
    color: els.color.value || "#000000",
    display: els.display.checked,
  };
}

function render() {
  const opts = currentOptions();
  let r;
  try {
    r = renderEquation(state.latex, opts);
  } catch (e) {
    r = { error: String(e.message || e) };
  }
  state.rendered = r;
  els.preview.innerHTML = "";
  els.error.hidden = true;
  if (r.error) {
    els.error.textContent = r.error;
    els.error.hidden = false;
  } else if (r.svg) {
    const shown = r.svg.cloneNode(true);
    shown.style.color = opts.color;
    els.preview.appendChild(shown);
  } else {
    els.preview.innerHTML = '<span class="muted small">ここにプレビューが表示されます</span>';
  }
  updateButtons();
}

// ---- モード表示 ----------------------------------------------------------------
function updateButtons() {
  const ok = !!(state.rendered && state.rendered.svgText) && state.inOffice && state.apiOk && !state.busy;
  els.insertBtn.disabled = !ok;
  els.updateBtn.disabled = !ok;
  els.insertBtn.hidden = !!state.editing;
  els.updateBtn.hidden = !state.editing;
  els.modeBadge.textContent = state.editing ? "編集中" : "新規";
  els.modeBadge.className = "badge " + (state.editing ? "edit" : "new");
  els.modeText.textContent = state.editing
    ? `スライド ${state.editing.slideIndex + 1} の数式`
    : "スライド上の数式を選択すると編集できます";
  for (const li of els.eqList.children) {
    li.classList.toggle("current", !!state.editing && li.dataset.shapeId === state.editing.shapeId);
  }
}

function enterEdit(info) {
  state.editing = { slideId: info.slideId, shapeId: info.shapeId, slideIndex: info.slideIndex };
  if (info.size) els.fontSize.value = info.size;
  if (info.color) els.color.value = info.color;
  if (info.display != null) els.display.checked = info.display;
  setLatex(info.latex);
  state.dirty = false;
  updateButtons();
}

function exitEdit() {
  state.editing = null;
  state.dirty = false;
  updateButtons();
}

function setStatus(msg, isError) {
  els.status.textContent = msg || "";
  els.status.classList.toggle("err", !!isError);
}

// ---- PowerPoint 連携 -----------------------------------------------------------
function readTags(ctx, shape) {
  const t = {};
  for (const [k, key] of Object.entries(TAG)) {
    t[k] = shape.tags.getItemOrNullObject(key);
    t[k].load("value");
  }
  return t;
}

function tagsToInfo(t) {
  if (t.LATEX.isNullObject) return null;
  const v = (k) => (t[k].isNullObject ? null : t[k].value);
  return {
    latex: v("LATEX"),
    size: v("SIZE") ? Number(v("SIZE")) : null,
    color: v("COLOR"),
    display: v("DISPLAY") == null ? null : v("DISPLAY") === "1",
    storedH: v("H") ? Number(v("H")) : null,
  };
}

async function onSelectionChanged() {
  if (!state.apiOk || state.busy) return;
  const seq = ++state.selSeq;
  try {
    const info = await PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items/id");
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items/id");
      const all = ctx.presentation.slides;
      all.load("items/id");
      await ctx.sync();
      if (shapes.items.length !== 1 || slides.items.length < 1) return null;
      const t = readTags(ctx, shapes.items[0]);
      await ctx.sync();
      const base = tagsToInfo(t);
      if (!base) return null;
      const slideId = slides.items[0].id;
      return { ...base, shapeId: shapes.items[0].id, slideId, slideIndex: all.items.findIndex((s) => s.id === slideId) };
    });
    if (seq !== state.selSeq || state.busy) return;
    if (info) {
      // 同じ数式を選び直しただけなら入力中の内容は触らない
      if (!state.editing || state.editing.shapeId !== info.shapeId) enterEdit(info);
    } else if (state.editing && !state.dirty) {
      exitEdit();
    }
  } catch (e) {
    console.warn("selection read failed", e);
  }
}

function setSelectedSvg(svgText) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(svgText, { coercionType: Office.CoercionType.XmlSvg }, (res) => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

async function getShapeIds(slideId) {
  return PowerPoint.run(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    await ctx.sync();
    return new Set(shapes.items.map((s) => s.id));
  });
}

/** 挿入 or 更新 */
async function commit({ auto = false } = {}) {
  const r = state.rendered;
  if (state.busy || !state.inOffice || !state.apiOk || !r || !r.svgText) return;
  if (auto && !state.editing) return;
  const opts = currentOptions();
  const latex = state.latex;
  const key = contentKey();
  // 入力中のカーソル位置を、更新後に戻すため覚えておく
  const hadFocus = document.activeElement === els.src;
  const caret = [els.src.selectionStart, els.src.selectionEnd];
  state.busy = true;
  updateButtons();
  if (!auto) setStatus(state.editing ? "更新中…" : "挿入中…");
  try {
    // 1. 対象スライドと、更新の場合は元の図形の位置・倍率を取得
    const target = await PowerPoint.run(async (ctx) => {
      let slideId, slideIndex, old = null;
      const all = ctx.presentation.slides;
      all.load("items/id");
      if (state.editing) {
        slideId = state.editing.slideId;
        const shape = ctx.presentation.slides.getItem(slideId).shapes.getItemOrNullObject(state.editing.shapeId);
        shape.load("left,top,width,height");
        const t = readTags(ctx, shape);
        await ctx.sync();
        if (shape.isNullObject) throw new Error("編集中の数式が見つかりません（削除されましたか？）。「新規」で挿入してください。");
        const info = tagsToInfo(t) || {};
        const scale = info.storedH ? shape.height / info.storedH : 1;
        old = { id: state.editing.shapeId, left: shape.left, top: shape.top, scale: isFinite(scale) && scale > 0 ? scale : 1 };
        ctx.presentation.setSelectedSlides([slideId]);
      } else {
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        if (!sel.items.length) throw new Error("スライドを選択してください。");
        slideId = sel.items[0].id;
      }
      // 挿入時に既存図形が置き換わらないよう選択解除
      ctx.presentation.slides.getItem(slideId).setSelectedShapes([]);
      await ctx.sync();
      slideIndex = all.items.findIndex((s) => s.id === slideId);
      return { slideId, slideIndex, old };
    });

    // 2. SVG を挿入し、新しくできた図形を特定
    const before = await getShapeIds(target.slideId);
    await setSelectedSvg(r.svgText);
    const after = await getShapeIds(target.slideId);
    const newId = [...after].find((id) => !before.has(id));
    if (!newId) throw new Error("挿入した図形を特定できませんでした。");

    // 3. 位置・サイズ・タグを設定し、古い図形を削除
    const scale = target.old ? target.old.scale : 1;
    const w = r.widthPt * scale;
    const h = r.heightPt * scale;
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItem(target.slideId);
      const shape = slide.shapes.getItem(newId);
      shape.load("left,top,width,height");
      await ctx.sync();
      if (target.old) {
        shape.left = target.old.left;
        shape.top = target.old.top;
      } else {
        // PowerPoint が置いた位置の中心を保つ
        shape.left = shape.left + (shape.width - w) / 2;
        shape.top = shape.top + (shape.height - h) / 2;
      }
      shape.width = w;
      shape.height = h;
      shape.name = SHAPE_NAME_PREFIX + latex.replace(/\s+/g, " ").slice(0, 60);
      shape.tags.add(TAG.LATEX, latex);
      shape.tags.add(TAG.SIZE, String(opts.fontSize));
      shape.tags.add(TAG.COLOR, opts.color);
      shape.tags.add(TAG.DISPLAY, opts.display ? "1" : "0");
      shape.tags.add(TAG.H, String(h));
      if (target.old) slide.shapes.getItem(target.old.id).delete();
      await ctx.sync();
      slide.setSelectedShapes([newId]);
      await ctx.sync();
    });

    const wasEditing = !!target.old;
    state.editing = { slideId: target.slideId, shapeId: newId, slideIndex: target.slideIndex };
    // 更新中にさらに入力されていれば、まだ未反映
    state.dirty = contentKey() !== key;
    setStatus(auto ? "スライドに反映しました" : wasEditing ? "数式を更新しました" : "数式を挿入しました");
  } catch (e) {
    console.error(e);
    setStatus("失敗: " + (e.message || e), true);
  } finally {
    state.busy = false;
    updateButtons();
    refreshList();
    if (hadFocus && document.activeElement !== els.src) {
      els.src.focus();
      els.src.setSelectionRange(caret[0], caret[1]);
    }
    if (state.autoPending || state.dirty) {
      state.autoPending = false;
      scheduleAuto();
    }
  }
}

// ---- 数式一覧 ------------------------------------------------------------------
async function refreshList() {
  if (!state.apiOk) {
    els.eqList.innerHTML = '<li class="empty">PowerPoint 内でのみ利用できます</li>';
    return;
  }
  try {
    const entries = await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const colls = slides.items.map((s) => {
        const c = s.shapes;
        c.load("items/id");
        return c;
      });
      await ctx.sync();
      const found = [];
      colls.forEach((c, i) => c.items.forEach((sh) => {
        found.push({ slideIndex: i, slideId: slides.items[i].id, shapeId: sh.id, t: readTags(ctx, sh) });
      }));
      await ctx.sync();
      return found
        .map((f) => ({ ...f, info: tagsToInfo(f.t) }))
        .filter((f) => f.info)
        .map(({ slideIndex, slideId, shapeId, info }) => ({ slideIndex, slideId, shapeId, ...info }));
    });
    drawList(entries);
  } catch (e) {
    console.warn(e);
    els.eqList.innerHTML = '<li class="empty">一覧を取得できませんでした</li>';
  }
}

function drawList(entries) {
  els.eqList.innerHTML = "";
  if (!entries.length) {
    els.eqList.innerHTML = '<li class="empty">まだ数式はありません</li>';
    return;
  }
  for (const e of entries) {
    const li = document.createElement("li");
    li.dataset.shapeId = e.shapeId;
    const no = document.createElement("span");
    no.className = "slide-no";
    no.textContent = "#" + (e.slideIndex + 1);
    const eq = document.createElement("span");
    eq.className = "eq";
    try {
      const r = renderEquation(e.latex, { fontSize: 12, color: "#000", display: false });
      if (r.svg) eq.appendChild(r.svg);
      else eq.textContent = e.latex;
    } catch (_) {
      eq.textContent = e.latex;
    }
    li.title = e.latex;
    li.append(no, eq);
    li.addEventListener("click", () => selectEquation(e));
    els.eqList.appendChild(li);
  }
  updateButtons();
}

async function selectEquation(e) {
  try {
    state.busy = true;
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.setSelectedSlides([e.slideId]);
      await ctx.sync();
      ctx.presentation.slides.getItem(e.slideId).setSelectedShapes([e.shapeId]);
      await ctx.sync();
    });
  } catch (err) {
    console.warn(err);
  } finally {
    state.busy = false;
  }
  enterEdit(e);
  els.src.focus();
}
