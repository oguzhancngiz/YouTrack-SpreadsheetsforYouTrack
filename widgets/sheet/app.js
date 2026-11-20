/*
 * Copyright 2026 Codeo
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/* Spreadsheets for YouTrack – the spreadsheet widget.
 *
 * mode-inline  ISSUE_ABOVE_ACTIVITY_STREAM: the spreadsheet in the issue, under the section header that YouTrack draws
 *              for the widget (YouTrack collapses and expands that section itself and removes the frame while it is
 *              collapsed). An issue without a spreadsheet shows one compact row.
 * mode-window  ISSUE_OPTIONS_MENU_ITEM: the same spreadsheet in a large window
 *
 * Everyone who opens the issue works on the same workbook: changes are saved as they are made and arrive at the
 * other users within a couple of seconds; their selections show up in their colors.
 */
import {
  Engine, parseInput, formatValue, editText, isErr, colName, parseAddr, addrName, formulaRefs, tokenize, translateFormula,
  FUNCTION_NAMES, FUNCTION_HELP, formulaSyntaxError, numToText, literal, localeInfo, quoteSheet, parseFormula, defaultCurrency,
  separatorStyle, formulaToLocal, formulaFromLocal, helpToLocal
} from './formula.js';
import { Store } from './sync.js';
import * as M from './model.js';
import { h, esc, icon, toast, menu, popover, closePopups, dialog, confirmDialog, colorPicker, contrastText, themedColors, paintBar, select, initials, peerColor, relTime } from './ui.js';
import * as D from './dialogs.js';
import * as W from './whatif.js';
import { renderChart, chartData } from './charts.js';
import * as IO from './io.js';

const Core = window.SheetCore;

/* ================================================================ host */

let host = null, ENTITY = null, LOCALE = null;
// @preview-start: outside YouTrack the development build runs on a simulated host (preview.js); build.py removes this
if (typeof window.YTApp === 'undefined') {
  try { host = await (await import('./preview.js')).createPreviewHost(); } catch (e) { host = null; }
  ENTITY = host ? host.entity : { id: '0-0', type: 'issue' };
  LOCALE = host && host.locale;
}
// @preview-end
if (typeof window.YTApp !== 'undefined') {
  host = await window.YTApp.register();
  ENTITY = window.YTApp.entity;
  LOCALE = window.YTApp.locale;
}
LOCALE = String(LOCALE || navigator.language || 'en-US').replace(/_/g, '-');
try { Intl.getCanonicalLocales(LOCALE); } catch (e) { LOCALE = navigator.language || 'en-US'; } // a code the browser does not know
const MODE = document.body.classList.contains('mode-window') ? 'window' : 'inline';
const CURRENCY = defaultCurrency(LOCALE);

/** Native controls follow the YouTrack theme (Ring UI variables), not the OS theme. */
function syncColorScheme() {
  const m = getComputedStyle(document.body).backgroundColor.match(/[\d.]+/g);
  let dark = false;
  if (m && m.length >= 3 && !(m.length > 3 && Number(m[3]) === 0)) {
    const [r, g, b] = m.map(Number);
    dark = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.5;
  }
  const scheme = dark ? 'dark' : 'light';
  const root = document.documentElement;
  if (root.style.colorScheme !== scheme) root.style.colorScheme = scheme;
  if (root.dataset.theme !== scheme) root.dataset.theme = scheme;
}
syncColorScheme();
{
  const obs = new MutationObserver(syncColorScheme);
  obs.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'style'] });
  obs.observe(document.body, { attributes: true, attributeFilter: ['class', 'style'] });
  obs.observe(document.head, { childList: true, subtree: true, characterData: true });
}

async function storageGet(key) { try { return host && host.storage ? await host.storage.getItem(key) : null; } catch (e) { return null; } }
async function storageSet(key, val) { try { if (host && host.storage) await host.storage.setItem(key, val); } catch (e) { /* ignore */ } }

function errorText(e) {
  if (!e) return 'Request failed';
  const d = e.data || e.body || e.response;
  if (d && d.error) return d.error;
  if (typeof d === 'string') { try { return JSON.parse(d).error || d; } catch (x) { return d; } }
  return e.message || String(e);
}
/** Calls the backend. Errors carry .status and .data. */
async function api(path, opts = {}) {
  if (!host) throw Object.assign(new Error('Not connected to YouTrack'), { status: 0 });
  try {
    const res = await host.fetchApp('backend/' + path, { scope: true, ...opts });
    if (res && res.ok === false && res.error) throw Object.assign(new Error(res.error), { status: res.status || 400, data: res });
    return res;
  } catch (e) {
    if (e.status !== undefined && e.data) throw e;
    let data = e.data || e.body || null;
    if (typeof data === 'string') { try { data = JSON.parse(data); } catch (x) { data = { error: data }; } }
    const status = e.status || e.statusCode || (data && data.status) || (e.response && e.response.status) || 0;
    throw Object.assign(new Error(errorText(e)), { status, data: data || {} });
  }
}

/* ================================================================ state */

const Z_ROW_H = 24, Z_COL_W = 100, Z_HDR_H = 24;
const DEFAULT_HEIGHT = 520, MIN_HEIGHT = 260;
const $ = (id) => document.getElementById(id);

const S = {
  sheetId: null,
  sels: {},              // sheet id → selection
  sel: { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 },
  editing: null,         // { sheetId, r, c, text, mode: 'enter'|'edit', from: 'cell'|'bar' }
  undo: [], redo: [],
  ants: null,            // { sheetId, rect, cut }
  presence: [],
  version: 0,
  zoom: 1,
  height: DEFAULT_HEIGHT,
  showFormulas: false,
  sepPref: 'auto', // formula separators: 'auto' (from the language) | 'comma' | 'semicolon'; kept in the browser
  layout: null,
  filtered: null,
  chartSel: null,
  find: null,
  saveState: '',
  meta: null,
  painter: null,         // format painter: style to apply
  tabStart: null         // column where Tab-typing a row started
};
const cfCache = new Map();
const peerMoves = new Map(); // login → when their selection last moved (name tags show for a moment)
const PEER_TAG_MS = 4000;
let peerTimer = null;

const store = new Store(api, {
  onChange: (why) => { dataChanged(why); },
  onPresence: (list) => {
    const now = Date.now();
    for (const p of list) {
      const prev = peerMoves.get(p.login);
      if (!prev || prev.k !== p.k || prev.k2 !== p.k2) peerMoves.set(p.login, { k: p.k, k2: p.k2, at: now });
    }
    const changed = JSON.stringify(list) !== JSON.stringify(S.presence);
    S.presence = list;
    if (!changed) return;
    renderPeers();
    if (store.view) renderGrid();
    clearTimeout(peerTimer);
    peerTimer = setTimeout(() => { if (store.view) renderGrid(); }, PEER_TAG_MS + 100);
  },
  onStatus: (st) => { S.saveState = st; renderSaveState(); },
  onGone: () => { toast('The spreadsheet was deleted.'); S.undo = []; S.redo = []; renderAll(); },
  onReadonly: (msg) => { toast(msg || 'You can no longer edit this spreadsheet.', 'error'); renderAll(); },
  onPermission: (canEditNow) => {
    if (!canEditNow && S.editing) { S.editing = null; $('cellEditor').hidden = true; }
    toast(canEditNow ? 'You can now edit this spreadsheet.' : 'You can no longer edit this spreadsheet. You can still view it.');
    renderAll();
  },
  onNotice: (msg) => toast(msg, 'error')
});

const view = () => store.view;
const sheet = () => (store.view ? store.view.sheets.find((s) => s.id === S.sheetId) || store.view.sheets[0] : null);
const ix = (sh) => M.indexOf(sh || sheet(), S.version);
const canEdit = () => store.canEdit && !!store.view;

const book = {
  get sheets() { return store.view ? store.view.sheets : []; },
  get names() { return store.view ? store.view.names : {}; },
  findTable: (name) => (store.view ? M.findTable(engine, store.view, name, S.version, LOCALE) : null),
  tableAt: (sh, r, c) => M.tableAt(engine, sh, r, c, S.version, LOCALE),
  /** 'manual' (hidden row), 'filter' (hidden by a table filter) or null; used by SUBTOTAL and AGGREGATE */
  rowHidden(sh, r) {
    if (sh.hr && sh.hr[sh.rows[r]]) return 'manual';
    return filteredFor(sh).has(r) ? 'filter' : null;
  }
};
const filterCache = new Map(); // sheet id -> { version, rows }
let filtering = false;
/** Rows hidden by table filters, computed once per change. */
function filteredFor(sh) {
  const hit = filterCache.get(sh.id);
  if (hit && hit.version === S.version) return hit.rows;
  if (filtering) return new Set(); // a filter that reads a SUBTOTAL: don't recurse
  filtering = true;
  try {
    const rows = M.filteredRows(engine, sh, S.version, LOCALE);
    filterCache.set(sh.id, { version: S.version, rows });
    return rows;
  } finally { filtering = false; }
}
const engine = new Engine(book, LOCALE);

/** Everything dialogs.js needs. */
/** The separator style of formulas for this person, and conversions between stored and shown formulas. */
const fmlStyle = () => separatorStyle(S.sepPref, LOCALE);
const fmlLocal = (raw) => formulaToLocal(raw, fmlStyle());
const fmlStored = (text) => formulaFromLocal(text, fmlStyle());
const fnHelp = (f) => helpToLocal(FUNCTION_HELP[f], fmlStyle());
/** What the formula bar and the cell editor show for a cell. */
const cellEditText = (raw, st) => (raw && raw[0] === '=' ? fmlLocal(raw) : editText(raw, st, LOCALE));

export const A = {
  S, store, engine, LOCALE, api, fmlLocal, fmlStored, fnHelp,
  view, sheet, ix, canEdit,
  commit, render: () => renderAll(), renderGrid: () => renderGrid(), selectRect, selRect, focusGrid,
  scrollIntoView, toast, goToSheet, version: () => S.version, cfCache,
  importFile: (f) => importFile(f)
};

/* ================================================================ data changes */

function dataChanged(why) {
  S.version++;
  engine.invalidate();
  cfCache.clear();
  S.layout = null;
  S.filtered = null;
  const sh = sheet();
  if (sh && sh.id !== S.sheetId) S.sheetId = sh.id;
  if (why === 'reset' && sh && !S.sels[sh.id]) S.sels[sh.id] = S.sel;
  clampSelection();
  renderAll();
}

function clampSelection() {
  const sh = sheet();
  if (!sh) return;
  const maxR = sh.rows.length - 1, maxC = sh.cols.length - 1;
  const s = S.sel;
  const cl = (v, m) => Math.max(0, Math.min(m, v));
  S.sel = { r: cl(s.r, maxR), c: cl(s.c, maxC), r1: cl(s.r1, maxR), c1: cl(s.c1, maxC), r2: cl(s.r2, maxR), c2: cl(s.c2, maxC) };
}

/**
 * Applies a change made here. opts: { label, regen, keepAnts }
 * Returns false when it was not allowed.
 */
function commit(ops, opts = {}) {
  if (!canEdit()) { toast('You can view this spreadsheet but not edit it.'); return false; }
  if (!ops || !ops.length) return true;
  if (opts.noUndo) {
    try { store.apply(ops, { regen: opts.regen }); } catch (e) { toast(e.message || String(e), 'error'); return false; }
    return true;
  }
  const entry = { label: opts.label || '', sheetId: S.sheetId, sel: { ...S.sel }, ops: null };
  try {
    entry.ops = store.apply(ops, { regen: opts.regen, undoEntry: entry });
  } catch (e) {
    toast(e.message || String(e), 'error');
    return false;
  }
  S.undo.push(entry);
  if (S.undo.length > 200) S.undo.shift();
  S.redo = [];
  if (!opts.keepAnts && S.ants && !S.ants.cut) S.ants = null;
  return true;
}

function undoRedo(isRedo) {
  if (!canEdit()) return;
  const from = isRedo ? S.redo : S.undo, to = isRedo ? S.undo : S.redo;
  let entry = from.pop();
  while (entry && entry.dead) entry = from.pop();
  if (!entry) return;
  const back = { label: entry.label, sheetId: S.sheetId, sel: { ...S.sel }, ops: null };
  try {
    back.ops = store.apply(entry.ops, { undoEntry: back });
  } catch (e) {
    toast('This change can no longer be undone: ' + e.message, 'error');
    return;
  }
  to.push(back);
  if (entry.sheetId && view().sheets.some((s) => s.id === entry.sheetId)) {
    if (entry.sheetId !== S.sheetId) goToSheet(entry.sheetId);
    S.sel = { ...entry.sel };
    clampSelection();
    scrollIntoView(S.sel.r, S.sel.c);
  }
  renderAll();
}

/* ================================================================ layout */

function computeLayout() {
  const sh = sheet();
  if (!sh) return null;
  const z = S.zoom;
  if (!S.filtered) S.filtered = filteredFor(sh);
  const fr = Math.min(sh.freeze ? sh.freeze.r : 0, sh.rows.length), fc = Math.min(sh.freeze ? sh.freeze.c : 0, sh.cols.length);
  const rowH = sh.rows.map((id, r) => (sh.hr[id] || S.filtered.has(r) ? 0 : Math.round((sh.rowH[id] || Z_ROW_H) * z)));
  const colW = sh.cols.map((id) => (sh.hc[id] ? 0 : Math.round((sh.colW[id] || Z_COL_W) * z)));
  const colLeft = [0];
  colW.forEach((w, i) => colLeft.push(colLeft[i] + w));
  // rows below the frozen ones scroll; their tops are measured from the first scrolling row
  const rowTop = new Array(sh.rows.length + 1);
  rowTop[0] = 0;
  for (let r = 0; r < sh.rows.length; r++) rowTop[r + 1] = rowTop[r] + rowH[r];
  const digits = String(sh.rows.length).length;
  const rhW = Math.round(Math.max(36, 14 + digits * 8) * z);
  const hdrH = Math.round(Z_HDR_H * z);
  return { sh, z, fr, fc, rowH, colW, colLeft, rowTop, rhW, hdrH, frozenH: rowTop[fr], frozenW: colLeft[fc], totalW: rhW + colLeft[sh.cols.length], totalH: hdrH + rowTop[sh.rows.length] };
}
function layout() { if (!S.layout || S.layout.sh !== sheet()) S.layout = computeLayout(); return S.layout; }

/** Merge containing (r, c) as positions, or null. */
function mergeAt(sh, r, c) {
  const x = ix(sh);
  for (const m of sh.merges) { const rect = M.objRect(m, x); if (M.inRect(rect, r, c)) return rect; }
  return null;
}
function mergeRects(sh) { const x = ix(sh); return sh.merges.map((m) => M.objRect(m, x)).filter(Boolean); }

/** Grows a rectangle so that it covers every merge it touches (selections never cut a merge in half). */
function expandForMerges(sh, rect) {
  const merges = mergeRects(sh);
  if (!merges.length) return rect;
  let r = { ...rect }, changed = true;
  while (changed) {
    changed = false;
    for (const m of merges) {
      if (m.r2 < r.r1 || m.r1 > r.r2 || m.c2 < r.c1 || m.c1 > r.c2) continue;
      const n = { r1: Math.min(r.r1, m.r1), c1: Math.min(r.c1, m.c1), r2: Math.max(r.r2, m.r2), c2: Math.max(r.c2, m.c2) };
      if (n.r1 !== r.r1 || n.c1 !== r.c1 || n.r2 !== r.r2 || n.c2 !== r.c2) { r = n; changed = true; }
    }
  }
  return r;
}

/* ================================================================ rendering */

let menubarFor = null; // what the menubar was built for: people who can only view get no greyed-out menus
const menubarKind = () => (canEdit() ? 'edit' : store.view ? 'view' : 'none');
function renderAll() {
  renderBar();
  const exists = store.exists && !!store.view;
  const failed = !!$('barMeta').dataset.error;
  document.body.classList.toggle('has-sheet', exists);
  // in the issue, an issue without a spreadsheet shows one compact row instead of the bar and the spreadsheet
  document.body.classList.toggle('no-sheet', store.loaded && !exists && !failed);
  document.body.classList.toggle('readonly', exists && !store.canEdit);
  $('roTag').hidden = !(exists && !store.canEdit);
  if (menubarFor !== menubarKind()) buildMenubar();
  // in the issue the bar alone shows the loading state and load errors; the large window shows its body from the start
  const showBody = MODE === 'window' || (store.loaded && !failed);
  $('body').hidden = !showBody;
  if (!showBody) return;
  $('empty').hidden = exists;
  $('sheetUi').hidden = !exists;
  if (MODE === 'inline') $('app').style.setProperty('--panel-h', S.height + 'px');
  if (!exists) { renderEmpty(); return; }
  $('addSheet').hidden = !canEdit();
  renderToolbarState();
  renderFormulaBar();
  renderGrid();
  renderTabs();
  renderStatus();
}

function renderBar() {
  const meta = $('barMeta');
  const fresh = store.meta || {};
  const m = store.view ? {
    sheets: store.view.sheets.length,
    updatedAt: fresh.updatedAt || store.view.updatedAt, updatedByName: fresh.updatedByName || store.view.updatedByName
  } : null;
  const loading = !store.loaded && !meta.dataset.error;
  meta.classList.toggle('loading', loading);
  if (meta.dataset.error) { /* keeps the load error */ }
  else if (loading) { if (!meta.querySelector('.loader')) { meta.innerHTML = ''; meta.append(h('span', { class: 'loader', 'aria-hidden': 'true' }), 'Loading…'); } }
  else if (!store.exists) meta.textContent = store.canEdit ? 'No spreadsheet yet' : 'No spreadsheet';
  else if (m) {
    const edited = m.updatedAt ? (Date.now() - Date.parse(m.updatedAt) < 60000 ? 'just now' : relTime(m.updatedAt, 'en')) : '';
    meta.textContent = (m.sheets === 1 ? '1 sheet' : m.sheets + ' sheets') + (edited ? ' · edited ' + edited + (m.updatedByName ? ' by ' + m.updatedByName : '') : '');
  }
  // the text is cut off when the bar is narrow: the whole text as a tooltip
  meta.title = loading ? '' : meta.textContent;
  $('moreBtn').hidden = !store.loaded;
}

let savedTimer = 0, shownState = null;
function renderSaveState() {
  const el = $('saveState');
  const st = store.exists ? S.saveState : '';
  // every poll reports the state again: "All changes saved" shows once, for 3 s, after saving
  if (st === shownState) return;
  const prev = shownState;
  shownState = st;
  clearTimeout(savedTimer);
  el.innerHTML = '';
  el.title = '';
  if (st === 'saving') el.append(h('span', { class: 'loader', 'aria-hidden': 'true' }), 'Saving…');
  else if (st === 'offline') el.textContent = 'Offline – retrying';
  else if (st === 'saved' && (prev === 'saving' || prev === 'offline')) {
    el.textContent = 'All changes saved';
    savedTimer = setTimeout(() => { if (S.saveState === 'saved') el.textContent = ''; }, 3000);
  }
  el.className = 'save-state ' + st;
}

function renderPeers() {
  const el = $('peers');
  el.innerHTML = '';
  const seen = new Set();
  for (const p of S.presence) {
    if (seen.has(p.login)) continue;
    seen.add(p.login);
    const sh = store.view && store.view.sheets.find((s) => s.id === p.s);
    let where = '';
    if (sh && p.k) { const label = Core.cellLabel(sh, p.k); where = ' · ' + sh.name + '!' + label; }
    el.appendChild(h('span', {
      class: 'peer', style: { background: peerColor(p.login) }, title: p.name + where + (p.editing ? ' (typing)' : ''),
      onclick: () => { if (sh && p.k) { goToSheet(sh.id); const [rid, cid] = p.k.split(':'); const r = ix(sh).rows.get(rid), c = ix(sh).cols.get(cid); if (r !== undefined) { selectCell(r, c); scrollIntoView(r, c); } } }
    }, initials(p.name)));
  }
}

function renderEmpty() {
  const el = $('empty');
  el.innerHTML = '';
  if (!store.loaded) { el.appendChild(h('div', { class: 'loading' }, h('span', { class: 'loader', 'aria-hidden': 'true' }), 'Loading…')); return; }
  if (MODE === 'inline') {
    // one compact row under the section header that YouTrack draws for the widget
    el.appendChild(h('span', { class: 'empty-text' }, 'No spreadsheet yet'));
    if (store.canEdit) {
      el.appendChild(h('button', { class: 'btn primary', type: 'button', onclick: () => createWorkbook() }, 'Create spreadsheet'));
      el.appendChild(h('button', { class: 'btn', type: 'button', title: 'Import an .xlsx, .csv or .tsv file', onclick: () => $('fileInput').click() }, 'Import file…'));
    }
    return;
  }
  if (!store.canEdit) {
    el.appendChild(h('div', { class: 'empty-title' }, 'No spreadsheet yet'));
    el.appendChild(h('div', { class: 'muted' }, 'People who can edit this issue can add one.'));
    return;
  }
  el.appendChild(h('div', { class: 'empty-title' }, 'Add a spreadsheet to this issue'));
  el.appendChild(h('div', { class: 'muted' }, 'Everyone who can see this issue can view it. People who can edit the issue can work on it together.'));
  el.appendChild(h('div', { class: 'empty-actions' },
    h('button', { class: 'btn primary', type: 'button', onclick: () => createWorkbook() }, 'Create spreadsheet'),
    h('button', { class: 'btn', type: 'button', onclick: () => $('fileInput').click() }, 'Import file…')));
  el.appendChild(h('div', { class: 'muted empty-hint' }, '.xlsx, .csv or .tsv'));
}

async function createWorkbook(sheets) {
  try {
    const res = await api('create', { method: 'POST', body: sheets ? { sheets } : {} });
    store.confirmed = Core.upgrade(res.doc);
    store.rev = res.doc.rev;
    store.exists = true;
    store.pending = [];
    store.rebuild();
    S.sheetId = store.view.sheets[0].id;
    S.sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    dataChanged('reset');
    if (res.existed) toast('Someone else just created this spreadsheet.');
    focusGrid();
  } catch (e) {
    toast(errorText(e), 'error');
  }
}

/* ---------------------------------------------------------------- toolbar */

const TB = {};
/** The currency symbol of the locale ("$", "€", "₺") for the currency button. */
function currencySymbol() {
  for (const display of ['narrowSymbol', 'symbol']) {
    try {
      const part = new Intl.NumberFormat(LOCALE, { style: 'currency', currency: CURRENCY, currencyDisplay: display }).formatToParts(1).find((p) => p.type === 'currency');
      if (part && part.value) return part.value;
    } catch (e) { /* try the next form */ }
  }
  return CURRENCY.length <= 3 && !/^[A-Z]{3}$/.test(CURRENCY) ? CURRENCY : '$';
}

/*
 * The toolbar is one row. When it does not fit, tools are hidden from the end by priority (4 first, then 3, then 2;
 * priority 1 always stays) and the "More tools" button lists the hidden ones.
 */
function buildToolbar() {
  const tb = $('toolbar');
  tb.innerHTML = '';
  const track = (key, b, prio, label, iconName) => {
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.dataset.prio = String(prio);
    b._label = label;
    b._icon = iconName || null;
    TB[key] = b;
    tb.appendChild(b);
    return b;
  };
  const btn = (key, title, iconName, action, opts = {}) => {
    const b = h('button', { class: 'tb', type: 'button', title, 'aria-label': title, html: opts.html || icon(iconName), 'data-popup-anchor': opts.popup ? '' : null });
    b.addEventListener('click', (e) => action(b, e));
    b._run = (a) => action(a || b);
    return track(key, b, opts.prio || 4, title, iconName);
  };
  const dd = (key, title, cls, options, prio, onChange) => {
    const s = select(options, options[0][0], { class: 'tb-dd ' + cls, title, 'aria-label': title });
    s.addEventListener('change', () => { onChange(s.value); focusGrid(); });
    s._options = options;
    s._pick = (v) => { s.value = v; onChange(v); focusGrid(); };
    return track(key, s, prio, title, null);
  };
  const sep = () => tb.appendChild(h('span', { class: 'tb-sep', 'aria-hidden': 'true' }));
  btn('undo', 'Undo (Ctrl+Z)', 'undo', () => undoRedo(false), { prio: 1 });
  btn('redo', 'Redo (Ctrl+Y)', 'redo', () => undoRedo(true), { prio: 1 });
  btn('painter', 'Format painter', 'paint', () => togglePainter(), { prio: 3 });
  sep();
  const nf = h('button', { class: 'tb tb-text', type: 'button', title: 'Number format', 'aria-label': 'Number format', 'data-popup-anchor': '' }, '123', h('span', { html: icon('down', 12) }));
  nf.addEventListener('click', () => toggleMenu(nf, numberFormatItems));
  track('nf', nf, 2, 'Number format', null);
  btn('cur', 'Currency (' + CURRENCY + ')', 'currency', () => setFormat({ nf: 'cur', cur: CURRENCY, dp: 2 }), { prio: 2, html: `<span class="tb-sym" aria-hidden="true">${esc(currencySymbol())}</span>` });
  btn('pct', 'Percent', 'pct', () => setFormat({ nf: 'pct', dp: 0 }), { prio: 2 });
  btn('decDec', 'Fewer decimals', 'decDec', () => changeDecimals(-1), { prio: 4 });
  btn('decInc', 'More decimals', 'decInc', () => changeDecimals(1), { prio: 4 });
  sep();
  dd('font', 'Font', 'font', [['sans', 'Default'], ['serif', 'Serif'], ['mono', 'Monospace']], 3, (v) => applyStyle({ ff: v === 'sans' ? null : v }));
  dd('size', 'Font size', 'size', [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36, 48].map((n) => [String(n), String(n)]), 3, (v) => applyStyle({ fs: +v === 13 ? null : +v }));
  sep();
  btn('b', 'Bold (Ctrl+B)', 'bold', () => toggleStyle('b'), { prio: 1 });
  btn('i', 'Italic (Ctrl+I)', 'italic', () => toggleStyle('i'), { prio: 1 });
  btn('u', 'Underline (Ctrl+U)', 'underline', () => toggleStyle('u'), { prio: 3 });
  btn('x', 'Strikethrough (Ctrl+5)', 'strike', () => toggleStyle('x'), { prio: 3 });
  btn('fc', 'Text color', 'textColor', (b) => togglePopup(b, () => colorPicker(b.getBoundingClientRect(), curStyle().fc, (c) => applyStyle({ fc: c }))), { popup: true, prio: 1 });
  btn('bg', 'Fill color', 'fill', (b) => togglePopup(b, () => colorPicker(b.getBoundingClientRect(), curStyle().bg, (c) => applyStyle({ bg: c }), 'No fill')), { popup: true, prio: 1 });
  // the current colors show as a bar under the icons
  TB.fc.appendChild(h('span', { class: 'cbar', 'aria-hidden': 'true' }));
  TB.bg.appendChild(h('span', { class: 'cbar', 'aria-hidden': 'true' }));
  btn('bd', 'Borders', 'borders', (b) => toggleMenu(b, borderItems), { popup: true, prio: 2 });
  btn('merge', 'Merge cells', 'merge', (b) => toggleMenu(b, mergeItems), { popup: true, prio: 3 });
  sep();
  btn('left', 'Align left', 'left', () => applyStyle({ ha: curStyle().ha === 'left' ? null : 'left' }), { prio: 2 });
  btn('center', 'Align center', 'center', () => applyStyle({ ha: curStyle().ha === 'center' ? null : 'center' }), { prio: 2 });
  btn('right', 'Align right', 'right', () => applyStyle({ ha: curStyle().ha === 'right' ? null : 'right' }), { prio: 2 });
  btn('va', 'Vertical alignment', 'vmid', (b) => toggleMenu(b, vAlignItems), { popup: true, prio: 3 });
  btn('w', 'Wrap text', 'wrap', () => toggleStyle('w'), { prio: 3 });
  sep();
  btn('sum', 'AutoSum (Alt+=)', 'sum', (b) => toggleMenu(b, autoSumItems), { popup: true, prio: 3 });
  btn('filter', 'Filter / format as table (Ctrl+Shift+L)', 'filter', () => toggleFilter(), { prio: 2 });
  btn('sortAsc', 'Sort A → Z', 'sortAsc', () => quickSort(false), { prio: 3 });
  btn('sortDesc', 'Sort Z → A', 'sortDesc', () => quickSort(true), { prio: 3 });
  btn('table', 'Format as table', 'table', () => D.tableDialog(A), { prio: 4 });
  btn('chart', 'Insert chart', 'chart', () => D.chartDialog(A), { prio: 4 });
  btn('note', 'Note (Shift+F2)', 'note', () => D.noteDialog(A), { prio: 4 });
  btn('cf', 'Conditional formatting', 'cf', () => D.cfDialog(A), { prio: 4 });
  btn('find', 'Find and replace (Ctrl+F)', 'search', () => D.findDialog(A), { prio: 4 });
  const more = h('button', { class: 'tb tb-more', type: 'button', title: 'More tools', 'aria-label': 'More tools', 'aria-haspopup': 'menu', html: icon('more'), 'data-popup-anchor': '' });
  more.addEventListener('mousedown', (e) => e.preventDefault());
  more.addEventListener('click', () => toggleMenu(more, moreToolItems));
  TB.more = more;
  tb.appendChild(more);
  fitToolbar();
}
const vAlignItems = () => [
  { label: 'Top', icon: 'vtop', action: () => applyStyle({ va: 'top' }) },
  { label: 'Middle', icon: 'vmid', action: () => applyStyle({ va: 'middle' }) },
  { label: 'Bottom', icon: 'vbottom', action: () => applyStyle({ va: null }) }];
const autoSumItems = () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((f) => ({ label: f[0] + f.slice(1).toLowerCase(), action: () => autoSum(f) }));

/** Hides tools that do not fit (lowest priority first, from the end); "More tools" lists them (and the formula separators). */
function fitToolbar() {
  const tb = $('toolbar');
  if (!tb || !TB.more) return;
  const items = [...tb.children].filter((el) => el !== TB.more);
  items.forEach((el) => el.classList.remove('tb-off'));
  TB.more.hidden = false; // always there: it also holds the formula separators
  const seps = () => {
    // no separators at the ends or next to each other
    let prevSep = null, seen = false;
    for (const el of items) {
      if (el.classList.contains('tb-sep')) {
        el.classList.remove('tb-off');
        if (!seen || prevSep) el.classList.add('tb-off');
        else prevSep = el;
        continue;
      }
      if (el.classList.contains('tb-off')) continue;
      seen = true;
      prevSep = null;
    }
    if (prevSep) prevSep.classList.add('tb-off');
  };
  const overflows = () => tb.scrollWidth > tb.clientWidth + 1;
  if (!tb.clientWidth || !overflows()) return;
  for (const prio of ['4', '3', '2']) {
    const group = items.filter((el) => el.dataset.prio === prio).reverse();
    for (const el of group) {
      if (!overflows()) break;
      el.classList.add('tb-off');
      seps();
    }
    if (!overflows()) break;
  }
  seps();
}
let fitRaf = 0;
function scheduleFit() {
  if (fitRaf) return;
  fitRaf = requestAnimationFrame(() => { fitRaf = 0; fitToolbar(); });
}

/** The hidden tools for the "More tools" menu, in toolbar order. */
function moreToolItems() {
  const out = [];
  let lastSep = true;
  for (const el of $('toolbar').children) {
    if (el === TB.more) continue;
    if (el.classList.contains('tb-sep')) { if (!lastSep && out.length) { out.push('-'); lastSep = true; } continue; }
    if (!el.classList.contains('tb-off')) continue;
    const title = el._label || el.title || '';
    const m = /^(.*?)\s*\(([^)]*)\)\s*$/.exec(title);
    const label = m ? m[1] : title;
    const kbd = m && /^(Ctrl|Alt|Shift|Cmd|F\d)/.test(m[2]) ? m[2] : '';
    let item;
    if (el === TB.nf) item = { label, icon: null, sub: numberFormatItems(), disabled: el.disabled };
    else if (el._options) item = { label, sub: el._options.map(([v, l]) => ({ label: l, checked: el.value === v, action: () => el._pick(v) })), disabled: el.disabled };
    else item = { label, kbd, icon: el._icon, checked: el.classList.contains('on') ? true : undefined, disabled: el.disabled, action: () => el._run(TB.more) };
    if (el === TB.cur) item.icon = 'currency';
    out.push(item);
    lastSep = false;
  }
  while (out.length && out[out.length - 1] === '-') out.pop();
  // after the tools that do not fit: the person's formula separators
  if (out.length) out.push('-');
  out.push({ label: 'Formula separators', sub: sepItems() });
  return out;
}

function curStyle() {
  const sh = sheet();
  if (!sh) return {};
  const cell = M.cellAt(sh, S.sel.r, S.sel.c);
  return (cell && cell.st) || {};
}

function renderToolbarState() {
  const st = curStyle();
  for (const k of ['b', 'i', 'u', 'x', 'w']) TB[k] && TB[k].classList.toggle('on', !!st[k]);
  TB.left.classList.toggle('on', st.ha === 'left');
  TB.center.classList.toggle('on', st.ha === 'center');
  TB.right.classList.toggle('on', st.ha === 'right');
  TB.pct.classList.toggle('on', st.nf === 'pct');
  TB.cur.classList.toggle('on', st.nf === 'cur');
  TB.painter.classList.toggle('on', !!S.painter);
  TB.font.value = st.ff || 'sans';
  TB.size.value = String(st.fs || 13);
  paintBar(TB.fc, st.fc, true);
  paintBar(TB.bg, st.bg, false);
  const ro = !canEdit();
  $('toolbar').querySelectorAll('button,select').forEach((b) => { if (b !== TB.find) b.disabled = ro; });
  TB.undo.disabled = ro || !S.undo.some((e) => !e.dead);
  TB.redo.disabled = ro || !S.redo.some((e) => !e.dead);
  $('menubar').querySelectorAll('button').forEach((b) => { b.disabled = false; });
}

/* ---------------------------------------------------------------- formula bar */

function renderFormulaBar() {
  const sh = sheet();
  if (!sh) return;
  const nb = $('nameBox');
  if (document.activeElement !== nb) {
    const s = S.sel;
    nb.value = s.r1 === s.r2 && s.c1 === s.c2 ? addrName(s.r, s.c) : (s.r1 === s.r2 && s.c1 === s.c2 ? addrName(s.r, s.c) : rectLabel(s));
    const named = Object.entries(view().names || {}).find(([, ref]) => ref.replace(/\$/g, '').toLowerCase() === ('=' + quoteSheet(sh.name) + '!' + M.rectText(s)).toLowerCase());
    if (named) nb.value = named[0];
  }
  const fb = $('fbar');
  if (S.editing) { if (fb.value !== S.editing.text) fb.value = S.editing.text; }
  else {
    const cell = M.cellAt(sh, S.sel.r, S.sel.c);
    const anchor = engine.spillAnchorOf(sh, S.sel.r, S.sel.c);
    if (!cell && anchor) {
      const ac = M.cellAt(sh, anchor.r, anchor.c);
      fb.value = ac ? ac.v : '';
      fb.classList.add('ghost');
    } else {
      fb.value = cell ? cellEditText(cell.v, cell.st) : '';
      fb.classList.remove('ghost');
    }
  }
  fb.readOnly = !canEdit();
  autoGrowBar();
}
function rectLabel(s) { return M.rectText({ r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }); }
function autoGrowBar() {
  const fb = $('fbar');
  fb.classList.toggle('formula', fb.value[0] === '='); // monospace only for formulas
  if (document.activeElement !== fb) { fb.style.height = ''; return; }
  fb.style.height = 'auto';
  fb.style.height = Math.min(120, Math.max(28, fb.scrollHeight)) + 'px';
}

/* ---------------------------------------------------------------- grid */

const styleCache = new Map();
const HEX_COLOR = /^#[0-9a-f]{6}$/i;
/**
 * Inline style and classes of a cell: { css, cls }.
 * Fills and text colors that people pick (#rrggbb, in cells and in the formats of conditional formatting rules) are
 * stored for the light theme. They go into custom properties (class f: --cbg, class c: --cfc) together with a darker fill
 * and a lighter text for the dark theme (themedColors, app.css), so that colored cells stay readable there.
 * Table bands, data bars and color-scale fills (scaleFill: cf.bg comes from a color scale) keep their colors in both
 * themes: the order of a scale's colors is its meaning. Text on a scale fill gets a color that is readable on it.
 */
function cellCss(st, cf, band, z, scaleFill) {
  const key = JSON.stringify([st, cf, band, z, !!scaleFill]);
  let out = styleCache.get(key);
  if (out !== undefined) return out;
  const s = st || {};
  const parts = [];
  let cls = '';
  const bg = (cf && cf.bg) || s.bg || band || null;
  const fc = (cf && cf.fc) || s.fc || null;
  if (scaleFill && bg) {
    parts.push('background:' + bg);
    const t = fc || contrastText(bg);
    if (t) parts.push('color:' + t);
  } else {
    const userFill = bg && bg !== band && HEX_COLOR.test(bg) ? bg : null;
    const tc = themedColors(userFill, fc);
    cls += tc.cls;
    parts.push(...tc.parts);
    if (!userFill && bg) parts.push('background:' + bg);
    if (fc && !HEX_COLOR.test(fc)) parts.push('color:' + fc);
  }
  if (s.b || (cf && cf.b)) parts.push('font-weight:600');
  if (s.i || (cf && cf.i)) parts.push('font-style:italic');
  const deco = [s.u ? 'underline' : '', s.x || (cf && cf.x) ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) parts.push('text-decoration:' + deco);
  if (s.fs) parts.push('font-size:' + Math.round(s.fs * z) + 'px');
  if (s.ff) parts.push('font-family:var(--font-' + s.ff + ')');
  if (s.bd) {
    const c = s.bc || 'var(--cell-border)';
    const w = (ch) => (s.bd.includes(ch.toUpperCase()) ? 2 : s.bd.includes(ch) ? 1 : 0);
    const shadows = [];
    if (w('t')) shadows.push(`inset 0 ${w('t')}px 0 ${c}`);
    if (w('l')) shadows.push(`inset ${w('l')}px 0 0 ${c}`);
    if (w('r')) shadows.push(`inset -${w('r')}px 0 0 ${c}`);
    if (w('b')) shadows.push(`inset 0 -${w('b')}px 0 ${c}`);
    if (shadows.length) parts.push('box-shadow:' + shadows.join(','));
  }
  out = { css: parts.join(';'), cls };
  if (styleCache.size > 4000) styleCache.clear();
  styleCache.set(key, out);
  return out;
}

function hexAlpha(hex, a) {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || '');
  return m ? `rgba(${parseInt(m[1], 16)}, ${parseInt(m[2], 16)}, ${parseInt(m[3], 16)}, ${a})` : null;
}

let lastWindow = null;
function visibleRowRange(L) {
  const wrap = $('gridWrap');
  const top = wrap.scrollTop, hgt = wrap.clientHeight || 400;
  const sh = L.sh;
  const n = sh.rows.length;
  const scrollTop = Math.max(0, top);
  // scrolling rows start at L.fr; their y (relative to the first scrolling row) is rowTop[r] - rowTop[fr]
  const base = L.rowTop[L.fr];
  const y0 = scrollTop, y1 = scrollTop + hgt;
  let lo = L.fr, hi = n;
  while (lo < hi) { const mid = (lo + hi) >> 1; if (L.rowTop[mid + 1] - base <= y0) lo = mid + 1; else hi = mid; }
  let a = Math.max(L.fr, lo - 8);
  let b = a;
  while (b < n && L.rowTop[b] - base < y1) b++;
  b = Math.min(n - 1, b + 8);
  // include the first row of merges that reach into the window
  for (const m of mergeRects(sh)) if (m.r1 < a && m.r2 >= a && m.r1 >= L.fr) a = m.r1;
  return [a, b];
}

function selectionDecos(sh) {
  const out = [];
  const s = S.sel;
  if (!S.editing || S.editing.sheetId === sh.id) {
    const rect = expandForMerges(sh, { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 });
    const single = rect.r1 === rect.r2 && rect.c1 === rect.c2 || (mergeAt(sh, s.r, s.c) && JSON.stringify(mergeAt(sh, s.r, s.c)) === JSON.stringify(rect));
    out.push({ rect, cls: single ? 'd-sel single' : 'd-sel', fill: !single });
    const act = mergeAt(sh, s.r, s.c) || { r1: s.r, c1: s.c, r2: s.r, c2: s.c };
    out.push({ rect: act, cls: 'd-act', act: true });
  }
  if (S.ants && S.ants.sheetId === sh.id) out.push({ rect: S.ants.rect, cls: 'd-ants' });
  for (const p of S.presence) {
    if (p.s !== sh.id || !p.k) continue;
    const x = ix(sh);
    const [r1id, c1id] = p.k.split(':');
    const [r2id, c2id] = (p.k2 || p.k).split(':');
    const r1 = x.rows.get(r1id), c1 = x.cols.get(c1id), r2 = x.rows.get(r2id), c2 = x.cols.get(c2id);
    if (r1 === undefined || c1 === undefined) continue;
    const rect = { r1: Math.min(r1, r2 ?? r1), c1: Math.min(c1, c2 ?? c1), r2: Math.max(r1, r2 ?? r1), c2: Math.max(c1, c2 ?? c1) };
    const moved = peerMoves.get(p.login);
    out.push({ rect, cls: 'd-peer', color: peerColor(p.login), label: p.name, labelAt: { r: r1, c: c1 }, recent: !!moved && Date.now() - moved.at < PEER_TAG_MS });
  }
  if (S.editing && S.editing.text[0] === '=') {
    const refs = formulaRefs(S.editing.text);
    const editSheet = view().sheets.find((x) => x.id === S.editing.sheetId);
    refs.forEach((ref, i) => {
      const target = ref.sheet ? engine.sheetByName(ref.sheet) : editSheet;
      if (target !== sh || ref.r1 === null || ref.c1 === null) return;
      out.push({ rect: { r1: Math.min(ref.r1, ref.r2), c1: Math.min(ref.c1, ref.c2), r2: Math.max(ref.r1, ref.r2), c2: Math.max(ref.c1, ref.c2) }, cls: 'd-ref', color: REF_COLORS[i % REF_COLORS.length] });
    });
  }
  if (S.fillPreview) out.push({ rect: S.fillPreview, cls: 'd-fillpre' });
  if (!S.editing) {
    const sp = engine.spillAt(sh, s.r, s.c);
    if (sp) out.push({ rect: { r1: s.r, c1: s.c, r2: s.r + sp.rows - 1, c2: s.c + sp.cols - 1 }, cls: 'd-spill' });
  }
  if (S.find && S.find.hit && S.find.hit.sheetId === sh.id) out.push({ rect: { r1: S.find.hit.r, c1: S.find.hit.c, r2: S.find.hit.r, c2: S.find.hit.c }, cls: 'd-found' });
  return out;
}
const REF_COLORS = ['#3369d6', '#d9480f', '#7a4fd6', '#1f8039', '#c2255c', '#0b7285', '#e67700'];

function decoHtml(decos, r1, c1, r2, c2, isAnchorBottomRight) {
  let out = '';
  for (const d of decos) {
    const R = d.rect;
    if (R.r2 < r1 || R.r1 > r2 || R.c2 < c1 || R.c1 > c2) continue;
    const t = R.r1 >= r1 && R.r1 <= r2 ? 1 : 0, b = R.r2 >= r1 && R.r2 <= r2 ? 1 : 0;
    const l = R.c1 >= c1 && R.c1 <= c2 ? 1 : 0, rr = R.c2 >= c1 && R.c2 <= c2 ? 1 : 0;
    if (d.act) {
      if (!(R.r1 === r1 && R.c1 === c1)) continue;
      out += `<div class="deco d-act"></div>`;
      continue;
    }
    const w = d.cls === 'd-spill' || d.cls === 'd-fillpre' ? 1 : 2;
    const style = `border-width:${t * w}px ${rr * w}px ${b * w}px ${l * w}px` + (d.color ? `;--dc:${d.color}` : '');
    out += `<div class="deco ${d.cls}${d.fill ? ' fill' : ''}" style="${style}"></div>`;
    if (d.label && d.labelAt && d.labelAt.r >= r1 && d.labelAt.r <= r2 && d.labelAt.c >= c1 && d.labelAt.c <= c2) out += `<div class="peer-tag${d.recent ? ' recent' : ''}" style="--dc:${d.color}">${esc(d.label)}</div>`;
  }
  if (isAnchorBottomRight) out += '<div class="fill-handle" data-fill="1"></div>';
  return out;
}

function renderGrid() {
  const sh = sheet();
  if (!sh) return;
  const L = layout();
  const z = L.z;
  const table = $('grid');
  const visCols = [];
  for (let c = 0; c < sh.cols.length; c++) if (L.colW[c] > 0) visCols.push(c);
  const [a, b] = visibleRowRange(L);
  lastWindow = [a, b];
  const decos = selectionDecos(sh);
  const s = S.sel;
  const selRectX = expandForMerges(sh, { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 });
  const showHandle = canEdit() && !S.editing;
  const merges = mergeRects(sh);
  const covered = new Set();
  const anchors = new Map();
  for (const m of merges) {
    anchors.set(m.r1 + ':' + m.c1, m);
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(r + ':' + c);
  }
  const x = ix(sh);
  // tables (header, bands, totals, filter buttons)
  const tables = (sh.tables || []).map((t) => ({ t, rect: M.objRect(t, x) })).filter((o) => o.rect);
  const hasCf = (sh.cf || []).length > 0;
  // fills that come from a rule's format (people picked them); any other conditional fill comes from a color scale
  const cfStyleRects = (sh.cf || []).filter((rule) => rule.style && rule.style.bg)
    .map((rule) => ({ bg: String(rule.style.bg).toLowerCase(), rect: M.objRect(rule, x) })).filter((o) => o.rect);
  const isScaleFill = (cf, r, c) => {
    if (!cf || !cf.bg) return false;
    const bg = String(cf.bg).toLowerCase();
    return !cfStyleRects.some((o) => o.bg === bg && M.inRect(o.rect, r, c));
  };
  const editable = canEdit();
  const dvRects = (sh.dv || []).map((d) => ({ d, rect: M.objRect(d, x) })).filter((o) => o.rect);

  const colgroup = `<colgroup><col style="width:${L.rhW}px">${visCols.map((c) => `<col style="width:${L.colW[c]}px">`).join('')}</colgroup>`;
  // column header row
  let head = `<tr class="hdr-row" style="height:${L.hdrH}px"><th class="corner" style="top:0;left:0" data-corner="1"></th>`;
  // the headers of the last frozen column and row carry the divider
  let lastFzC = -1, lastFzR = -1;
  for (const c of visCols) if (c < L.fc) lastFzC = c;
  for (let r = 0; r < L.fr; r++) if (L.rowH[r] > 0) lastFzR = r;
  for (const c of visCols) {
    const inSel = c >= selRectX.c1 && c <= selRectX.c2;
    const full = inSel && selRectX.r1 === 0 && selRectX.r2 === sh.rows.length - 1;
    const frozen = c < L.fc;
    const hiddenBefore = c > 0 && L.colW[c - 1] === 0;
    head += `<th class="ch${inSel ? ' hl' : ''}${full ? ' full' : ''}${frozen ? ' fz' : ''}${c === lastFzC ? ' fz-last' : ''}${hiddenBefore ? ' hid-before' : ''}" data-ch="${c}" style="top:0${frozen ? `;left:${L.rhW + L.colLeft[c]}px` : ''}">${colName(c)}<span class="col-resize" data-cres="${c}"></span></th>`;
  }
  head += '</tr>';
  // frozen rows
  let frozenTop = L.hdrH;
  const rowsHtml = (r, sticky) => {
    const hgt = L.rowH[r];
    if (!hgt) return '';
    const inSel = r >= selRectX.r1 && r <= selRectX.r2;
    const full = inSel && selRectX.c1 === 0 && selRectX.c2 === sh.cols.length - 1;
    const hiddenBefore = r > 0 && L.rowH[r - 1] === 0;
    const topCss = sticky ? `top:${frozenTop}px;` : '';
    let html = `<tr style="height:${hgt}px" data-r="${r}"><th class="rh${inSel ? ' hl' : ''}${full ? ' full' : ''}${sticky ? ' fz' : ''}${r === lastFzR ? ' fz-last' : ''}${hiddenBefore ? ' hid-before' : ''}" data-rh="${r}" style="${topCss}left:0">${r + 1}<span class="row-resize" data-rres="${r}"></span></th>`;
    let skipOverflow = 0;
    for (let vi = 0; vi < visCols.length; vi++) {
      const c = visCols[vi];
      if (covered.has(r + ':' + c)) continue;
      const m = anchors.get(r + ':' + c);
      let colspan = 1, rowspan = 1, r2 = r, c2 = c;
      if (m) {
        r2 = m.r2; c2 = m.c2;
        colspan = visCols.filter((cc) => cc >= m.c1 && cc <= m.c2).length;
        for (let rr = m.r1; rr <= m.r2; rr++) if (L.rowH[rr] > 0 && (sticky ? rr < L.fr : rr >= L.fr)) rowspan += rr === r ? 0 : 1;
      }
      const cell = sh.cells[sh.rows[r] + ':' + sh.cols[c]];
      const st = cell && cell.st;
      const frozenC = c < L.fc;
      let v = engine.valueAt(sh, r, c);
      const raw = cell ? cell.v : undefined;
      let text = S.showFormulas && raw && raw[0] === '=' ? fmlLocal(raw) : formatValue(v, st, LOCALE);
      const cf = hasCf ? M.cfFor(engine, sh, r, c, x, cfCache) : null;
      // table styling: a tinted header row with a line in the table color, light bands
      let band = null, tclass = '', filterBtn = '', headLine = null;
      for (const { t, rect } of tables) {
        if (!M.inRect(rect, r, c)) continue;
        const color = M.TABLE_STYLES[t.style || 'blue'];
        const isHeader = t.header !== false && r === rect.r1;
        const isTotal = t.totals && r === rect.r2;
        if (isHeader) {
          if (color) { band = hexAlpha(color, 0.14); headLine = color; }
          tclass += ' t-head';
          if (t.filter !== false) {
            const active = t.filters && t.filters[sh.cols[c]];
            // people who can only view see which columns are filtered, without a button they cannot use
            if (editable) filterBtn = `<button class="flt-btn${active ? ' on' : ''}" type="button" data-flt="${esc(t.id)}" data-fc="${c}" tabindex="-1" aria-label="Filter">${icon(active ? 'filter' : 'dropdown', 12)}</button>`;
            else if (active) filterBtn = `<span class="flt-btn on" aria-hidden="true">${icon('filter', 12)}</span>`;
          }
        } else if (isTotal) { tclass += ' t-total'; if (color) band = hexAlpha(color, 0.16); }
        else if (color && t.banded !== false) {
          // bands follow the visible rows, so filtered tables stay striped
          let idx = 0;
          for (let rr = rect.r1 + (t.header !== false ? 1 : 0); rr < r; rr++) if (L.rowH[rr] > 0) idx++;
          if (idx % 2 === 0) band = hexAlpha(color, 0.09);
        }
        if (r === rect.r1) tclass += ' t-top';
        if (r === rect.r2) tclass += ' t-bottom';
        if (c === rect.c1) tclass += ' t-left';
        if (c === rect.c2) tclass += ' t-right';
        tclass += ' t-' + (t.style || 'blue');
        break;
      }
      const cc = cellCss(st, cf, band, z, isScaleFill(cf, r, c));
      let css = cc.css;
      if (headLine) css += ';border-bottom-color:' + headLine;
      // alignment: numbers right, booleans/errors centered, text left (like Excel)
      const isNum = typeof v === 'number' && !(S.showFormulas && raw && raw[0] === '=');
      const ha = (st && st.ha) || (isNum ? 'right' : (typeof v === 'boolean' || isErr(v)) ? 'center' : 'left');
      const va = (st && st.va) || 'bottom';
      const wrap = st && st.w;
      let cls = 'cell a-' + ha + ' v-' + va + (wrap ? ' wrap' : '') + tclass + cc.cls + (filterBtn ? ' has-flt' : '');
      if (frozenC) cls += ' fzc';
      if (sticky) cls += ' fzr';
      if (sh.nogrid) cls += ' nogrid';
      if (isErr(v)) cls += ' err';
      const spillAnchor = !cell && v !== null ? engine.spillAnchorOf(sh, r, c) : null;
      if (spillAnchor) cls += ' spilled';
      if (skipOverflow > 0) { cls += ' ovl'; skipOverflow--; }
      // overflow of long left-aligned text into empty neighbours
      let width = L.colW[c];
      if (colspan > 1) width = visCols.filter((cc) => cc >= c && cc <= c2).reduce((sum, cc) => sum + L.colW[cc], 0);
      let ovW = 0;
      if (typeof v === 'string' && ha === 'left' && !wrap && !m && text.length * 7 * z > width - 8) {
        let k = vi + 1, extra = 0;
        while (k < visCols.length && extra < 1200 * z) {
          const cc = visCols[k];
          const ncell = sh.cells[sh.rows[r] + ':' + sh.cols[cc]];
          if ((ncell && ncell.v !== undefined) || covered.has(r + ':' + cc) || anchors.has(r + ':' + cc) || engine.spillAnchorOf(sh, r, cc)) break;
          extra += L.colW[cc];
          k++;
          if (text.length * 7 * z <= width + extra - 8) break;
        }
        if (extra) { ovW = width + extra; cls += ' ov'; skipOverflow = k - vi - 1; }
      }
      // numbers that don't fit show ####
      if (isNum && !wrap && text.length * 7.2 * z > width - 6 && width > 0) text = '#'.repeat(Math.max(1, Math.floor((width - 6) / (7.2 * z))));
      // content
      let inner = '';
      const link0 = typeof v === 'string' && /^(https?:\/\/|mailto:)[^\s]+$/i.test(v) ? v : engine.linkAt(sh, r, c);
      const link = link0 && /^(https?:\/\/|mailto:)/i.test(link0) ? link0 : null; // never javascript: or data: links
      const dvo = dvRects.find((o) => M.inRect(o.rect, r, c));
      if (dvo && dvo.d.type === 'checkbox' && (v === true || v === false || v === null)) {
        inner = `<span class="cbx${v === true ? ' on' : ''}" data-cbx="1" role="checkbox" aria-checked="${v === true}">${v === true ? icon('check', 12) : ''}</span>`;
        cls = cls.replace(/a-\w+/, 'a-center');
      } else if (link && !S.showFormulas) {
        inner = `<a class="lnk" href="${esc(link)}" target="_blank" rel="noopener noreferrer" data-link="1" title="Ctrl+click to open">${esc(text)}</a>`;
      } else inner = esc(text);
      if (cf && cf.bar) {
        const left = cf.bar.neg ? cf.bar.zero - cf.bar.pct : cf.bar.zero;
        inner = `<span class="dbar" style="left:${left}%;width:${cf.bar.pct}%;background:${hexAlpha(cf.bar.color, 0.45)}"></span>` + inner;
      }
      if (cf && cf.icon) inner = `<span class="cf-icon ${cf.icon}">${cf.icon === 'up' ? '▲' : cf.icon === 'down' ? '▼' : '▶'}</span>` + inner;
      const innerH = (rowspan > 1 ? Array.from({ length: r2 - r + 1 }, (_, i) => L.rowH[r + i]).reduce((s2, x2) => s2 + x2, 0) : L.rowH[r]) - 1;
      const cv = `<div class="cv" style="height:${innerH}px${ovW ? `;width:${ovW - 6}px` : ''}">${inner}</div>`;
      // extras: note marker, list arrow, decorations
      let extras = '';
      if (cell && cell.n) extras += '<span class="note-mark"></span>';
      if (filterBtn) extras += filterBtn;
      const isActive = r === s.r && c === s.c;
      if (dvo && dvo.d.type === 'list' && (isActive || dvo.d.arrows === 'always') && canEdit()) extras += `<button class="dv-btn" type="button" data-dv="1" tabindex="-1" aria-label="Choose from list">${icon('dropdown', 12)}</button>`;
      const brR = m ? m.r2 : r, brC = m ? m.c2 : c;
      extras += decoHtml(decos, r, c, brR, brC, showHandle && brR === selRectX.r2 && brC === selRectX.c2);
      const leftCss = frozenC ? `left:${L.rhW + L.colLeft[c]}px;` : '';
      const title = cell && cell.n ? ` title="${esc(cell.n)}"` : isErr(v) ? ` title="${esc(errorHelp(v.code))}"` : '';
      html += `<td class="${cls}" data-r="${r}" data-c="${c}"${colspan > 1 ? ` colspan="${colspan}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''} style="${topCss}${leftCss}${css}"${title}>${cv}${extras}</td>`;
    }
    html += '</tr>';
    if (sticky) frozenTop += hgt;
    return html;
  };
  for (let r = 0; r < L.fr; r++) head += rowsHtml(r, true);
  let body = '';
  const base = L.rowTop[L.fr];
  const topSpace = L.rowTop[a] - base, bottomSpace = L.rowTop[sh.rows.length] - L.rowTop[b + 1];
  if (topSpace > 0) body += `<tr class="spacer" style="height:${topSpace}px"><td colspan="${visCols.length + 1}"></td></tr>`;
  for (let r = a; r <= b; r++) body += rowsHtml(r, false);
  if (bottomSpace > 0) body += `<tr class="spacer" style="height:${bottomSpace}px"><td colspan="${visCols.length + 1}"></td></tr>`;
  table.style.width = L.totalW + 'px';
  table.style.setProperty('--z', z);
  table.innerHTML = colgroup + `<thead>${head}</thead><tbody>${body}</tbody>`;
  renderCharts();
  if (S.editing && S.editing.sheetId === sh.id) positionEditor();
}

function errorHelp(code) {
  return {
    '#DIV/0!': 'Division by zero', '#N/A': 'Value not available (nothing matched)', '#NAME?': 'Unknown function or name',
    '#REF!': 'The formula refers to cells that no longer exist', '#VALUE!': 'Wrong type of value in the formula',
    '#NUM!': 'Invalid number', '#SPILL!': 'The result needs empty cells to spill into', '#CALC!': 'The result is empty',
    '#CIRC!': 'Circular reference: the formula depends on its own cell', '#ERROR!': 'The formula has a syntax error', '#NULL!': 'Ranges do not intersect'
  }[code] || code;
}

/* ---------------------------------------------------------------- charts layer */

function renderCharts() {
  const sh = sheet();
  const layer = $('chartsLayer');
  const charts = (sh && sh.charts) || [];
  if (!charts.length) { layer.innerHTML = ''; return; }
  const L = layout();
  const x = ix(sh);
  const base = L.rowTop[L.fr];
  let html = '';
  for (const ch of charts) {
    const ar = x.rows.get(ch.ar), ac = x.cols.get(ch.ac);
    if (ar === undefined || ac === undefined) continue;
    const top = L.hdrH + L.frozenH + (ar >= L.fr ? L.rowTop[ar] - base : L.rowTop[ar] - L.frozenH) + (ch.oy || 0) * L.z;
    const left = L.rhW + L.colLeft[ac] + (ch.ox || 0) * L.z;
    const w = Math.round((ch.w || 480) * L.z), hh = Math.round((ch.h || 280) * L.z);
    const rect = M.objRect(ch, x);
    let svg = '';
    if (rect) {
      const values = [], texts = [];
      for (let r = rect.r1; r <= rect.r2; r++) {
        if (L.rowH[r] === 0) continue;
        const vr = [], tr = [];
        for (let c = rect.c1; c <= rect.c2; c++) {
          if (L.colW[c] === 0) continue;
          const v = engine.valueAt(sh, r, c);
          const cell = M.cellAt(sh, r, c);
          vr.push(isErr(v) ? null : v === '' ? null : v);
          tr.push(formatValue(v, cell && cell.st, LOCALE));
        }
        values.push(vr); texts.push(tr);
      }
      svg = renderChart(ch, chartData(values, texts), w, hh, LOCALE);
    }
    const selected = S.chartSel === ch.id;
    html += `<div class="chart${selected ? ' sel' : ''}" data-chart="${esc(ch.id)}" style="top:${top}px;left:${left}px;width:${w}px;height:${hh}px">${svg}`
      + (selected && canEdit() ? `<div class="chart-tools"><button class="btn small" type="button" data-chart-edit="1">Edit</button><button class="btn small" type="button" data-chart-del="1">Delete</button></div><div class="chart-resize" data-chart-resize="1"></div>` : '')
      + '</div>';
  }
  layer.innerHTML = html;
}

/* ---------------------------------------------------------------- tabs, status */

function renderTabs() {
  const el = $('tabs');
  el.innerHTML = '';
  view().sheets.forEach((sh, i) => {
    const tab = h('button', {
      class: 'tab' + (sh.id === S.sheetId ? ' on' : ''), type: 'button', role: 'tab', 'aria-selected': String(sh.id === S.sheetId),
      draggable: canEdit() ? 'true' : null, 'data-sheet': sh.id, 'data-label': sh.name
    }, sh.name);
    // a custom property needs setProperty (h() assigns style objects, which ignores them)
    if (sh.color) { tab.classList.add('colored'); tab.style.setProperty('--tab-color', sh.color); }
    tab.addEventListener('click', () => {
      if (S.editing && S.editing.text[0] === '=' ) { S.sheetId = sh.id; S.layout = null; S.filtered = null; renderAll(); $('fbar').focus(); return; }
      goToSheet(sh.id);
    });
    tab.addEventListener('dblclick', () => renameSheet(sh));
    tab.addEventListener('contextmenu', (e) => { e.preventDefault(); menu(sheetMenuItems(sh, i), { left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY }, { onDone: () => focusGrid(), hideDisabled: !canEdit() }); });
    tab.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-sheet', sh.id); });
    tab.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('text/x-sheet')) e.preventDefault(); });
    tab.addEventListener('drop', (e) => {
      const id = e.dataTransfer.getData('text/x-sheet');
      if (id && id !== sh.id) commit([{ t: 'moveSheet', s: id, to: view().sheets.indexOf(sh) }], { label: 'Move sheet' });
    });
    el.appendChild(tab);
  });
}

const TAB_COLORS = [['Blue', '#3369d6'], ['Green', '#1f8039'], ['Orange', '#d9730d'], ['Red', '#cc3645'], ['Purple', '#7a4fd6'], ['Teal', '#0f8b8d'], ['Gray', '#6c707e'], ['Yellow', '#d4a72c']];
function sheetMenuItems(sh, i) {
  const ro = !canEdit();
  return [
    { label: 'Rename…', action: () => renameSheet(sh), disabled: ro },
    { label: 'Duplicate', action: () => commit([M.duplicateSheet(view(), sh)], { label: 'Duplicate sheet' }), disabled: ro },
    { label: 'Delete', danger: true, action: () => deleteSheet(sh), disabled: ro || view().sheets.length < 2 },
    '-',
    { label: 'Move left', action: () => commit([{ t: 'moveSheet', s: sh.id, to: i - 1 }], { label: 'Move sheet' }), disabled: ro || i === 0 },
    { label: 'Move right', action: () => commit([{ t: 'moveSheet', s: sh.id, to: i + 1 }], { label: 'Move sheet' }), disabled: ro || i === view().sheets.length - 1 },
    { label: 'Tab color', disabled: ro, sub: [{ label: 'None', checked: !sh.color, action: () => commit([{ t: 'sheet', s: sh.id, color: null }]) }].concat(
      TAB_COLORS.map(([name, c]) => ({ label: name, swatch: c, checked: (sh.color || '').toLowerCase() === c, action: () => commit([{ t: 'sheet', s: sh.id, color: c }]) }))) }
  ];
}

async function renameSheet(sh) {
  if (!canEdit()) return;
  const name = await D.prompt('Rename sheet', 'Sheet name', sh.name);
  if (name === null) return;
  const n = name.replace(/[\\/?*[\]:]/g, '').trim().slice(0, 31);
  if (!n || n === sh.name) return;
  if (view().sheets.some((s) => s !== sh && s.name.toLowerCase() === n.toLowerCase())) { toast('Another sheet is already named "' + n + '".', 'error'); return; }
  commit(M.renameSheetOps(view(), sh, n), { label: 'Rename sheet' });
}
async function deleteSheet(sh) {
  if (!canEdit() || view().sheets.length < 2) return;
  const ok = await confirmDialog('Delete sheet', `Delete "${sh.name}" and everything on it? Formulas on other sheets that use it will show #REF!. You can undo this with Ctrl+Z.`, 'Delete', true);
  if (!ok) return;
  const i = view().sheets.indexOf(sh);
  if (commit(M.deleteSheetOps(view(), sh), { label: 'Delete sheet' })) goToSheet(view().sheets[Math.max(0, i - 1)].id);
}
function addSheet() {
  if (!canEdit()) return;
  const n = view().sheets.length;
  let name = 'Sheet' + (n + 1), k = n + 1;
  while (view().sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = 'Sheet' + ++k;
  const sh = Core.newSheet(name);
  if (commit([{ t: 'addSheet', at: n, sheet: sh }], { label: 'Add sheet' })) goToSheet(sh.id);
}

function goToSheet(id) {
  if (S.editing) commitEdit();
  S.sels[S.sheetId] = { ...S.sel };
  S.sheetId = id;
  S.sel = S.sels[id] ? { ...S.sels[id] } : { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
  S.layout = null;
  S.filtered = null;
  S.chartSel = null;
  clampSelection();
  $('gridWrap').scrollTop = 0;
  $('gridWrap').scrollLeft = 0;
  renderAll();
  sendSelection();
  scrollIntoView(S.sel.r, S.sel.c);
}

function renderStatus() {
  const sh = sheet();
  const s = S.sel;
  const el = $('status');
  if (s.r1 === s.r2 && s.c1 === s.c2) { el.textContent = ''; return; }
  let sum = 0, count = 0, nums = 0, min = Infinity, max = -Infinity;
  const cells = (s.r2 - s.r1 + 1) * (s.c2 - s.c1 + 1);
  if (cells > 200000) { el.textContent = ''; return; }
  const L = layout();
  for (let r = s.r1; r <= s.r2; r++) {
    if (L.rowH[r] === 0) continue;
    for (let c = s.c1; c <= s.c2; c++) {
      const v = engine.valueAt(sh, r, c);
      if (v === null || v === '') continue;
      count++;
      if (typeof v === 'number') { nums++; sum += v; min = Math.min(min, v); max = Math.max(max, v); }
    }
  }
  // in the number format of the active cell when it has one, otherwise grouped: whole numbers exactly (up to 15 digits),
  // other numbers with up to 10 significant digits
  // (no rounding to a fixed number of decimals: 0.001 + 0.002 shows 0.003)
  const st = curStyle();
  const own = st.nf && st.nf !== 'text' && !(st.nf === 'gen' && st.dp === undefined);
  const nf = (digits) => { try { return new Intl.NumberFormat(LOCALE, { maximumSignificantDigits: digits }); } catch (e) { return new Intl.NumberFormat('en-US', { maximumSignificantDigits: digits }); } };
  const whole = nf(15), frac = nf(10);
  const fmt = (n) => { const x = +n.toPrecision(15); return own ? formatValue(x, st, LOCALE) : (Number.isInteger(x) ? whole : frac).format(x); };
  const parts = [];
  if (nums) parts.push('Sum: ' + fmt(sum), 'Average: ' + fmt(sum / nums));
  if (nums > 1) parts.push('Min: ' + fmt(min), 'Max: ' + fmt(max));
  parts.push('Count: ' + whole.format(count));
  el.textContent = parts.join(' · ');
}

/* ================================================================ selection */

function selRect() { const s = S.sel; return { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }; }

function selectRect(rect, active) {
  const sh = sheet();
  const R = expandForMerges(sh, rect);
  const a = active || { r: rect.r1, c: rect.c1 };
  S.sel = { r: a.r, c: a.c, r1: R.r1, c1: R.c1, r2: R.r2, c2: R.c2 };
  clampSelection();
  afterSelection();
}
function selectCell(r, c, extend, keepTabStart) {
  if (!keepTabStart) S.tabStart = null;
  const sh = sheet();
  if (!sh) return;
  r = Math.max(0, Math.min(sh.rows.length - 1, r));
  c = Math.max(0, Math.min(sh.cols.length - 1, c));
  if (extend) {
    const a = { r: S.sel.r, c: S.sel.c };
    const R = expandForMerges(sh, { r1: Math.min(a.r, r), c1: Math.min(a.c, c), r2: Math.max(a.r, r), c2: Math.max(a.c, c) });
    S.sel = { ...a, ...R };
  } else {
    const m = mergeAt(sh, r, c);
    if (m) S.sel = { r: m.r1, c: m.c1, ...m };
    else S.sel = { r, c, r1: r, c1: c, r2: r, c2: c };
  }
  afterSelection();
}
function afterSelection() {
  S.chartSel = null;
  if (S.painter && !S.painterDragging) applyPainter();
  renderGrid();
  renderFormulaBar();
  renderToolbarState();
  renderStatus();
  sendSelection();
}
function sendSelection() {
  const sh = sheet();
  if (!sh) return;
  const s = S.sel;
  const k = M.keyAt(sh, s.r1, s.c1), k2 = M.keyAt(sh, s.r2, s.c2);
  store.setSelection({ s: sh.id, k, k2: k2 === k ? '' : k2, editing: !!S.editing });
}

/* ---------------------------------------------------------------- an open-ended sheet
 * A sheet starts with 100 rows and 26 columns. Moving or scrolling past the last row or column adds more, as in an
 * endless spreadsheet, up to the limits (20,000 rows and 200 columns). Added rows and columns are blank; they are not
 * an undo step and do not show in the change history. */
const GROW_ROWS = 100, GROW_COLS = 10;
let limitToastAt = 0;
function growTo(r, c) {
  const sh = sheet();
  if (!sh || !canEdit()) return false;
  const L = Core.LIMITS;
  const needR = r > sh.rows.length - 1, needC = c > sh.cols.length - 1;
  if (!needR && !needC) return false;
  const ops = [];
  if (needR && sh.rows.length < L.rows) ops.push({ t: 'rows', s: sh.id, at: -1, add: Core.ids(Math.min(L.rows - sh.rows.length, r + 1 - sh.rows.length + GROW_ROWS)) });
  if (needC && sh.cols.length < L.cols) ops.push({ t: 'cols', s: sh.id, at: -1, add: Core.ids(Math.min(L.cols - sh.cols.length, c + 1 - sh.cols.length + GROW_COLS)) });
  if (!ops.length) {
    if (Date.now() - limitToastAt > 5000) { limitToastAt = Date.now(); toast(`A sheet can have at most ${L.rows.toLocaleString(LOCALE || undefined)} rows and ${L.cols} columns (up to column ${colName(L.cols - 1)}).`); }
    return false;
  }
  return commit(ops, { noUndo: true });
}
/** Scrolled to within a screen of the end: add rows or columns so that scrolling can go on. */
function growOnScroll(wrap) {
  const sh = sheet();
  if (!sh || !canEdit() || S.growing) return;
  const nearBottom = wrap.scrollHeight - wrap.scrollTop - wrap.clientHeight < wrap.clientHeight;
  const nearRight = wrap.scrollWidth - wrap.scrollLeft - wrap.clientWidth < wrap.clientWidth / 2;
  const scrollsV = wrap.scrollHeight > wrap.clientHeight + 4, scrollsH = wrap.scrollWidth > wrap.clientWidth + 4;
  if (!(nearBottom && scrollsV && sh.rows.length < Core.LIMITS.rows) && !(nearRight && scrollsH && sh.cols.length < Core.LIMITS.cols)) return;
  S.growing = true;
  try {
    growTo(nearBottom && scrollsV ? sh.rows.length - 1 + 1 : -1, nearRight && scrollsH ? sh.cols.length - 1 + 1 : -1);
  } finally { setTimeout(() => { S.growing = false; }, 150); }
}

/** Next visible row/col in a direction (skips hidden and filtered rows). */
function step(r, c, dr, dc) {
  // moving down or right from the last row or column adds more
  { const s0 = sheet(); if (s0 && ((dr > 0 && r + dr > s0.rows.length - 1) || (dc > 0 && c + dc > s0.cols.length - 1))) growTo(dr > 0 ? r + dr : -1, dc > 0 ? c + dc : -1); }
  const sh = sheet(), L = layout();
  if (dr) { let n = r + dr; while (n >= 0 && n < sh.rows.length && L.rowH[n] === 0) n += dr; if (n >= 0 && n < sh.rows.length) r = n; }
  if (dc) { let n = c + dc; while (n >= 0 && n < sh.cols.length && L.colW[n] === 0) n += dc; if (n >= 0 && n < sh.cols.length) c = n; }
  // skip over merges
  const m = mergeAt(sh, r, c);
  if (m && (dr > 0 && r > m.r1 || dc > 0 && c > m.c1)) {
    if (dr > 0 && m.r2 + 1 < sh.rows.length) r = m.r2 + 1;
    if (dc > 0 && m.c2 + 1 < sh.cols.length) c = m.c2 + 1;
  }
  return [r, c];
}
/** Ctrl+arrow: jump to the edge of the data block. */
function jump(r, c, dr, dc) {
  const sh = sheet();
  const has = (rr, cc) => { const cell = M.cellAt(sh, rr, cc); return !!(cell && cell.v !== undefined && cell.v !== ''); };
  const maxR = sh.rows.length - 1, maxC = sh.cols.length - 1;
  const inside = (rr, cc) => rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC;
  let nr = r + dr, nc = c + dc;
  if (!inside(nr, nc)) return [r, c];
  if (has(r, c) && has(nr, nc)) {
    while (inside(nr + dr, nc + dc) && has(nr + dr, nc + dc)) { nr += dr; nc += dc; }
    return [nr, nc];
  }
  while (inside(nr, nc) && !has(nr, nc)) {
    if (!inside(nr + dr, nc + dc)) return [nr, nc];
    nr += dr; nc += dc;
  }
  return [nr, nc];
}

function scrollIntoView(r, c) {
  const L = layout();
  if (!L) return;
  const wrap = $('gridWrap');
  const base = L.rowTop[L.fr];
  if (r >= L.fr) {
    const top = L.rowTop[r] - base, bottom = L.rowTop[r + 1] - base;
    const viewH = wrap.clientHeight - L.hdrH - L.frozenH;
    if (top < wrap.scrollTop) wrap.scrollTop = top;
    else if (bottom > wrap.scrollTop + viewH) wrap.scrollTop = bottom - viewH;
  }
  if (c >= L.fc) {
    const left = L.colLeft[c] - L.frozenW, right = L.colLeft[c + 1] - L.frozenW;
    const viewW = wrap.clientWidth - L.rhW - L.frozenW;
    if (left < wrap.scrollLeft) wrap.scrollLeft = left;
    else if (right > wrap.scrollLeft + viewW) wrap.scrollLeft = Math.min(right - viewW, left);
  }
}

/* ================================================================ editing */

function startEdit(text, mode, from = 'cell') {
  if (!canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  const sh = sheet();
  const s = S.sel;
  const anchor = engine.spillAnchorOf(sh, s.r, s.c);
  const cell = M.cellAt(sh, s.r, s.c);
  const dv = M.dvFor(sh, s.r, s.c, ix(sh));
  if (dv && dv.type === 'checkbox' && text === null) { toggleCheckbox(s.r, s.c); return; }
  const current = cell ? cellEditText(cell.v, cell.st) : '';
  S.editing = { sheetId: sh.id, r: s.r, c: s.c, text: text === null ? current : text, mode, from, original: current };
  if (anchor && !cell) S.editing.spillOf = anchor;
  renderGrid();
  renderFormulaBar();
  renderToolbarState();
  sendSelection();
  if (from === 'bar') { $('fbar').focus(); return; }
  const ed = $('cellEditor');
  ed.hidden = false;
  ed.value = S.editing.text;
  positionEditor();
  ed.focus({ preventScroll: true });
  const end = ed.value.length;
  ed.setSelectionRange(end, end);
  updateAutocomplete(ed);
}

function positionEditor() {
  const ed = $('cellEditor');
  if (!S.editing || S.editing.sheetId !== S.sheetId || S.editing.from === 'bar' && ed.hidden) { if (!S.editing || S.editing.sheetId !== S.sheetId) ed.hidden = true; return; }
  const td = $('grid').querySelector(`td[data-r="${S.editing.r}"][data-c="${S.editing.c}"]`);
  if (!td) { ed.style.visibility = 'hidden'; return; }
  ed.style.visibility = '';
  const inner = $('gridInner').getBoundingClientRect();
  const rc = td.getBoundingClientRect();
  const cell = M.cellAt(sheet(), S.editing.r, S.editing.c);
  const st = (cell && cell.st) || {};
  ed.style.left = (rc.left - inner.left) + 'px';
  ed.style.top = (rc.top - inner.top) + 'px';
  ed.style.minWidth = rc.width + 'px';
  ed.style.minHeight = rc.height + 'px';
  ed.style.maxWidth = Math.max(rc.width, inner.right - rc.left - 4, 200) + 'px';
  ed.style.fontWeight = st.b ? '600' : '';
  ed.style.fontStyle = st.i ? 'italic' : '';
  ed.style.fontSize = st.fs ? Math.round(st.fs * S.zoom) + 'px' : '';
  ed.style.textAlign = st.ha || 'left';
  ed.style.width = 'auto';
  ed.style.height = 'auto';
  ed.style.width = Math.min(parseFloat(ed.style.maxWidth), Math.max(rc.width, ed.scrollWidth + 4)) + 'px';
  ed.style.height = Math.max(rc.height, ed.scrollHeight) + 'px';
}

function editorInput(src) {
  if (!S.editing) return;
  S.editing.text = src.value;
  if (src.id === 'cellEditor') { $('fbar').value = src.value; autoGrowBar(); }
  else { const ed = $('cellEditor'); ed.value = src.value; src.classList.toggle('formula', src.value[0] === '='); }
  positionEditor();
  updateAutocomplete(src);
  renderGrid();
}

function cancelEdit() {
  S.editing = null;
  $('cellEditor').hidden = true;
  closeAutocomplete();
  renderAll();
  focusGrid();
}

/** Saves the edited cell; then moves the selection by (dr, dc). Returns false when the value was rejected. */
function commitEdit(dr = 0, dc = 0, fillSelection = false) {
  const ed = S.editing;
  if (!ed) return true;
  const sh = view().sheets.find((x) => x.id === ed.sheetId);
  if (!sh) { cancelEdit(); return true; }
  let text = ed.text;
  closeAutocomplete();
  // auto-close parentheses in formulas, like Excel; formulas are stored with "," between arguments and "." in numbers
  if (text[0] === '=') {
    text = fmlStored(text);
    const open = (text.match(/\(/g) || []).length - (text.match(/\)/g) || []).length;
    if (open > 0 && open < 5) text += ')'.repeat(open);
    const err = formulaSyntaxError(text);
    if (err) { toast('There is a problem with this formula: ' + err, 'error'); return false; }
  }
  const p = parseInput(text, LOCALE);
  const x = ix(sh);
  const dv = M.dvFor(sh, ed.r, ed.c, x);
  if (dv && p.kind !== 'empty') {
    const msg = dv.type === 'formula' ? checkCustomRule(sh, ed.r, ed.c, dv, p.raw) : M.validate(engine, view(), sh, ed.r, ed.c, dv, p, LOCALE);
    if (msg) {
      if (dv.strict !== false) { toast(msg, 'error'); return false; }
      toast(msg);
    }
  }
  const cell = M.cellAt(sh, ed.r, ed.c);
  const ops = [];
  const targets = [];
  if (fillSelection && (S.sel.r1 !== S.sel.r2 || S.sel.c1 !== S.sel.c2)) {
    for (let r = S.sel.r1; r <= S.sel.r2; r++) for (let c = S.sel.c1; c <= S.sel.c2; c++) if (layout().rowH[r] > 0) targets.push([r, c]);
  } else targets.push([ed.r, ed.c]);
  for (const [r, c] of targets) {
    const k = M.keyAt(sh, r, c);
    let raw = p.raw === '' ? null : p.raw;
    if (raw && raw[0] === '=' && (r !== ed.r || c !== ed.c)) raw = translateFormula(raw, r - ed.r, c - ed.c);
    const op = { t: 'set', s: sh.id, k, v: raw };
    const c0 = M.cellAt(sh, r, c);
    const st0 = (c0 && c0.st) || {};
    if (p.fmt && (!st0.nf || st0.nf === 'gen') && st0.nf !== 'text') op.st = { ...st0, ...p.fmt };
    if (p.fmt && st0.nf === 'text') op.v = text;
    if (text.includes('\n') && !st0.w) op.st = { ...(op.st || st0), w: 1 };
    if (st0.nf === 'text' && p.kind !== 'formula') op.v = text === '' ? null : text;
    ops.push(op);
  }
  // tables: a row typed right under a table (or a column right of it) extends it; a formula fills its column
  for (const t of sh.tables || []) {
    const rect = M.objRect(t, x);
    if (!rect || p.kind === 'empty') continue;
    if (ed.r === rect.r2 + 1 && ed.c >= rect.c1 && ed.c <= rect.c2 && !t.totals) ops.push({ t: 'obj', s: sh.id, kind: 'tables', id: t.id, val: { ...t, r2: sh.rows[ed.r] } });
    else if (ed.c === rect.c2 + 1 && ed.r >= rect.r1 && ed.r <= rect.r2) {
      ops.push({ t: 'obj', s: sh.id, kind: 'tables', id: t.id, val: { ...t, c2: sh.cols[ed.c] } });
      if (ed.r === rect.r1 && t.header !== false) break;
    }
    if (p.kind === 'formula' && M.inRect(rect, ed.r, ed.c) && ed.r > rect.r1 + (t.header !== false ? 0 : -1)) {
      const top = rect.r1 + (t.header !== false ? 1 : 0), bottom = rect.r2 - (t.totals ? 1 : 0);
      let empty = true;
      for (let r = top; r <= bottom; r++) if (r !== ed.r) { const cc = M.cellAt(sh, r, ed.c); if (cc && cc.v !== undefined) { empty = false; break; } }
      if (empty && bottom > top) for (let r = top; r <= bottom; r++) if (r !== ed.r) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, ed.c), v: translateFormula(p.raw, r - ed.r, 0) });
    }
    break;
  }
  const changed = ops.length > 1 || (cell ? cell.v : undefined) !== (ops[0].v ?? undefined) || ops[0].st !== undefined;
  S.editing = null;
  $('cellEditor').hidden = true;
  if (S.sheetId !== sh.id) { S.sheetId = sh.id; S.layout = null; S.filtered = null; }
  if (changed) commit(ops, { label: 'Edit cell' });
  if (dr || dc) {
    const s = S.sel;
    const multi = s.r1 !== s.r2 || s.c1 !== s.c2;
    if (multi && !fillSelection) moveWithinSelection(dr, dc);
    else {
      // Tab, Tab, Enter goes back to the column where the row was started, like Excel
      let base = ed.c;
      if (dc && S.tabStart === null) S.tabStart = ed.c;
      if (dr === 1 && !dc && S.tabStart !== null) { base = S.tabStart; S.tabStart = null; }
      else if (!dc) S.tabStart = null;
      const [r, c] = step(ed.r, base, dr, dc);
      selectCell(r, c, false, true);
      scrollIntoView(r, c);
    }
  } else renderAll();
  focusGrid();
  return true;
}

/**
 * Custom formula validation: the formula (written for the top-left cell of the rule's range) is evaluated with the
 * new value in place; anything but TRUE or a non-zero number rejects the value.
 */
function checkCustomRule(sh, r, c, dv, raw) {
  if (!dv.formula) return null;
  const rect = M.objRect(dv, ix(sh));
  if (!rect) return null;
  const f = translateFormula(dv.formula[0] === '=' ? dv.formula : '=' + dv.formula, r - rect.r1, c - rect.c1);
  const k = M.keyAt(sh, r, c);
  const before = sh.cells[k];
  sh.cells[k] = { ...(before || {}), v: raw };
  engine.invalidate();
  let res;
  try { res = engine.evalAt(f, sh, r, c); } finally {
    if (before) sh.cells[k] = before; else delete sh.cells[k];
    engine.invalidate();
    cfCache.clear();
  }
  return res === true || (typeof res === 'number' && res !== 0) ? null : dv.msg || 'This value does not meet the validation rule of the cell.';
}

/** Enter / Tab inside a selected range cycle through it, like Excel. */
function moveWithinSelection(dr, dc) {
  const s = S.sel;
  let r = s.r + dr, c = s.c + dc;
  if (c > s.c2) { c = s.c1; r++; } else if (c < s.c1) { c = s.c2; r--; }
  if (r > s.r2) { r = s.r1; c = dc ? c : (c + 1 > s.c2 ? s.c1 : c + 1); } else if (r < s.r1) { r = s.r2; c = dc ? c : (c - 1 < s.c1 ? s.c2 : c - 1); }
  S.sel = { ...s, r, c };
  afterSelection();
  scrollIntoView(r, c);
}

/* ---- point mode: clicking cells while typing a formula inserts references */

function canInsertRef(text, pos) {
  if (text[0] !== '=') return false;
  const before = text.slice(0, pos).replace(/\s+$/, '');
  if (before === '=') return true;
  if (S.editing && S.editing.lastRef && S.editing.lastRef.end === pos) return true;
  return /[=(,;+\-*/^&<>:]$/.test(before);
}
function insertRef(r1, c1, r2, c2, extending) {
  const ed = S.editing;
  const input = ed.from === 'bar' || ed.sheetId !== S.sheetId ? $('fbar') : $('cellEditor');
  const pos = input.selectionStart;
  let start = pos;
  if (ed.lastRef && ed.lastRef.end === pos) start = ed.lastRef.start;
  const sh = sheet();
  const prefix = ed.sheetId !== sh.id ? quoteSheet(sh.name) + '!' : '';
  const ref = prefix + (r1 === r2 && c1 === c2 ? addrName(r1, c1) : M.rectText({ r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }));
  const text = input.value.slice(0, start) + ref + input.value.slice(extending ? start + (ed.lastRef ? ed.lastRef.end - ed.lastRef.start : 0) : pos);
  input.value = text;
  const end = start + ref.length;
  input.setSelectionRange(end, end);
  ed.lastRef = { start, end, r1, c1 };
  editorInput(input);
  input.focus({ preventScroll: true });
}

/* ---- autocomplete */

let ac = null;
function updateAutocomplete(input) {
  const text = input.value, pos = input.selectionStart;
  closeAutocomplete(true);
  if (text[0] !== '=') { hideHint(); return; }
  const before = text.slice(0, pos);
  const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
  const hint = currentFunction(before);
  if (hint) showHint(input, hint); else hideHint();
  if (!m || /[A-Za-z]{1,3}\d+$/.test(m[1]) && !/^[A-Za-z]+$/.test(m[1])) return;
  const prefix = m[1].toUpperCase();
  if (/[\d"]$/.test(before.slice(0, -m[1].length))) return;
  const items = FUNCTION_NAMES.filter((f) => f.startsWith(prefix)).slice(0, 12);
  const names = Object.keys(view().names || {}).filter((n) => n.startsWith(prefix)).slice(0, 4);
  const tables = view().sheets.flatMap((s) => (s.tables || []).map((t) => t.name)).filter((n) => n.toUpperCase().startsWith(prefix)).slice(0, 4);
  const all = items.map((f) => ({ label: f, insert: f + '(', help: fnHelp(f) })).concat(names.map((n) => ({ label: n, insert: n, help: 'Named range' })), tables.map((n) => ({ label: n, insert: n + '[', help: 'Table' })));
  if (!all.length || (all.length === 1 && all[0].label === prefix)) return;
  const list = h('div', { class: 'ac-list', role: 'listbox' });
  all.forEach((it, i) => {
    const row = h('div', { class: 'ac-item' + (i === 0 ? ' on' : ''), role: 'option' }, h('b', null, it.label), it.help ? h('span', { class: 'muted' }, it.help.replace(/^\w[\w.]*/, '')) : null);
    row.addEventListener('mousedown', (e) => { e.preventDefault(); acceptAutocomplete(input, it, m[1].length); });
    list.appendChild(row);
  });
  const rc = input.getBoundingClientRect();
  ac = popover(list, { left: rc.left, right: rc.right, top: rc.top, bottom: rc.bottom }, { keep: true, class: 'ac-pop' });
  ac._items = all; ac._index = 0; ac._len = m[1].length; ac._input = input;
}
function acceptAutocomplete(input, it, len) {
  const pos = input.selectionStart;
  input.value = input.value.slice(0, pos - len) + it.insert + input.value.slice(pos);
  const np = pos - len + it.insert.length;
  input.setSelectionRange(np, np);
  closeAutocomplete();
  editorInput(input);
}
function closeAutocomplete(keepHint) {
  if (ac) { ac.remove(); ac = null; }
  if (!keepHint) hideHint();
}
function acKey(e) {
  if (!ac) return false;
  const items = ac.querySelectorAll('.ac-item');
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    items[ac._index].classList.remove('on');
    ac._index = (ac._index + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
    items[ac._index].classList.add('on');
    return true;
  }
  if (e.key === 'Tab') { e.preventDefault(); acceptAutocomplete(ac._input, ac._items[ac._index], ac._len); return true; }
  if (e.key === 'Escape') { e.preventDefault(); closeAutocomplete(); return true; }
  return false;
}
function currentFunction(before) {
  let depth = 0;
  for (let i = before.length - 1; i >= 0; i--) {
    const ch = before[i];
    if (ch === '"') { const j = before.lastIndexOf('"', i - 1); if (j < 0) return null; i = j; continue; }
    if (ch === ')') depth++;
    else if (ch === '(') {
      if (depth === 0) { const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before.slice(0, i)); return m ? m[1].toUpperCase() : null; }
      depth--;
    }
  }
  return null;
}
let hintEl = null;
function showHint(input, fn) {
  const text = fnHelp(fn) || (FUNCTION_NAMES.includes(fn) ? fn + '(…)' : null);
  if (!text) { hideHint(); return; }
  if (!hintEl) { hintEl = h('div', { class: 'fn-hint' }); document.body.appendChild(hintEl); }
  hintEl.textContent = text;
  const rc = input.getBoundingClientRect();
  hintEl.style.left = rc.left + 'px';
  hintEl.style.top = Math.max(2, rc.top - 26) + 'px';
  hintEl.hidden = false;
}
function hideHint() { if (hintEl) hintEl.hidden = true; }

/* ================================================================ commands */

function applyStyle(patch) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const ops = M.styleOps(sh, selRect(), patch);
  commit(ops, { label: 'Format' });
}
function toggleStyle(key) { applyStyle({ [key]: curStyle()[key] ? null : 1 }); }
function setFormat(fmt) {
  const base = { nf: null, dp: null, cur: null, pat: null };
  applyStyle({ ...base, ...fmt });
}
function changeDecimals(delta) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const v = engine.valueAt(sh, S.sel.r, S.sel.c);
  const st = curStyle();
  let cur = st.dp;
  if (cur === undefined) {
    if (st.nf === 'num' || st.nf === 'cur' || st.nf === 'acc') cur = 2;
    else if (typeof v === 'number') { const t = numToText(v); cur = t.includes('.') ? Math.min(10, t.split('.')[1].length) : 0; }
    else cur = 0;
  }
  const dp = Math.max(0, Math.min(10, cur + delta));
  commit(M.styleOps(sh, selRect(), (s) => ({ dp, nf: s.nf || 'gen' })), { label: 'Decimals' });
}
function numberFormatItems() {
  const sample = (fmt) => { const v = 1234.5678; return formatValue(fmt.nf && ['date', 'ldate', 'dt', 'time'].includes(fmt.nf) ? 46000.5 : v, fmt, LOCALE); };
  const f = (label, fmt) => ({ label, kbd: sample(fmt), action: () => setFormat(fmt), checked: curStyle().nf === fmt.nf && (!fmt.cur || curStyle().cur === fmt.cur) });
  return [
    { label: 'Automatic', action: () => setFormat({}), checked: !curStyle().nf },
    f('Plain text', { nf: 'text' }),
    '-',
    f('Number', { nf: 'num', dp: 2 }),
    f('Percent', { nf: 'pct', dp: 2 }),
    f('Scientific', { nf: 'sci', dp: 2 }),
    '-',
    f('Accounting', { nf: 'acc', dp: 2 }),
    { label: 'Currency', sub: [...new Set([CURRENCY, '$', '€', '£', '₺', '¥'])].map((c) => f('Currency (' + c + ')', { nf: 'cur', cur: c, dp: 2 })) },
    '-',
    f('Date', { nf: 'date' }),
    f('Long date', { nf: 'ldate' }),
    f('Time', { nf: 'time' }),
    f('Date time', { nf: 'dt' }),
    f('Duration', { nf: 'dur' }),
    '-',
    { label: 'Custom format…', action: () => D.customFormatDialog(A, setFormat) }
  ];
}
function borderItems() {
  const go = (mode) => () => commit(M.borderOps(sheet(), selRect(), mode, S.borderColor || null), { label: 'Borders' });
  return [
    { label: 'All borders', action: go('all') }, { label: 'Outside borders', action: go('outer') }, { label: 'Inside borders', action: go('inner') },
    { label: 'Thick outside border', action: go('thick') },
    '-',
    { label: 'Top border', action: go('top') }, { label: 'Bottom border', action: go('bottom') }, { label: 'Left border', action: go('left') }, { label: 'Right border', action: go('right') },
    '-',
    { label: 'No borders', action: go('none') },
    { label: 'Border color…', swatch: S.borderColor || '#000000', action: () => colorPicker($('toolbar').getBoundingClientRect(), S.borderColor, (c) => { S.borderColor = c; }, 'Default') }
  ];
}
function mergeItems() {
  return [
    { label: 'Merge all', action: () => mergeCells('all') },
    { label: 'Merge and center', action: () => mergeCells('center') },
    { label: 'Merge horizontally', action: () => mergeCells('across') },
    { label: 'Unmerge', action: () => mergeCells('none') }
  ];
}
async function mergeCells(mode) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const rect = expandForMerges(sh, selRect());
  const x = ix(sh);
  let merges = sh.merges.filter((m) => { const r = M.objRect(m, x); return r && (r.r2 < rect.r1 || r.r1 > rect.r2 || r.c2 < rect.c1 || r.c1 > rect.c2); });
  const ops = [];
  if (mode !== 'none') {
    const parts = mode === 'across' ? Array.from({ length: rect.r2 - rect.r1 + 1 }, (_, i) => ({ r1: rect.r1 + i, r2: rect.r1 + i, c1: rect.c1, c2: rect.c2 })) : [rect];
    let lost = false;
    for (const p of parts) {
      if (p.r1 === p.r2 && p.c1 === p.c2) continue;
      merges.push(M.rectIds(sh, p));
      for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) {
        if (r === p.r1 && c === p.c1) continue;
        const cell = M.cellAt(sh, r, c);
        if (cell && cell.v !== undefined) { lost = true; ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), v: null }); }
      }
    }
    if (lost && !(await confirmDialog('Merge cells', 'Merging keeps only the value of the top-left cell. The other values will be deleted.', 'Merge'))) return;
    if (mode === 'center') ops.push(...M.styleOps(sh, { r1: rect.r1, c1: rect.c1, r2: rect.r1, c2: rect.c1 }, { ha: 'center', va: 'middle' }));
  }
  ops.push({ t: 'sheet', s: sh.id, merges });
  commit(ops, { label: mode === 'none' ? 'Unmerge' : 'Merge' });
}

function togglePainter() {
  if (S.painter) { S.painter = null; renderToolbarState(); return; }
  S.painter = { st: curStyle() };
  renderToolbarState();
  toast('Select the cells to apply the formatting to.');
}
function applyPainter() {
  const p = S.painter;
  S.painter = null;
  const sh = sheet();
  const ops = [];
  for (let r = S.sel.r1; r <= S.sel.r2; r++) for (let c = S.sel.c1; c <= S.sel.c2; c++) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), st: Object.keys(p.st).length ? p.st : null });
  commit(ops, { label: 'Paste format' });
}

function clearCells(what = 'values') {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const ops = [];
  const s = selRect();
  for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
    const k = M.keyAt(sh, r, c);
    const cell = sh.cells[k];
    if (!cell) continue;
    const dv = M.dvFor(sh, r, c, ix(sh));
    if (what === 'values' && cell.v !== undefined) ops.push({ t: 'set', s: sh.id, k, v: dv && dv.type === 'checkbox' ? 'FALSE' : null });
    else if (what === 'formats' && cell.st) ops.push({ t: 'set', s: sh.id, k, st: null });
    else if (what === 'notes' && cell.n) ops.push({ t: 'set', s: sh.id, k, n: null });
    else if (what === 'all') ops.push({ t: 'set', s: sh.id, k, v: null, st: null, n: null });
  }
  commit(ops, { label: 'Clear' });
}

function toggleCheckbox(r, c) {
  const sh = sheet();
  if (!canEdit()) return;
  const v = engine.valueAt(sh, r, c);
  const cell = M.cellAt(sh, r, c);
  if (cell && cell.v && cell.v[0] === '=') return;
  commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), v: v === true ? 'FALSE' : 'TRUE' }], { label: 'Checkbox' });
}

/** Insert or delete rows/columns of the selection. */
function structure(kind) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const s = selRect();
  const sid = sh.id;
  let axis, at, count;
  switch (kind) {
    case 'rowAbove': axis = 'r'; at = s.r1; count = s.r2 - s.r1 + 1; break;
    case 'rowBelow': axis = 'r'; at = s.r2 + 1; count = s.r2 - s.r1 + 1; break;
    case 'colLeft': axis = 'c'; at = s.c1; count = s.c2 - s.c1 + 1; break;
    case 'colRight': axis = 'c'; at = s.c2 + 1; count = s.c2 - s.c1 + 1; break;
    case 'delRows': axis = 'r'; at = s.r1; count = -(s.r2 - s.r1 + 1); break;
    case 'delCols': axis = 'c'; at = s.c1; count = -(s.c2 - s.c1 + 1); break;
    default: return;
  }
  const list = axis === 'r' ? sh.rows : sh.cols;
  if (count < 0 && -count >= list.length) { toast('A sheet needs at least one ' + (axis === 'r' ? 'row.' : 'column.'), 'error'); return; }
  const make = () => {
    const cur = view().sheets.find((x) => x.id === sid);
    return cur ? M.structuralOps(view(), cur, axis, at, count) : [];
  };
  if (commit(make(), { label: kind, regen: make })) {
    if (count < 0) clampSelection();
    else if (kind === 'rowBelow') selectRect({ r1: s.r2 + 1, c1: s.c1, r2: s.r2 + count, c2: s.c2 });
    else if (kind === 'colRight') selectRect({ r1: s.r1, c1: s.c2 + 1, r2: s.r2, c2: s.c2 + count });
    renderAll();
  }
}

function setHidden(axis, hide) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const s = selRect();
  const patch = {};
  if (axis === 'r') for (let r = s.r1; r <= s.r2; r++) patch[sh.rows[r]] = hide ? 1 : null;
  else for (let c = s.c1; c <= s.c2; c++) patch[sh.cols[c]] = hide ? 1 : null;
  if (hide) {
    const list = axis === 'r' ? sh.rows : sh.cols, hidden = axis === 'r' ? sh.hr : sh.hc;
    if (list.every((id) => hidden[id] || patch[id])) { toast('At least one ' + (axis === 'r' ? 'row' : 'column') + ' must stay visible.', 'error'); return; }
  }
  commit([{ t: 'sheet', s: sh.id, [axis === 'r' ? 'hr' : 'hc']: patch }], { label: hide ? 'Hide' : 'Unhide' });
}
function unhideAll(axis) {
  const sh = sheet();
  const hidden = axis === 'r' ? sh.hr : sh.hc;
  const patch = {};
  Object.keys(hidden).forEach((id) => { patch[id] = null; });
  commit([{ t: 'sheet', s: sh.id, [axis === 'r' ? 'hr' : 'hc']: patch }], { label: 'Unhide' });
}

function freeze(rows, cols) {
  const sh = sheet();
  commit([{ t: 'sheet', s: sh.id, freeze: { r: rows ?? sh.freeze.r, c: cols ?? sh.freeze.c } }], { label: 'Freeze' });
}

function autoFitColumns(cols) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const ctx = document.createElement('canvas').getContext('2d');
  const font = getComputedStyle($('grid')).fontFamily;
  const patch = {};
  for (const c of cols) {
    let w = 30;
    for (let r = 0; r < sh.rows.length; r++) {
      const cell = M.cellAt(sh, r, c);
      if (!cell && !engine.spillAnchorOf(sh, r, c)) continue;
      const st = (cell && cell.st) || {};
      ctx.font = `${st.b ? '600 ' : ''}${st.i ? 'italic ' : ''}${st.fs || 13}px ${font}`;
      const text = formatValue(engine.valueAt(sh, r, c), st, LOCALE);
      const lines = st.w ? [text] : String(text).split('\n');
      for (const line of lines) w = Math.max(w, ctx.measureText(line).width + 14);
    }
    patch[sh.cols[c]] = Math.min(600, Math.ceil(w));
  }
  commit([{ t: 'sheet', s: sh.id, colW: patch }], { label: 'Column width' });
}
function autoFitRows(rows) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const ctx = document.createElement('canvas').getContext('2d');
  const font = getComputedStyle($('grid')).fontFamily;
  const patch = {};
  for (const r of rows) {
    let hgt = Z_ROW_H;
    for (let c = 0; c < sh.cols.length; c++) {
      const cell = M.cellAt(sh, r, c);
      if (!cell || cell.v === undefined) continue;
      const st = cell.st || {};
      const size = st.fs || 13;
      ctx.font = `${st.b ? '600 ' : ''}${size}px ${font}`;
      const text = formatValue(engine.valueAt(sh, r, c), st, LOCALE);
      let lines = String(text).split('\n').length;
      if (st.w) {
        const w = (sh.colW[sh.cols[c]] || Z_COL_W) - 8;
        lines = String(text).split('\n').reduce((n, line) => n + Math.max(1, Math.ceil(ctx.measureText(line).width / w)), 0);
      }
      hgt = Math.max(hgt, Math.ceil(lines * size * 1.35 + 6));
    }
    patch[sh.rows[r]] = hgt === Z_ROW_H ? null : Math.min(600, hgt);
  }
  commit([{ t: 'sheet', s: sh.id, rowH: patch }], { label: 'Row height' });
}

function autoSum(fn = 'SUM') {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const s = selRect();
  // a selected block: put totals under each column
  if (s.r2 > s.r1 || s.c2 > s.c1) {
    const ops = [];
    let r = s.r2 + 1;
    ops.push(...M.growOps(sh, r, s.c2));
    for (let c = s.c1; c <= s.c2; c++) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, c) || (sh.rows[r] + ':' + sh.cols[c]), v: `=${fn}(${M.rectText({ r1: s.r1, c1: c, r2: s.r2, c2: c })})` });
    if (sh.rows[r]) commit(ops, { label: 'AutoSum' });
    return;
  }
  // a single cell: guess the numbers above (or to the left)
  const isNum = (r, c) => typeof engine.valueAt(sh, r, c) === 'number';
  let r1 = s.r1 - 1;
  while (r1 >= 0 && isNum(r1, s.c1)) r1--;
  if (r1 < s.r1 - 1) { startEdit(`=${fn}(${M.rectText({ r1: r1 + 1, c1: s.c1, r2: s.r1 - 1, c2: s.c1 })})`, 'edit'); return; }
  let c1 = s.c1 - 1;
  while (c1 >= 0 && isNum(s.r1, c1)) c1--;
  if (c1 < s.c1 - 1) { startEdit(`=${fn}(${M.rectText({ r1: s.r1, c1: c1 + 1, r2: s.r1, c2: s.c1 - 1 })})`, 'edit'); return; }
  startEdit(`=${fn}()`, 'edit');
  const ed = $('cellEditor');
  ed.setSelectionRange(ed.value.length - 1, ed.value.length - 1);
}

function quickSort(desc) {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const x = ix(sh);
  const t = (sh.tables || []).find((tb) => M.inRect(M.objRect(tb, x), S.sel.r, S.sel.c));
  let rect, header;
  if (t) { const r = M.objRect(t, x); rect = { r1: r.r1, c1: r.c1, r2: r.r2 - (t.totals ? 1 : 0), c2: r.c2 }; header = t.header !== false; }
  else if (S.sel.r1 === S.sel.r2 && S.sel.c1 === S.sel.c2) {
    rect = M.currentRegion(sh, S.sel.r, S.sel.c);
    const first = engine.valueAt(sh, rect.r1, S.sel.c), second = engine.valueAt(sh, rect.r1 + 1, S.sel.c);
    header = typeof first === 'string' && typeof second !== typeof first || (typeof first === 'string' && (sh.cells[M.keyAt(sh, rect.r1, S.sel.c)] || {}).st?.b);
  } else { rect = selRect(); header = false; }
  const ops = M.sortOps(engine, sh, rect, [{ c: S.sel.c, desc }], header);
  commit(ops, { label: 'Sort' });
}

function toggleFilter() {
  const sh = sheet();
  if (!sh || !canEdit()) return;
  const x = ix(sh);
  const t = (sh.tables || []).find((tb) => M.inRect(M.objRect(tb, x), S.sel.r, S.sel.c));
  if (t) { commit([{ t: 'obj', s: sh.id, kind: 'tables', id: t.id, val: { ...t, filter: t.filter === false, filters: t.filter === false ? t.filters : {} } }], { label: 'Filter' }); return; }
  const rect = S.sel.r1 === S.sel.r2 && S.sel.c1 === S.sel.c2 ? M.currentRegion(sh, S.sel.r, S.sel.c) : selRect();
  if (rect.r1 === rect.r2) { toast('Select a range with a header row first.'); return; }
  const id = Core.randomId();
  commit([{ t: 'obj', s: sh.id, kind: 'tables', id, val: { name: M.uniqueTableName(view(), 'Filter'), ...M.rectIds(sh, rect), header: true, totals: false, style: 'none', banded: false, filter: true, filters: {} } }], { label: 'Filter' });
}

/* ---- clipboard */

function onCopy(e, cut) {
  if (isTextTarget(e.target)) return;
  const sh = sheet();
  if (!sh || S.editing) return;
  if (S.chartSel) return;
  e.preventDefault();
  const rect = expandForMerges(sh, selRect());
  const res = M.copyRect(engine, sh, rect, LOCALE, cut);
  e.clipboardData.setData('text/plain', res.text);
  e.clipboardData.setData('text/html', res.html);
  S.clip = { text: res.text, html: res.html };
  S.ants = { sheetId: sh.id, rect, cut: !!cut && canEdit() };
  renderGrid();
}
function onPaste(e) {
  if (isTextTarget(e.target)) return;
  if (!sheet() || S.editing) return;
  e.preventDefault();
  if (!canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  const html = e.clipboardData.getData('text/html');
  const text = e.clipboardData.getData('text/plain');
  pasteData(html, text, S.pasteMode || 'all');
  S.pasteMode = null;
}
function pasteData(html, text, mode, transpose) {
  const sh = sheet();
  const parsed = M.parseClipboardHtml(html);
  let data = null;
  if (parsed && parsed.ours) data = parsed.ours;
  else if (parsed && parsed.rows && parsed.rows.length) data = { rows: parsed.rows };
  else if (text) {
    const rows = M.parseTSV(text);
    data = { rows: rows.map((r) => r.map((t) => ({ text: t, st: null }))) };
  }
  if (!data) return;
  const target = selRect();
  const ops = M.pasteOps(engine, view(), sh, target, data, { mode, transpose, locale: LOCALE });
  if (data === M.getInternalClip() && data.cut && mode === 'all') {
    ops.push(...M.cutClearOps(view(), data, sh.id, target));
    M.clearInternalClip();
    S.ants = null;
  }
  if (parsed && parsed.spans && parsed.spans.length && mode === 'all') {
    const merges = sh.merges.slice();
    parsed.spans.forEach((sp) => { if (sh.rows[target.r1 + sp.r2] && sh.cols[target.c1 + sp.c2]) merges.push(M.rectIds(sh, { r1: target.r1 + sp.r1, c1: target.c1 + sp.c1, r2: target.r1 + sp.r2, c2: target.c1 + sp.c2 })); });
    ops.push({ t: 'sheet', s: sh.id, merges });
  }
  if (commit(ops, { label: 'Paste', keepAnts: true })) {
    const h2 = (transpose ? data.rows[0].length : data.rows.length), w2 = (transpose ? data.rows.length : data.rows[0].length);
    selectRect({ r1: target.r1, c1: target.c1, r2: Math.max(target.r2, target.r1 + h2 - 1), c2: Math.max(target.c2, target.c1 + w2 - 1) }, { r: target.r1, c: target.c1 });
  }
}
async function pasteSpecial(mode, transpose) {
  if (!canEdit()) return;
  const clip = M.getInternalClip();
  if (clip) {
    const sh = sheet();
    const ops = M.pasteOps(engine, view(), sh, selRect(), clip, { mode, transpose, locale: LOCALE });
    commit(ops, { label: 'Paste special', keepAnts: true });
    return;
  }
  try {
    const items = await navigator.clipboard.read();
    let html = '', text = '';
    for (const it of items) {
      if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
      if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
    }
    pasteData(html, text, mode, transpose);
  } catch (e) {
    if (S.clip) { pasteData(S.clip.html, S.clip.text, mode, transpose); return; }
    S.pasteMode = mode;
    toast('Press Ctrl+V to paste.');
  }
}
function isTextTarget(t) {
  if (!t) return false;
  if (t.id === 'kbd') return false;
  return t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable;
}
async function clipboardAction(kind) {
  focusGrid();
  if (kind === 'paste') {
    try {
      const items = await navigator.clipboard.read();
      let html = '', text = '';
      for (const it of items) {
        if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
        if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
      }
      if (!canEdit()) return;
      pasteData(html, text, 'all');
    } catch (e) {
      // the browser blocks reading the clipboard in the frame: paste what was copied in this spreadsheet
      if (S.clip && canEdit()) pasteData(S.clip.html, S.clip.text, 'all');
      else toast('Your browser blocks pasting from the menu here. Use Ctrl+V instead.');
    }
    return;
  }
  const sh = sheet();
  const rect = expandForMerges(sh, selRect());
  const res = M.copyRect(engine, sh, rect, LOCALE, kind === 'cut');
  // copied for this spreadsheet in any case; the system clipboard (other apps) only when the browser allows it
  S.clip = { text: res.text, html: res.html };
  S.ants = { sheetId: sh.id, rect, cut: kind === 'cut' && canEdit() };
  renderGrid();
  try {
    await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([res.text], { type: 'text/plain' }), 'text/html': new Blob([res.html], { type: 'text/html' }) })]);
  } catch (e) {
    toast(`Copied for this spreadsheet. To paste into other apps, use ${kind === 'cut' ? 'Ctrl+X' : 'Ctrl+C'} instead.`);
  }
}

/* ---- fill */

function fillTo(dest, copyOnly) {
  const sh = sheet();
  const src = selRect();
  let dir;
  if (dest.r2 > src.r2) dir = 'down';
  else if (dest.r1 < src.r1) dir = 'up';
  else if (dest.c2 > src.c2) dir = 'right';
  else if (dest.c1 < src.c1) dir = 'left';
  else return;
  const area = dir === 'down' ? { r1: src.r2 + 1, r2: dest.r2, c1: src.c1, c2: src.c2 }
    : dir === 'up' ? { r1: dest.r1, r2: src.r1 - 1, c1: src.c1, c2: src.c2 }
      : dir === 'right' ? { r1: src.r1, r2: src.r2, c1: src.c2 + 1, c2: dest.c2 }
        : { r1: src.r1, r2: src.r2, c1: dest.c1, c2: src.c1 - 1 };
  if (commit(M.fillOps(sh, src, area, dir, copyOnly), { label: 'Fill' })) selectRect({ r1: Math.min(src.r1, dest.r1), c1: Math.min(src.c1, dest.c1), r2: Math.max(src.r2, dest.r2), c2: Math.max(src.c2, dest.c2) }, { r: S.sel.r, c: S.sel.c });
}
function fillDirection(dir) {
  const s = selRect();
  const sh = sheet();
  if (dir === 'down') {
    if (s.r1 === s.r2) { if (s.r1 === 0) return; commit(M.fillOps(sh, { ...s, r1: s.r1 - 1, r2: s.r1 - 1 }, s, 'down', true), { label: 'Fill down' }); }
    else commit(M.fillOps(sh, { ...s, r2: s.r1 }, { ...s, r1: s.r1 + 1 }, 'down', true), { label: 'Fill down' });
  } else {
    if (s.c1 === s.c2) { if (s.c1 === 0) return; commit(M.fillOps(sh, { ...s, c1: s.c1 - 1, c2: s.c1 - 1 }, s, 'right', true), { label: 'Fill right' }); }
    else commit(M.fillOps(sh, { ...s, c2: s.c1 }, { ...s, c1: s.c1 + 1 }, 'right', true), { label: 'Fill right' });
  }
}

/* ================================================================ mouse */

function cellFromEvent(e) {
  const td = e.target.closest && e.target.closest('td[data-r]');
  if (!td) return null;
  return { r: +td.dataset.r, c: +td.dataset.c, td };
}
/** Cell under a point, also outside the rendered cells (for drag selection). */
function cellAtPoint(x, y) {
  const el = document.elementFromPoint(x, y);
  const td = el && el.closest && el.closest('td[data-r]');
  if (td) return { r: +td.dataset.r, c: +td.dataset.c };
  const L = layout();
  const wrap = $('gridWrap').getBoundingClientRect();
  const sx = x - wrap.left + $('gridWrap').scrollLeft - L.rhW, sy = y - wrap.top + $('gridWrap').scrollTop - L.hdrH - L.frozenH + L.rowTop[L.fr];
  let r = 0, c = 0;
  while (r < L.sh.rows.length - 1 && L.rowTop[r + 1] <= sy) r++;
  while (c < L.sh.cols.length - 1 && L.colLeft[c + 1] <= sx) c++;
  return { r, c };
}

let drag = null;
let lastCellDown = null; // { r, c, sheetId, t, x, y } of the last click on a cell, to recognise a double-click
function onGridMouseDown(e) {
  if (e.button === 2) return;
  const t = e.target;
  // charts
  const chartEl = t.closest && t.closest('.chart');
  if (chartEl) { chartMouseDown(e, chartEl); return; }
  const fltBtn = t.closest && t.closest('[data-flt]');
  if (fltBtn) { e.preventDefault(); if (S.editing) commitEdit(); D.filterPopover(A, fltBtn.dataset.flt, +fltBtn.dataset.fc, fltBtn.getBoundingClientRect()); return; }
  if (t.closest && t.closest('[data-dv]')) { e.preventDefault(); openListPicker(); return; }
  if (t.closest && t.closest('[data-cbx]')) {
    const cc = cellFromEvent(e);
    if (cc) { e.preventDefault(); selectCell(cc.r, cc.c); toggleCheckbox(cc.r, cc.c); focusGrid(); }
    return;
  }
  if (t.dataset && t.dataset.fill) { e.preventDefault(); startFillDrag(e); return; }
  if (t.dataset && (t.dataset.cres !== undefined)) { e.preventDefault(); startResize(e, 'c', +t.dataset.cres); return; }
  if (t.dataset && (t.dataset.rres !== undefined)) { e.preventDefault(); startResize(e, 'r', +t.dataset.rres); return; }
  if (t.closest && t.closest('[data-link]') && (e.ctrlKey || e.metaKey)) return;
  const sh = sheet();
  const L = layout();
  // headers
  const th = t.closest && t.closest('th');
  if (th) {
    e.preventDefault();
    if (S.editing && S.editing.text[0] === '=' && (th.dataset.ch || th.dataset.rh)) {
      if (th.dataset.ch !== undefined) insertRef(0, +th.dataset.ch, sh.rows.length - 1, +th.dataset.ch);
      return;
    }
    if (S.editing) commitEdit();
    if (th.dataset.corner) { selectRect({ r1: 0, c1: 0, r2: sh.rows.length - 1, c2: sh.cols.length - 1 }, { r: S.sel.r, c: S.sel.c }); focusGrid(); return; }
    if (th.dataset.ch !== undefined) {
      const c = +th.dataset.ch;
      const from = e.shiftKey ? S.sel.c : c;
      selectRect({ r1: 0, r2: sh.rows.length - 1, c1: Math.min(from, c), c2: Math.max(from, c) }, { r: firstVisibleRow(), c: from });
      drag = { kind: 'cols', from };
    } else if (th.dataset.rh !== undefined) {
      const r = +th.dataset.rh;
      const from = e.shiftKey ? S.sel.r : r;
      selectRect({ c1: 0, c2: sh.cols.length - 1, r1: Math.min(from, r), r2: Math.max(from, r) }, { r: from, c: 0 });
      drag = { kind: 'rows', from };
    }
    focusGrid();
    return;
  }
  const hit = cellFromEvent(e);
  if (!hit) return;
  e.preventDefault();
  // a second click on the same cell within the double-click time edits the cell. The browser's dblclick event is not
  // reliable here because the grid is redrawn after the first click.
  const now = Date.now();
  const again = !e.shiftKey && !e.ctrlKey && !e.metaKey && lastCellDown && lastCellDown.sheetId === S.sheetId && lastCellDown.r === hit.r &&
    lastCellDown.c === hit.c && now - lastCellDown.t < 500 && Math.abs(e.clientX - lastCellDown.x) < 6 && Math.abs(e.clientY - lastCellDown.y) < 6;
  lastCellDown = again ? null : { r: hit.r, c: hit.c, sheetId: S.sheetId, t: now, x: e.clientX, y: e.clientY };
  if (again && !S.editing) {
    drag = null;
    if (!canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
    startEdit(null, 'edit');
    return;
  }
  if (S.editing) {
    const ed = S.editing;
    const input = ed.from === 'bar' || ed.sheetId !== S.sheetId ? $('fbar') : $('cellEditor');
    if (ed.text[0] === '=' && canInsertRef(input.value, input.selectionStart) && !(ed.sheetId === sh.id && hit.r === ed.r && hit.c === ed.c)) {
      insertRef(hit.r, hit.c, hit.r, hit.c);
      drag = { kind: 'ref', r: hit.r, c: hit.c };
      return;
    }
    if (ed.sheetId === sh.id && hit.r === ed.r && hit.c === ed.c) return;
    if (!commitEdit()) return;
  }
  if (e.shiftKey) selectCell(hit.r, hit.c, true);
  else selectCell(hit.r, hit.c);
  drag = { kind: 'cells' };
  S.painterDragging = !!S.painter;
  focusGrid();
  void L;
}
function firstVisibleRow() { const L = layout(); let r = 0; while (r < L.sh.rows.length - 1 && L.rowH[r] === 0) r++; return r; }

function onMouseMove(e) {
  if (!drag) return;
  autoScroll(e);
  const sh = sheet();
  const p = cellAtPoint(e.clientX, e.clientY);
  if (!p) return;
  if (drag.kind === 'cells') {
    if (p.r !== drag.lr || p.c !== drag.lc) { drag.lr = p.r; drag.lc = p.c; selectCell(p.r, p.c, true); }
  } else if (drag.kind === 'cols') {
    selectRect({ r1: 0, r2: sh.rows.length - 1, c1: Math.min(drag.from, p.c), c2: Math.max(drag.from, p.c) }, { r: S.sel.r, c: drag.from });
  } else if (drag.kind === 'rows') {
    selectRect({ c1: 0, c2: sh.cols.length - 1, r1: Math.min(drag.from, p.r), r2: Math.max(drag.from, p.r) }, { r: drag.from, c: 0 });
  } else if (drag.kind === 'ref') {
    if (p.r !== drag.lr || p.c !== drag.lc) { drag.lr = p.r; drag.lc = p.c; insertRef(drag.r, drag.c, p.r, p.c, true); }
  } else if (drag.kind === 'fill') {
    const s = selRect();
    const dy = p.r > s.r2 ? p.r - s.r2 : p.r < s.r1 ? p.r - s.r1 : 0;
    const dx = p.c > s.c2 ? p.c - s.c2 : p.c < s.c1 ? p.c - s.c1 : 0;
    let prev;
    if (Math.abs(dy) >= Math.abs(dx) && dy) prev = dy > 0 ? { ...s, r2: p.r } : { ...s, r1: p.r };
    else if (dx) prev = dx > 0 ? { ...s, c2: p.c } : { ...s, c1: p.c };
    else prev = null;
    if (JSON.stringify(prev) !== JSON.stringify(S.fillPreview)) { S.fillPreview = prev; renderGrid(); }
  } else if (drag.kind === 'resize') {
    const L = layout();
    const delta = (drag.axis === 'c' ? e.clientX : e.clientY) - drag.start;
    const size = Math.max(drag.axis === 'c' ? 16 : 10, Math.round((drag.size + delta) / L.z));
    drag.newSize = size;
    const guide = $('resizeGuide');
    if (drag.axis === 'c') guide.style.left = (drag.edge + delta) + 'px';
    else guide.style.top = (drag.edge + delta) + 'px';
  } else if (drag.kind === 'chartMove' || drag.kind === 'chartSize') chartMouseMove(e);
}
function onMouseUp(e) {
  if (!drag) return;
  stopAutoScroll();
  const d = drag;
  drag = null;
  if (d.kind === 'cells' && S.painterDragging) { S.painterDragging = false; if (S.painter) { applyPainter(); renderAll(); } }
  if (d.kind === 'fill') {
    const prev = S.fillPreview;
    S.fillPreview = null;
    if (prev) fillTo(prev, e.ctrlKey || e.metaKey); else renderGrid();
  }
  if (d.kind === 'resize') {
    $('resizeGuide').hidden = true;
    if (d.newSize && d.newSize !== Math.round(d.size / layout().z)) {
      const sh = sheet();
      const s = selRect();
      const patch = {};
      const full = d.axis === 'c' ? s.r1 === 0 && s.r2 === sh.rows.length - 1 && d.index >= s.c1 && d.index <= s.c2 : s.c1 === 0 && s.c2 === sh.cols.length - 1 && d.index >= s.r1 && d.index <= s.r2;
      const range = full ? (d.axis === 'c' ? [s.c1, s.c2] : [s.r1, s.r2]) : [d.index, d.index];
      for (let i = range[0]; i <= range[1]; i++) patch[(d.axis === 'c' ? sh.cols : sh.rows)[i]] = d.newSize;
      commit([{ t: 'sheet', s: sh.id, [d.axis === 'c' ? 'colW' : 'rowH']: patch }], { label: 'Resize' });
    }
  }
}

let scrollTimer = null;
function autoScroll(e) {
  const wrap = $('gridWrap');
  const r = wrap.getBoundingClientRect();
  const L = layout();
  let dx = 0, dy = 0;
  if (e.clientY > r.bottom - 10) dy = 20; else if (e.clientY < r.top + L.hdrH + L.frozenH + 4 && drag.kind !== 'resize') dy = -20;
  if (e.clientX > r.right - 10) dx = 20; else if (e.clientX < r.left + L.rhW + L.frozenW + 4 && drag.kind !== 'resize') dx = -20;
  stopAutoScroll();
  if ((dx || dy) && drag && drag.kind !== 'resize' && !drag.kind.startsWith('chart')) {
    scrollTimer = setInterval(() => { wrap.scrollTop += dy; wrap.scrollLeft += dx; onMouseMove(e); }, 50);
  }
}
function stopAutoScroll() { clearInterval(scrollTimer); scrollTimer = null; }

function startFillDrag(e) { drag = { kind: 'fill' }; S.fillPreview = null; }
function startResize(e, axis, index) {
  if (!canEdit()) return;
  const L = layout();
  const wrap = $('gridWrap');
  const wr = wrap.getBoundingClientRect();
  let guide = $('resizeGuide');
  if (!guide) { guide = h('div', { id: 'resizeGuide', class: 'resize-guide' }); $('gridInner').appendChild(guide); }
  guide.className = 'resize-guide ' + (axis === 'c' ? 'v' : 'h');
  guide.hidden = false;
  const thEl = e.target.closest('th');
  const tr = thEl.getBoundingClientRect();
  const ir = $('gridInner').getBoundingClientRect();
  if (axis === 'c') { guide.style.left = (tr.right - ir.left) + 'px'; guide.style.top = wrap.scrollTop + 'px'; guide.style.height = wr.height + 'px'; }
  else { guide.style.top = (tr.bottom - ir.top) + 'px'; guide.style.left = wrap.scrollLeft + 'px'; guide.style.width = wr.width + 'px'; }
  drag = { kind: 'resize', axis, index, start: axis === 'c' ? e.clientX : e.clientY, size: axis === 'c' ? L.colW[index] : L.rowH[index], edge: axis === 'c' ? tr.right - ir.left : tr.bottom - ir.top };
}

function onGridDblClick(e) {
  const t = e.target;
  if (t.dataset && t.dataset.cres !== undefined) {
    const s = selRect(), sh = sheet(), c = +t.dataset.cres;
    const full = s.r1 === 0 && s.r2 === sh.rows.length - 1 && c >= s.c1 && c <= s.c2;
    autoFitColumns(full ? Array.from({ length: s.c2 - s.c1 + 1 }, (_, i) => s.c1 + i) : [c]);
    return;
  }
  if (t.dataset && t.dataset.rres !== undefined) {
    const s = selRect(), sh = sheet(), r = +t.dataset.rres;
    const full = s.c1 === 0 && s.c2 === sh.cols.length - 1 && r >= s.r1 && r <= s.r2;
    autoFitRows(full ? Array.from({ length: s.r2 - s.r1 + 1 }, (_, i) => s.r1 + i) : [r]);
    return;
  }
  const chartEl = t.closest && t.closest('.chart');
  if (chartEl) { if (canEdit()) D.chartDialog(A, chartEl.dataset.chart); return; }
  if (t.dataset && t.dataset.fill) {
    // double-click the fill handle: fill down as far as the column to the left has data
    const sh = sheet(), s = selRect();
    let r = s.r2;
    const ref = s.c1 > 0 ? s.c1 - 1 : s.c2 + 1;
    while (r + 1 < sh.rows.length && engine.valueAt(sh, r + 1, ref) !== null) r++;
    if (r > s.r2) fillTo({ ...s, r2: r }, false);
    return;
  }
  const hit = cellFromEvent(e);
  if (!hit || S.editing) return;
  if (t.closest('[data-cbx]')) return;
  startEdit(null, 'edit');
}

function onGridContextMenu(e) {
  const t = e.target;
  if (t.closest('.chart')) return;
  e.preventDefault();
  const th = t.closest('th');
  const hit = cellFromEvent(e);
  const sh = sheet();
  if (hit && !M.inRect(selRect(), hit.r, hit.c)) selectCell(hit.r, hit.c);
  if (th && th.dataset.ch !== undefined && !(S.sel.r1 === 0 && S.sel.r2 === sh.rows.length - 1 && +th.dataset.ch >= S.sel.c1 && +th.dataset.ch <= S.sel.c2)) selectRect({ r1: 0, r2: sh.rows.length - 1, c1: +th.dataset.ch, c2: +th.dataset.ch });
  if (th && th.dataset.rh !== undefined && !(S.sel.c1 === 0 && S.sel.c2 === sh.cols.length - 1 && +th.dataset.rh >= S.sel.r1 && +th.dataset.rh <= S.sel.r2)) selectRect({ c1: 0, c2: sh.cols.length - 1, r1: +th.dataset.rh, r2: +th.dataset.rh });
  const pt = { left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY };
  menu(contextItems(th ? (th.dataset.ch !== undefined ? 'col' : th.dataset.rh !== undefined ? 'row' : 'cell') : 'cell'), pt, { onDone: () => focusGrid(), hideDisabled: !canEdit() });
}

function contextItems(kind) {
  const ro = !canEdit();
  const s = selRect();
  const nR = s.r2 - s.r1 + 1, nC = s.c2 - s.c1 + 1;
  const sh = sheet();
  const cell = M.cellAt(sh, S.sel.r, S.sel.c);
  const items = [
    { label: 'Cut', kbd: 'Ctrl+X', icon: null, action: () => clipboardAction('cut'), disabled: ro },
    { label: 'Copy', kbd: 'Ctrl+C', action: () => clipboardAction('copy') },
    { label: 'Paste', kbd: 'Ctrl+V', action: () => clipboardAction('paste'), disabled: ro },
    { label: 'Paste special', disabled: ro, sub: [
      { label: 'Values only', kbd: 'Ctrl+Shift+V', action: () => pasteSpecial('values') },
      { label: 'Formulas only', action: () => pasteSpecial('formulas') },
      { label: 'Formatting only', action: () => pasteSpecial('formats') },
      { label: 'Transposed', action: () => pasteSpecial('all', true) }] },
    '-'
  ];
  // inserting and deleting rows and columns are submenus, so that the menu fits a small frame
  const ins = [], del = [];
  if (kind !== 'col') ins.push({ label: `${nR === 1 ? 'Row' : nR + ' rows'} above`, action: () => structure('rowAbove') }, { label: `${nR === 1 ? 'Row' : nR + ' rows'} below`, action: () => structure('rowBelow') });
  if (kind !== 'row') ins.push({ label: `${nC === 1 ? 'Column' : nC + ' columns'} left`, action: () => structure('colLeft') }, { label: `${nC === 1 ? 'Column' : nC + ' columns'} right`, action: () => structure('colRight') });
  if (kind !== 'col') del.push({ label: nR === 1 ? 'Row ' + (s.r1 + 1) : 'Rows ' + (s.r1 + 1) + '–' + (s.r2 + 1), action: () => structure('delRows') });
  if (kind !== 'row') del.push({ label: nC === 1 ? 'Column ' + colName(s.c1) : 'Columns ' + colName(s.c1) + '–' + colName(s.c2), action: () => structure('delCols') });
  items.push({ label: 'Insert', sub: ins, disabled: ro }, { label: 'Delete', sub: del, disabled: ro });
  items.push('-');
  if (kind === 'row') items.push({ label: 'Hide rows', action: () => setHidden('r', true), disabled: ro }, { label: 'Unhide rows', action: () => unhideAll('r'), disabled: ro }, { label: 'Row height…', action: () => D.sizeDialog(A, 'r'), disabled: ro }, { label: 'Fit to data', action: () => autoFitRows(range(s.r1, s.r2)), disabled: ro }, '-');
  if (kind === 'col') items.push({ label: 'Hide columns', action: () => setHidden('c', true), disabled: ro }, { label: 'Unhide columns', action: () => unhideAll('c'), disabled: ro }, { label: 'Column width…', action: () => D.sizeDialog(A, 'c'), disabled: ro }, { label: 'Fit to data', action: () => autoFitColumns(range(s.c1, s.c2)), disabled: ro }, '-');
  items.push(
    { label: 'Clear contents', kbd: 'Delete', action: () => clearCells('values'), disabled: ro },
    { label: 'Clear formatting', action: () => clearCells('formats'), disabled: ro },
    '-',
    { label: 'Sort A → Z', action: () => quickSort(false), disabled: ro },
    { label: 'Sort Z → A', action: () => quickSort(true), disabled: ro },
    { label: cell && cell.n ? 'Edit note' : 'Insert note', kbd: 'Shift+F2', action: () => D.noteDialog(A), disabled: ro },
    { label: 'Data validation…', action: () => D.dvDialog(A), disabled: ro },
    { label: 'Conditional formatting…', action: () => D.cfDialog(A), disabled: ro },
    { label: 'Define named range…', action: () => D.namesDialog(A, true), disabled: ro }
  );
  return items;
}
const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

/* ---- charts: select, move, resize */

function chartMouseDown(e, el) {
  const id = el.dataset.chart;
  const t = e.target;
  if (t.dataset.chartEdit) { D.chartDialog(A, id); return; }
  if (t.dataset.chartDel) { const sh = sheet(); commit([{ t: 'obj', s: sh.id, kind: 'charts', id, val: null }], { label: 'Delete chart' }); S.chartSel = null; return; }
  e.preventDefault();
  if (S.editing) commitEdit();
  if (S.chartSel !== id) { S.chartSel = id; renderCharts(); }
  focusGrid();
  if (!canEdit()) return;
  const sh = sheet();
  const ch = sh.charts.find((c) => c.id === id);
  drag = { kind: t.dataset.chartResize ? 'chartSize' : 'chartMove', id, x0: e.clientX, y0: e.clientY, ch: { ...ch }, el: $('chartsLayer').querySelector(`[data-chart="${CSS.escape(id)}"]`) };
}
function chartMouseMove(e) {
  const d = drag;
  const z = layout().z;
  const dx = e.clientX - d.x0, dy = e.clientY - d.y0;
  if (d.kind === 'chartMove') { d.el.style.transform = `translate(${dx}px, ${dy}px)`; d.dx = dx; d.dy = dy; }
  else { d.el.style.width = Math.max(200, (d.ch.w || 480) * z + dx) + 'px'; d.el.style.height = Math.max(140, (d.ch.h || 280) * z + dy) + 'px'; d.dx = dx; d.dy = dy; }
}
function finishChartDrag(d) {
  if (!d || (!d.dx && !d.dy)) { renderCharts(); return; }
  const sh = sheet();
  const L = layout();
  const ch = d.ch;
  if (d.kind === 'chartSize') {
    commit([{ t: 'obj', s: sh.id, kind: 'charts', id: d.id, val: { ...ch, w: Math.max(200, Math.round((ch.w || 480) + d.dx / L.z)), h: Math.max(140, Math.round((ch.h || 280) + d.dy / L.z)) } }], { label: 'Resize chart' });
    return;
  }
  const x = ix(sh);
  const base = L.rowTop[L.fr];
  const ar = x.rows.get(ch.ar), ac = x.cols.get(ch.ac);
  const top = (ar >= L.fr ? L.rowTop[ar] - base + L.frozenH : L.rowTop[ar]) + (ch.oy || 0) * L.z + d.dy;
  const left = L.colLeft[ac] + (ch.ox || 0) * L.z + d.dx;
  // find the cell under the new top-left corner
  let r = 0, c = 0;
  const ytarget = Math.max(0, top - L.frozenH) + base;
  while (r < sh.rows.length - 1 && L.rowTop[r + 1] <= ytarget) r++;
  while (c < sh.cols.length - 1 && L.colLeft[c + 1] <= Math.max(0, left)) c++;
  const oy = Math.max(0, Math.round((ytarget - L.rowTop[r]) / L.z)), ox = Math.max(0, Math.round((Math.max(0, left) - L.colLeft[c]) / L.z));
  commit([{ t: 'obj', s: sh.id, kind: 'charts', id: d.id, val: { ...ch, ar: sh.rows[r], ac: sh.cols[c], ox, oy } }], { label: 'Move chart' });
}

/* ---- list picker (data validation) */

function openListPicker() {
  const sh = sheet();
  const dv = M.dvFor(sh, S.sel.r, S.sel.c, ix(sh));
  if (!dv || dv.type !== 'list' || !canEdit()) return;
  const items = M.listItems(engine, view(), sh, dv, LOCALE);
  const td = $('grid').querySelector(`td[data-r="${S.sel.r}"][data-c="${S.sel.c}"]`);
  if (!td) return;
  const current = formatValue(engine.valueAt(sh, S.sel.r, S.sel.c), (M.cellAt(sh, S.sel.r, S.sel.c) || {}).st, LOCALE);
  const search = h('input', { class: 'input small', placeholder: 'Search', 'aria-label': 'Search' });
  const list = h('div', { class: 'pick-list', role: 'listbox' });
  const pick = (val) => {
    closePopups();
    const p = parseInput(val, LOCALE);
    commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, S.sel.r, S.sel.c), v: p.raw === '' ? null : p.raw }], { label: 'Choose' });
    focusGrid();
  };
  const fill = () => {
    list.innerHTML = '';
    const q = search.value.toLowerCase();
    items.filter((it) => it.toLowerCase().includes(q)).slice(0, 300).forEach((it) => list.appendChild(h('button', { class: 'pick-item' + (it === current ? ' on' : ''), type: 'button', role: 'option', onclick: () => pick(it) }, it)));
    if (!list.children.length) list.appendChild(h('div', { class: 'muted pad' }, 'No items'));
  };
  fill();
  search.addEventListener('input', fill);
  search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const f = list.querySelector('.pick-item'); if (f) f.click(); } });
  popover(h('div', { class: 'pick' }, items.length > 8 ? search : null, list, h('button', { class: 'btn small', type: 'button', onclick: () => pick('') }, 'Clear')), td.getBoundingClientRect());
  if (items.length > 8) setTimeout(() => search.focus(), 0);
}

/* ================================================================ keyboard */

function focusGrid() {
  if (S.editing || document.querySelector('.modal-back')) return;
  const k = $('kbd');
  if (document.activeElement !== k) k.focus({ preventScroll: true });
}

function onGridKey(e) {
  if (!store.view) return;
  const sh = sheet();
  const s = S.sel;
  const ctrl = e.ctrlKey || e.metaKey;
  const key = e.key;
  // chart selected
  if (S.chartSel) {
    if ((key === 'Delete' || key === 'Backspace') && canEdit()) { e.preventDefault(); commit([{ t: 'obj', s: sh.id, kind: 'charts', id: S.chartSel, val: null }], { label: 'Delete chart' }); S.chartSel = null; return; }
    if (key === 'Escape') { S.chartSel = null; renderCharts(); return; }
  }
  if (ctrl && !e.altKey) {
    const k = key.toLowerCase();
    if (k === 'z' && !e.shiftKey) { e.preventDefault(); undoRedo(false); return; }
    if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); undoRedo(true); return; }
    if (k === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (k === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (k === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (k === '5') { e.preventDefault(); toggleStyle('x'); return; }
    if (k === 'a') { e.preventDefault(); selectAll(); return; }
    if (k === 'f' || k === 'h') { e.preventDefault(); D.findDialog(A, k === 'h' && canEdit()); return; }
    if (k === 'd') { e.preventDefault(); fillDirection('down'); return; }
    if (k === 'r') { e.preventDefault(); fillDirection('right'); return; }
    if (k === 't') { e.preventDefault(); D.tableDialog(A); return; }
    if (k === 'k') { e.preventDefault(); D.linkDialog(A); return; }
    if (k === 'l' && e.shiftKey) { e.preventDefault(); toggleFilter(); return; }
    if (k === 'v' && e.shiftKey) { S.pasteMode = 'values'; return; }
    if (k === '`') { e.preventDefault(); S.showFormulas = !S.showFormulas; renderGrid(); return; }
    if (key === ';' || key === ':') {
      e.preventDefault();
      const n = new Date();
      if (e.shiftKey) startEdit(String(n.getHours()).padStart(2, '0') + ':' + String(n.getMinutes()).padStart(2, '0'), 'edit');
      else startEdit(formatValue(Math.floor(engine.evalAt('=TODAY()', sh, 0, 0)), { nf: 'date' }, LOCALE), 'edit');
      return;
    }
    if (key === ' ') { e.preventDefault(); selectRect({ r1: 0, r2: sh.rows.length - 1, c1: s.c1, c2: s.c2 }, { r: s.r, c: s.c }); return; }
    if (key === '-' ) { e.preventDefault(); if (s.c1 === 0 && s.c2 === sh.cols.length - 1) structure('delRows'); else if (s.r1 === 0 && s.r2 === sh.rows.length - 1) structure('delCols'); return; }
    if (key === '+' || key === '=' && e.shiftKey) { e.preventDefault(); if (s.c1 === 0 && s.c2 === sh.cols.length - 1) structure('rowAbove'); else if (s.r1 === 0 && s.r2 === sh.rows.length - 1) structure('colLeft'); return; }
    if (key === 'PageDown' || key === 'PageUp') { e.preventDefault(); const i = view().sheets.findIndex((x) => x.id === sh.id); const n = view().sheets[i + (key === 'PageDown' ? 1 : -1)]; if (n) goToSheet(n.id); return; }
    if (e.shiftKey && (key === '$' || key === '4')) { e.preventDefault(); setFormat({ nf: 'cur', cur: CURRENCY, dp: 2 }); return; }
    if (e.shiftKey && (key === '%' || key === '5')) { e.preventDefault(); setFormat({ nf: 'pct', dp: 0 }); return; }
    if (e.shiftKey && (key === '!' || key === '1')) { e.preventDefault(); setFormat({ nf: 'num', dp: 2 }); return; }
    if (e.shiftKey && (key === '#' || key === '3')) { e.preventDefault(); setFormat({ nf: 'date' }); return; }
  }
  if (e.altKey && (key === '=' || key === '0' && e.shiftKey)) { e.preventDefault(); autoSum('SUM'); return; }
  if (e.altKey && key === 'ArrowDown') { e.preventDefault(); openListPicker(); return; }
  if (e.altKey && (key === 'PageDown' || key === 'PageUp')) { e.preventDefault(); const i = view().sheets.findIndex((x) => x.id === sh.id); const n = view().sheets[i + (key === 'PageDown' ? 1 : -1)]; if (n) goToSheet(n.id); return; }
  if (key === 'F2' && e.shiftKey) { e.preventDefault(); D.noteDialog(A); return; }
  const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
  if (arrows[key]) {
    e.preventDefault();
    const [dr, dc] = arrows[key];
    if (e.shiftKey) {
      const end = { r: s.r === s.r1 ? s.r2 : s.r1, c: s.c === s.c1 ? s.c2 : s.c1 };
      const [r, c] = ctrl ? jump(end.r, end.c, dr, dc) : step(end.r, end.c, dr, dc);
      const R = expandForMerges(sh, { r1: Math.min(s.r, r), c1: Math.min(s.c, c), r2: Math.max(s.r, r), c2: Math.max(s.c, c) });
      S.sel = { r: s.r, c: s.c, ...R };
      afterSelection();
      scrollIntoView(r, c);
    } else {
      const [r, c] = ctrl ? jump(s.r, s.c, dr, dc) : step(s.r, s.c, dr, dc);
      selectCell(r, c);
      scrollIntoView(r, c);
    }
    return;
  }
  if (key === 'Tab') {
    e.preventDefault();
    if (s.r1 !== s.r2 || s.c1 !== s.c2) moveWithinSelection(0, e.shiftKey ? -1 : 1);
    else { if (S.tabStart === null) S.tabStart = s.c; const [r, c] = step(s.r, s.c, 0, e.shiftKey ? -1 : 1); selectCell(r, c, false, true); scrollIntoView(r, c); }
    return;
  }
  if (key === 'Enter') {
    e.preventDefault();
    if (s.r1 !== s.r2 || s.c1 !== s.c2) moveWithinSelection(e.shiftKey ? -1 : 1, 0);
    else {
      const base = !e.shiftKey && S.tabStart !== null ? S.tabStart : s.c;
      const [r, c] = step(s.r, base, e.shiftKey ? -1 : 1, 0);
      selectCell(r, c);
      scrollIntoView(r, c);
    }
    return;
  }
  if (key === 'Home') { e.preventDefault(); const r = ctrl ? 0 : s.r; selectCell(r, 0); scrollIntoView(r, 0); return; }
  if (key === 'End' && ctrl) { e.preventDefault(); const [r, c] = lastUsed(sh); selectCell(r, c); scrollIntoView(r, c); return; }
  if (key === 'PageDown' || key === 'PageUp') {
    e.preventDefault();
    const rows = Math.max(1, Math.floor(($('gridWrap').clientHeight - 60) / (Z_ROW_H * S.zoom)));
    const r = Math.max(0, Math.min(sh.rows.length - 1, s.r + (key === 'PageDown' ? rows : -rows)));
    selectCell(r, s.c, e.shiftKey);
    scrollIntoView(r, s.c);
    return;
  }
  if (key === 'Delete' || key === 'Backspace') { e.preventDefault(); clearCells('values'); return; }
  if (key === 'F2') { e.preventDefault(); startEdit(null, 'edit'); return; }
  if (key === 'Escape') { if (S.ants) { S.ants = null; M.clearInternalClip(); renderGrid(); } if (S.painter) { S.painter = null; renderToolbarState(); } return; }
  if (key === ' ' && e.shiftKey) { e.preventDefault(); selectRect({ c1: 0, c2: sh.cols.length - 1, r1: s.r1, r2: s.r2 }, { r: s.r, c: s.c }); return; }
  if (key === ' ' && !ctrl) {
    const dv = M.dvFor(sh, s.r, s.c, ix(sh));
    if (dv && dv.type === 'checkbox') { e.preventDefault(); toggleCheckbox(s.r, s.c); return; }
  }
}

function selectAll() {
  const sh = sheet();
  const s = S.sel;
  const region = M.currentRegion(sh, s.r, s.c);
  const isRegion = region.r1 === s.r1 && region.r2 === s.r2 && region.c1 === s.c1 && region.c2 === s.c2;
  if (!isRegion && (region.r2 > region.r1 || region.c2 > region.c1)) selectRect(region, { r: s.r, c: s.c });
  else selectRect({ r1: 0, c1: 0, r2: sh.rows.length - 1, c2: sh.cols.length - 1 }, { r: s.r, c: s.c });
}
function lastUsed(sh) {
  const x = ix(sh);
  let r = 0, c = 0;
  for (const k in sh.cells) { if (sh.cells[k].v === undefined) continue; const [rid, cid] = k.split(':'); r = Math.max(r, x.rows.get(rid) ?? 0); c = Math.max(c, x.cols.get(cid) ?? 0); }
  return [r, c];
}

/** Typing in the keyboard trap starts editing (works with IME and dead keys too). */
function onKbdInput(e) {
  const k = $('kbd');
  const text = k.value;
  k.value = '';
  if (!text || S.editing || !store.view) return;
  if (!canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  startEdit(text, 'enter');
}

function onEditorKey(e) {
  const input = e.target;
  if (acKey(e)) return;
  const ed = S.editing;
  if (!ed) return;
  const ctrl = e.ctrlKey || e.metaKey;
  if (e.key === 'Enter' && e.altKey) {
    e.preventDefault();
    const p = input.selectionStart;
    input.value = input.value.slice(0, p) + '\n' + input.value.slice(input.selectionEnd);
    input.setSelectionRange(p + 1, p + 1);
    editorInput(input);
    return;
  }
  if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0, ctrl); return; }
  if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); return; }
  if (e.key === 'Escape') { e.preventDefault(); if (S.sheetId !== ed.sheetId) { S.sheetId = ed.sheetId; S.layout = null; } cancelEdit(); return; }
  if (e.key === 'F4' && ed.text[0] === '=') { e.preventDefault(); cycleAbsolute(input); return; }
  if (e.key === 'F2') { e.preventDefault(); ed.mode = ed.mode === 'enter' ? 'edit' : 'enter'; return; }
  // in "enter" mode, arrows finish the edit and move (or insert a reference while typing a formula)
  if (ed.mode === 'enter' && input.id === 'cellEditor' && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (ed.text[0] === '=' && canInsertRef(input.value, input.selectionStart)) {
      e.preventDefault();
      const base = ed.lastRef && ed.lastRef.end === input.selectionStart ? ed.pointer || { r: ed.r, c: ed.c } : { r: ed.r, c: ed.c };
      const [r, c] = step(base.r, base.c, d[0], d[1]);
      ed.pointer = { r, c };
      insertRef(r, c, r, c);
      scrollIntoView(r, c);
      return;
    }
    if (ed.text[0] !== '=') { e.preventDefault(); commitEdit(d[0], d[1]); }
  }
  if (ctrl && e.key.toLowerCase() === 'z' && ed.text === ed.original) { e.preventDefault(); cancelEdit(); }
}

/** F4: A1 → $A$1 → A$1 → $A1 → A1 for the reference at the caret. */
function cycleAbsolute(input) {
  const pos = input.selectionStart;
  const refs = formulaRefs(input.value);
  const ref = refs.find((r) => pos >= r.s && pos <= r.e);
  if (!ref) return;
  const txt = input.value.slice(ref.s, ref.e);
  const m = /^((?:'[^']+'|[^!]+)!)?(\$?)([A-Za-z]+)(\$?)(\d+)(?::(\$?)([A-Za-z]+)(\$?)(\d+))?$/.exec(txt);
  if (!m) return;
  const cyc = (ca, ra) => (!ca && !ra ? ['$', '$'] : ca && ra ? ['', '$'] : !ca && ra ? ['$', ''] : ['', '']);
  const [c1, r1] = cyc(m[2], m[4]);
  let out = (m[1] || '') + c1 + m[3] + r1 + m[5];
  if (m[7]) { const [c2, r2] = cyc(m[6], m[8]); out += ':' + c2 + m[7] + r2 + m[9]; }
  input.value = input.value.slice(0, ref.s) + out + input.value.slice(ref.e);
  input.setSelectionRange(ref.s + out.length, ref.s + out.length);
  editorInput(input);
}

/* ================================================================ menus */

/** True when a menu has something that can be used (not only disabled items). */
function hasUsable(items) {
  return items.some((it) => it && it !== '-' && !it.header && !it.disabled && (!Array.isArray(it.sub) || hasUsable(it.sub)));
}
function buildMenubar() {
  const mb = $('menubar');
  mb.innerHTML = '';
  menubarFor = menubarKind();
  const add = (label, items) => {
    // people who can only view get the menus they can use, without greyed-out items
    if (!canEdit() && store.view) { try { if (!hasUsable(items())) return; } catch (e) { /* show it */ } }
    const b = h('button', { class: 'mb-item', type: 'button', role: 'menuitem', 'aria-haspopup': 'menu', 'data-popup-anchor': '' }, label);
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => {
      if (b.classList.contains('open')) {
        if (Date.now() - (b._hoverAt || 0) < 500) return; // the pointer moved here from another open menu and clicked
        closePopups(); return;
      }
      const m = menu(items(), b.getBoundingClientRect(), { onDone: () => focusGrid(), hideDisabled: !canEdit(), onClose: () => b.classList.remove('open') });
      if (m) b.classList.add('open');
    });
    // moving over the menubar switches menus only while one of its menus is open (not a dropdown or a context menu)
    b.addEventListener('mouseenter', () => { const cur = mb.querySelector('.mb-item.open'); if (cur && cur !== b) { b.click(); b._hoverAt = Date.now(); } });
    mb.appendChild(b);
  };
  add('Edit', () => {
    const ro = !canEdit();
    return [
      { label: 'Undo', kbd: 'Ctrl+Z', icon: 'undo', action: () => undoRedo(false), disabled: ro || !S.undo.length },
      { label: 'Redo', kbd: 'Ctrl+Y', icon: 'redo', action: () => undoRedo(true), disabled: ro || !S.redo.length },
      '-',
      { label: 'Cut', kbd: 'Ctrl+X', action: () => clipboardAction('cut'), disabled: ro },
      { label: 'Copy', kbd: 'Ctrl+C', action: () => clipboardAction('copy') },
      { label: 'Paste', kbd: 'Ctrl+V', action: () => clipboardAction('paste'), disabled: ro },
      { label: 'Paste special', disabled: ro, sub: [
        { label: 'Values only', action: () => pasteSpecial('values') }, { label: 'Formulas only', action: () => pasteSpecial('formulas') },
        { label: 'Formatting only', action: () => pasteSpecial('formats') }, { label: 'Transposed', action: () => pasteSpecial('all', true) }] },
      '-',
      { label: ro ? 'Find' : 'Find and replace', kbd: ro ? 'Ctrl+F' : 'Ctrl+H', icon: 'search', action: () => D.findDialog(A, canEdit()) },
      { label: 'Fill down', kbd: 'Ctrl+D', action: () => fillDirection('down'), disabled: ro },
      { label: 'Fill right', kbd: 'Ctrl+R', action: () => fillDirection('right'), disabled: ro },
      '-',
      { label: 'Delete values', kbd: 'Delete', action: () => clearCells('values'), disabled: ro },
      { label: 'Delete rows', action: () => structure('delRows'), disabled: ro },
      { label: 'Delete columns', action: () => structure('delCols'), disabled: ro },
      { label: 'Delete notes', action: () => clearCells('notes'), disabled: ro },
      { label: 'Clear all', action: () => clearCells('all'), disabled: ro }
    ];
  });
  add('View', () => {
    const sh = sheet();
    return [
      { label: 'Freeze', sub: freezeItems() },
      { label: 'Gridlines', checked: !sh.nogrid, action: () => commit([{ t: 'sheet', s: sh.id, nogrid: !sh.nogrid }], { label: 'Gridlines' }), disabled: !canEdit() },
      { label: 'Show formulas', kbd: 'Ctrl+`', checked: S.showFormulas, action: () => { S.showFormulas = !S.showFormulas; renderGrid(); } },
      { label: 'Formula separators', sub: sepItems() },
      '-',
      { label: 'Hidden rows', sub: [{ label: 'Hide selected rows', action: () => setHidden('r', true) }, { label: 'Unhide selected rows', action: () => setHidden('r', false) }, { label: 'Unhide all rows', action: () => unhideAll('r') }], disabled: !canEdit() },
      { label: 'Hidden columns', sub: [{ label: 'Hide selected columns', action: () => setHidden('c', true) }, { label: 'Unhide selected columns', action: () => setHidden('c', false) }, { label: 'Unhide all columns', action: () => unhideAll('c') }], disabled: !canEdit() },
      '-',
      { label: 'Zoom', sub: [0.75, 0.9, 1, 1.1, 1.25, 1.5].map((z) => ({ label: Math.round(z * 100) + '%', checked: S.zoom === z, action: () => setZoom(z) })) },
      { label: 'Change history', icon: 'history', action: () => D.historyPanel(A) }
    ];
  });
  add('Insert', () => {
    const ro = !canEdit();
    const s = selRect();
    const nR = s.r2 - s.r1 + 1, nC = s.c2 - s.c1 + 1;
    return [
      { label: `${nR === 1 ? 'Row' : nR + ' rows'} above`, action: () => structure('rowAbove'), disabled: ro },
      { label: `${nR === 1 ? 'Row' : nR + ' rows'} below`, action: () => structure('rowBelow'), disabled: ro },
      { label: `${nC === 1 ? 'Column' : nC + ' columns'} left`, action: () => structure('colLeft'), disabled: ro },
      { label: `${nC === 1 ? 'Column' : nC + ' columns'} right`, action: () => structure('colRight'), disabled: ro },
      { label: 'Sheet', action: () => addSheet(), disabled: ro },
      '-',
      { label: 'Table', icon: 'table', action: () => D.tableDialog(A), disabled: ro },
      { label: 'Chart', icon: 'chart', action: () => D.chartDialog(A), disabled: ro },
      { label: 'Function', icon: 'fx', action: () => D.functionDialog(A, insertFunction), disabled: ro },
      { label: 'Link', kbd: 'Ctrl+K', icon: 'link', action: () => D.linkDialog(A), disabled: ro },
      { label: 'Note', kbd: 'Shift+F2', icon: 'note', action: () => D.noteDialog(A), disabled: ro },
      '-',
      { label: 'Checkbox', icon: 'checkbox', action: () => insertCheckboxes(), disabled: ro },
      { label: 'Dropdown list', icon: 'dropdown', action: () => D.dvDialog(A, 'list'), disabled: ro },
      { label: "Today's date", kbd: 'Ctrl+;', icon: 'calendar', action: () => { const sh = sheet(); commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, S.sel.r, S.sel.c), v: numToText(Math.floor(engine.evalAt('=TODAY()', sh, 0, 0))), st: { ...curStyle(), nf: 'date' } }]); }, disabled: ro }
    ];
  });
  add('Format', () => {
    const ro = !canEdit();
    return [
      { label: 'Number', sub: numberFormatItems(), disabled: ro },
      { label: 'Text', disabled: ro, sub: [
        { label: 'Bold', kbd: 'Ctrl+B', icon: 'bold', action: () => toggleStyle('b') }, { label: 'Italic', kbd: 'Ctrl+I', icon: 'italic', action: () => toggleStyle('i') },
        { label: 'Underline', kbd: 'Ctrl+U', icon: 'underline', action: () => toggleStyle('u') }, { label: 'Strikethrough', kbd: 'Ctrl+5', icon: 'strike', action: () => toggleStyle('x') }] },
      { label: 'Alignment', disabled: ro, sub: [
        { label: 'Left', icon: 'left', action: () => applyStyle({ ha: 'left' }) }, { label: 'Center', icon: 'center', action: () => applyStyle({ ha: 'center' }) },
        { label: 'Right', icon: 'right', action: () => applyStyle({ ha: 'right' }) }, '-',
        { label: 'Top', icon: 'vtop', action: () => applyStyle({ va: 'top' }) }, { label: 'Middle', icon: 'vmid', action: () => applyStyle({ va: 'middle' }) },
        { label: 'Bottom', icon: 'vbottom', action: () => applyStyle({ va: null }) }] },
      { label: 'Wrap text', icon: 'wrap', checked: !!curStyle().w, action: () => toggleStyle('w'), disabled: ro },
      { label: 'Borders', sub: borderItems(), disabled: ro },
      { label: 'Merge cells', sub: mergeItems(), disabled: ro },
      '-',
      { label: 'Conditional formatting…', action: () => D.cfDialog(A), disabled: ro },
      { label: 'Table style…', action: () => D.tableDialog(A), disabled: ro },
      '-',
      { label: 'Column width…', action: () => D.sizeDialog(A, 'c'), disabled: ro },
      { label: 'Row height…', action: () => D.sizeDialog(A, 'r'), disabled: ro },
      { label: 'Fit columns to data', action: () => autoFitColumns(range(selRect().c1, selRect().c2)), disabled: ro },
      { label: 'Fit rows to data', action: () => autoFitRows(range(selRect().r1, selRect().r2)), disabled: ro },
      '-',
      { label: 'Clear formatting', icon: 'clear', action: () => clearCells('formats'), disabled: ro }
    ];
  });
  add('Data', () => {
    const ro = !canEdit();
    return [
      { label: 'Sort A → Z', icon: 'sortAsc', action: () => quickSort(false), disabled: ro },
      { label: 'Sort Z → A', icon: 'sortDesc', action: () => quickSort(true), disabled: ro },
      { label: 'Sort range…', action: () => D.sortDialog(A), disabled: ro },
      '-',
      { label: 'Filter', kbd: 'Ctrl+Shift+L', icon: 'filter', action: () => toggleFilter(), disabled: ro },
      { label: 'Format as table…', icon: 'table', action: () => D.tableDialog(A), disabled: ro },
      { label: 'Remove duplicates…', action: () => D.dedupeDialog(A), disabled: ro },
      { label: 'Split text to columns…', action: () => D.splitDialog(A), disabled: ro },
      '-',
      { label: 'Data validation…', action: () => D.dvDialog(A), disabled: ro },
      { label: 'Named ranges…', action: () => D.namesDialog(A) },
      '-',
      { label: 'What-if analysis', icon: 'whatif', disabled: ro, sub: [
        { label: 'Scenario Manager…', action: () => W.scenarioManager(A) },
        { label: 'Goal Seek…', action: () => W.goalSeek(A) }] },
      { label: 'Solver…', icon: 'solver', action: () => W.solverDialog(A), disabled: ro },
      '-',
      { label: 'Import .xlsx or .csv file…', icon: 'upload', action: () => $('fileInput').click(), disabled: ro },
      { label: 'Download as .xlsx', icon: 'download', action: () => exportXlsx() },
      { label: 'Download sheet as CSV', icon: 'download', action: () => exportCsv() }
    ];
  });
}
function freezeItems() {
  const sh = sheet();
  const fr = sh.freeze || { r: 0, c: 0 };
  return [
    { label: 'No rows', checked: !fr.r, action: () => freeze(0, null) },
    { label: '1 row', checked: fr.r === 1, action: () => freeze(1, null) },
    { label: '2 rows', checked: fr.r === 2, action: () => freeze(2, null) },
    { label: `Up to row ${S.sel.r + 1}`, checked: fr.r === S.sel.r + 1 && fr.r > 2, action: () => freeze(Math.min(50, S.sel.r + 1), null) },
    '-',
    { label: 'No columns', checked: !fr.c, action: () => freeze(null, 0) },
    { label: '1 column', checked: fr.c === 1, action: () => freeze(null, 1) },
    { label: '2 columns', checked: fr.c === 2, action: () => freeze(null, 2) },
    { label: `Up to column ${colName(S.sel.c)}`, checked: fr.c === S.sel.c + 1 && fr.c > 2, action: () => freeze(null, Math.min(26, S.sel.c + 1)) }
  ].map((it) => (typeof it === 'object' ? { ...it, disabled: !canEdit() } : it));
}

function insertCheckboxes() {
  const sh = sheet();
  const s = selRect();
  const ops = [{ t: 'obj', s: sh.id, kind: 'dv', id: Core.randomId(), val: { ...M.rectIds(sh, s), type: 'checkbox', strict: true } }];
  for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
    const v = engine.valueAt(sh, r, c);
    if (v !== true && v !== false) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), v: 'FALSE' });
  }
  commit(ops, { label: 'Checkbox' });
}
function insertFunction(name) {
  if (!canEdit()) return;
  if (S.editing) {
    const input = S.editing.from === 'bar' ? $('fbar') : $('cellEditor');
    const p = input.selectionStart;
    input.value = input.value.slice(0, p) + name + '(' + input.value.slice(p);
    input.setSelectionRange(p + name.length + 1, p + name.length + 1);
    editorInput(input);
    input.focus();
  } else startEdit('=' + name + '(', 'enter');
}

function moreMenuItems() {
  const exists = store.exists && store.view;
  const ro = !store.canEdit;
  return [
    exists ? { label: 'Change history', icon: 'history', action: () => D.historyPanel(A) } : null,
    exists ? { label: 'Download as .xlsx', icon: 'download', action: () => exportXlsx() } : null,
    exists ? { label: 'Download sheet as CSV', icon: 'download', action: () => exportCsv() } : null,
    !ro ? { label: 'Import .xlsx or .csv file…', icon: 'upload', action: () => $('fileInput').click() } : null,
    exists && !ro ? '-' : null,
    exists && !ro ? { label: 'Delete spreadsheet…', icon: 'trash', danger: true, action: () => deleteWorkbook() } : null
  ].filter(Boolean);
}

async function deleteWorkbook() {
  const ok = await confirmDialog('Delete spreadsheet', 'Delete the spreadsheet of this issue with all its sheets and its change history? This cannot be undone. Download it as an .xlsx file first if you may need it later.', 'Delete spreadsheet', true);
  if (!ok) return;
  try {
    await api('delete', { method: 'POST', body: {} });
    store.exists = false;
    store.confirmed = store.view = null;
    store.pending = [];
    S.undo = []; S.redo = [];
    toast('The spreadsheet was deleted.');
    renderAll();
  } catch (e) { toast(errorText(e), 'error'); }
}

/* ================================================================ import / export */

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: name });
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 2000);
}
let readableId = null;
async function loadReadableId() {
  if (ENTITY && ENTITY.idReadable) { readableId = ENTITY.idReadable; return; }
  try {
    const res = await host.fetchYouTrack('issues/' + encodeURIComponent(ENTITY.id) + '?fields=idReadable');
    if (res && res.idReadable) readableId = res.idReadable;
  } catch (e) { /* the internal id is used instead */ }
}
function fileBase() {
  const id = readableId || (ENTITY && ENTITY.id) || 'issue';
  return String(id).replace(/[^\w.-]+/g, '_');
}
async function exportXlsx() {
  try {
    const blob = await IO.writeXlsx(view(), engine, S.version);
    download(blob, fileBase() + '.xlsx');
  } catch (e) { toast('Could not create the file: ' + e.message, 'error'); }
}
/** Opens the menu of a button, or closes it when this button's menu is already open (a second click closes it). */
function toggleMenu(b, items, opts = {}) {
  if (b._menu && b._menu.isConnected) { closePopups(); return null; }
  const m = menu(typeof items === 'function' ? items() : items, b.getBoundingClientRect(), {
    ...opts,
    onClose: () => { if (b._menu === m) b._menu = null; b.setAttribute('aria-expanded', 'false'); if (opts.onClose) opts.onClose(); }
  });
  b._menu = m;
  b.setAttribute('aria-expanded', m ? 'true' : 'false');
  return m;
}

/** Like toggleMenu for a popover (the color pickers): open() opens it, a second click on the button closes it. */
function togglePopup(b, open) {
  if (b._menu && b._menu.isConnected) { closePopups(); return null; }
  b._menu = open();
  return b._menu;
}

/** The choice of formula separators (View menu and the toolbar's ⋯ More tools menu). */
function sepItems() {
  return [
    { header: 'Separators in formulas' },
    { label: 'Automatic (' + (separatorStyle('auto', LOCALE) === 'semicolon' ? 'semicolon' : 'comma') + ', from your language)', checked: S.sepPref === 'auto', action: () => setSepPref('auto') },
    { label: 'Comma: =ROUND(1.5, 0)', checked: S.sepPref === 'comma', action: () => setSepPref('comma') },
    { label: 'Semicolon: =ROUND(1,5; 0)', checked: S.sepPref === 'semicolon', action: () => setSepPref('semicolon') }
  ];
}

/** Formula separators for this person (only in this browser; the stored formulas do not change). */
function setSepPref(v) {
  S.sepPref = v;
  storageSet('sep', v);
  renderAll();
  const st = fmlStyle();
  toast(st === 'semicolon' ? 'Formulas now use ; between arguments and , in numbers: =ROUND(1,5; 0)' : 'Formulas now use , between arguments and . in numbers: =ROUND(1.5, 0)');
}
async function exportCsv() {
  const sh = sheet();
  let saved = null;
  try { saved = JSON.parse(await storageGet('csv') || 'null'); } catch (e) { saved = null; }
  // default: what spreadsheet apps in the person's language expect (semicolons and decimal commas where 1,5 is written)
  const local = localeInfo(LOCALE).dec === ',';
  D.csvDialog(A, saved || { delim: local ? ';' : ',', dec: local ? ',' : '.' }, (opts) => {
    storageSet('csv', JSON.stringify(opts));
    download(IO.writeCsv(engine, sh, LOCALE, opts), fileBase() + ' - ' + sh.name.replace(/[^\w.-]+/g, '_') + '.csv');
  });
}

async function importFile(file) {
  if (!file) return;
  if (!store.canEdit) { toast('You can view this spreadsheet but not edit it.'); return; }
  let imported;
  try {
    if (/\.xlsx$/i.test(file.name) || file.type.includes('spreadsheetml')) imported = await IO.readXlsx(await file.arrayBuffer(), { translate: translateFormula });
    else imported = IO.readCsv(await file.text(), LOCALE, file.name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1');
  } catch (e) {
    toast('Could not read the file: ' + e.message, 'error');
    return;
  }
  const sheets = IO.toSheets(imported);
  if (!sheets.length) { toast('The file has no sheets.', 'error'); return; }
  if (!store.exists) {
    await createWorkbook(sheets);
    for (const [n, ref] of Object.entries(imported.names || {})) commit([{ t: 'name', name: n, ref }]);
    return;
  }
  D.importChoice(A, file.name, async (mode) => {
    const ops = [];
    if (mode === 'replace') {
      // delete all but one current sheet, rename that one out of the way, add the first imported sheet, delete the
      // last current sheet, then add the rest: names stay as in the file and the workbook never exceeds the limit
      const old = view().sheets.slice();
      old.slice(1).forEach((sh) => ops.push({ t: 'delSheet', s: sh.id }));
      ops.push({ t: 'sheet', s: old[0].id, name: '~replaced ' + Core.randomId().slice(0, 6) });
      ops.push({ t: 'addSheet', at: 0, sheet: sheets[0] });
      ops.push({ t: 'delSheet', s: old[0].id });
      sheets.slice(1).forEach((sh, i) => ops.push({ t: 'addSheet', at: i + 1, sheet: sh }));
      for (const n of Object.keys(view().names || {})) ops.push({ t: 'name', name: n, ref: null });
    } else {
      sheets.forEach((sh) => { sh.name = M.uniqueSheetName(view(), sh.name); ops.push({ t: 'addSheet', at: view().sheets.length, sheet: sh }); });
    }
    for (const [n, ref] of Object.entries(imported.names || {})) ops.push({ t: 'name', name: n, ref });
    // table names must stay unique
    const taken = new Set(mode === 'replace' ? [] : view().sheets.flatMap((s) => (s.tables || []).map((t) => t.name.toLowerCase())));
    sheets.forEach((s) => (s.tables || []).forEach((t) => { let n = t.name, k = 2; while (taken.has(n.toLowerCase())) n = t.name + '_' + k++; t.name = n; taken.add(n.toLowerCase()); }));
    if (commit(ops, { label: 'Import' })) { goToSheet(sheets[0].id); toast('Imported ' + file.name + '.'); }
  });
}

/* ================================================================ height, zoom */

function setZoom(z) {
  S.zoom = z;
  $('zoom').value = String(z);
  S.layout = null;
  styleCache.clear();
  document.documentElement.style.setProperty('--z', z);
  storageSet('zoom', String(z));
  renderGrid();
}

function initResize() {
  const handle = $('vresize');
  let start = null;
  handle.addEventListener('pointerdown', (e) => {
    start = { y: e.clientY, h: S.height };
    handle.setPointerCapture(e.pointerId);
    document.body.classList.add('vresizing');
  });
  handle.addEventListener('pointermove', (e) => {
    if (!start) return;
    S.height = Math.max(MIN_HEIGHT, Math.min(2000, start.h + e.clientY - start.y));
    $('app').style.setProperty('--panel-h', S.height + 'px');
  });
  const end = () => { if (!start) return; start = null; document.body.classList.remove('vresizing'); storageSet('height', String(S.height)); renderGrid(); };
  handle.addEventListener('pointerup', end);
  handle.addEventListener('pointercancel', end);
  handle.addEventListener('dblclick', () => { S.height = DEFAULT_HEIGHT; $('app').style.setProperty('--panel-h', S.height + 'px'); storageSet('height', String(S.height)); renderGrid(); });
  handle.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      S.height = Math.max(MIN_HEIGHT, S.height + (e.key === 'ArrowDown' ? 40 : -40));
      $('app').style.setProperty('--panel-h', S.height + 'px');
      storageSet('height', String(S.height));
      renderGrid();
    }
  });
}

/* ================================================================ wiring */

function wire() {
  $('moreBtn').innerHTML = icon('more');
  $('moreBtn').addEventListener('click', (e) => toggleMenu(e.currentTarget, moreMenuItems));
  $('addSheet').innerHTML = icon('plus');
  $('addSheet').addEventListener('mousedown', (e) => e.preventDefault()); // the focus stays in the grid, so typing does not add more sheets
  $('addSheet').addEventListener('click', () => { addSheet(); focusGrid(); });
  $('allSheets').innerHTML = icon('sheets');
  $('allSheets').addEventListener('click', (e) => toggleMenu(e.currentTarget, () => view().sheets.map((sh) => ({ label: sh.name, checked: sh.id === S.sheetId, swatch: sh.color || null, action: () => goToSheet(sh.id) }))));
  $('fxBtn').innerHTML = icon('fx');
  $('fxBtn').addEventListener('click', () => D.functionDialog(A, insertFunction));
  $('zoom').addEventListener('change', (e) => setZoom(+e.target.value));
  // in the issue the menus sit in the bar, which saves a row; the large window keeps a menu row
  if (MODE === 'inline') $('roTag').before($('menubar'));
  // the toolbar is one row: tools that don't fit move into "More tools"
  if (typeof ResizeObserver !== 'undefined') new ResizeObserver(() => scheduleFit()).observe($('toolbar'));

  const wrap = $('gridWrap');
  let raf = 0;
  wrap.addEventListener('scroll', () => {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      const L = layout();
      if (!L) return;
      growOnScroll(wrap);
      const [a, b] = visibleRowRange(L);
      if (!lastWindow || a < lastWindow[0] || b > lastWindow[1] || Math.abs(a - lastWindow[0]) > 6) renderGrid();
      else if (S.editing) positionEditor();
    });
  });
  const grid = $('gridInner');
  grid.addEventListener('mousedown', onGridMouseDown);
  grid.addEventListener('dblclick', onGridDblClick);
  grid.addEventListener('contextmenu', onGridContextMenu);
  grid.addEventListener('click', (e) => {
    const a = e.target.closest('[data-link]');
    if (a && !(e.ctrlKey || e.metaKey)) e.preventDefault();
  });
  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', (e) => {
    if (drag && (drag.kind === 'chartMove' || drag.kind === 'chartSize')) { const d = drag; drag = null; finishChartDrag(d); return; }
    onMouseUp(e);
  });
  // drop files to import
  wrap.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) { e.preventDefault(); wrap.classList.add('drop'); } });
  wrap.addEventListener('dragleave', () => wrap.classList.remove('drop'));
  wrap.addEventListener('drop', (e) => { wrap.classList.remove('drop'); const f = e.dataTransfer.files[0]; if (f) { e.preventDefault(); importFile(f); } });
  $('empty').addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files')) e.preventDefault(); });
  $('empty').addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) { e.preventDefault(); importFile(f); } });
  $('bar').addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('Files') && store.canEdit) e.preventDefault(); });
  $('bar').addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) { e.preventDefault(); importFile(f); } });
  $('fileInput').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; importFile(f); });

  const kbd = $('kbd');
  kbd.addEventListener('keydown', (e) => { if (e.isComposing) return; onGridKey(e); });
  kbd.addEventListener('input', onKbdInput);
  kbd.addEventListener('compositionend', () => onKbdInput());

  const ed = $('cellEditor');
  ed.addEventListener('input', () => editorInput(ed));
  ed.addEventListener('keydown', onEditorKey);
  ed.addEventListener('click', () => updateAutocomplete(ed));
  ed.addEventListener('blur', () => setTimeout(() => { if (S.editing && document.activeElement !== ed && document.activeElement !== $('fbar') && !document.querySelector('.modal-back') && !drag) commitEdit(); }, 150));

  const fb = $('fbar');
  fb.addEventListener('focus', () => { if (!S.editing && canEdit() && store.view) { startEdit(null, 'edit', 'bar'); } });
  fb.addEventListener('input', () => { if (!S.editing) startEdit(fb.value, 'edit', 'bar'); else editorInput(fb); autoGrowBar(); });
  fb.addEventListener('focus', () => autoGrowBar());
  fb.addEventListener('blur', () => setTimeout(autoGrowBar, 0));
  fb.addEventListener('keydown', onEditorKey);
  fb.addEventListener('blur', () => setTimeout(() => { if (S.editing && S.editing.from === 'bar' && document.activeElement !== fb && document.activeElement !== ed && !document.querySelector('.modal-back') && !drag) commitEdit(); }, 150));

  const nb = $('nameBox');
  nb.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); goToName(nb.value); focusGrid(); }
    if (e.key === 'Escape') { e.preventDefault(); renderFormulaBar(); focusGrid(); }
  });
  nb.addEventListener('focus', () => nb.select());

  // keys pressed while nothing has focus (after a menu or a click on the bar) still go to the sheet
  document.addEventListener('keydown', (e) => {
    const a = document.activeElement;
    if (a && a !== document.body && a !== document.documentElement) return;
    if (!store.view || document.querySelector('.modal-back') || document.querySelector('.menu, .popover')) return;
    $('kbd').focus({ preventScroll: true });
    if (e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey) onGridKey(e);
  }, true);
  document.addEventListener('copy', (e) => onCopy(e, false));
  document.addEventListener('cut', (e) => onCopy(e, true));
  document.addEventListener('paste', onPaste);
  document.addEventListener('visibilitychange', () => { if (!document.hidden && store.exists) store.pull(); });
  window.addEventListener('resize', () => { if (store.view) renderGrid(); });
  initResize();
}

/** Name box: go to "B12", "A1:C5", "Sheet2!A1", a named range or a table. */
function goToName(text) {
  const t = text.trim();
  if (!t) return;
  const named = view().names[t.toUpperCase()];
  const table = M.findTable(engine, view(), t, S.version, LOCALE);
  if (table) { goToSheet(table.sheet.id); selectRect(table); scrollIntoView(table.r1, table.c1); return; }
  const rect = M.parseRect(named ? named.replace(/^=/, '') : t);
  if (!rect) {
    if (canEdit() && Core.NAME_RE.test(t) && !/^[A-Za-z]{1,3}\d+$/.test(t)) {
      // typing a new name defines it for the selection, like Excel
      commit([{ t: 'name', name: t, ref: '=' + quoteSheet(sheet().name) + '!' + M.rectText(selRect(), true) }], { label: 'Define name' });
      toast('Defined the name ' + t.toUpperCase() + '.');
      return;
    }
    toast('There is no cell or name "' + t + '".', 'error');
    return;
  }
  if (rect.sheet) { const sh = engine.sheetByName(rect.sheet); if (!sh) { toast('There is no sheet "' + rect.sheet + '".', 'error'); return; } goToSheet(sh.id); }
  let sh = sheet();
  if (isFinite(rect.r2) && isFinite(rect.c2) && (rect.r2 >= sh.rows.length || rect.c2 >= sh.cols.length)) { growTo(rect.r2, rect.c2); sh = sheet(); }
  const r2 = Math.min(rect.r2, sh.rows.length - 1), c2 = Math.min(rect.c2, sh.cols.length - 1);
  if (rect.r1 >= sh.rows.length || rect.c1 >= sh.cols.length) { toast('That cell is outside the sheet.', 'error'); return; }
  selectRect({ r1: rect.r1, c1: rect.c1, r2, c2 });
  scrollIntoView(rect.r1, rect.c1);
}

/* ================================================================ start */

async function init() {
  $('zoom').replaceWith(select([['0.75', '75%'], ['0.9', '90%'], ['1', '100%'], ['1.1', '110%'], ['1.25', '125%'], ['1.5', '150%']], '1', { id: 'zoom', class: 'small zoom', title: 'Zoom', 'aria-label': 'Zoom' }));
  wire();
  buildToolbar();
  buildMenubar();
  if (!host) {
    document.body.innerHTML = '<div class="empty"><div class="empty-title">Open this page in YouTrack</div></div>';
    return;
  }
  const [height, zoom, sep] = await Promise.all([storageGet('height'), storageGet('zoom'), storageGet('sep')]);
  if (sep === 'comma' || sep === 'semicolon') S.sepPref = sep;
  if (height && +height >= MIN_HEIGHT) S.height = +height;
  if (zoom && +zoom > 0) { S.zoom = +zoom; $('zoom').value = zoom; document.documentElement.style.setProperty('--z', zoom); }
  renderAll();
  await loadAndStart();
}

/** Loads the spreadsheet and starts polling; a failed load shows the error in the bar with Try again. */
async function loadAndStart() {
  const meta = $('barMeta');
  try {
    await store.load();
    store.loaded = true;
    if (store.view) S.sheetId = store.view.sheets[0].id;
  } catch (e) {
    store.loaded = true;
    meta.dataset.error = '1';
    meta.classList.remove('loading');
    meta.textContent = 'Could not load the spreadsheet: ' + errorText(e) + ' ';
    meta.title = meta.textContent;
    const again = h('button', { class: 'btn small', type: 'button' }, 'Try again');
    again.addEventListener('click', () => {
      delete meta.dataset.error;
      store.loaded = false;
      renderAll();
      loadAndStart();
    });
    meta.after(again);
    meta._again && meta._again.remove();
    meta._again = again;
    renderAll();
    return;
  }
  if (meta._again) { meta._again.remove(); meta._again = null; }
  renderAll();
  renderPeers();
  sendSelection();
  loadReadableId();
  store.startPolling(() => {
    if (document.hidden) return 30000;
    // an issue without a spreadsheet only checks now and then whether someone has added one
    if (MODE === 'inline' && !store.exists) return 20000;
    return S.presence.length ? 1500 : 2500;
  });
}

init();
