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
/* Spreadsheets for YouTrack – workbook helpers: positions, structural changes with formula rewriting, clipboard, fill series,
 * sort, tables and filters, conditional formatting, data validation. Everything here produces change operations
 * (see sheet-core.js) and never changes the workbook directly. */
import {
  shiftFormula, translateFormula, renameSheetInFormula, rewriteRefs, parseInput, formatValue, isErr, compare,
  colName, literal, numToText, serialToParts, dateToSerial, parseRefText, parseFormula, budgetSpent, isSlowCalc, charge, cellIndependent
} from './formula.js';

const Core = window.SheetCore;
export const newIds = (n) => Core.ids(n);

/* ---------------------------------------------------------------- positions */

/** Row/column id → index maps for a sheet. `version` changes whenever the workbook changes. */
const indexCache = new WeakMap();
export function indexOf(sheet, version) {
  const hit = indexCache.get(sheet);
  if (hit && hit.version === version) return hit;
  const ix = { version, rows: new Map(), cols: new Map() };
  sheet.rows.forEach((id, i) => ix.rows.set(id, i));
  sheet.cols.forEach((id, i) => ix.cols.set(id, i));
  indexCache.set(sheet, ix);
  return ix;
}
export const keyAt = (sheet, r, c) => sheet.rows[r] + ':' + sheet.cols[c];
export function cellAt(sheet, r, c) {
  const rid = sheet.rows[r], cid = sheet.cols[c];
  return rid && cid ? sheet.cells[rid + ':' + cid] || null : null;
}
/** Range object {r1, c1, r2, c2} with ids → positions, or null when it no longer exists. */
export function objRect(o, ix) {
  const r1 = ix.rows.get(o.r1), r2 = ix.rows.get(o.r2), c1 = ix.cols.get(o.c1), c2 = ix.cols.get(o.c2);
  if ([r1, r2, c1, c2].some((x) => x === undefined)) return null;
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}
export function rectIds(sheet, rect) {
  return { r1: sheet.rows[rect.r1], c1: sheet.cols[rect.c1], r2: sheet.rows[rect.r2], c2: sheet.cols[rect.c2] };
}
export const inRect = (rect, r, c) => rect && r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
export function rectText(rect, abs) {
  const d = abs ? '$' : '';
  const a = d + colName(rect.c1) + d + (rect.r1 + 1);
  if (rect.r1 === rect.r2 && rect.c1 === rect.c2) return a;
  return a + ':' + d + colName(rect.c2) + d + (rect.r2 + 1);
}
export function parseRect(text) {
  const ref = parseRefText(text);
  if (!ref || ref.kind === 'cols' || ref.kind === 'rows') {
    if (ref && ref.kind === 'cols') return { sheet: ref.sheet, r1: 0, r2: Infinity, c1: Math.min(ref.c1, ref.c2), c2: Math.max(ref.c1, ref.c2) };
    if (ref && ref.kind === 'rows') return { sheet: ref.sheet, c1: 0, c2: Infinity, r1: Math.min(ref.r1, ref.r2), r2: Math.max(ref.r1, ref.r2) };
    return null;
  }
  return { sheet: ref.sheet, r1: Math.min(ref.r1, ref.r2), r2: Math.max(ref.r1, ref.r2), c1: Math.min(ref.c1, ref.c2), c2: Math.max(ref.c1, ref.c2) };
}

/* ---------------------------------------------------------------- formula rewriting across the workbook */

/** Calls fn(formula, formulaSheet) for every formula in the workbook; returns set/obj/name ops for changed ones. */
function rewriteAll(doc, fn) {
  const ops = [];
  for (const sh of doc.sheets) {
    for (const k in sh.cells) {
      const v = sh.cells[k].v;
      if (typeof v !== 'string' || v[0] !== '=') continue;
      const nv = fn(v, sh);
      if (nv !== v) ops.push({ t: 'set', s: sh.id, k, v: nv });
    }
    for (const kind of ['cf', 'dv']) {
      for (const o of sh[kind] || []) {
        let changed = false;
        const copy = JSON.parse(JSON.stringify(o));
        for (const f of ['formula', 'a', 'b', 'source']) {
          if (typeof copy[f] === 'string' && copy[f][0] === '=') {
            const nv = fn(copy[f], sh);
            if (nv !== copy[f]) { copy[f] = nv; changed = true; }
          }
        }
        if (changed) ops.push({ t: 'obj', s: sh.id, kind, id: o.id, val: copy });
      }
    }
    // Solver models and scenarios: objective, variable cells, both sides of constraints, changing cells
    for (const kind of ['solver', 'scenarios']) {
      for (const o of sh[kind] || []) {
        let changed = false;
        const fix = (v) => {
          if (typeof v !== 'string' || v[0] !== '=') return v;
          const nv = fn(v, sh);
          if (nv !== v) changed = true;
          return nv;
        };
        const copy = JSON.parse(JSON.stringify(o));
        if ('obj' in copy) copy.obj = fix(copy.obj);
        for (const k of ['vars', 'cells']) if (Array.isArray(copy[k])) copy[k] = copy[k].map(fix);
        if (Array.isArray(copy.cons)) copy.cons.forEach((c) => { if (c) { c.l = fix(c.l); c.r = fix(c.r); } });
        if (changed) ops.push({ t: 'obj', s: sh.id, kind, id: o.id, val: copy });
      }
    }
  }
  for (const name in doc.names || {}) {
    const def = doc.names[name];
    const nv = fn(def[0] === '=' ? def : '=' + def, null);
    if (nv !== (def[0] === '=' ? def : '=' + def)) ops.push({ t: 'name', name, ref: nv });
  }
  return ops;
}
const targets = (sheet) => (ref, fsheet) => (ref.sheet == null ? fsheet === sheet : ref.sheet.toLowerCase() === sheet.name.toLowerCase());

/** Insert (count > 0) or delete (count < 0) rows ('r') or columns ('c') at index `at`. */
export function structuralOps(doc, sheet, axis, at, count) {
  const isT = targets(sheet);
  const ops = rewriteAll(doc, (f, fs) => shiftFormula(f, axis, at, count, (ref) => isT(ref, fs)));
  const t = axis === 'r' ? 'rows' : 'cols';
  if (count > 0) ops.push({ t, s: sheet.id, at, add: Core.ids(count) });
  else {
    const list = axis === 'r' ? sheet.rows : sheet.cols;
    ops.push({ t, s: sheet.id, del: list.slice(at, at - count) });
  }
  const fr = sheet.freeze || { r: 0, c: 0 };
  const key = axis === 'r' ? 'r' : 'c';
  if (fr[key] > at) {
    const n = count > 0 ? fr[key] + count : Math.max(at, fr[key] + count);
    ops.push({ t: 'sheet', s: sheet.id, freeze: { ...fr, [key]: Math.min(n, key === 'r' ? 50 : 26) } });
  }
  return ops;
}

export function renameSheetOps(doc, sheet, newName) {
  const ops = rewriteAll(doc, (f) => renameSheetInFormula(f, sheet.name, newName));
  ops.unshift({ t: 'sheet', s: sheet.id, name: newName });
  return ops;
}

export function deleteSheetOps(doc, sheet) {
  const ops = rewriteAll(doc, (f, fs) => (fs === sheet ? f : renameSheetInFormula(f, sheet.name, null)));
  ops.push({ t: 'delSheet', s: sheet.id });
  return ops;
}

export function uniqueSheetName(doc, base) {
  let name = base, n = 2;
  while (doc.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = base.replace(/ \(\d+\)$/, '') + ' (' + n++ + ')';
  return name.slice(0, 31);
}
export function uniqueTableName(doc, base = 'Table') {
  const all = new Set(doc.sheets.flatMap((s) => (s.tables || []).map((t) => String(t.name).toLowerCase())));
  let n = 1;
  while (all.has((base + n).toLowerCase())) n++;
  return base + n;
}

export function duplicateSheet(doc, sheet) {
  const copy = JSON.parse(JSON.stringify(sheet));
  copy.id = Core.randomId();
  copy.name = uniqueSheetName(doc, sheet.name + ' (2)');
  for (const t of copy.tables || []) t.name = uniqueTableName(doc, String(t.name).replace(/\d+$/, '') || 'Table');
  // formulas that pointed at the original sheet by name keep doing so; same-sheet refs follow the copy
  return { t: 'addSheet', at: doc.sheets.indexOf(sheet) + 1, sheet: copy };
}

/* ---------------------------------------------------------------- clipboard */

let internalClip = null;
const CLIP_MARK = 'data-issue-sheets-clip';

function escHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function tsvCell(s) { return /[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }

/**
 * Copies a rectangle: text for other apps, HTML table for Excel and Google Sheets, and an internal copy.
 * Values the formula work budget left out are never copied as if they were values: when values read in this pass may
 * be missing, it returns { pending: true } and keeps the earlier copy (the caller computes them for the next try). A
 * formula that needs more calculation than the spreadsheet does at once (#CALC!, isSlowCalc) is copied without its
 * value, and the result says so with slow: true.
 */
export function copyRect(engine, sheet, rect, locale, cut) {
  return copyCells(engine, sheet, span(rect.r1, rect.r2), span(rect.c1, rect.c2), locale, cut);
}
const span = (a, b) => Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);
/**
 * The rows and columns that a copy of several selected ranges takes, as in Excel: ranges over the same rows are copied
 * side by side (their columns in sheet order), ranges over the same columns one above the other (their rows in sheet
 * order). A cell in more than one of them is copied once. null when the ranges are neither.
 */
export function copyLines(rects) {
  const a = rects[0];
  const union = (list) => [...new Set(list)].sort((x, y) => x - y);
  if (rects.every((x) => x.r1 === a.r1 && x.r2 === a.r2)) return { rows: span(a.r1, a.r2), cols: union(rects.flatMap((x) => span(x.c1, x.c2))) };
  if (rects.every((x) => x.c1 === a.c1 && x.c2 === a.c2)) return { rows: union(rects.flatMap((x) => span(x.r1, x.r2))), cols: span(a.c1, a.c2) };
  return null;
}
/**
 * Copies the cells where the rows cross the columns (sorted lists of row and column numbers): a rectangle, or several
 * ranges copied together (copyLines), which are then pasted as one block. Hidden rows are left out when more than one
 * row is copied. The internal copy keeps the row and column each copied cell came from (src.rows, src.cols): a pasted
 * formula moves by the distance from its own cell, and Values only reads that cell.
 */
export function copyCells(engine, sheet, lines, cols, locale, cut) {
  const rows = [], text = [], html = [], srcRows = [];
  let slow = false;
  for (const r of lines) {
    if (budgetSpent()) return { pending: true };
    if (sheet.hr && sheet.hr[sheet.rows[r]] && lines.length > 1) continue;
    srcRows.push(r);
    const row = [], trow = [], hrow = [];
    for (const c of cols) {
      const cell = cellAt(sheet, r, c);
      row.push(cell ? { v: cell.v, st: cell.st, n: cell.n } : null);
      const v = engine.valueAt(sheet, r, c);
      if (isSlowCalc(v)) slow = true;
      const shown = isSlowCalc(v) ? '' : formatValue(v, cell && cell.st, locale);
      trow.push(tsvCell(shown));
      const st = cell && cell.st ? cell.st : {};
      // checked values only (#rrggbb colors, known alignments), and the attribute is escaped as well
      const fc = typeof st.fc === 'string' && HEX.test(st.fc) ? st.fc : null, bg = typeof st.bg === 'string' && HEX.test(st.bg) ? st.bg : null;
      const ha = ['left', 'center', 'right'].includes(st.ha) ? st.ha : null;
      const css = [st.b ? 'font-weight:bold' : '', st.i ? 'font-style:italic' : '', st.u ? 'text-decoration:underline' : '',
        fc ? 'color:' + fc : '', bg ? 'background:' + bg : '', ha ? 'text-align:' + ha : ''].filter(Boolean).join(';');
      hrow.push(`<td${css ? ` style="${escHtml(css)}"` : ''}>${escHtml(shown)}</td>`);
    }
    rows.push(row);
    text.push(trow.join('\t'));
    html.push('<tr>' + hrow.join('') + '</tr>');
  }
  if (budgetSpent()) return { pending: true };
  const token = Math.random().toString(36).slice(2);
  internalClip = { token, rows, src: { sheetId: sheet.id, r1: lines[0], c1: cols[0], rows: srcRows, cols: cols.slice() }, cut: !!cut };
  return {
    text: text.join('\r\n'),
    html: `<table ${CLIP_MARK}="${token}"><tbody>${html.join('')}</tbody></table>`,
    clip: internalClip,
    slow
  };
}
export function clearInternalClip() { internalClip = null; }
export function getInternalClip() { return internalClip; }

export function parseTSV(text) {
  const rows = [];
  let row = [], cur = '', q = false;
  const t = text.replace(/\r\n/g, '\n').replace(/\n$/, '');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"' && cur === '') q = true;
    else if (ch === '\t') { row.push(cur); cur = ''; }
    else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  row.push(cur);
  rows.push(row);
  return rows;
}

function colorToHex(c) {
  if (!c) return null;
  c = c.trim();
  if (/^#[0-9a-f]{6}$/i.test(c)) return c.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(c)) return '#' + c.slice(1).split('').map((x) => x + x).join('').toLowerCase();
  const m = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/.exec(c);
  if (m) { if (m[4] !== undefined && +m[4] === 0) return null; return '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join(''); }
  return null;
}

/**
 * Parses clipboard HTML. Returns { ours: clip } for our own copy, { rows: [[{text, st}]] } for a table,
 * { tooBig: true } for a table larger than one batch of changes, or null. DOMParser builds an inert document:
 * nothing in the pasted markup runs or loads.
 */
export function parseClipboardHtml(html) {
  if (!html) return null;
  const m = new RegExp(CLIP_MARK + '="([a-z0-9]+)"').exec(html);
  if (m && internalClip && internalClip.token === m[1]) return { ours: internalClip };
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const L = Core.LIMITS;
  // every cell, including the ones under row and column spans, counts against one batch of changes
  let budget = L.opsPerBatch, tooBig = false;
  const out = [];
  const spans = [];
  table.querySelectorAll('tr').forEach((tr, ri) => {
    if (tooBig) return;
    if (ri >= L.rows) { tooBig = true; return; }
    const row = out[ri] || (out[ri] = []);
    let ci = 0;
    tr.querySelectorAll('td,th').forEach((td) => {
      if (tooBig) return;
      while (row[ci] !== undefined) ci++;
      if (ci >= L.cols) { tooBig = true; return; }
      const style = td.getAttribute('style') || '';
      const get = (prop) => { const mm = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([^;]+)', 'i').exec(style); return mm ? mm[1] : null; };
      const st = {};
      const fw = get('font-weight');
      if (td.tagName === 'TH' || (fw && (fw === 'bold' || +fw >= 600)) || td.querySelector('b,strong')) st.b = 1;
      if ((get('font-style') || '').includes('italic') || td.querySelector('i,em')) st.i = 1;
      if ((get('text-decoration') || '').includes('underline') || td.querySelector('u')) st.u = 1;
      if ((get('text-decoration') || '').includes('line-through') || td.querySelector('s,strike')) st.x = 1;
      const fc = colorToHex(get('color'));
      if (fc && fc !== '#000000') st.fc = fc;
      const bg = colorToHex(get('background-color') || get('background'));
      if (bg && bg !== '#ffffff') st.bg = bg;
      const ta = get('text-align');
      if (ta && ['left', 'center', 'right'].includes(ta.trim())) st.ha = ta.trim();
      let text = td.innerText !== undefined && td.innerText !== '' ? td.innerText : td.textContent;
      text = text.replace(/ /g, ' ').replace(/\n$/, '');
      // whole numbers of at least 1 that fit in the sheet (attributes are text: NaN, negative and huge values are clamped)
      const span = (name, max) => { const n = Math.floor(Number(td.getAttribute(name))); return n > 1 ? Math.min(n, max) : 1; };
      const rs = span('rowspan', L.rows - ri), cs = span('colspan', L.cols - ci);
      if ((budget -= rs * cs) < 0) { tooBig = true; return; }
      row[ci] = { text, st: Object.keys(st).length ? st : null };
      if (rs > 1 || cs > 1) spans.push({ r1: ri, c1: ci, r2: ri + rs - 1, c2: ci + cs - 1 });
      for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) {
        if (!a && !b) continue;
        const rr = out[ri + a] || (out[ri + a] = []);
        rr[ci + b] = { text: '', st: null };
      }
      ci += cs;
    });
  });
  if (tooBig) return { tooBig: true };
  let w = 0;
  for (const r of out) if (r && r.length > w) w = r.length;
  if (out.length * w > L.opsPerBatch) return { tooBig: true };
  return { rows: Array.from(out, (r) => Array.from({ length: w }, (_, i) => (r && r[i]) || { text: '', st: null })), spans };
}

/** Ops that grow the sheet so that rows/cols up to (r, c) exist. */
export function growOps(sheet, r, c) {
  const ops = [];
  if (r >= sheet.rows.length) ops.push({ t: 'rows', s: sheet.id, at: -1, add: Core.ids(Math.min(r + 1 - sheet.rows.length + 20, Core.LIMITS.rows - sheet.rows.length)) });
  if (c >= sheet.cols.length) ops.push({ t: 'cols', s: sheet.id, at: -1, add: Core.ids(Math.min(c + 1 - sheet.cols.length + 2, Core.LIMITS.cols - sheet.cols.length)) });
  return ops;
}

/**
 * Paste. `data` is our clip ({rows: [[{v, st, n}]], src}) or external ({rows: [[{text, st}]]}).
 * mode: 'all' | 'values' | 'formats' | 'formulas'; transpose: boolean.
 * When the target selection is larger than the copied block and a multiple of it, the block is repeated.
 * Returns null when it is more than one batch of changes can hold, and false when values of copied formulas are not
 * computed (the formula work budget left them out): they are never pasted as if they were values.
 */
export function pasteOps(engine, doc, sheet, target, data, opts) {
  const { mode = 'all', transpose = false, locale } = opts || {};
  let rows = data.rows;
  if (transpose) rows = rows[0].map((_, j) => rows.map((r) => r[j]));
  const h = rows.length, w = rows[0] ? rows[0].length : 0;
  if (!h || !w) return [];
  const th = target.r2 - target.r1 + 1, tw = target.c2 - target.c1 + 1;
  const reps = th % h === 0 && tw % w === 0 && (th > h || tw > w) ? [th / h, tw / w] : [1, 1];
  // more cells than one batch of changes can hold: the caller says so instead of trying
  if (h * w * reps[0] * reps[1] > Core.LIMITS.opsPerBatch) return null;
  const ops = growOps(sheet, target.r1 + h * reps[0] - 1, target.c1 + w * reps[1] - 1);
  // grow a scratch copy so keys exist
  const rowIds = sheet.rows.slice(), colIds = sheet.cols.slice();
  for (const op of ops) { if (op.t === 'rows') rowIds.push(...op.add); if (op.t === 'cols') colIds.push(...op.add); }
  const ext = !!data.rows[0][0] && 'text' in (data.rows[0][0] || {});
  // the row and column a copied cell came from (copies of several ranges, and copies that left hidden rows out, are
  // not one rectangle of the sheet)
  const srcR = (i) => (data.src && Array.isArray(data.src.rows) ? data.src.rows[i] : data.src.r1 + i);
  const srcC = (j) => (data.src && Array.isArray(data.src.cols) ? data.src.cols[j] : data.src.c1 + j);
  let read = false, unknown = false;
  for (let a = 0; a < reps[0]; a++) for (let b = 0; b < reps[1]; b++) {
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
      const r = target.r1 + a * h + i, c = target.c1 + b * w + j;
      const k = rowIds[r] + ':' + colIds[c];
      const src = rows[i][j];
      const op = { t: 'set', s: sheet.id, k };
      if (ext) {
        if (mode !== 'formats') {
          const p = parseInput(src ? src.text : '', locale);
          op.v = p.raw === '' ? null : p.raw;
          if (mode === 'all') op.st = mergeFmt(src && src.st, p.fmt);
        } else op.st = src ? src.st : null;
      } else {
        if (mode !== 'formats') {
          let v = src ? src.v ?? null : null;
          if (typeof v === 'string' && v[0] === '=') {
            if (mode === 'values') {
              const sr = srcR(transpose ? j : i), sc = srcC(transpose ? i : j);
              const srcSheet = doc.sheets.find((s) => s.id === data.src.sheetId);
              const val = srcSheet ? engine.valueAt(srcSheet, sr, sc) : null;
              read = true;
              if (isSlowCalc(val)) unknown = true;
              v = val === null || isErr(val) ? (isErr(val) ? val.code : null) : typeof val === 'number' ? numToText(val) : typeof val === 'boolean' ? (val ? 'TRUE' : 'FALSE') : String(val).match(/^[=+'-]/) ? "'" + val : String(val);
            } else {
              const dr = r - srcR(transpose ? j : i), dc = c - srcC(transpose ? i : j);
              v = translateFormula(v, dr, dc);
            }
          }
          op.v = v;
          if (mode === 'all') { op.n = src ? src.n || null : null; }
        }
        if (mode === 'all' || mode === 'formats') op.st = src ? src.st || null : null;
      }
      ops.push(op);
    }
  }
  // values read after the work budget ran out may be wrong without showing #CALC!
  if (unknown || (read && budgetSpent())) return false;
  return ops;
}
function mergeFmt(st, fmt) {
  if (!fmt) return st || null;
  return { ...(st || {}), ...fmt };
}

/** Ops that clear the source of a cut (cells that the paste didn't overwrite). */
export function cutClearOps(doc, clip, targetSheetId, target) {
  const sheet = doc.sheets.find((s) => s.id === clip.src.sheetId);
  if (!sheet || !clip.rows.length) return [];
  const ops = [];
  const h = clip.rows.length, w = clip.rows[0].length;
  // the cells that were copied (rows hidden at the time were left out of the copy, and stay)
  const srcR = (i) => (Array.isArray(clip.src.rows) ? clip.src.rows[i] : clip.src.r1 + i);
  const srcC = (j) => (Array.isArray(clip.src.cols) ? clip.src.cols[j] : clip.src.c1 + j);
  for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
    const r = srcR(i), c = srcC(j);
    if (sheet.id === targetSheetId && r >= target.r1 && r < target.r1 + h && c >= target.c1 && c < target.c1 + w) continue;
    if (!sheet.rows[r] || !sheet.cols[c]) continue;
    const k = keyAt(sheet, r, c);
    if (sheet.cells[k]) ops.push({ t: 'set', s: sheet.id, k, v: null, st: null, n: null });
  }
  return ops;
}

/* ---------------------------------------------------------------- fill (drag handle, Ctrl+D, Ctrl+R) */

const MONTHS = [
  ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'],
  ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'],
  ['ocak', 'şubat', 'mart', 'nisan', 'mayıs', 'haziran', 'temmuz', 'ağustos', 'eylül', 'ekim', 'kasım', 'aralık'],
  ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'],
  ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'],
  ['pazartesi', 'salı', 'çarşamba', 'perşembe', 'cuma', 'cumartesi', 'pazar'],
  ['q1', 'q2', 'q3', 'q4']
];
function matchCase(word, sample) {
  if (sample === sample.toUpperCase()) return word.toLocaleUpperCase('tr');
  if (sample[0] === sample[0].toUpperCase()) return word[0].toLocaleUpperCase('tr') + word.slice(1);
  return word;
}
/** Generator of the n-th fill value from the source values (array of raw strings). */
function seriesFor(raws, styles) {
  const n = raws.length;
  const nums = raws.map((v) => (v !== undefined && v !== null && v !== '' && v[0] !== '=' ? literal(v) : null));
  // numbers (and dates): linear trend from 2+ values, +1 day for a single date, copy for a single number
  if (nums.every((x) => typeof x === 'number')) {
    if (n === 1) {
      const isDate = styles[0] && ['date', 'ldate', 'dt'].includes(styles[0].nf);
      return (i) => (isDate ? numToText(nums[0] + i + 1) : raws[0]);
    }
    const xs = nums.map((_, i) => i), mx = (n - 1) / 2, my = nums.reduce((s, x) => s + x, 0) / n;
    let sxy = 0, sxx = 0;
    xs.forEach((x, i) => { sxy += (x - mx) * (nums[i] - my); sxx += (x - mx) ** 2; });
    const slope = sxx ? sxy / sxx : 0;
    const allInt = nums.every(Number.isInteger) && Number.isInteger(slope);
    return (i) => { const v = my + slope * (n + i - mx); return numToText(allInt ? Math.round(v) : +v.toPrecision(12)); };
  }
  // month / day names
  for (const list of MONTHS) {
    const idx = raws.map((v) => (typeof v === 'string' ? list.indexOf(v.toLocaleLowerCase('tr')) : -1));
    if (idx.every((x) => x >= 0)) {
      const step = n > 1 ? (idx[1] - idx[0] + list.length) % list.length || 1 : 1;
      return (i) => matchCase(list[(idx[n - 1] + step * (i + 1)) % list.length], raws[n - 1]);
    }
  }
  // text with a trailing number ("Item 1", "Q1-2026" not)
  const tm = raws.map((v) => (typeof v === 'string' && v[0] !== '=' ? /^(.*?)(\d+)$/.exec(v) : null));
  if (tm.every((m) => m) && tm.every((m) => m[1] === tm[0][1])) {
    const vals = tm.map((m) => +m[2]);
    const step = n > 1 ? vals[n - 1] - vals[n - 2] : 1;
    const width = tm[0][2].length;
    return (i) => tm[0][1] + String(vals[n - 1] + step * (i + 1)).padStart(tm[0][2][0] === '0' ? width : 0, '0');
  }
  return null;
}

/**
 * Fills `dest` (a rectangle next to `src`) the way Excel's fill handle does.
 * dir: 'down' | 'up' | 'right' | 'left'. copyOnly: Ctrl held (no series).
 */
export function fillOps(sheet, src, dest, dir, copyOnly) {
  const ops = growOps(sheet, dest.r2, dest.c2);
  const rowIds = sheet.rows.slice(), colIds = sheet.cols.slice();
  for (const op of ops) { if (op.t === 'rows') rowIds.push(...op.add); if (op.t === 'cols') colIds.push(...op.add); }
  const vertical = dir === 'down' || dir === 'up';
  const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
  const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
  for (let line = lines[0]; line <= lines[1]; line++) {
    const raws = [], styles = [], notes = [], pos = [];
    for (let i = 0; i < len; i++) {
      const r = vertical ? src.r1 + i : line, c = vertical ? line : src.c1 + i;
      const cell = cellAt(sheet, r, c);
      raws.push(cell ? cell.v : undefined); styles.push(cell ? cell.st : undefined); notes.push(cell ? cell.n : undefined); pos.push([r, c]);
    }
    const forward = dir === 'down' || dir === 'right';
    const ordered = forward ? raws : raws.slice().reverse();
    const series = copyOnly || raws.some((v) => typeof v === 'string' && v[0] === '=') ? null : seriesFor(ordered.filter((v) => v !== undefined), styles);
    const count = vertical ? dest.r2 - dest.r1 + 1 : dest.c2 - dest.c1 + 1;
    for (let i = 0; i < count; i++) {
      const r = vertical ? (forward ? dest.r1 + i : dest.r2 - i) : line;
      const c = vertical ? line : (forward ? dest.c1 + i : dest.c2 - i);
      const si = forward ? i % len : len - 1 - (i % len);
      const [sr, sc] = pos[si];
      let v = raws[si];
      if (typeof v === 'string' && v[0] === '=') v = translateFormula(v, r - sr, c - sc);
      else if (series && raws.every((x) => x !== undefined)) v = series(i);
      ops.push({ t: 'set', s: sheet.id, k: rowIds[r] + ':' + colIds[c], v: v ?? null, st: styles[si] || null });
    }
  }
  return ops;
}

/* ---------------------------------------------------------------- sort, dedupe */

/* Sort and dedupe never decide on values that the formula work budget left out (#CALC!, isSlowCalc) or that were read
 * after it ran out (they may be wrong without showing it): they give null, and the caller computes the values first
 * (app.js computeValues) or says why nothing was done. */

/** keys: [{ c: column index, desc }]. Null when sort keys are not computed. */
export function sortOps(engine, sheet, rect, keys, hasHeader) {
  const top = rect.r1 + (hasHeader ? 1 : 0);
  if (top > rect.r2) return [];
  const rows = [];
  let unknown = false;
  for (let r = top; r <= rect.r2; r++) {
    const vals = keys.map((k) => engine.valueAt(sheet, r, k.c));
    if (vals.some(isSlowCalc)) unknown = true;
    rows.push({ r, cells: Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, j) => cellAt(sheet, r, rect.c1 + j)), vals });
  }
  if (unknown || budgetSpent()) return null;
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const x = a.vals[i], y = b.vals[i];
      const ex = x === null || x === '', ey = y === null || y === '';
      if (ex || ey) { if (ex && ey) continue; return ex ? 1 : -1; } // blanks last, either order
      const c = compare(isErr(x) ? x.code : x, isErr(y) ? y.code : y);
      if (c) return keys[i].desc ? -c : c;
    }
    return a.r - b.r;
  });
  const ops = [];
  rows.forEach((row, i) => {
    const r = top + i;
    row.cells.forEach((cell, j) => {
      const c = rect.c1 + j;
      let v = cell ? cell.v ?? null : null;
      if (typeof v === 'string' && v[0] === '=') v = translateFormula(v, r - row.r, 0);
      const cur = cellAt(sheet, r, c);
      if (JSON.stringify(cur || null) === JSON.stringify(cell ? { ...cell, v: v ?? undefined } : null) && row.r === r) return;
      ops.push({ t: 'set', s: sheet.id, k: keyAt(sheet, r, c), v, st: cell ? cell.st || null : null, n: cell ? cell.n || null : null });
    });
  });
  return ops;
}

/** { ops, removed }, or null when the compared values are not computed. */
export function dedupeOps(engine, sheet, rect, cols, hasHeader) {
  const top = rect.r1 + (hasHeader ? 1 : 0);
  const seen = new Set();
  const keep = [];
  let removed = 0, unknown = false;
  for (let r = top; r <= rect.r2; r++) {
    const key = JSON.stringify(cols.map((c) => {
      const v = engine.valueAt(sheet, r, c);
      if (isSlowCalc(v)) unknown = true;
      return typeof v === 'string' ? v.toLowerCase() : isErr(v) ? v.code : v;
    }));
    if (seen.has(key)) { removed++; continue; }
    seen.add(key);
    keep.push(r);
  }
  if (unknown || budgetSpent()) return null;
  const ops = [];
  keep.forEach((src, i) => {
    const r = top + i;
    if (src === r) return;
    for (let c = rect.c1; c <= rect.c2; c++) {
      const cell = cellAt(sheet, src, c);
      let v = cell ? cell.v ?? null : null;
      if (typeof v === 'string' && v[0] === '=') v = translateFormula(v, r - src, 0);
      ops.push({ t: 'set', s: sheet.id, k: keyAt(sheet, r, c), v, st: cell ? cell.st || null : null, n: cell ? cell.n || null : null });
    }
  });
  for (let r = top + keep.length; r <= rect.r2; r++) {
    for (let c = rect.c1; c <= rect.c2; c++) if (cellAt(sheet, r, c)) ops.push({ t: 'set', s: sheet.id, k: keyAt(sheet, r, c), v: null, st: null, n: null });
  }
  return { ops, removed };
}

/** The block of non-empty cells around (r, c), like Excel's "current region". */
export function currentRegion(sheet, r, c, version) {
  const has = (rr, cc) => { const cell = cellAt(sheet, rr, cc); return !!(cell && cell.v !== undefined && cell.v !== ''); };
  let rect = { r1: r, r2: r, c1: c, c2: c };
  let grew = true, guard = 0;
  while (grew && guard++ < 200) {
    grew = false;
    const { r1, r2, c1, c2 } = rect;
    const tryGrow = (nr1, nr2, nc1, nc2, testR1, testR2, testC1, testC2) => {
      for (let rr = testR1; rr <= testR2; rr++) for (let cc = testC1; cc <= testC2; cc++) {
        if (rr < 0 || cc < 0 || rr >= sheet.rows.length || cc >= sheet.cols.length) continue;
        if (has(rr, cc)) { rect = { r1: nr1, r2: nr2, c1: nc1, c2: nc2 }; grew = true; return true; }
      }
      return false;
    };
    if (r1 > 0 && tryGrow(r1 - 1, r2, c1, c2, r1 - 1, r1 - 1, c1 - 1, c2 + 1)) continue;
    if (r2 < sheet.rows.length - 1 && tryGrow(r1, r2 + 1, c1, c2, r2 + 1, r2 + 1, c1 - 1, c2 + 1)) continue;
    if (c1 > 0 && tryGrow(r1, r2, c1 - 1, c2, r1 - 1, r2 + 1, c1 - 1, c1 - 1)) continue;
    if (c2 < sheet.cols.length - 1 && tryGrow(r1, r2, c1, c2 + 1, r1 - 1, r2 + 1, c2 + 1, c2 + 1)) continue;
  }
  // drop empty edge rows and columns (the start cell itself may be empty)
  const rowEmpty = (r) => { for (let c = rect.c1; c <= rect.c2; c++) if (has(r, c)) return false; return true; };
  const colEmpty = (c) => { for (let r = rect.r1; r <= rect.r2; r++) if (has(r, c)) return false; return true; };
  while (rect.r2 > rect.r1 && rowEmpty(rect.r2)) rect.r2--;
  while (rect.r1 < rect.r2 && rowEmpty(rect.r1)) rect.r1++;
  while (rect.c2 > rect.c1 && colEmpty(rect.c2)) rect.c2--;
  while (rect.c1 < rect.c2 && colEmpty(rect.c1)) rect.c1++;
  return rect;
}

/* ---------------------------------------------------------------- tables and filters */

export const TABLE_STYLES = {
  blue: '#3369d6', teal: '#0f8b8d', green: '#1f8039', orange: '#d9730d', red: '#cc3645', purple: '#7a4fd6', gray: '#6c707e', none: null
};

/** Positions and column names of a table, or null. */
export function tableInfo(engine, sheet, t, version, locale) {
  const rect = objRect(t, indexOf(sheet, version));
  if (!rect) return null;
  const columns = [];
  for (let c = rect.c1; c <= rect.c2; c++) {
    const v = t.header !== false ? engine.valueAt(sheet, rect.r1, c) : null;
    columns.push(v === null || v === '' ? 'Column' + (c - rect.c1 + 1) : formatValue(v, null, locale));
  }
  return { ...rect, name: t.name, sheet, header: t.header !== false, totals: !!t.totals, columns, t };
}

export function findTable(engine, doc, name, version, locale) {
  const low = String(name).toLowerCase();
  for (const sh of doc.sheets) for (const t of sh.tables || []) if (String(t.name).toLowerCase() === low) return tableInfo(engine, sh, t, version, locale);
  return null;
}
export function tableAt(engine, sheet, r, c, version, locale) {
  const ix = indexOf(sheet, version);
  for (const t of sheet.tables || []) { const rect = objRect(t, ix); if (inRect(rect, r, c)) return tableInfo(engine, sheet, t, version, locale); }
  return null;
}

/** Display text used by filters for a value. */
export const filterText = (v, st, locale) => (v === null || v === '' ? '' : formatValue(v, st, locale));

/**
 * Row indexes hidden by table filters. A value that is not computed (the formula work budget left it out, or it was
 * read in the step that ran out) is unknown: its row stays visible. The widget works the filters out again once the
 * values are computed.
 */
export function filteredRows(engine, sheet, version, locale) {
  const hidden = new Set();
  const ix = indexOf(sheet, version);
  for (const t of sheet.tables || []) {
    const f = t.filters;
    if (!f || typeof f !== 'object' || !Object.keys(f).length) continue;
    const rect = objRect(t, ix);
    if (!rect) continue;
    const top = rect.r1 + (t.header !== false ? 1 : 0), bottom = rect.r2 - (t.totals ? 1 : 0);
    for (const colId of Object.keys(f)) {
      const c = ix.cols.get(colId);
      if (c === undefined) continue;
      const spec = f[colId];
      // stored data: a filter that is not in the shape the app writes is ignored instead of breaking the sheet
      if (!spec || typeof spec !== 'object') continue;
      const allowed = Array.isArray(spec.vals) ? new Set(spec.vals) : null;
      for (let r = top; r <= bottom; r++) {
        if (hidden.has(r)) continue;
        const cell = cellAt(sheet, r, c);
        const was = budgetSpent();
        const v = engine.valueAt(sheet, r, c);
        if (isSlowCalc(v) || (!was && budgetSpent())) {
          // not known: the row stays shown, and formulas that ask whether it is hidden are not computed either (app.js
          // rowHidden). Read after the budget ran out, the whole result is computed again later (app.js filteredFor).
          (hidden.unknown || (hidden.unknown = new Set())).add(r);
          continue;
        }
        if (allowed && !allowed.has(filterText(v, cell && cell.st, locale))) { hidden.add(r); continue; }
        if (spec.cond && !condMatch(spec.cond, v, filterText(v, cell && cell.st, locale))) hidden.add(r);
      }
    }
  }
  return hidden;
}
function condMatch(cond, v, text) {
  const a = cond.a ?? '', low = text.toLowerCase(), al = String(a).toLowerCase();
  const num = typeof v === 'number' ? v : null, an = Number(String(a).replace(',', '.'));
  switch (cond.op) {
    case 'contains': return low.includes(al);
    case 'notContains': return !low.includes(al);
    case 'begins': return low.startsWith(al);
    case 'ends': return low.endsWith(al);
    case 'eq': return num !== null && !isNaN(an) ? num === an : low === al;
    case 'ne': return num !== null && !isNaN(an) ? num !== an : low !== al;
    case 'gt': return num !== null && num > an;
    case 'lt': return num !== null && num < an;
    case 'ge': return num !== null && num >= an;
    case 'le': return num !== null && num <= an;
    case 'blank': return text === '';
    case 'notBlank': return text !== '';
    default: return true;
  }
}

/* ---------------------------------------------------------------- conditional formatting */

export const CF_TYPES = {
  gt: 'Greater than', lt: 'Less than', ge: 'Greater than or equal to', le: 'Less than or equal to', eq: 'Equal to', ne: 'Not equal to',
  between: 'Between', notBetween: 'Not between', contains: 'Text contains', notContains: 'Text does not contain', begins: 'Text begins with',
  ends: 'Text ends with', blank: 'Is empty', notBlank: 'Is not empty', dup: 'Duplicate values', unique: 'Unique values',
  top: 'Top N items', bottom: 'Bottom N items', aboveAvg: 'Above average', belowAvg: 'Below average', today: 'Date is today',
  past: 'Date is in the past', future: 'Date is in the future', last7: 'Date in the last 7 days', thisMonth: 'Date in this month',
  formula: 'Custom formula is true', scale2: 'Color scale (2 colors)', scale3: 'Color scale (3 colors)', bar: 'Data bar', icons: 'Icon set (arrows)'
};

function ruleStats(engine, sheet, rect, cache, id) {
  if (cache.has(id)) return cache.get(id);
  const nums = [], counts = new Map();
  let cut = false;
  try {
    for (let r = rect.r1; r <= rect.r2; r++) {
      // every cell read counts against the formula work budget, so that rules over huge ranges cannot freeze the page
      charge(rect.c2 - rect.c1 + 1);
      for (let c = rect.c1; c <= rect.c2; c++) {
        const v = engine.valueAt(sheet, r, c);
        if (typeof v === 'number') nums.push(v);
        if (v !== null && v !== '') { const k = typeof v === 'string' ? v.toLowerCase() : isErr(v) ? v.code : String(v); counts.set(k, (counts.get(k) || 0) + 1); }
      }
    }
  } catch (e) {
    if (!isSlowCalc(e)) throw e;
    cut = true;
  }
  // the budget ran out (in charge, or in a value read last): no statistics yet (the rule is not drawn), the next
  // redraw reads them again
  if (cut || budgetSpent()) return null;
  const sorted = nums.slice().sort((a, b) => a - b);
  const stats = { min: sorted[0], max: sorted[sorted.length - 1], avg: nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0, sorted, counts };
  stats.mid = sorted.length ? (sorted.length % 2 ? sorted[sorted.length >> 1] : (sorted[(sorted.length >> 1) - 1] + sorted[sorted.length >> 1]) / 2) : 0;
  // kept statistics count as progress, so that the widget draws again for the rules after them (app.js continueCalc)
  cache.set(id, stats);
  engine.noteProgress();
  return stats;
}
function mix(c1, c2, t) {
  const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const a = p(c1), b = p(c2);
  return '#' + a.map((x, i) => Math.round(x + (b[i] - x) * t).toString(16).padStart(2, '0')).join('');
}
/* The group of a conditional formatting rule's formula for Engine.evalAt: one name for all the cells of the rule (the
 * formula text, translated for each cell, is not repeated in the engine's keys). Each rule object gets its own. */
const ruleGroups = new WeakMap();
let ruleGroupSeq = 0;
function ruleGroup(rule, part) {
  let g = ruleGroups.get(rule);
  if (!g) { g = 'cf' + (++ruleGroupSeq); ruleGroups.set(rule, g); }
  return g + part;
}
/** A number of a rule: as written, or the value of its formula (group: see ruleGroup; none for validation). */
function cfNumber(engine, sheet, x, r, c, group) {
  if (typeof x === 'string' && x[0] === '=') return engine.evalAt(x, sheet, r, c, group, !!group && cellIndependent(x, engine.book.names));
  const n = Number(String(x).replace(',', '.'));
  return x === '' || x === undefined || isNaN(n) ? x : n;
}

const HEX = /^#[0-9a-f]{6}$/i;
/** A #rrggbb color from a rule, or null: rules are stored data, so only checked colors leave this function. */
const ruleColor = (c) => (typeof c === 'string' && HEX.test(c) ? c : null);

/**
 * Conditional formatting result for a cell: { bg, fc, b, i, x, bar: {pct, color, neg}, icon }. Colors in the result
 * are always #rrggbb. A rule that cannot be evaluated (malformed stored data) is skipped instead of breaking the grid.
 */
export function cfFor(engine, sheet, r, c, ix, cache) {
  let out = null;
  for (const rule of sheet.cf || []) {
    if (!rule || typeof rule.type !== 'string' || !Object.prototype.hasOwnProperty.call(CF_TYPES, rule.type)) continue;
    const rect = objRect(rule, ix);
    if (!inRect(rect, r, c)) continue;
    let res;
    try { res = cfRule(engine, sheet, rule, rect, r, c, cache); } catch (e) { res = null; }
    if (!res) continue;
    out = out || {};
    if (rule.style && typeof rule.style === 'object') {
      for (const k of ['bg', 'fc']) { const col = ruleColor(rule.style[k]); if (col && out[k] === undefined) out[k] = col; }
      for (const k of ['b', 'i', 'x']) if (rule.style[k] && out[k] === undefined) out[k] = 1;
    }
    for (const k in res.extra) if (out[k] === undefined) out[k] = res.extra[k];
    if (rule.stop === true) break;
  }
  return out;
}

/** Whether one rule applies to the cell: null, or { extra } with the rule's own fill, data bar or icon. */
function cfRule(engine, sheet, rule, rect, r, c, cache) {
  if (rule.id === undefined || rule.id === null) return null;
  {
    const v = engine.valueAt(sheet, r, c);
    const text = v === null ? '' : isErr(v) ? v.code : typeof v === 'string' ? v : String(v);
    let hit = false;
    const extra = {};
    // a threshold that is not computed (the work budget's #CALC!) applies to nothing: compare() orders errors above every
    // number, so 'lt' would hit every cell. Thrown, it makes cfFor skip the rule for the cell.
    const known = (x) => { if (isSlowCalc(x)) throw x; return x; };
    const A = () => known(cfNumber(engine, sheet, rule.a, r, c, ruleGroup(rule, 'a'))), B = () => known(cfNumber(engine, sheet, rule.b, r, c, ruleGroup(rule, 'b')));
    switch (rule.type) {
      case 'gt': hit = v !== null && compare(v, A()) > 0; break;
      case 'lt': hit = v !== null && compare(v, A()) < 0; break;
      case 'ge': hit = v !== null && compare(v, A()) >= 0; break;
      case 'le': hit = v !== null && compare(v, A()) <= 0; break;
      case 'eq': hit = compare(v, A()) === 0; break;
      case 'ne': hit = compare(v, A()) !== 0; break;
      case 'between': case 'notBetween': {
        const a = A(), b = B(); const lo = compare(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
        const inside = v !== null && compare(v, lo) >= 0 && compare(v, hi) <= 0;
        hit = rule.type === 'between' ? inside : v !== null && !inside;
        break;
      }
      case 'contains': hit = text.toLowerCase().includes(String(rule.a || '').toLowerCase()); break;
      case 'notContains': hit = !text.toLowerCase().includes(String(rule.a || '').toLowerCase()); break;
      case 'begins': hit = text.toLowerCase().startsWith(String(rule.a || '').toLowerCase()); break;
      case 'ends': hit = text.toLowerCase().endsWith(String(rule.a || '').toLowerCase()); break;
      case 'blank': hit = text === ''; break;
      case 'notBlank': hit = text !== ''; break;
      case 'dup': case 'unique': {
        if (text === '') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        const n = s.counts.get(typeof v === 'string' ? v.toLowerCase() : isErr(v) ? v.code : String(v)) || 0;
        hit = rule.type === 'dup' ? n > 1 : n === 1;
        break;
      }
      case 'top': case 'bottom': {
        if (typeof v !== 'number') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        let n = Math.max(1, Number(rule.a) || 10);
        if (rule.pct) n = Math.max(1, Math.floor(s.sorted.length * n / 100));
        const lim = rule.type === 'top' ? s.sorted[Math.max(0, s.sorted.length - n)] : s.sorted[Math.min(s.sorted.length - 1, n - 1)];
        hit = rule.type === 'top' ? v >= lim : v <= lim;
        break;
      }
      case 'aboveAvg': case 'belowAvg': {
        if (typeof v !== 'number') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        hit = rule.type === 'aboveAvg' ? v > s.avg : v < s.avg;
        break;
      }
      case 'today': case 'past': case 'future': case 'last7': case 'thisMonth': {
        if (typeof v !== 'number') break;
        const now = new Date(); const today = dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()); const d = Math.floor(v);
        if (rule.type === 'today') hit = d === today;
        else if (rule.type === 'past') hit = d < today;
        else if (rule.type === 'future') hit = d > today;
        else if (rule.type === 'last7') hit = d <= today && d > today - 7;
        else { const p = serialToParts(d); hit = p.y === now.getFullYear() && p.m === now.getMonth() + 1; }
        break;
      }
      case 'formula': {
        if (typeof rule.formula !== 'string' || !rule.formula) break;
        const g = ruleGroup(rule, 'f');
        let res = engine.keptAt(sheet, r, c, g);
        if (res === undefined) {
          // relative references are relative to the top-left cell of the range, as in Excel. Translating a long formula
          // for each drawn cell takes time too, so it counts against the work budget (once the result is kept, not)
          const src = rule.formula[0] === '=' ? rule.formula : '=' + rule.formula;
          charge(1 + (src.length >> 6));
          res = engine.evalAt(translateFormula(src, r - rect.r1, c - rect.c1), sheet, r, c, g, cellIndependent(src, engine.book.names));
        }
        hit = res === true || (typeof res === 'number' && res !== 0);
        break;
      }
      case 'scale2': case 'scale3': {
        if (typeof v !== 'number') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        // the rule's own colors only when they are all #rrggbb and as many as the scale needs
        const def = rule.type === 'scale2' ? ['#ffffff', '#57bb8a'] : ['#e67c73', '#ffd666', '#57bb8a'];
        const cols = Array.isArray(rule.colors) && rule.colors.length === def.length && rule.colors.every(ruleColor) ? rule.colors : def;
        if (s.max === s.min) { extra.bg = cols[cols.length - 1]; hit = true; break; }
        if (rule.type === 'scale2') extra.bg = mix(cols[0], cols[1], (v - s.min) / (s.max - s.min));
        else extra.bg = v <= s.mid ? mix(cols[0], cols[1], s.mid === s.min ? 1 : (v - s.min) / (s.mid - s.min)) : mix(cols[1], cols[2], s.max === s.mid ? 1 : (v - s.mid) / (s.max - s.mid));
        if (!ruleColor(extra.bg)) delete extra.bg;
        hit = true;
        break;
      }
      case 'bar': {
        if (typeof v !== 'number') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        const lo = Math.min(0, s.min), hi = Math.max(0, s.max);
        extra.bar = { pct: hi === lo ? 100 : Math.abs(v) / (hi - lo) * 100, color: ruleColor(rule.color) || '#3369d6', neg: v < 0, zero: hi === lo ? 0 : -lo / (hi - lo) * 100 };
        hit = true;
        break;
      }
      case 'icons': {
        if (typeof v !== 'number') break;
        const s = ruleStats(engine, sheet, rect, cache, rule.id);
        if (!s) break;
        const t = s.max === s.min ? 1 : (v - s.min) / (s.max - s.min);
        extra.icon = t >= 0.67 ? 'up' : t >= 0.33 ? 'side' : 'down';
        hit = true;
        break;
      }
      default: break;
    }
    return hit ? { extra } : null;
  }
}

/* ---------------------------------------------------------------- data validation */

export const DV_TYPES = { list: 'List of items', checkbox: 'Checkbox', number: 'Number', int: 'Whole number', date: 'Date', length: 'Text length', formula: 'Custom formula' };
export const DV_OPS = { between: 'between', notBetween: 'not between', eq: 'equal to', ne: 'not equal to', gt: 'greater than', lt: 'less than', ge: 'greater than or equal to', le: 'less than or equal to' };

export function dvFor(sheet, r, c, ix) {
  for (const rule of sheet.dv || []) if (inRect(objRect(rule, ix), r, c)) return rule;
  return null;
}

/** Items of a list rule: "a, b, c" or a range ("=$A$1:$A$10", "=Lists!A:A"). */
export function listItems(engine, doc, sheet, rule, locale, names = new Set()) {
  const src = String(rule.source || '');
  if (src[0] === '=') {
    const rect = parseRect(src.slice(1));
    if (!rect) {
      // a defined name; names can point at other names, so a cycle (or a very long chain) ends with no items
      const key = src.slice(1).toUpperCase();
      if (names.has(key) || names.size >= 16) return [];
      names.add(key);
      const def = doc.names && Object.prototype.hasOwnProperty.call(doc.names, key) ? doc.names[key] : null;
      return typeof def === 'string' && def ? listItems(engine, doc, sheet, { source: def[0] === '=' ? def : '=' + def }, locale, names) : [];
    }
    const sh = rect.sheet ? doc.sheets.find((s) => s.name.toLowerCase() === rect.sheet.toLowerCase()) : sheet;
    if (!sh) return [];
    const out = [], seen = new Set();
    for (let r = rect.r1; r <= Math.min(rect.r2, sh.rows.length - 1); r++) for (let c = rect.c1; c <= Math.min(rect.c2, sh.cols.length - 1); c++) {
      const v = engine.valueAt(sh, r, c);
      // a value the formula work budget left out is no item: the list is then marked incomplete
      if (isSlowCalc(v)) { out.incomplete = true; continue; }
      if (v === null || v === '') continue;
      const cell = cellAt(sh, r, c);
      const t = formatValue(v, cell && cell.st, locale);
      if (!seen.has(t)) { seen.add(t); out.push(t); }
    }
    if (budgetSpent()) out.incomplete = true;
    return out;
  }
  return src.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
}

/** Returns null when `parsed` (from parseInput) is valid for the rule, else a message. */
export function validate(engine, doc, sheet, r, c, rule, parsed, locale) {
  if (!rule || parsed.kind === 'empty') return null;
  const msg = rule.msg || null;
  const v = parsed.kind === 'formula' ? null : parsed.num !== undefined ? parsed.num : parsed.raw;
  const cmp = (x) => {
    const a = cfNumber(engine, sheet, rule.a, r, c), b = cfNumber(engine, sheet, rule.b, r, c);
    switch (rule.op || 'between') {
      case 'between': return x >= a && x <= b;
      case 'notBetween': return x < a || x > b;
      case 'eq': return x === a; case 'ne': return x !== a;
      case 'gt': return x > a; case 'lt': return x < a; case 'ge': return x >= a; case 'le': return x <= a;
      default: return true;
    }
  };
  switch (rule.type) {
    case 'list': {
      if (parsed.kind === 'formula') return null;
      const list = listItems(engine, doc, sheet, rule, locale);
      // items that are not computed (the work budget) are unknown: no value is rejected for lack of them
      if (list.incomplete) return null;
      const items = list.map((x) => x.toLowerCase());
      const text = String(parsed.raw).replace(/^'/, '');
      const shown = parsed.num !== undefined ? formatValue(parsed.num, parsed.fmt, locale).toLowerCase() : text.toLowerCase();
      return items.includes(text.toLowerCase()) || items.includes(shown) ? null : msg || 'Choose a value from the list.';
    }
    case 'checkbox': return parsed.kind === 'bool' ? null : msg || 'This cell takes TRUE or FALSE.';
    case 'number': return typeof v === 'number' && cmp(v) ? null : msg || 'Enter a number ' + DV_OPS[rule.op || 'between'] + ' ' + (rule.a ?? '') + (rule.op === 'between' || rule.op === 'notBetween' || !rule.op ? ' and ' + (rule.b ?? '') : '') + '.';
    case 'int': return typeof v === 'number' && Number.isInteger(v) && cmp(v) ? null : msg || 'Enter a whole number ' + DV_OPS[rule.op || 'between'] + ' ' + (rule.a ?? '') + (rule.op === 'between' || rule.op === 'notBetween' || !rule.op ? ' and ' + (rule.b ?? '') : '') + '.';
    case 'date': {
      if (!(parsed.kind === 'date' || parsed.kind === 'datetime')) return msg || 'Enter a date.';
      const toSerial = (x) => { if (typeof x === 'number') return x; const p = parseInput(String(x ?? ''), locale); return p.num ?? x; };
      const a = toSerial(cfNumber(engine, sheet, rule.a, r, c)), b = toSerial(cfNumber(engine, sheet, rule.b, r, c));
      const x = v;
      const ok = (() => { switch (rule.op || 'between') { case 'between': return x >= a && x <= b; case 'notBetween': return x < a || x > b; case 'gt': return x > a; case 'lt': return x < a; case 'ge': return x >= a; case 'le': return x <= a; case 'eq': return Math.floor(x) === Math.floor(a); default: return true; } })();
      return ok ? null : msg || 'The date is outside the allowed range.';
    }
    case 'length': {
      const len = String(parsed.raw).replace(/^'/, '').length;
      return cmp(len) ? null : msg || 'The text length is not allowed.';
    }
    case 'formula': return null; // checked after the value is stored
    default: return null;
  }
}

/* ---------------------------------------------------------------- misc */

/** Applies a style patch to every cell in a rectangle (null values remove keys). */
export function styleOps(sheet, rect, patch, opts = {}) {
  const ops = [];
  for (let r = rect.r1; r <= rect.r2; r++) {
    if (opts.skipHidden && sheet.hr[sheet.rows[r]]) continue;
    for (let c = rect.c1; c <= rect.c2; c++) {
      const k = keyAt(sheet, r, c);
      const cell = sheet.cells[k];
      const st = { ...(cell && cell.st ? cell.st : {}) };
      const p = typeof patch === 'function' ? patch(st, r, c) : patch;
      for (const key in p) { if (p[key] === null || p[key] === undefined || p[key] === false) delete st[key]; else st[key] = p[key]; }
      const next = Object.keys(st).length ? st : null;
      if (JSON.stringify(next) !== JSON.stringify(cell && cell.st ? cell.st : null)) ops.push({ t: 'set', s: sheet.id, k, st: next });
    }
  }
  return ops;
}

/**
 * Style ops for several rectangles (the ranges of a selection) as one change: patchFor(rect) is the patch of each
 * rectangle (an object, or a function of the cell's style, row, column and whether an earlier rectangle of this change
 * already styled the cell). A cell in more than one of them gets one op, made from what the earlier rectangles left (a
 * border around each of two overlapping ranges keeps both).
 */
export function styleOpsMulti(sheet, rects, patchFor, opts = {}) {
  const next = new Map(); // cell key → its style after the rectangles so far
  for (const rect of rects) {
    const patch = patchFor(rect);
    for (let r = rect.r1; r <= rect.r2; r++) {
      if (opts.skipHidden && sheet.hr[sheet.rows[r]]) continue;
      for (let c = rect.c1; c <= rect.c2; c++) {
        const k = keyAt(sheet, r, c);
        const cell = sheet.cells[k];
        const again = next.has(k);
        const before = again ? next.get(k) : (cell && cell.st ? cell.st : null);
        const st = { ...(before || {}) };
        const p = typeof patch === 'function' ? patch(st, r, c, again) : patch;
        for (const key in p) { if (p[key] === null || p[key] === undefined || p[key] === false) delete st[key]; else st[key] = p[key]; }
        next.set(k, Object.keys(st).length ? st : null);
      }
    }
  }
  const ops = [];
  for (const [k, st] of next) {
    const cell = sheet.cells[k];
    if (JSON.stringify(st) !== JSON.stringify(cell && cell.st ? cell.st : null)) ops.push({ t: 'set', s: sheet.id, k, st });
  }
  return ops;
}

/** Border ops: mode 'all' | 'outer' | 'inner' | 'top' | 'bottom' | 'left' | 'right' | 'none' | 'thick'. */
export function borderOps(sheet, rect, mode, color) {
  return styleOps(sheet, rect, borderPatch(rect, mode, color));
}
/**
 * The style patch of a border mode for the cells of a rectangle (see borderOps). again: an earlier rectangle of the
 * same change already gave the cell its borders (several selected ranges, styleOpsMulti), which a thick outline keeps.
 */
export function borderPatch(rect, mode, color) {
  return (st, r, c, again) => {
    let sides = mode === 'none' ? '' : (st.bd || '');
    const add = (s) => { if (!sides.toLowerCase().includes(s)) sides += s; };
    const thick = (s) => { sides = sides.replace(s, '').replace(s.toUpperCase(), '') + s.toUpperCase(); };
    const top = r === rect.r1, bottom = r === rect.r2, left = c === rect.c1, right = c === rect.c2;
    switch (mode) {
      case 'all': sides = 'tlrb'; break;
      case 'outer': if (top) add('t'); if (bottom) add('b'); if (left) add('l'); if (right) add('r'); break;
      case 'thick': if (!again) sides = ''; if (top) thick('t'); if (bottom) thick('b'); if (left) thick('l'); if (right) thick('r'); break;
      case 'inner': if (!bottom) add('b'); if (!right) add('r'); break;
      case 'top': if (top) add('t'); break;
      case 'bottom': if (bottom) add('b'); break;
      case 'left': if (left) add('l'); break;
      case 'right': if (right) add('r'); break;
      default: break;
    }
    return { bd: sides || null, bc: sides && color ? color : (sides ? st.bc : null) };
  };
}

/** Rewrites references to a defined name or text in formulas (used by Find & Replace in formulas). */
export function replaceInFormula(f, find, repl) {
  return f.split(find).join(repl);
}

export { rewriteRefs, parseFormula };
