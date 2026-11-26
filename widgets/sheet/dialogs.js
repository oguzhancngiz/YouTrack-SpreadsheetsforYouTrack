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
/* Spreadsheets for YouTrack – dialogs and side panels: tables, charts, notes, conditional formatting, data validation, sort,
 * remove duplicates, split text, find and replace, named ranges, sizes, number formats, functions, links, filters,
 * import options and the change history. Every change goes through A.commit() (shared, undoable). */
import { h, esc, icon, dialog, confirmDialog, promptDialog, field, select, checkbox, colorPicker, popover, closePopups, toast, relTime, initials, peerColor, contrastText, luminance, themedColors, paintBar, hexColor } from './ui.js';
import * as M from './model.js';
import { colName, formatValue, parseInput, isErr, FUNCTION_NAMES, FUNCTION_HELP, formatWithPattern, quoteSheet, editText, formulaSyntaxError, numToText, isSlowCalc } from './formula.js';
import { CHART_TYPES } from './charts.js';

const Core = window.SheetCore;
/** The stored objects of one kind (tables, rules, …): anything that is not an object counts as absent. */
const objs = (v) => (Array.isArray(v) ? v.filter((o) => o && typeof o === 'object' && !Array.isArray(o)) : []);
/** A stored value that should be text or a number: objects and arrays count as absent. */
const plain = (v) => (v !== null && typeof v === 'object' ? undefined : v);

export const prompt = (title, label, value, opts) => promptDialog(title, label, value, opts);

function rangeInput(A, rect, sheetPrefix) {
  const sh = A.sheet();
  return h('input', { class: 'input mono', value: (sheetPrefix ? quoteSheet(sh.name) + '!' : '') + M.rectText(rect), spellcheck: 'false', 'aria-label': 'Range' });
}
function readRange(A, input) {
  const rect = M.parseRect(input.value);
  if (!rect) { toast('Enter a range such as A1:D20.', 'error'); return null; }
  const sh = rect.sheet ? A.engine.sheetByName(rect.sheet) : A.sheet();
  if (!sh) { toast('There is no sheet "' + rect.sheet + '".', 'error'); return null; }
  return { sheet: sh, r1: rect.r1, c1: rect.c1, r2: Math.min(rect.r2, sh.rows.length - 1), c2: Math.min(rect.c2, sh.cols.length - 1) };
}
function defaultRegion(A) {
  const s = A.selRect();
  if (s.r1 === s.r2 && s.c1 === s.c2) return M.currentRegion(A.sheet(), s.r1, s.c1);
  return s;
}
/**
 * Computes the values a command reads before it acts on them (the formula work budget may have left some out, see
 * A.computeValues in app.js). True when they are all computed and the dialog `dlg` (when given) is still open; false
 * when it was closed meanwhile (nothing is done) or after saying why the values are not there.
 */
async function valuesReady(A, areas, dlg) {
  const status = await A.computeValues(areas, dlg ? { live: () => dlg.el.isConnected } : undefined);
  if (dlg && !dlg.el.isConnected) return false;
  if (status === 'ok') return true;
  A.notComputed(status);
  return false;
}
/** Sorts a range once the values of its sort columns are computed. False when they are not (nothing was done). */
async function sortRange(A, sh, rect, keys, hasHeader, dlg) {
  const top = rect.r1 + (hasHeader ? 1 : 0);
  if (!(await valuesReady(A, keys.map((k) => ({ sheet: sh, r1: top, c1: k.c, r2: rect.r2, c2: k.c })), dlg))) return false;
  const ops = M.sortOps(A.engine, sh, rect, keys, hasHeader);
  if (!ops) { A.notComputed('pending'); return false; }
  A.commit(ops, { label: 'Sort' });
  return true;
}

/* ================================================================ tables */

export function tableDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const x = A.ix(sh);
  const existing = objs(sh.tables).find((t) => M.inRect(M.objRect(t, x), A.S.sel.r, A.S.sel.c));
  const rect = existing ? M.objRect(existing, x) : defaultRegion(A);
  const t = existing || { name: M.uniqueTableName(A.view()), header: true, totals: false, style: 'blue', banded: true, filter: true, filters: {} };
  const name = h('input', { class: 'input', value: t.name, spellcheck: 'false' });
  const range = rangeInput(A, rect);
  const header = checkbox('The first row has headers', t.header !== false);
  const totals = checkbox('Total row', !!t.totals);
  const banded = checkbox('Banded rows', t.banded !== false);
  const filter = checkbox('Filter buttons', t.filter !== false);
  let style = t.style || 'blue';
  const swatches = h('div', { class: 'style-row' });
  const paint = () => {
    swatches.innerHTML = '';
    Object.entries(M.TABLE_STYLES).forEach(([k, color]) => {
      const b = h('button', { class: 'tstyle' + (k === style ? ' on' : ''), type: 'button', title: k, 'aria-label': 'Style ' + k, onclick: () => { style = k; paint(); } },
        h('span', { class: 'th', style: { background: color || 'transparent', borderColor: color || 'var(--border)' } }),
        h('span', { class: 'tb1', style: { background: color ? color + '22' : 'transparent' } }), h('span', { class: 'tb2' }));
      swatches.appendChild(b);
    });
  };
  paint();
  // the help names this table and one of its columns (a numeric one when there is one)
  const sample = (() => {
    let first = null;
    if (t.header !== false && rect) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        const head = formatValue(A.engine.valueAt(sh, rect.r1, c), null, A.LOCALE);
        if (!head) continue;
        if (first === null) first = head;
        for (let r = rect.r1 + 1; r <= Math.min(rect.r2, rect.r1 + 20); r++) if (typeof A.engine.valueAt(sh, r, c) === 'number') return ['SUM', head];
      }
    }
    return ['COUNTA', first || 'Column1'];
  })();
  const helpText = 'Use it in formulas: =' + sample[0] + '(' + t.name + '[' + sample[1].replace(/([\[\]'#@])/g, "'$1") + '])';
  const body = h('div', { class: 'form' },
    h('div', { class: 'row2' }, field('Table name', name, helpText), field('Range', range)),
    h('div', { class: 'checks' }, header.el, totals.el, banded.el, filter.el),
    field('Style', swatches));
  const buttons = [{ label: 'Cancel' }];
  const dlg = dialog(existing ? 'Table ' + t.name : 'Format as table', body, buttons.concat([{
    label: existing ? 'Save' : 'Create table', primary: true, action: async () => {
      const R = readRange(A, range);
      if (!R) return false;
      if (R.sheet !== sh) { toast('A table must be on the current sheet.', 'error'); return false; }
      const n = name.value.trim();
      if (!Core.NAME_RE.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n)) { toast('Table names start with a letter and contain only letters, digits, "_" and "." (not a cell address).', 'error'); return false; }
      if (A.view().sheets.some((s2) => objs(s2.tables).some((o) => o.id !== t.id && typeof o.name === 'string' && o.name.toLowerCase() === n.toLowerCase()))) { toast('Another table is already named ' + n + '.', 'error'); return false; }
      for (const o of objs(sh.tables)) {
        if (o.id === t.id) continue;
        const r2 = M.objRect(o, x);
        if (r2 && !(r2.r2 < R.r1 || r2.r1 > R.r2 || r2.c2 < R.c1 || r2.c1 > R.c2)) { toast('Tables cannot overlap (' + o.name + ').', 'error'); return false; }
      }
      // a new totals row adds up the numeric columns: their values are read once they are computed
      if (totals.input.checked && !t.totals && !(await valuesReady(A, [R], dlg))) return false;
      const ops = [];
      let r2 = R.r2;
      const id = t.id || Core.randomId();
      if (totals.input.checked && !t.totals) {
        // add a totals row under the data with SUBTOTAL formulas for numeric columns
        r2 = R.r2 + 1;
        ops.push(...M.growOps(sh, r2, R.c2));
        const rowIds = sh.rows.slice();
        ops.forEach((op) => { if (op.t === 'rows') rowIds.push(...op.add); });
        const rowEmpty = !Array.from({ length: R.c2 - R.c1 + 1 }, (_, j) => M.cellAt(sh, r2, R.c1 + j)).some((c) => c && c.v !== undefined);
        if (!rowEmpty) { toast('The row under the table is not empty. Insert a row first.', 'error'); return false; }
        for (let c = R.c1; c <= R.c2; c++) {
          const k = rowIds[r2] + ':' + sh.cols[c];
          const top = R.r1 + (header.input.checked ? 1 : 0);
          let numeric = false;
          for (let r = top; r <= R.r2; r++) if (typeof A.engine.valueAt(sh, r, c) === 'number') { numeric = true; break; }
          const colLabel = header.input.checked ? formatValue(A.engine.valueAt(sh, R.r1, c), null, A.LOCALE) || 'Column' + (c - R.c1 + 1) : 'Column' + (c - R.c1 + 1);
          if (c === R.c1 && !numeric) ops.push({ t: 'set', s: sh.id, k, v: 'Total', st: { b: 1 } });
          else if (numeric) ops.push({ t: 'set', s: sh.id, k, v: `=SUBTOTAL(109,${n}[${colLabel.replace(/([\[\]'#@])/g, "'$1")}])`, st: { b: 1 } });
        }
        const ids = { r1: sh.rows[R.r1], c1: sh.cols[R.c1], r2: rowIds[r2], c2: sh.cols[R.c2] };
        ops.push({ t: 'obj', s: sh.id, kind: 'tables', id, val: { ...t, id: undefined, name: n, ...ids, header: header.input.checked, totals: true, banded: banded.input.checked, filter: filter.input.checked, style, filters: t.filters || {} } });
      } else {
        let R2 = R;
        if (!totals.input.checked && t.totals && existing) {
          // remove the totals row contents
          for (let c = R.c1; c <= R.c2; c++) { const cell = M.cellAt(sh, R.r2, c); if (cell) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, R.r2, c), v: null, st: null }); }
          R2 = { ...R, r2: R.r2 - 1 };
        }
        ops.push({ t: 'obj', s: sh.id, kind: 'tables', id, val: { ...t, id: undefined, name: n, ...M.rectIds(sh, R2), header: header.input.checked, totals: totals.input.checked, banded: banded.input.checked, filter: filter.input.checked, style, filters: filter.input.checked ? t.filters || {} : {} } });
      }
      // a header row needs headers: fill empty ones
      if (header.input.checked && !existing) {
        for (let c = R.c1; c <= R.c2; c++) { const cell = M.cellAt(sh, R.r1, c); if (!cell || cell.v === undefined) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, R.r1, c), v: 'Column' + (c - R.c1 + 1) }); }
      }
      if (existing && n !== t.name && typeof t.name === 'string' && Core.NAME_RE.test(t.name)) {
        // rename structured references (the old name is stored data: every regex character in it is escaped)
        const re = new RegExp('(^|[^\\w.])' + t.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=\\[|\\b)', 'gi');
        for (const s2 of A.view().sheets) for (const k in s2.cells) { const v = (s2.cells[k] || {}).v; if (typeof v === 'string' && v[0] === '=' && re.test(v)) { re.lastIndex = 0; ops.push({ t: 'set', s: s2.id, k, v: v.replace(re, '$1' + n) }); } re.lastIndex = 0; }
      }
      return A.commit(ops, { label: existing ? 'Table' : 'Create table' }) ? undefined : false;
    }
  }]), { footLeft: existing ? h('button', { class: 'btn danger', type: 'button', onclick: async (e) => {
    const ok = await confirmDialog('Convert to range', 'Remove the table ' + t.name + '? The data stays; table formatting, filters and the table name go away. Formulas that use ' + t.name + '[…] will show #REF!.', 'Convert to range');
    if (!ok) return;
    A.commit([{ t: 'obj', s: sh.id, kind: 'tables', id: t.id, val: null }], { label: 'Remove table' });
    document.querySelector('.modal-back') && document.querySelector('.modal-back').remove();
  } }, 'Convert to range') : null });
}

/* ---- filter popover in a table header */

export async function filterPopover(A, tableId, c, anchor) {
  const sh = A.sheet();
  const t = objs(sh.tables).find((o) => o.id === tableId);
  if (!t) return;
  const x = A.ix(sh);
  const rect = M.objRect(t, x);
  if (!rect) return;
  const colId = sh.cols[c];
  const spec = (t.filters || {})[colId] || {};
  const top = rect.r1 + (t.header !== false ? 1 : 0), bottom = rect.r2 - (t.totals ? 1 : 0);
  // the list offers the column's values, so they are computed first. A formula that needs more calculation than the
  // spreadsheet does at once has no value to offer: its row stays visible whatever is chosen.
  const sel = JSON.stringify(A.S.sel);
  const status = await A.computeValues([{ sheet: sh, r1: top, c1: c, r2: bottom, c2: c }]);
  if (status === 'pending' || status === 'changed') { A.notComputed(status); return; }
  // the button may have been drawn again meanwhile; when it scrolled away, or the person went on to something else
  // while the values were computed, the list does not open
  const btn = [...document.querySelectorAll('[data-flt]')].find((b) => b.dataset.flt === tableId && +b.dataset.fc === c);
  if (!btn || JSON.stringify(A.S.sel) !== sel || document.querySelector('.modal-back')) return;
  anchor = btn.getBoundingClientRect();
  const values = new Map();
  let notCalculated = false;
  for (let r = top; r <= bottom; r++) {
    const cell = M.cellAt(sh, r, c);
    const v = A.engine.valueAt(sh, r, c);
    if (isSlowCalc(v)) { notCalculated = true; continue; }
    const txt = M.filterText(v, cell && cell.st, A.LOCALE);
    if (!values.has(txt)) values.set(txt, v);
  }
  const sorted = [...values.entries()].sort((a, b) => (a[0] === '' ? 1 : b[0] === '' ? -1 : typeof a[1] === 'number' && typeof b[1] === 'number' ? a[1] - b[1] : a[0].localeCompare(b[0], A.LOCALE, { numeric: true })));
  const allowed = new Set(Array.isArray(spec.vals) ? spec.vals : sorted.map((e) => e[0]));
  const ro = !A.canEdit();
  const list = h('div', { class: 'flt-list' });
  const search = h('input', { class: 'input small', placeholder: 'Search', 'aria-label': 'Search values' });
  const boxes = [];
  const fill = () => {
    list.innerHTML = '';
    boxes.length = 0;
    const q = search.value.toLowerCase();
    sorted.filter(([txt]) => txt.toLowerCase().includes(q)).slice(0, 1000).forEach(([txt]) => {
      const cb = h('input', { type: 'checkbox', checked: allowed.has(txt), disabled: ro });
      cb.addEventListener('change', () => { if (cb.checked) allowed.add(txt); else allowed.delete(txt); });
      boxes.push(cb);
      list.appendChild(h('label', { class: 'check' }, cb, h('span', null, txt === '' ? '(Empty)' : txt)));
    });
  };
  fill();
  search.addEventListener('input', fill);
  const condOp = select({ '': 'No condition', contains: 'Text contains', notContains: 'Does not contain', begins: 'Begins with', ends: 'Ends with', eq: 'Equal to', ne: 'Not equal to', gt: 'Greater than', ge: 'Greater or equal', lt: 'Less than', le: 'Less or equal', blank: 'Is empty', notBlank: 'Is not empty' }, spec.cond ? spec.cond.op : '', { disabled: ro, 'aria-label': 'Condition' });
  const condVal = h('input', { class: 'input small', value: spec.cond ? spec.cond.a || '' : '', disabled: ro, placeholder: 'Value', 'aria-label': 'Condition value' });
  const apply = (filters) => {
    closePopups();
    A.commit([{ t: 'obj', s: sh.id, kind: 'tables', id: t.id, val: { ...t, filters } }], { label: 'Filter' });
  };
  const sortBtn = (desc) => h('button', { class: 'menu-item', type: 'button', disabled: ro, onclick: () => {
    closePopups();
    sortRange(A, sh, { ...rect, r2: bottom }, [{ c, desc }], t.header !== false);
  } }, h('span', { class: 'mi-icon', html: icon(desc ? 'sortDesc' : 'sortAsc') }), h('span', { class: 'mi-label' }, desc ? 'Sort Z → A' : 'Sort A → Z'));
  const content = h('div', { class: 'flt' },
    sortBtn(false), sortBtn(true),
    h('div', { class: 'menu-sep' }),
    h('div', { class: 'flt-cond' }, condOp, condVal),
    h('div', { class: 'flt-tools' },
      h('button', { class: 'link-btn', type: 'button', disabled: ro, onclick: () => { sorted.forEach(([v]) => allowed.add(v)); boxes.forEach((b) => { b.checked = true; }); } }, 'Select all'),
      h('button', { class: 'link-btn', type: 'button', disabled: ro, onclick: () => { allowed.clear(); boxes.forEach((b) => { b.checked = false; }); } }, 'Clear')),
    sorted.length > 8 ? search : null,
    list,
    notCalculated ? h('div', { class: 'muted small' }, 'Rows whose formula is not calculated (#CALC!) are not filtered.') : null,
    h('div', { class: 'flt-foot' },
      h('button', { class: 'btn small', type: 'button', disabled: ro || !spec.vals && !spec.cond, onclick: () => { const f = { ...(t.filters || {}) }; delete f[colId]; apply(f); } }, 'Clear filter'),
      h('button', { class: 'btn small primary', type: 'button', disabled: ro, onclick: () => {
        const f = { ...(t.filters || {}) };
        const next = {};
        if (allowed.size < sorted.length) next.vals = [...allowed];
        if (condOp.value) next.cond = { op: condOp.value, a: condVal.value };
        if (Object.keys(next).length) f[colId] = next; else delete f[colId];
        apply(f);
      } }, 'OK')));
  popover(content, anchor, { class: 'flt-pop' });
}

/* ================================================================ charts */

export function chartDialog(A, chartId) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const x = A.ix(sh);
  const existing = chartId ? objs(sh.charts).find((c) => c.id === chartId) : null;
  const rect = existing ? M.objRect(existing, x) : defaultRegion(A);
  const type = select(CHART_TYPES, existing ? existing.type : 'column');
  const title = h('input', { class: 'input', value: existing ? existing.title || '' : '' });
  const range = rangeInput(A, rect || A.selRect());
  const legend = checkbox('Show legend', existing ? existing.legend !== false : true);
  dialog(existing ? 'Edit chart' : 'Insert chart', h('div', { class: 'form' },
    h('div', { class: 'row2' }, field('Chart type', type), field('Data range', range, 'The first row and column are used as labels when they contain text.')),
    field('Title', title), legend.el), [
    { label: 'Cancel' },
    { label: existing ? 'Save' : 'Insert', primary: true, action: () => {
      const R = readRange(A, range);
      if (!R) return false;
      if (R.sheet !== sh) { toast('The data must be on the current sheet.', 'error'); return false; }
      const id = existing ? existing.id : Core.randomId();
      const anchorC = Math.min(sh.cols.length - 1, R.c2 + 1);
      const val = {
        ...(existing || {}), id: undefined, type: type.value, title: title.value.trim(), legend: legend.input.checked, ...M.rectIds(sh, R),
        ar: existing ? existing.ar : sh.rows[R.r1], ac: existing ? existing.ac : sh.cols[anchorC], ox: existing ? existing.ox || 0 : 12, oy: existing ? existing.oy || 0 : 0,
        w: existing ? existing.w : 480, h: existing ? existing.h : 280
      };
      return A.commit([{ t: 'obj', s: sh.id, kind: 'charts', id, val }], { label: existing ? 'Edit chart' : 'Insert chart' }) ? undefined : false;
    } }]);
}

/* ================================================================ notes, links */

export function noteDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const { r, c } = A.S.sel;
  const cell = M.cellAt(sh, r, c);
  const text = h('textarea', { class: 'input', rows: 5, autofocus: true }, cell && cell.n ? cell.n : '');
  text.value = cell && cell.n ? cell.n : '';
  dialog('Note for ' + colName(c) + (r + 1), field('Note', text, 'Notes show when the pointer rests on the cell (marked in the top-right corner).'), [
    { label: 'Cancel' },
    cell && cell.n ? { label: 'Delete note', danger: true, action: () => { A.commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), n: null }], { label: 'Note' }); } } : null,
    { label: 'Save', primary: true, action: () => { A.commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), n: text.value.trim() || null }], { label: 'Note' }); } }
  ].filter(Boolean));
}

export function linkDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const { r, c } = A.S.sel;
  const v = A.engine.valueAt(sh, r, c);
  const url = h('input', { class: 'input', placeholder: 'https://', autofocus: true });
  const label = h('input', { class: 'input', value: typeof v === 'string' && !/^https?:/i.test(v) ? v : '' });
  dialog('Insert link', h('div', { class: 'form' }, field('Link', url), field('Text to show', label)), [
    { label: 'Cancel' },
    { label: 'Insert', primary: true, action: () => {
      const u = url.value.trim();
      if (!/^(https?:\/\/|mailto:)/i.test(u)) { toast('Enter a link that starts with https:// or mailto:', 'error'); return false; }
      const l = label.value.trim();
      const q = (s) => '"' + s.replace(/"/g, '""') + '"';
      A.commit([{ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), v: l ? `=HYPERLINK(${q(u)},${q(l)})` : u }], { label: 'Link' });
    } }]);
}

/* ================================================================ conditional formatting */

const CF_PRESETS = {
  red: { bg: '#f4cccc', fc: '#990000', label: 'Light red fill, dark red text' },
  yellow: { bg: '#fff2cc', fc: '#7f6000', label: 'Yellow fill, dark yellow text' },
  green: { bg: '#d9ead3', fc: '#274e13', label: 'Green fill, dark green text' },
  blue: { bg: '#c9daf8', fc: '#1c4587', label: 'Blue fill, dark blue text' },
  bold: { b: 1, label: 'Bold text' },
  redText: { fc: '#cc0000', label: 'Red text' }
};
/** The colors of a color scale: the stored ones when they are a #rrggbb color for each point, otherwise the defaults. */
function scaleColors(colors, type) {
  const def = type === 'scale2' ? ['#ffffff', '#57bb8a'] : ['#e67c73', '#ffd666', '#57bb8a'];
  return Array.isArray(colors) && colors.length === def.length && colors.every((c) => hexColor(c)) ? colors : def;
}

export function cfDialog(A) {
  const sh = A.sheet();
  const box = h('div', { class: 'cf' });
  const d = dialog('Conditional formatting', box, [{ label: 'Close' }], { wide: true });
  const x = A.ix(sh);
  const renderList = () => {
    box.innerHTML = '';
    const rules = objs(sh.cf);
    const list = h('div', { class: 'rule-list' });
    if (!rules.length) list.appendChild(h('div', { class: 'muted pad' }, 'No rules on this sheet yet.'));
    rules.forEach((rule, i) => {
      const rect = M.objRect(rule, x);
      // the rule's format in the colors the cells get in the current theme (as cellCss in app.js); scales and data bars
      // keep their colors in both themes, as in the cells. Rules are stored data: only #rrggbb colors reach the style,
      // anything else counts as no color.
      const rs = rule.style && typeof rule.style === 'object' ? rule.style : {};
      const type = typeof rule.type === 'string' ? rule.type : '';
      const bg = hexColor(rs.bg);
      const tc = themedColors(bg, hexColor(rs.fc));
      const css = tc.parts.slice();
      if (!bg && type.startsWith('scale')) {
        const cols = scaleColors(rule.colors, type);
        const L = cols.reduce((sum, c) => sum + (luminance(c) ?? 1), 0) / cols.length;
        css.push('background:linear-gradient(90deg,' + cols.join(',') + ')', 'color:' + (L > 0.3 ? '#1f2023' : '#f5f6f8'));
      }
      else if (!bg && type === 'bar') css.push('background:linear-gradient(90deg,color-mix(in srgb,' + (hexColor(rule.color) || '#3369d6') + ' 45%,transparent) 60%,transparent 60%)');
      if (rs.b) css.push('font-weight:600');
      if (rs.i) css.push('font-style:italic');
      const sample = h('span', { class: 'cf-sample' + tc.cls, style: css.join(';') || null }, 'Aa 123');
      list.appendChild(h('div', { class: 'rule' },
        sample,
        h('div', { class: 'rule-text' }, h('b', null, M.CF_TYPES[type] || type), ' ', ruleArgs(rule), h('div', { class: 'muted small' }, rect ? M.rectText(rect) : '(range deleted)')),
        h('button', { class: 'icon-btn', type: 'button', title: 'Move up', 'aria-label': 'Move up', disabled: i === 0 || !A.canEdit(), html: icon('up'), onclick: () => reorder(i, -1) }),
        h('button', { class: 'btn small', type: 'button', disabled: !A.canEdit(), onclick: () => edit(rule) }, 'Edit'),
        h('button', { class: 'btn small', type: 'button', disabled: !A.canEdit(), onclick: () => { A.commit([{ t: 'obj', s: sh.id, kind: 'cf', id: rule.id, val: null }], { label: 'Delete rule' }); setTimeout(renderList, 0); } }, 'Delete')));
    });
    box.appendChild(list);
    if (A.canEdit()) box.appendChild(h('button', { class: 'btn primary', type: 'button', onclick: () => edit(null) }, 'Add rule'));
  };
  const ruleArgs = (rule) => {
    const a = plain(rule.a), b = plain(rule.b);
    if (['between', 'notBetween'].includes(rule.type)) return `${a ?? ''} and ${b ?? ''}`;
    if (rule.type === 'formula') return A.fmlLocal(typeof rule.formula === 'string' ? rule.formula : '');
    if (['top', 'bottom'].includes(rule.type)) return `${a || 10}${rule.pct ? '%' : ''}`;
    if (a !== undefined && a !== '') return String(a);
    return '';
  };
  const reorder = (i, delta) => {
    const rules = objs(sh.cf);
    const [r] = rules.splice(i, 1);
    rules.splice(i + delta, 0, r);
    const ops = objs(sh.cf).map((o) => ({ t: 'obj', s: sh.id, kind: 'cf', id: o.id, val: null })).concat(rules.map((o) => ({ t: 'obj', s: sh.id, kind: 'cf', id: o.id, val: { ...o } })));
    A.commit(ops, { label: 'Reorder rules' });
    setTimeout(renderList, 0);
  };
  const edit = (rule) => {
    const rect = rule ? M.objRect(rule, x) || A.selRect() : A.selRect();
    const range = rangeInput(A, rect);
    const type = select(M.CF_TYPES, rule ? plain(rule.type) : 'gt');
    const a = h('input', { class: 'input', value: rule ? plain(rule.a) ?? '' : '' });
    const b = h('input', { class: 'input', value: rule ? plain(rule.b) ?? '' : '' });
    const formula = h('input', { class: 'input mono', value: rule ? A.fmlLocal(typeof rule.formula === 'string' ? rule.formula : '') : '=' + colName(rect.c1) + (rect.r1 + 1) + '>100', spellcheck: 'false' });
    const pct = checkbox('Percent of items', rule ? !!rule.pct : false);
    const preset = select(Object.fromEntries(Object.entries(CF_PRESETS).map(([k, v]) => [k, v.label]).concat([['custom', 'Custom…']])), rule && rule.style ? (Object.entries(CF_PRESETS).find(([, p]) => p.bg === rule.style.bg && p.fc === rule.style.fc && !!p.b === !!rule.style.b) || ['custom'])[0] : 'red');
    // only #rrggbb colors of a stored format are taken over (they reach the color bars and the saved rule)
    let custom = rule && rule.style ? { ...rule.style, bg: hexColor(rule.style.bg), fc: hexColor(rule.style.fc) } : { bg: '#f4cccc', fc: '#990000' };
    // the picked colors show as a bar under the labels, as on the toolbar (the labels stay readable in both themes)
    const bgBtn = h('button', { class: 'btn small has-cbar', type: 'button', 'data-popup-anchor': '' }, 'Fill', h('span', { class: 'cbar', 'aria-hidden': 'true' }));
    const fcBtn = h('button', { class: 'btn small has-cbar', type: 'button', 'data-popup-anchor': '' }, 'Text', h('span', { class: 'cbar', 'aria-hidden': 'true' }));
    const boldCb = checkbox('Bold', !!custom.b);
    bgBtn.onclick = () => colorPicker(bgBtn.getBoundingClientRect(), custom.bg, (c2) => { custom.bg = c2; paintCustom(); }, 'No fill');
    fcBtn.onclick = () => colorPicker(fcBtn.getBoundingClientRect(), custom.fc, (c2) => { custom.fc = c2; paintCustom(); });
    const paintCustom = () => { paintBar(bgBtn, custom.bg, false); paintBar(fcBtn, custom.fc, true); };
    paintCustom();
    // stored scale and bar colors are used only when they are #rrggbb (they reach the swatch buttons' style)
    const colors = rule && Array.isArray(rule.colors) && rule.colors.every((c) => hexColor(c)) ? rule.colors.slice() : null;
    const scaleRow = h('div', { class: 'row2' });
    const barColor = { v: hexColor(rule && rule.color) || '#3369d6' };
    const update = () => {
      const t = type.value;
      const needsA = ['gt', 'lt', 'ge', 'le', 'eq', 'ne', 'between', 'notBetween', 'contains', 'notContains', 'begins', 'ends', 'top', 'bottom'].includes(t);
      a.parentElement.hidden = !needsA;
      b.parentElement.hidden = !['between', 'notBetween'].includes(t);
      formula.parentElement.hidden = t !== 'formula';
      pct.el.hidden = !['top', 'bottom'].includes(t);
      const visual = ['scale2', 'scale3', 'bar', 'icons'].includes(t);
      styleRow.hidden = visual;
      scaleRow.hidden = !['scale2', 'scale3', 'bar'].includes(t);
      scaleRow.innerHTML = '';
      if (t === 'scale2' || t === 'scale3') {
        // the colors picked so far (kept when the dialog repaints), else the rule's stored ones, else the defaults
        const want = t === 'scale2' ? 2 : 3;
        const cs = (colorsRef.v && colorsRef.v.length === want ? colorsRef.v : colors && colors.length === want ? colors : (t === 'scale2' ? ['#ffffff', '#57bb8a'] : ['#e67c73', '#ffd666', '#57bb8a'])).slice();
        cs.forEach((col, i) => {
          const btn = h('button', { class: 'btn small swatch-btn', type: 'button', style: { background: col, color: contrastText(col) || '' }, 'data-popup-anchor': '' }, ['Min', t === 'scale3' ? 'Mid' : 'Max', 'Max'][i]);
          btn.onclick = () => colorPicker(btn.getBoundingClientRect(), col, (c2) => { cs[i] = c2 || col; colorsRef.v = cs; update(); });
          scaleRow.appendChild(btn);
        });
        colorsRef.v = cs;
      } else if (t === 'bar') {
        const btn = h('button', { class: 'btn small swatch-btn', type: 'button', style: { background: barColor.v, color: contrastText(barColor.v) || '' }, 'data-popup-anchor': '' }, 'Bar color');
        btn.onclick = () => colorPicker(btn.getBoundingClientRect(), barColor.v, (c2) => { barColor.v = c2 || barColor.v; update(); });
        scaleRow.appendChild(btn);
      }
      customRow.hidden = preset.value !== 'custom';
    };
    const colorsRef = { v: colors };
    const customRow = h('div', { class: 'row2' }, bgBtn, fcBtn, boldCb.el);
    const styleRow = h('div', { class: 'form' }, field('Format', preset), customRow);
    preset.addEventListener('change', update);
    type.addEventListener('change', update);
    const body = h('div', { class: 'form' },
      h('div', { class: 'row2' }, field('Apply to range', range), field('Format cells if…', type)),
      h('div', { class: 'row2' }, field('Value', a, 'A number, text, or a formula such as =$B$1'), field('and', b)),
      field('Formula', formula, 'Written for the top-left cell of the range; relative references move with each cell.'),
      pct.el, styleRow, scaleRow);
    dialog(rule ? 'Edit rule' : 'New rule', body, [
      { label: 'Cancel' },
      { label: 'Save', primary: true, action: () => {
        const R = readRange(A, range);
        if (!R) return false;
        if (R.sheet !== sh) { toast('Rules apply to the current sheet.', 'error'); return false; }
        const t = type.value;
        const cfFormula = A.fmlStored(formula.value.startsWith('=') ? formula.value : '=' + formula.value);
        if (t === 'formula') { const err = formulaSyntaxError(cfFormula); if (err) { toast('Formula: ' + err, 'error'); return false; } }
        const p = preset.value === 'custom' ? { bg: custom.bg || null, fc: custom.fc || null, b: boldCb.input.checked ? 1 : null } : { ...CF_PRESETS[preset.value] };
        delete p.label;
        // color scales, data bars and icons have no cell format (a saved format's fill would cover the scale's colors)
        const visual = ['scale2', 'scale3', 'bar', 'icons'].includes(t);
        const val = { ...M.rectIds(sh, R), type: t, a: a.value.trim(), b: b.value.trim(), pct: pct.input.checked, style: visual ? null : p };
        if (t === 'formula') val.formula = cfFormula;
        if (t === 'scale2' || t === 'scale3') val.colors = colorsRef.v;
        if (t === 'bar') val.color = barColor.v;
        A.commit([{ t: 'obj', s: sh.id, kind: 'cf', id: rule ? rule.id : Core.randomId(), val }], { label: 'Conditional formatting' });
        setTimeout(renderList, 0);
      } }]);
    update();
  };
  renderList();
  return d;
}

/* ================================================================ data validation */

export function dvDialog(A, presetType) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const x = A.ix(sh);
  const existing = M.dvFor(sh, A.S.sel.r, A.S.sel.c, x);
  const rule = existing || { type: presetType || 'list', strict: true };
  const rect = existing ? M.objRect(existing, x) : A.selRect();
  const range = rangeInput(A, rect);
  const type = select(M.DV_TYPES, rule.type);
  const source = h('input', { class: 'input', value: rule.source || '', placeholder: 'Open, In progress, Done   or   =$H$2:$H$10' });
  const op = select(M.DV_OPS, rule.op || 'between');
  const a = h('input', { class: 'input', value: rule.a ?? '' });
  const b = h('input', { class: 'input', value: rule.b ?? '' });
  const formula = h('input', { class: 'input mono', value: A.fmlLocal(rule.formula || ''), spellcheck: 'false' });
  const strict = checkbox('Reject invalid values (otherwise show a warning)', rule.strict !== false);
  const arrows = checkbox('Always show the dropdown arrow', rule.arrows === 'always');
  const msg = h('input', { class: 'input', value: rule.msg || '', placeholder: 'Optional' });
  const update = () => {
    const t = type.value;
    source.parentElement.hidden = t !== 'list';
    arrows.el.hidden = t !== 'list';
    op.parentElement.hidden = a.parentElement.hidden = !['number', 'int', 'date', 'length'].includes(t);
    b.parentElement.hidden = !['number', 'int', 'date', 'length'].includes(t) || !['between', 'notBetween'].includes(op.value);
    formula.parentElement.hidden = t !== 'formula';
  };
  type.addEventListener('change', update);
  op.addEventListener('change', update);
  const body = h('div', { class: 'form' },
    h('div', { class: 'row2' }, field('Cells', range), field('Allow', type)),
    field('Items', source, 'Separate items with commas, or refer to a range: =Lists!A2:A20'),
    h('div', { class: 'row2' }, field('Condition', op), field('Value', a), field('and', b)),
    field('Custom formula', formula),
    field('Message for invalid values', msg),
    strict.el, arrows.el);
  const dlg = dialog('Data validation', body, [
    { label: 'Cancel' },
    existing ? { label: 'Remove', danger: true, action: () => { A.commit([{ t: 'obj', s: sh.id, kind: 'dv', id: existing.id, val: null }], { label: 'Data validation' }); } } : null,
    { label: 'Save', primary: true, action: async () => {
      const R = readRange(A, range);
      if (!R) return false;
      if (R.sheet !== sh) { toast('Validation applies to the current sheet.', 'error'); return false; }
      const t = type.value;
      if (t === 'list' && !source.value.trim()) { toast('Enter the list items.', 'error'); return false; }
      const val = { ...M.rectIds(sh, R), type: t, strict: strict.input.checked, msg: msg.value.trim() || null };
      if (t === 'list') { val.source = source.value.trim(); if (arrows.input.checked) val.arrows = 'always'; }
      if (['number', 'int', 'date', 'length'].includes(t)) Object.assign(val, { op: op.value, a: a.value.trim(), b: b.value.trim() });
      if (t === 'formula') val.formula = A.fmlStored(formula.value.trim());
      const ops = [];
      // one rule per cell: rules this one covers are replaced
      for (const o of objs(sh.dv)) {
        if (existing && o.id === existing.id) continue;
        const r2 = M.objRect(o, x);
        if (r2 && r2.r1 >= R.r1 && r2.r2 <= R.r2 && r2.c1 >= R.c1 && r2.c2 <= R.c2) ops.push({ t: 'obj', s: sh.id, kind: 'dv', id: o.id, val: null });
      }
      ops.push({ t: 'obj', s: sh.id, kind: 'dv', id: existing ? existing.id : Core.randomId(), val });
      if (t === 'checkbox') {
        // cells that are TRUE or FALSE keep their content: their values are read once they are computed
        if (!(await valuesReady(A, [R], dlg))) return false;
        for (let r = R.r1; r <= R.r2; r++) for (let c = R.c1; c <= R.c2; c++) {
          const v = A.engine.valueAt(sh, r, c);
          if (v !== true && v !== false) ops.push({ t: 'set', s: sh.id, k: M.keyAt(sh, r, c), v: 'FALSE' });
        }
      }
      A.commit(ops, { label: 'Data validation' });
    } }
  ].filter(Boolean));
  update();
}

/* ================================================================ sort, duplicates, split */

export function sortDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const rect = defaultRegion(A);
  const range = rangeInput(A, rect);
  const header = checkbox('The data has a header row', typeof A.engine.valueAt(sh, rect.r1, rect.c1) === 'string' && typeof A.engine.valueAt(sh, rect.r1 + 1, rect.c1) !== 'string');
  const levels = h('div', { class: 'levels' });
  const colOptions = (R) => {
    const out = {};
    for (let c = R.c1; c <= R.c2; c++) {
      const label = header.input.checked ? formatValue(A.engine.valueAt(sh, R.r1, c), null, A.LOCALE) : '';
      out[c] = 'Column ' + colName(c) + (label ? ' – ' + label : '');
    }
    return out;
  };
  const addLevel = (c, desc) => {
    const R = M.parseRect(range.value) || rect;
    const col = select(colOptions(R), c ?? R.c1);
    const order = select({ asc: 'A → Z', desc: 'Z → A' }, desc ? 'desc' : 'asc');
    const row = h('div', { class: 'row2 level' }, field(levels.children.length ? 'Then by' : 'Sort by', col), field('Order', order),
      levels.children.length ? h('button', { class: 'icon-btn', type: 'button', 'aria-label': 'Remove', html: icon('close'), onclick: () => row.remove() }) : null);
    row._get = () => ({ c: +col.value, desc: order.value === 'desc' });
    levels.appendChild(row);
  };
  addLevel(A.S.sel.c >= rect.c1 && A.S.sel.c <= rect.c2 ? A.S.sel.c : rect.c1, false);
  header.input.addEventListener('change', () => { const keep = [...levels.children].map((r) => r._get()); levels.innerHTML = ''; keep.forEach((k) => addLevel(k.c, k.desc)); });
  const dlg = dialog('Sort range', h('div', { class: 'form' }, field('Range', range), header.el, levels,
    h('button', { class: 'btn small', type: 'button', onclick: () => addLevel() }, 'Add another sort column')), [
    { label: 'Cancel' },
    { label: 'Sort', primary: true, action: async () => {
      const R = readRange(A, range);
      if (!R) return false;
      const keys = [...levels.children].map((r) => r._get()).filter((k) => k.c >= R.c1 && k.c <= R.c2);
      if (!keys.length) { toast('Choose a column inside the range.', 'error'); return false; }
      // the dialog stays open while the sort columns are computed
      return (await sortRange(A, R.sheet, R, keys, header.input.checked, dlg)) ? undefined : false;
    } }]);
}

export function dedupeDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const rect = defaultRegion(A);
  const range = rangeInput(A, rect);
  const header = checkbox('The data has a header row', typeof A.engine.valueAt(sh, rect.r1, rect.c1) === 'string');
  const cols = h('div', { class: 'checks cols' });
  const fillCols = () => {
    cols.innerHTML = '';
    const R = M.parseRect(range.value) || rect;
    for (let c = R.c1; c <= Math.min(R.c2, R.c1 + 40); c++) {
      const label = header.input.checked ? formatValue(A.engine.valueAt(sh, R.r1, c), null, A.LOCALE) : '';
      const cb = checkbox('Column ' + colName(c) + (label ? ' – ' + label : ''), true, { 'data-c': c });
      cols.appendChild(cb.el);
    }
  };
  fillCols();
  range.addEventListener('change', fillCols);
  header.input.addEventListener('change', fillCols);
  const dlg = dialog('Remove duplicates', h('div', { class: 'form' }, field('Range', range), header.el, field('Compare these columns', cols)), [
    { label: 'Cancel' },
    { label: 'Remove duplicates', primary: true, action: async () => {
      const R = readRange(A, range);
      if (!R) return false;
      const chosen = [...cols.querySelectorAll('input:checked')].map((i) => +i.dataset.c);
      if (!chosen.length) { toast('Choose at least one column.', 'error'); return false; }
      const hasHeader = header.input.checked;
      // rows are compared by the values of the chosen columns: the dialog stays open while they are computed
      const top = R.r1 + (hasHeader ? 1 : 0);
      if (!(await valuesReady(A, chosen.map((c) => ({ sheet: R.sheet, r1: top, c1: c, r2: R.r2, c2: c })), dlg))) return false;
      const res = M.dedupeOps(A.engine, R.sheet, R, chosen, hasHeader);
      if (!res) { A.notComputed('pending'); return false; }
      const { ops, removed } = res;
      if (!removed) { toast('No duplicates found.'); return; }
      if (A.commit(ops, { label: 'Remove duplicates' })) toast(removed === 1 ? 'Removed 1 duplicate row.' : 'Removed ' + removed + ' duplicate rows.');
    } }]);
}

export function splitDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const s = A.selRect();
  const delim = select({ auto: 'Detect automatically', ',': 'Comma', ';': 'Semicolon', '\t': 'Tab', ' ': 'Space', custom: 'Custom' }, 'auto');
  const custom = h('input', { class: 'input small', placeholder: 'Separator' });
  const dlg = dialog('Split text to columns', h('div', { class: 'form' }, h('p', { class: 'dlg-text' }, `Splits the text in ${M.rectText({ ...s, c2: s.c1 })} into the columns to the right.`), h('div', { class: 'row2' }, field('Separator', delim), field('Custom separator', custom))), [
    { label: 'Cancel' },
    { label: 'Split', primary: true, action: async () => {
      // the texts are read once they are computed (a value left out is not skipped as if it were no text)
      if (!(await valuesReady(A, [{ sheet: sh, r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c1 }], dlg))) return false;
      const ops = [];
      let maxW = 0;
      const rows = [];
      for (let r = s.r1; r <= s.r2; r++) {
        const v = A.engine.valueAt(sh, r, s.c1);
        if (typeof v !== 'string') { rows.push(null); continue; }
        let d = delim.value === 'custom' ? custom.value : delim.value;
        if (d === 'auto') d = [',', ';', '\t', '|', ' '].find((x) => v.includes(x)) || ',';
        if (!d) { toast('Enter a separator.', 'error'); return false; }
        const parts = v.split(d).map((p) => p.trim());
        rows.push(parts);
        maxW = Math.max(maxW, parts.length);
      }
      ops.push(...M.growOps(sh, s.r2, s.c1 + maxW - 1));
      const colIds = sh.cols.slice();
      ops.forEach((op) => { if (op.t === 'cols') colIds.push(...op.add); });
      rows.forEach((parts, i) => {
        if (!parts) return;
        parts.forEach((p, j) => { const pi = parseInput(p, A.LOCALE); ops.push({ t: 'set', s: sh.id, k: sh.rows[s.r1 + i] + ':' + colIds[s.c1 + j], v: pi.raw || null }); });
      });
      A.commit(ops, { label: 'Split text' });
    } }]);
}

/* ================================================================ find and replace */

/** Download sheet as CSV: the separator between values and the decimal mark of plain numbers. onDownload(opts, isOpen)
 * may be async (the dialog waits for it, and stays open when it gives false); isOpen() tells whether the dialog is
 * still open. */
export function csvDialog(A, opts, onDownload) {
  const delim = select([[',', 'Comma ( , )'], [';', 'Semicolon ( ; )'], ['\t', 'Tab']], opts.delim || ',');
  const dec = select([['.', 'Point: 1.5'], [',', 'Comma: 1,5']], opts.dec || '.');
  const dlg = dialog('Download sheet as CSV', h('div', { class: 'form' },
    field('Separator between values', delim),
    field('Decimal separator in numbers', dec, 'Numbers with a number format are written as they are shown. Values that contain the separator are put in quotes.')), [
    { label: 'Download', primary: true, action: () => onDownload({ delim: delim.value, dec: dec.value }, () => dlg.el.isConnected) },
    { label: 'Cancel' }
  ]);
}

export function findDialog(A, replace) {
  const what = h('input', { class: 'input', autofocus: true, value: A.S.find ? A.S.find.q : '' });
  const repl = h('input', { class: 'input', value: A.S.find ? A.S.find.repl || '' : '' });
  const matchCase = checkbox('Match case', A.S.find ? A.S.find.case : false);
  const whole = checkbox('Match entire cell', A.S.find ? A.S.find.whole : false);
  const formulas = checkbox('Also search in formulas', A.S.find ? A.S.find.formulas : false);
  const scope = select({ sheet: 'This sheet', all: 'All sheets' }, A.S.find ? A.S.find.scope : 'sheet');
  const status = h('div', { class: 'muted small find-status', 'aria-live': 'polite' });
  const opts = () => ({ q: what.value, repl: repl.value, case: matchCase.input.checked, whole: whole.input.checked, formulas: formulas.input.checked, scope: scope.value });
  // values are searched once they are computed (the formula work budget may have left some out). False, with the
  // reason in the status line, when they are not.
  // only cells that exist are searched (matches), and only formulas need computing: not the whole used rectangle
  const formulaCells = (sheets) => {
    const out = [];
    for (const sh of sheets) {
      const x = A.ix(sh);
      for (const k in sh.cells) {
        const v = sh.cells[k] && sh.cells[k].v;
        if (typeof v !== 'string' || v[0] !== '=') continue;
        const [rid, cid] = k.split(':');
        const r = x.rows.get(rid), c = x.cols.get(cid);
        if (r !== undefined && c !== undefined) out.push({ sheet: sh, r1: r, c1: c, r2: r, c2: c });
      }
    }
    return out;
  };
  const searchable = async () => {
    const o = opts();
    if (!o.q) return true;
    const st = await A.computeValues(formulaCells(o.scope === 'all' ? A.view().sheets : [A.sheet()]), { live: () => dlg.el.isConnected });
    if (!dlg.el.isConnected) return false;
    if (st === 'pending' || st === 'changed') {
      status.textContent = st === 'changed' ? 'The spreadsheet changed while its values were being calculated. Try again.' : 'Some values are still being calculated. Try again in a moment.';
      return false;
    }
    return true;
  };
  let unsearched = false; // a value was not searched: its formula needs more calculation than the spreadsheet does at once
  const note = () => (unsearched ? ' Values that are not calculated (#CALC!) were not searched.' : '');
  const matches = () => {
    const o = opts();
    unsearched = false;
    if (!o.q) return [];
    const out = [];
    const sheets = o.scope === 'all' ? A.view().sheets : [A.sheet()];
    const norm = (s) => (o.case ? s : s.toLowerCase());
    const q = norm(o.q);
    for (const sh of sheets) {
      const x = A.ix(sh);
      for (const k in sh.cells) {
        const cell = sh.cells[k];
        const [rid, cid] = k.split(':');
        const r = x.rows.get(rid), c = x.cols.get(cid);
        if (r === undefined || c === undefined) continue;
        const v = A.engine.valueAt(sh, r, c);
        const texts = [];
        if (isSlowCalc(v)) unsearched = true;
        else texts.push(formatValue(v, cell.st, A.LOCALE));
        if (o.formulas && cell.v && cell.v[0] === '=') texts.push(cell.v);
        if (texts.some((t) => (o.whole ? norm(t) === q : norm(t).includes(q)))) out.push({ sheetId: sh.id, r, c, k });
      }
    }
    const order = A.view().sheets.map((s) => s.id);
    return out.sort((a, b) => order.indexOf(a.sheetId) - order.indexOf(b.sheetId) || a.r - b.r || a.c - b.c);
  };
  const next = async () => {
    if (!(await searchable())) return;
    const list = matches();
    A.S.find = { ...opts(), hit: null };
    if (!list.length) { status.textContent = 'No matches.' + note(); A.renderGrid(); return; }
    const cur = { sheetId: A.S.sheetId, r: A.S.sel.r, c: A.S.sel.c };
    const order = A.view().sheets.map((s) => s.id);
    const after = (m) => order.indexOf(m.sheetId) > order.indexOf(cur.sheetId) || (m.sheetId === cur.sheetId && (m.r > cur.r || (m.r === cur.r && m.c > cur.c)));
    const hit = list.find(after) || list[0];
    if (hit.sheetId !== A.S.sheetId) A.goToSheet(hit.sheetId);
    A.selectRect({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c });
    A.scrollIntoView(hit.r, hit.c);
    A.S.find.hit = hit;
    status.textContent = `${list.indexOf(hit) + 1} of ${list.length}` + note();
    A.renderGrid();
  };
  const replaceIn = (cell, o) => {
    const flags = o.case ? 'g' : 'gi';
    const re = new RegExp(o.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    if (o.whole) return null;
    return cell.v.replace(re, o.repl.replace(/\$/g, '$$$$'));
  };
  const replaceOne = () => {
    const o = opts();
    const hit = A.S.find && A.S.find.hit;
    if (!hit || !A.canEdit()) return next();
    const sh = A.view().sheets.find((s) => s.id === hit.sheetId);
    const cell = sh && sh.cells[hit.k];
    if (cell && cell.v !== undefined && (cell.v[0] !== '=' || o.formulas)) {
      const v = o.whole ? o.repl : replaceIn(cell, o);
      const p = cell.v[0] === '=' ? { raw: v } : parseInput(v, A.LOCALE);
      A.commit([{ t: 'set', s: sh.id, k: hit.k, v: p.raw || null }], { label: 'Replace' });
    }
    return next();
  };
  const replaceAll = async () => {
    if (!A.canEdit()) return;
    if (!(await searchable())) return;
    const o = opts();
    const list = matches();
    const ops = [];
    for (const m of list) {
      const sh = A.view().sheets.find((s) => s.id === m.sheetId);
      const cell = sh.cells[m.k];
      if (cell.v === undefined || (cell.v[0] === '=' && !o.formulas)) continue;
      const v = o.whole ? o.repl : replaceIn(cell, o);
      if (v === cell.v) continue;
      const p = cell.v[0] === '=' ? { raw: v } : parseInput(v, A.LOCALE);
      ops.push({ t: 'set', s: sh.id, k: m.k, v: p.raw || null });
    }
    if (A.commit(ops, { label: 'Replace all' })) status.textContent = (ops.length === 1 ? 'Replaced 1 cell.' : 'Replaced ' + ops.length + ' cells.') + note();
  };
  what.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); next(); } });
  const replaceRow = h('div', { class: 'row2' }, field('Replace with', repl));
  replaceRow.hidden = !replace;
  const body = h('div', { class: 'form' }, field('Find', what), replaceRow, h('div', { class: 'checks' }, matchCase.el, whole.el, formulas.el), field('Search', scope), status);
  // the buttons wait for the search (values may be computed first)
  const dlg = dialog(replace ? 'Find and replace' : 'Find', body, [
    replace || !A.canEdit() ? null : { label: 'Replace…', action: () => { replaceRow.hidden = false; repl.focus(); return false; } },
    replace ? { label: 'Replace', action: async () => { await replaceOne(); return false; } } : null,
    replace ? { label: 'Replace all', action: async () => { await replaceAll(); return false; } } : null,
    { label: 'Find next', primary: true, action: async () => { await next(); return false; } }
  ].filter(Boolean), { onClose: () => { if (A.S.find) A.S.find.hit = null; A.renderGrid(); A.focusGrid(); } });
}

/* ================================================================ named ranges */

export function namesDialog(A, defineFromSelection) {
  const box = h('div', { class: 'names' });
  const sh = A.sheet();
  const paint = () => {
    box.innerHTML = '';
    const names = Object.entries(A.view().names || {}).sort((a, b) => a[0].localeCompare(b[0]));
    const table = h('div', { class: 'name-list' });
    if (!names.length) table.appendChild(h('div', { class: 'muted pad' }, 'No named ranges yet. Names make formulas easier to read: =SUM(Sales) instead of =SUM(Sheet1!$B$2:$B$40).'));
    names.forEach(([n, ref]) => table.appendChild(h('div', { class: 'name-row' },
      h('b', { class: 'mono' }, n), h('span', { class: 'mono muted' }, ref),
      // people who can only view get the list without buttons they cannot use
      A.canEdit() ? h('button', { class: 'btn small', type: 'button', onclick: () => editName(n, ref) }, 'Edit') : null,
      A.canEdit() ? h('button', { class: 'btn small', type: 'button', onclick: () => { A.commit([{ t: 'name', name: n, ref: null }], { label: 'Delete name' }); setTimeout(paint, 0); } }, 'Delete') : null)));
    box.appendChild(table);
    if (A.canEdit()) box.appendChild(h('button', { class: 'btn primary', type: 'button', onclick: () => editName(null, '=' + quoteSheet(sh.name) + '!' + M.rectText(A.selRect(), true)) }, 'Add name'));
  };
  const editName = (name, ref) => {
    const n = h('input', { class: 'input mono', value: name || '', autofocus: true, spellcheck: 'false' });
    const r = h('input', { class: 'input mono', value: ref || '', spellcheck: 'false' });
    dialog(name ? 'Edit name' : 'New name', h('div', { class: 'form' }, field('Name', n, 'Letters, digits, "_" and "."; it must not look like a cell address.'), field('Refers to', r, 'A range such as =Sheet1!$A$2:$A$50, a value, or a formula.')), [
      { label: 'Cancel' },
      { label: 'Save', primary: true, action: () => {
        const nn = n.value.trim();
        if (!Core.validName(nn) || /^[A-Za-z]{1,3}\d+$/.test(nn) || /^(TRUE|FALSE)$/i.test(nn)) { toast('That name is not allowed.', 'error'); return false; }
        let rr = r.value.trim();
        if (!rr) { toast('Enter what the name refers to.', 'error'); return false; }
        if (rr[0] !== '=') rr = '=' + rr;
        const err = formulaSyntaxError(rr);
        if (err) { toast(err, 'error'); return false; }
        const ops = [];
        if (name && name.toUpperCase() !== nn.toUpperCase()) ops.push({ t: 'name', name, ref: null });
        ops.push({ t: 'name', name: nn, ref: rr });
        A.commit(ops, { label: 'Name' });
        setTimeout(paint, 0);
      } }]);
  };
  dialog('Named ranges', box, [{ label: 'Close' }], { wide: true });
  paint();
  if (defineFromSelection) editName(null, '=' + quoteSheet(sh.name) + '!' + M.rectText(A.selRect(), true));
}

/* ================================================================ sizes, formats, functions */

export async function sizeDialog(A, axis) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const s = A.selRect();
  const cur = axis === 'c' ? ((sh.colW || {})[sh.cols[s.c1]] || 100) : ((sh.rowH || {})[sh.rows[s.r1]] || 24);
  const v = await promptDialog(axis === 'c' ? 'Column width' : 'Row height', axis === 'c' ? 'Width in pixels' : 'Height in pixels', String(cur), { type: 'number', min: axis === 'c' ? 16 : 8, max: 2000 });
  if (v === null) return;
  const n = Math.round(+v);
  if (!(n >= (axis === 'c' ? 16 : 8) && n <= 2000)) { toast('Enter a number between ' + (axis === 'c' ? 16 : 8) + ' and 2000.', 'error'); return; }
  // the columns (rows) of every selected range
  const patch = {};
  for (const i of A.selLines(axis)) patch[(axis === 'c' ? sh.cols : sh.rows)[i]] = n;
  A.commit([{ t: 'sheet', s: sh.id, [axis === 'c' ? 'colW' : 'rowH']: patch }], { label: 'Size' });
}

const FORMAT_PRESETS = ['0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '#,##0.00 "₺"', '"$"#,##0.00', 'dd.mm.yyyy', 'yyyy-mm-dd', 'dd mmmm yyyy', 'dddd', 'hh:mm', 'hh:mm:ss', 'mmm yy', '0.00E+00', '#,##0;[Red]-#,##0'];
const PATTERN_HELP = '0 digit, # optional digit, , thousands, % percent; d m y h s for dates and times; "text" in quotes';

export function customFormatDialog(A, setFormat) {
  const sh = A.sheet();
  const v = A.engine.valueAt(sh, A.S.sel.r, A.S.sel.c);
  const sample = typeof v === 'number' ? v : 1234.5;
  const st = (M.cellAt(sh, A.S.sel.r, A.S.sel.c) || {}).st || {};
  const pat = h('input', { class: 'input mono', value: st.pat || '#,##0.00', spellcheck: 'false' });
  const preview = h('div', { class: 'preview mono' });
  const presets = FORMAT_PRESETS;
  const list = h('div', { class: 'preset-list' }, presets.map((p) => h('button', { class: 'chip-btn mono', type: 'button', onclick: () => { pat.value = p; upd(); } }, p)));
  const upd = () => { preview.textContent = 'Preview: ' + formatWithPattern(sample, pat.value, A.LOCALE); };
  pat.addEventListener('input', upd);
  upd();
  dialog('Custom number format', h('div', { class: 'form' }, field('Format', pat, PATTERN_HELP), preview, list), [
    { label: 'Cancel' },
    { label: 'Apply', primary: true, action: () => setFormat({ nf: 'custom', pat: pat.value }) }
  ]);
}

/* ---------------------------------------------------------------- format cells */

const NF_CATEGORIES = [['gen', 'Automatic'], ['num', 'Number'], ['cur', 'Currency'], ['acc', 'Accounting'], ['pct', 'Percent'], ['sci', 'Scientific'],
  ['date', 'Date'], ['ldate', 'Long date'], ['time', 'Time'], ['dt', 'Date time'], ['dur', 'Duration'], ['text', 'Plain text'], ['custom', 'Custom']];
const WITH_DECIMALS = ['num', 'cur', 'acc', 'pct', 'sci'];
const MAX_DP = 15; // as stored (sheet-core.js cleanStyle)
const DP_MSG = 'Decimal places must be a whole number from 0 to ' + MAX_DP + '.';
const BORDER_MODES = [['', 'No change'], ['all', 'All borders'], ['outer', 'Outside borders'], ['inner', 'Inside borders'], ['thick', 'Thick outside border'],
  ['top', 'Top border'], ['bottom', 'Bottom border'], ['left', 'Left border'], ['right', 'Right border'], ['none', 'No borders']];
const FONT_SIZES = [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36, 48];

/**
 * Format cells (the context menu, Format > Format cells…): number format, alignment, font, borders and fill of every
 * selected range, as one change. It shows the active cell's format; only what is changed here is applied, so cells
 * with other formats keep the rest of theirs.
 */
export function formatCellsDialog(A) {
  if (!A.canEdit()) return;
  const sh = A.sheet();
  const st = A.curStyle();
  const changed = new Set();
  const on = (el, key, ev = 'change') => { el.addEventListener(ev, () => { changed.add(key); update(); }); return el; };

  // number. Decimals go up to 15 as stored (sheet-core.js cleanStyle; .xlsx imports have them). Automatic with fixed
  // decimals (the toolbar's decimal buttons, imported 0.00 formats) shows them, and an empty field drops them.
  const nf0 = NF_CATEGORIES.some(([v]) => v === st.nf) ? st.nf : 'gen';
  const genDp = nf0 === 'gen' && Number.isInteger(st.dp);
  const category = select(NF_CATEGORIES, nf0, { 'aria-label': 'Category' });
  // a category with decimals starts from its usual number of them when the field is empty
  category.addEventListener('change', () => { if (WITH_DECIMALS.includes(category.value) && dp.value.trim() === '') dp.value = category.value === 'pct' ? '0' : '2'; });
  on(category, 'nf');
  const dp = on(h('input', { class: 'input', type: 'number', min: '0', max: String(MAX_DP), step: '1', value: String(Number.isInteger(st.dp) ? st.dp : nf0 === 'gen' ? '' : nf0 === 'pct' ? 0 : 2), 'aria-label': 'Decimal places' }), 'nf', 'input');
  const symbols = A.currencies();
  if (typeof st.cur === 'string' && st.cur && !symbols.includes(st.cur)) symbols.push(st.cur);
  const symbol = on(select(symbols, st.nf === 'cur' && st.cur ? st.cur : symbols[0], { 'aria-label': 'Symbol' }), 'nf');
  const pat = on(h('input', { class: 'input mono', value: typeof st.pat === 'string' ? st.pat : '#,##0.00', spellcheck: 'false', 'aria-label': 'Format' }), 'nf', 'input');
  const presets = h('div', { class: 'preset-list' }, FORMAT_PRESETS.map((p) => h('button', { class: 'chip-btn mono', type: 'button', onclick: () => { pat.value = p; changed.add('nf'); update(); } }, p)));
  const sample = h('div', { class: 'preview' });
  const v = A.engine.valueAt(sh, A.S.sel.r, A.S.sel.c);
  const decimals = () => Math.round(+dp.value);
  const dpShown = (nf) => WITH_DECIMALS.includes(nf) || (nf === 'gen' && genDp);
  // the decimals field holds a whole number from 0 to MAX_DP (for Automatic it may also be empty)
  const dpOk = (nf) => {
    if (!dpShown(nf)) return true;
    if (dp.validity && dp.validity.badInput) return false; // text that is not a number reads as '' in a number field
    const t = dp.value.trim();
    if (nf === 'gen' && t === '') return true;
    return t !== '' && Number.isInteger(+t) && +t >= 0 && +t <= MAX_DP;
  };
  const numberFormat = () => {
    const nf = category.value;
    if (nf === 'gen') return genDp && dp.value.trim() !== '' ? { nf, dp: decimals() } : {};
    if (nf === 'cur') return { nf, cur: symbol.value, dp: decimals() };
    if (WITH_DECIMALS.includes(nf)) return { nf, dp: decimals() };
    if (nf === 'custom') return { nf, pat: pat.value };
    return { nf };
  };
  const dpRow = field('Decimal places', dp), symbolRow = field('Symbol', symbol), patRow = h('div', { class: 'form' }, field('Format', pat, PATTERN_HELP), presets);

  // alignment
  const ha = on(select([['', 'General'], ['left', 'Left'], ['center', 'Center'], ['right', 'Right']], ['left', 'center', 'right'].includes(st.ha) ? st.ha : '', { 'aria-label': 'Horizontal' }), 'ha');
  const va = on(select([['', 'Bottom'], ['middle', 'Middle'], ['top', 'Top']], ['middle', 'top'].includes(st.va) ? st.va : '', { 'aria-label': 'Vertical' }), 'va');
  const wrap = checkbox('Wrap text', !!st.w);
  on(wrap.input, 'w');

  // font
  const ff = on(select([['', 'Default'], ['serif', 'Serif'], ['mono', 'Monospace']], ['serif', 'mono'].includes(st.ff) ? st.ff : '', { 'aria-label': 'Font' }), 'ff');
  const size0 = Number.isInteger(st.fs) && st.fs >= 6 && st.fs <= 72 ? st.fs : 13;
  const fs = on(select([...new Set([...FONT_SIZES, size0])].sort((a, b) => a - b).map((n) => [String(n), String(n)]), String(size0), { 'aria-label': 'Size' }), 'fs');
  const marks = [['b', 'Bold'], ['i', 'Italic'], ['u', 'Underline'], ['x', 'Strikethrough']].map(([k, label]) => { const cb = checkbox(label, !!st[k]); on(cb.input, k); return [k, cb]; });
  // colors: only #rrggbb ones are taken over (they reach the color bars and the preview's style)
  const colors = { fc: hexColor(st.fc), bg: hexColor(st.bg), bc: null };
  const colorBtn = (key, label, reset) => {
    const b = h('button', { class: 'btn small has-cbar', type: 'button', 'data-popup-anchor': '' }, label, h('span', { class: 'cbar', 'aria-hidden': 'true' }));
    b.onclick = () => colorPicker(b.getBoundingClientRect(), colors[key], (c) => { colors[key] = hexColor(c); changed.add(key); update(); }, reset);
    return b;
  };
  const fcBtn = colorBtn('fc', 'Text color', 'Automatic');
  const fontPreview = h('div', { class: 'fmt-sample' }, 'AaBbCc 123');

  // border and fill
  const border = on(select(BORDER_MODES, '', { 'aria-label': 'Borders' }), 'bd');
  const bcBtn = colorBtn('bc', 'Border color', 'Automatic');
  const bgBtn = colorBtn('bg', 'Fill color', 'No fill');
  const fillPreview = h('div', { class: 'fmt-sample' }, 'AaBbCc 123');

  const panes = {
    number: h('div', { class: 'form fmt-pane' }, field('Category', category), h('div', { class: 'row2' }, dpRow, symbolRow), patRow, sample),
    align: h('div', { class: 'form fmt-pane' }, h('div', { class: 'row2' }, field('Horizontal', ha), field('Vertical', va)), wrap.el),
    font: h('div', { class: 'form fmt-pane' }, h('div', { class: 'row2' }, field('Font', ff), field('Size', fs)), h('div', { class: 'checks cols' }, marks.map(([, cb]) => cb.el)), h('div', { class: 'row2' }, fcBtn), fontPreview),
    border: h('div', { class: 'form fmt-pane' }, field('Borders', border, 'Each selected range gets its own borders.'), h('div', { class: 'row2' }, bcBtn)),
    fill: h('div', { class: 'form fmt-pane' }, h('div', { class: 'row2' }, bgBtn), fillPreview)
  };
  const tabs = h('div', { class: 'opt-tabs', role: 'tablist' });
  const show = (k) => {
    Object.entries(panes).forEach(([key, p]) => { p.hidden = key !== k; });
    tabs.querySelectorAll('button').forEach((b) => { const sel = b.dataset.tab === k; b.classList.toggle('on', sel); b.setAttribute('aria-selected', sel ? 'true' : 'false'); });
  };
  [['number', 'Number'], ['align', 'Alignment'], ['font', 'Font'], ['border', 'Border'], ['fill', 'Fill']].forEach(([k, label]) =>
    tabs.appendChild(h('button', { class: 'opt-tab', type: 'button', role: 'tab', 'data-tab': k, onclick: () => show(k) }, label)));

  function update() {
    const nf = category.value;
    dpRow.hidden = !dpShown(nf);
    symbolRow.hidden = nf !== 'cur';
    patRow.hidden = nf !== 'custom';
    const x = typeof v === 'number' ? v : ['date', 'ldate', 'dt', 'time', 'dur'].includes(nf) ? 46000.5 : 1234.5678;
    sample.textContent = dpOk(nf) ? 'Sample: ' + formatValue(x, numberFormat(), A.LOCALE) : DP_MSG;
    paintBar(fcBtn, colors.fc, true);
    paintBar(bcBtn, colors.bc, true);
    paintBar(bgBtn, colors.bg, false);
    // the preview shows the font and colors picked here, with the colors as the cells show them in the current theme
    // (colors are #rrggbb, fonts and sizes come from the lists)
    const tc = themedColors(colors.bg, colors.fc);
    const css = tc.parts.slice();
    css.push('font-family:' + (ff.value === 'serif' || ff.value === 'mono' ? 'var(--font-' + ff.value + ')' : 'inherit'), 'font-size:' + Math.min(28, +fs.value || 13) + 'px');
    if (marks[0][1].input.checked) css.push('font-weight:600');
    if (marks[1][1].input.checked) css.push('font-style:italic');
    const deco = [marks[2][1].input.checked ? 'underline' : '', marks[3][1].input.checked ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) css.push('text-decoration:' + deco);
    for (const el of [fontPreview, fillPreview]) { el.className = 'fmt-sample' + tc.cls; el.setAttribute('style', css.join(';')); }
  }
  update();
  show('number');
  dialog('Format cells', h('div', { class: 'fmt-cells' }, tabs, ...Object.values(panes)), [
    { label: 'Cancel' },
    { label: 'Apply', primary: true, action: () => {
      const patch = {};
      if (changed.has('nf')) {
        const nf = category.value;
        if (!dpOk(nf)) { show('number'); toast(DP_MSG, 'error'); dp.focus(); return false; }
        if (nf === 'custom' && !pat.value.trim()) { show('number'); toast('Enter a format, such as #,##0.00.', 'error'); pat.focus(); return false; }
        Object.assign(patch, { nf: null, dp: null, cur: null, pat: null }, numberFormat());
      }
      if (changed.has('ha')) patch.ha = ha.value || null;
      if (changed.has('va')) patch.va = va.value || null;
      if (changed.has('w')) patch.w = wrap.input.checked ? 1 : null;
      if (changed.has('ff')) patch.ff = ff.value || null;
      if (changed.has('fs')) patch.fs = +fs.value === 13 ? null : +fs.value;
      for (const [k, cb] of marks) if (changed.has(k)) patch[k] = cb.input.checked ? 1 : null;
      if (changed.has('fc')) patch.fc = colors.fc || null;
      if (changed.has('bg')) patch.bg = colors.bg || null;
      const bd = border.value ? { mode: border.value, color: colors.bc } : null;
      if (Object.keys(patch).length || bd) A.formatCells(patch, bd);
    } }
  ], { class: 'fmt-dialog' });
}

const CATEGORIES = {
  'Most used': ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'IF', 'IFS', 'IFERROR', 'VLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'ROUND', 'CONCAT', 'TEXT', 'TODAY', 'NOW', 'FILTER', 'SORT', 'UNIQUE'],
  Math: ['ABS', 'CEILING', 'FLOOR', 'INT', 'MOD', 'POWER', 'PRODUCT', 'QUOTIENT', 'RAND', 'RANDBETWEEN', 'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'SIGN', 'SQRT', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUBTOTAL', 'AGGREGATE', 'TRUNC', 'SEQUENCE', 'MMULT', 'GCD', 'LCM', 'FACT', 'COMBIN', 'EXP', 'LN', 'LOG', 'LOG10', 'PI', 'SIN', 'COS', 'TAN', 'ROMAN', 'BASE'],
  Statistics: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MAX', 'MAXIFS', 'MIN', 'MINIFS', 'MEDIAN', 'MODE', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'LARGE', 'SMALL', 'RANK.EQ', 'PERCENTILE.INC', 'QUARTILE.INC', 'CORREL', 'SLOPE', 'INTERCEPT', 'FORECAST.LINEAR', 'TREND', 'NORM.DIST', 'NORM.INV', 'BINOM.DIST', 'FREQUENCY', 'GEOMEAN'],
  Logical: ['AND', 'OR', 'NOT', 'XOR', 'IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'CHOOSE', 'LET', 'TRUE', 'FALSE'],
  Text: ['CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'FIND', 'SEARCH', 'REPLACE', 'SUBSTITUTE', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'TEXT', 'VALUE', 'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER', 'REPT', 'EXACT', 'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'CHAR', 'CODE', 'NUMBERVALUE', 'FIXED'],
  'Date and time': ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'ISOWEEKNUM', 'EDATE', 'EOMONTH', 'DATEDIF', 'DAYS', 'DAYS360', 'YEARFRAC', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'WORKDAY', 'WORKDAY.INTL', 'DATEVALUE', 'TIMEVALUE'],
  'Lookup and arrays': ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'XMATCH', 'OFFSET', 'INDIRECT', 'ADDRESS', 'ROW', 'ROWS', 'COLUMN', 'COLUMNS', 'TRANSPOSE', 'UNIQUE', 'SORT', 'SORTBY', 'FILTER', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'HYPERLINK', 'FORMULATEXT'],
  Financial: ['PMT', 'IPMT', 'PPMT', 'FV', 'PV', 'NPV', 'XNPV', 'IRR', 'XIRR', 'MIRR', 'RATE', 'NPER', 'CUMIPMT', 'CUMPRINC', 'SLN', 'SYD', 'DB', 'DDB', 'EFFECT', 'NOMINAL', 'RRI', 'PDURATION'],
  Information: ['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISERR', 'ISNA', 'ISLOGICAL', 'ISFORMULA', 'ISEVEN', 'ISODD', 'NA', 'TYPE', 'ERROR.TYPE', 'CELL', 'SHEET', 'SHEETS', 'N', 'T']
};

export function functionDialog(A, onPick) {
  const search = h('input', { class: 'input', placeholder: 'Search ' + FUNCTION_NAMES.length + ' functions', autofocus: true });
  const cat = select(Object.keys(CATEGORIES).concat(['All']), 'Most used');
  const list = h('div', { class: 'fn-list', role: 'listbox' });
  const help = h('div', { class: 'fn-help mono' });
  let chosen = null;
  const paint = () => {
    list.innerHTML = '';
    const q = search.value.trim().toUpperCase();
    const names = q ? FUNCTION_NAMES.filter((f) => f.includes(q)) : cat.value === 'All' ? FUNCTION_NAMES : CATEGORIES[cat.value].filter((f) => FUNCTION_NAMES.includes(f));
    names.forEach((f, i) => {
      const b = h('button', { class: 'fn-item' + (i === 0 ? ' on' : ''), type: 'button', role: 'option' }, f);
      b.addEventListener('click', () => { list.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); chosen = f; help.textContent = A.fnHelp(f) || f + '(…)'; });
      b.addEventListener('dblclick', () => { document.querySelector('.modal-back .btn.primary').click(); });
      list.appendChild(b);
    });
    chosen = names[0] || null;
    help.textContent = chosen ? A.fnHelp(chosen) || chosen + '(…)' : 'No function matches.';
  };
  search.addEventListener('input', paint);
  cat.addEventListener('change', paint);
  paint();
  dialog('Insert function', h('div', { class: 'form' }, h('div', { class: 'row2' }, field('Search', search), field('Category', cat)), list, help), [
    { label: 'Cancel' },
    { label: 'Insert', primary: true, action: () => { if (chosen) setTimeout(() => onPick(chosen), 0); } }
  ]);
}

/* ================================================================ import choice */

export function importChoice(A, fileName, onChoose) {
  let mode = 'add';
  const add = h('input', { type: 'radio', name: 'imp', checked: true, onchange: () => { mode = 'add'; } });
  const rep = h('input', { type: 'radio', name: 'imp', onchange: () => { mode = 'replace'; } });
  dialog('Import ' + fileName, h('div', { class: 'form' },
    h('label', { class: 'check' }, add, h('span', null, 'Add its sheets to this spreadsheet')),
    h('label', { class: 'check' }, rep, h('span', null, 'Replace all sheets of this spreadsheet (you can undo this)'))), [
    { label: 'Cancel' },
    { label: 'Import', primary: true, action: () => { setTimeout(() => onChoose(mode), 0); } }
  ]);
}

/* ================================================================ history */

export function historyPanel(A) {
  document.querySelectorAll('.history').forEach((x) => x.remove());
  const panel = h('aside', { class: 'history', role: 'complementary', 'aria-label': 'Change history' });
  // right under the bar, whose height depends on the mode
  const bar = document.getElementById('bar');
  if (bar) panel.style.top = bar.offsetHeight + 'px';
  const list = h('div', { class: 'hist-list' });
  let offset = 0, total = 0, canEdit = false;
  const onKey = (e) => {
    if (e.key !== 'Escape' || e.defaultPrevented || !panel.isConnected) return;
    if (document.querySelector('.modal-back, .popover, .menu, .pick-bar') || document.querySelector('.cell-editor:not([hidden])')) return;
    e.preventDefault(); close();
  };
  const close = () => { document.removeEventListener('keydown', onKey); panel.remove(); A.focusGrid(); };
  document.addEventListener('keydown', onKey);
  const clearBtn = h('button', { class: 'btn small', type: 'button', onclick: async () => {
    if (!(await confirmDialog('Clear history', 'Delete the whole change history of this spreadsheet? The spreadsheet itself does not change.', 'Clear history', true))) return;
    try { await A.api('history-delete', { method: 'POST', body: { all: true } }); offset = 0; list.innerHTML = ''; load(); } catch (e) { toast(e.message, 'error'); }
  } }, 'Clear history');
  clearBtn.hidden = true; // shown once the server says that the user can edit
  panel.append(h('div', { class: 'hist-head' }, h('h2', null, 'Change history'), clearBtn,
    h('button', { class: 'icon-btn', type: 'button', 'aria-label': 'Close', html: icon('close'), onclick: close })), list);
  const more = h('button', { class: 'btn small more', type: 'button', onclick: () => load() }, 'Show older changes');
  const load = async () => {
    more.remove();
    const wait = h('div', { class: 'loading pad' }, h('span', { class: 'loader', 'aria-hidden': 'true' }), 'Loading…');
    list.appendChild(wait);
    try {
      const res = await A.api('history', { query: { offset: String(offset), limit: '30' } });
      wait.remove();
      total = res.total; canEdit = res.canEdit;
      clearBtn.hidden = !canEdit;
      if (!res.entries.length && offset === 0) list.appendChild(h('div', { class: 'muted pad' }, 'No changes recorded yet.'));
      res.entries.forEach((e) => list.appendChild(entryEl(e)));
      offset += res.entries.length;
      if (offset < total) list.appendChild(more);
    } catch (e) { wait.remove(); list.appendChild(h('div', { class: 'error pad' }, e.message)); }
  };
  const show = (c) => (c && c.v !== undefined ? c.v : '');
  const entryEl = (e) => {
    const el = h('div', { class: 'hist-item' });
    // display names can be changed by their owner and look alike: the login (unique) is shown on hover, as in the bar
    const byName = typeof e.byName === 'string' && e.byName ? e.byName : String(e.by || '');
    const byLogin = typeof e.by === 'string' && e.by && e.by !== byName ? byName + ' (' + e.by + ')' : null;
    const who = h('div', { class: 'hist-who' }, h('span', { class: 'peer small', style: { background: peerColor(e.by) } }, initials(byName)), h('b', { title: byLogin }, byName), h('span', { class: 'muted' }, relTime(e.at, 'en')));
    el.appendChild(who);
    (Array.isArray(e.events) ? e.events : []).forEach((ev) => el.appendChild(h('div', { class: 'hist-ev' }, ev)));
    const changed = objs(e.cells);
    if (changed.length) {
      const cells = h('div', { class: 'hist-cells' });
      changed.slice(0, 12).forEach((c) => {
        const from = show(c.from), to = show(c.to);
        const styleOnly = from === to;
        cells.appendChild(h('div', { class: 'hist-cell' }, h('span', { class: 'mono addr' }, (c.sn && A.view().sheets.length > 1 ? c.sn + '!' : '') + c.a),
          styleOnly ? h('span', { class: 'muted' }, c.to && c.to.st && !c.from ? 'formatted' : 'format changed')
            : [h('span', { class: 'old' }, from === '' ? '(empty)' : from), ' → ', h('span', { class: 'new' }, to === '' ? '(empty)' : to)]));
      });
      if (e.total > 12) cells.appendChild(h('div', { class: 'muted small' }, '… and ' + (e.total - 12) + ' more cells'));
      el.appendChild(cells);
    }
    const actions = h('div', { class: 'hist-actions' });
    if (canEdit && changed.length && !e.more) {
      actions.appendChild(h('button', { class: 'btn small', type: 'button', title: 'Put back the values these cells had before this change', onclick: () => {
        const view = A.view();
        const ops = [];
        let skipped = 0;
        for (const c of changed) {
          const sh = view.sheets.find((s) => s.id === c.s);
          if (!sh || c.trunc || typeof c.k !== 'string') { skipped++; continue; }
          const [rid, cid] = c.k.split(':');
          if (!sh.rows.includes(rid) || !sh.cols.includes(cid)) { skipped++; continue; }
          ops.push({ t: 'set', s: c.s, k: c.k, v: c.from && c.from.v !== undefined ? c.from.v : null, st: c.from && c.from.st ? c.from.st : null });
        }
        if (A.commit(ops, { label: 'Restore' })) toast(skipped ? `Restored. ${skipped} cell(s) no longer exist.` : 'Restored the earlier values.');
      } }, 'Restore these values'));
    }
    if (canEdit) actions.appendChild(h('button', { class: 'btn small', type: 'button', onclick: async () => {
      try { await A.api('history-delete', { method: 'POST', body: { id: e.id } }); el.remove(); } catch (err) { toast(err.message, 'error'); }
    } }, 'Delete entry'));
    if (actions.children.length) el.appendChild(actions);
    return el;
  };
  document.getElementById('app').appendChild(panel);
  load();
}
