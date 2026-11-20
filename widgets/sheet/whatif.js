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
/* Spreadsheets for YouTrack – What-if analysis: Solver, Goal Seek and Scenario Manager.
 *
 * Solver models and scenarios are saved with their sheet (objects of kind "solver" and "scenarios"), so everyone who
 * works on the issue sees and reuses them. References are stored as formulas ("=$B$10"), so they follow inserted and
 * deleted rows and columns and renamed sheets like any formula.
 * Solving works on a private copy of the workbook: the sheet changes only when the user keeps the solution, in one
 * undoable change together with the reports.
 */
import { h, icon, dialog, confirmDialog, promptDialog, field, select, checkbox, toast } from './ui.js';
import * as M from './model.js';
import { Engine, colName, numToText, isErr, quoteSheet, formulaSyntaxError, formatValue, literal, parseInput, localeInfo, rewriteRefs } from './formula.js';
import * as SV from './solver.js';

const Core = window.SheetCore;
const MODEL_ID = 'model';
const OP_TEXT = { le: '<=', eq: '=', ge: '>=', int: 'int', bin: 'bin', dif: 'dif' };
const OP_WORD = { int: 'integer', bin: 'binary', dif: 'AllDifferent' };
const METHODS = { grg: 'GRG Nonlinear', lp: 'Simplex LP', evo: 'Evolutionary' };
const METHOD_HELP = 'Select the GRG Nonlinear engine for Solver problems that are smooth nonlinear. Select the Simplex LP engine for linear Solver problems, and select the Evolutionary engine for Solver problems that are non-smooth.';
const MAX_VARS = 500;
const MAX_ROWS = 2000;
const BIG = 1e30;

/* ================================================================ small helpers */

const isFormula = (raw) => typeof raw === 'string' && raw[0] === '=';
const strip = (t) => String(t || '').trim().replace(/^=/, '').trim();
/** Numeric value of a cell for Solver: blank = 0, TRUE/FALSE = 1/0, text and errors are not numbers. */
const num = (v) => {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return NaN;
};
/** Stored cell text for a value (numbers in canonical form, text that looks like something else gets a quote). */
function rawOf(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isFinite(v) ? numToText(v) : '#NUM!';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) return v.code;
  const s = String(v);
  if (s === '') return '';
  return s[0] === '=' || s[0] === "'" || literal(s) !== s ? "'" + s : s;
}
/** Removes floating-point noise: 0.30000000000000004 → 0.3, -0 → 0. */
const tidy = (x) => (x === 0 ? 0 : +x.toPrecision(15));
const fmtNum = (v, locale) => (typeof v === 'number' ? formatValue(tidy(v), null, locale) : formatValue(v, null, locale));
/** A number for an input field: every significant digit, the locale's decimal separator, no grouping, so that reading
 * the field back gives exactly the same number (fmtNum rounds to 10 digits for display). */
function editNum(v, locale) {
  if (typeof v !== 'number' || !isFinite(v)) return '';
  const t = numToText(v);
  const dec = localeInfo(locale).dec;
  return dec === '.' ? t : t.replace('.', dec);
}

function addr(sheet, r, c, base) {
  return (base && sheet.id !== base.id ? quoteSheet(sheet.name) + '!' : '') + '$' + colName(c) + '$' + (r + 1);
}
function areaText(a, base) {
  const p = base && a.sheet.id !== base.id ? quoteSheet(a.sheet.name) + '!' : '';
  return p + M.rectText(a, true);
}
/** "A1:B2, 'My sheet'!C3; D4" → parts (commas and semicolons outside quotes separate them). */
function splitList(text) {
  const out = [];
  let cur = '', q = false;
  for (const ch of String(text || '')) {
    if (ch === "'") q = !q;
    if (!q && (ch === ',' || ch === ';')) { if (cur.trim()) out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}
/** A reference or a name → { sheet, r1, c1, r2, c2 } (sheets of `engine`'s workbook) or null. */
function resolveArea(engine, names, text, base) {
  const t = strip(text);
  if (!t) return null;
  let rect = M.parseRect(t);
  if (!rect) {
    const def = (names || {})[t.toUpperCase()];
    if (def) rect = M.parseRect(strip(def));
  }
  if (!rect || !isFinite(rect.r2) || !isFinite(rect.c2)) return null;
  const sheet = rect.sheet ? engine.sheetByName(rect.sheet) : base;
  if (!sheet || rect.r2 >= sheet.rows.length || rect.c2 >= sheet.cols.length) return null;
  return { sheet, r1: rect.r1, c1: rect.c1, r2: rect.r2, c2: rect.c2 };
}
function cellsOf(a) {
  const out = [];
  for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) out.push([r, c]);
  return out;
}
/** Report label of a cell: the nearest text to its left and above it, as Excel does. */
function labelOf(engine, sheet, r, c) {
  const textAt = (rr, cc) => { const v = engine.valueAt(sheet, rr, cc); return typeof v === 'string' && v.trim() ? v.trim() : null; };
  let left = null, up = null;
  for (let cc = c - 1; cc >= 0 && !left; cc--) left = textAt(r, cc);
  for (let rr = r - 1; rr >= 0 && !up; rr--) up = textAt(rr, c);
  return [left, up].filter(Boolean).join(' ');
}
function parseNumber(text, locale) {
  const t = String(text ?? '').trim().replace(/\u2212/g, '-');
  if (!t) return null;
  // the exact form editNum writes: digits, the locale's decimal separator, an optional exponent
  const dec = localeInfo(locale).dec;
  if (dec === '.' || !t.includes('.')) {
    const canon = dec === '.' ? t : t.replace(dec, '.');
    if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(canon)) return +canon;
  }
  const p = parseInput(t, locale);
  if (p.num !== undefined && isFinite(p.num)) return p.num;
  return null;
}
/** Cell references in a reference or formula made absolute ($B$3), as Solver keeps them; names stay as they are. */
function absRefs(text) {
  const t = strip(text);
  if (!t) return '';
  const f = rewriteRefs('=' + t, (ref) => {
    if (ref.kind !== 'rows') { ref.ca1 = true; ref.ca2 = true; }
    if (ref.kind !== 'cols') { ref.ra1 = true; ref.ra2 = true; }
    return ref;
  });
  return strip(f);
}
const sheetById = (A, id) => A.view().sheets.find((s) => s.id === id) || null;

/* ================================================================ private copy of the workbook */

let snapVersion = -1;
/** A deep copy of the workbook with its own engine; changing constants there never touches the shared workbook. */
function snapshot(A) {
  const doc = JSON.parse(JSON.stringify(A.view()));
  const ver = snapVersion--;
  const filtered = new Map();
  let se = null;
  const book = {
    sheets: doc.sheets,
    names: doc.names || {},
    findTable: (name) => M.findTable(se, doc, name, ver, A.LOCALE),
    tableAt: (s, r, c) => M.tableAt(se, s, r, c, ver, A.LOCALE),
    rowHidden(s, r) {
      if (s.hr && s.hr[s.rows[r]]) return 'manual';
      if (!filtered.has(s.id)) {
        const live = sheetById(A, s.id);
        filtered.set(s.id, live ? M.filteredRows(A.engine, live, A.version(), A.LOCALE) : new Set());
      }
      return filtered.get(s.id).has(r) ? 'filter' : null;
    }
  };
  se = new Engine(book, A.LOCALE);
  const sheet = (id) => doc.sheets.find((s) => s.id === id) || null;
  /** Sets constants (key → stored text), recalculates lazily. */
  const setRaw = (s, key, raw) => {
    const cell = { ...(s.cells[key] || {}) };
    if (raw === '' || raw === null) delete cell.v; else cell.v = raw;
    if (cell.v !== undefined || cell.st || cell.n) s.cells[key] = cell; else delete s.cells[key];
  };
  return { doc, se, sheet, setRaw };
}

/* ================================================================ Solver model */

/** Allowed ranges of the numeric options: [min, max, whole number]. */
const OPTION_RULES = {
  precision: [1e-12, 0.999999, false], intTolerance: [0, 100, false], maxTime: [0, 1e7, false], iterations: [0, 1e9, true],
  maxSubproblems: [0, 1e9, true], maxFeasible: [0, 1e9, true], convergence: [1e-12, 0.999999, false], population: [10, 1000, true],
  seed: [0, 2147483647, true], mutation: [0.0001, 0.999, false], maxNoImprove: [1, 1e7, false]
};
function cleanOptions(o) {
  const out = { ...SV.DEFAULT_OPTIONS };
  for (const k in SV.DEFAULT_OPTIONS) {
    const v = o ? o[k] : undefined;
    if (typeof SV.DEFAULT_OPTIONS[k] === 'boolean') { if (typeof v === 'boolean') out[k] = v; continue; }
    if (k === 'derivatives') { out[k] = v === 'central' ? 'central' : 'forward'; continue; }
    const rule = OPTION_RULES[k];
    if (typeof v === 'number' && isFinite(v) && v >= rule[0] && v <= rule[1] && (!rule[2] || Number.isInteger(v))) out[k] = v;
  }
  return out;
}

export function defaultModel() {
  return { obj: '', goal: 'max', value: '0', vars: [], cons: [], nonneg: true, method: 'grg', opts: { ...SV.DEFAULT_OPTIONS } };
}
function normModel(o) {
  const d = defaultModel();
  if (!o) return d;
  return {
    obj: typeof o.obj === 'string' ? o.obj : '',
    goal: ['max', 'min', 'value'].includes(o.goal) ? o.goal : 'max',
    value: typeof o.value === 'string' ? o.value : '0',
    vars: Array.isArray(o.vars) ? o.vars.filter((v) => typeof v === 'string') : [],
    cons: Array.isArray(o.cons) ? o.cons.filter((c) => c && typeof c.l === 'string' && OP_TEXT[c.op]).map((c) => ({ l: c.l, op: c.op, r: typeof c.r === 'string' ? c.r : '' })) : [],
    nonneg: o.nonneg !== false,
    method: METHODS[o.method] ? o.method : 'grg',
    opts: cleanOptions(o.opts)
  };
}
export function modelOf(sheet) {
  const o = (sheet && sheet.solver || []).find((x) => x.id === MODEL_ID);
  return o ? normModel(o) : null;
}
function saveModel(A, sheetId, m) {
  const sh = sheetById(A, sheetId);
  if (!sh || !A.canEdit()) return false;
  const val = normModel(m);
  const cur = modelOf(sh);
  if (cur && JSON.stringify(cur) === JSON.stringify(val)) return true;
  // an empty model is not worth a shared change
  if (!cur && JSON.stringify(val) === JSON.stringify(normModel(defaultModel()))) return true;
  return A.commit([{ t: 'obj', s: sheetId, kind: 'solver', id: MODEL_ID, val }], { label: 'Solver model' });
}
const MODEL_FIELDS = ['obj', 'goal', 'value', 'vars', 'cons', 'nonneg', 'method', 'opts'];
/**
 * The model to save after the dialog: the fields the user changed in the dialog, and for every other field the stored
 * model as it is now (it may have been rewritten meanwhile, e.g. rows inserted by another user moved its references).
 */
function mergeModel(A, sheetId, orig, edited) {
  const cur = modelOf(sheetById(A, sheetId));
  if (!cur) return normModel(edited);
  const o = JSON.parse(orig), e = normModel(edited), out = {};
  for (const f of MODEL_FIELDS) out[f] = JSON.stringify(e[f]) !== JSON.stringify(o[f]) ? e[f] : cur[f];
  return normModel(out);
}
const NUM_RE = /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i;
/** Right side for display and editing: stored numbers ("=2.125") in the user's locale ("2,125" in Turkish), with all
 * their digits; anything else as is. */
function rhsText(r, locale) {
  const t = strip(r);
  return NUM_RE.test(t) ? editNum(+t, locale) : t;
}
function consText(c, locale) {
  return strip(c.l) + ' ' + (OP_WORD[c.op] ? '= ' + OP_WORD[c.op] : OP_TEXT[c.op] + ' ' + rhsText(c.r, locale));
}

class InputError extends Error {}

/**
 * Turns the model into numbers and an evaluate(x) function on a private copy of the workbook.
 * Everything Solver needs for reports is kept in the returned context.
 */
function buildProblem(A, sheetId, model) {
  const fail = (msg) => { throw new InputError(msg); };
  const snap = snapshot(A);
  const { se, doc } = snap;
  const ssh = snap.sheet(sheetId);
  if (!ssh) fail('The sheet of this model no longer exists.');
  const areaOf = (text, what) => {
    const a = resolveArea(se, doc.names, text, ssh);
    if (!a) fail(`${what}: "${strip(text)}" is not a valid reference.`);
    return a;
  };
  const rawAt = (s, r, c) => se.raw(s, r, c);

  // objective
  let obj = null;
  if (strip(model.obj)) {
    const a = areaOf(model.obj, 'Objective');
    if (a.sheet !== ssh || a.r1 !== a.r2 || a.c1 !== a.c2) fail('The Objective must be a single cell on this sheet.');
    if (!isFormula(rawAt(ssh, a.r1, a.c1))) fail('The Objective cell must contain a formula.');
    obj = { r: a.r1, c: a.c1, addr: addr(ssh, a.r1, a.c1), name: labelOf(se, ssh, a.r1, a.c1) };
  }
  let target = null;
  if (model.goal === 'value') {
    if (!obj) fail('Enter the Objective cell to use "Value Of".');
    target = parseNumber(model.value, 'en-US'); // stored in canonical form
    if (target === null) fail('"Value Of" must be a number.');
  }
  const sense = model.goal === 'max' ? -1 : 1;

  // variable cells
  const vars = [], varIndex = new Map();
  for (const text of model.vars) {
    const a = areaOf(text, 'Variable cells');
    if (a.sheet !== ssh) fail('Variable cells must be on this sheet.');
    for (const [r, c] of cellsOf(a)) {
      const k = r + '|' + c;
      if (varIndex.has(k)) continue;
      if (isFormula(rawAt(ssh, r, c))) fail(`Variable cells must contain values, not formulas (${addr(ssh, r, c)}).`);
      const v = se.valueAt(ssh, r, c);
      varIndex.set(k, vars.length);
      vars.push({ r, c, key: M.keyAt(ssh, r, c), addr: addr(ssh, r, c), name: labelOf(se, ssh, r, c), orig: v, start: isFinite(num(v)) ? num(v) : 0 });
      if (vars.length > MAX_VARS) fail(`${SV.MESSAGES.tooLarge} Solver can change at most ${MAX_VARS} variable cells.`);
    }
  }
  if (!vars.length) fail('Enter the variable cells in "By Changing Variable Cells".');
  if (obj && varIndex.has(obj.r + '|' + obj.c)) fail('The Objective cell cannot also be a variable cell.');
  const n = vars.length;

  // constraints
  const implicitLb = new Array(n).fill(-Infinity), implicitUb = new Array(n).fill(Infinity), implicitLower = new Array(n).fill(false);
  const isInt = new Array(n).fill(false), intKind = new Array(n).fill(null);
  const difGroups = [], bounds = [], rows = [], listed = [];
  model.cons.forEach((con, ci) => {
    const L = areaOf(con.l, 'Constraint');
    if (L.sheet !== ssh) fail('Constraint cells must be on this sheet.');
    const cells = cellsOf(L);
    if (OP_WORD[con.op]) {
      const js = cells.map(([r, c]) => varIndex.get(r + '|' + c));
      if (js.some((j) => j === undefined)) fail(`"${strip(con.l)}": integer, binary and AllDifferent constraints can only be set on variable cells.`);
      if (con.op === 'dif') {
        difGroups.push(js);
        js.forEach((j) => { isInt[j] = true; intKind[j] = 'AllDiff'; implicitLb[j] = Math.max(implicitLb[j], 1); implicitUb[j] = Math.min(implicitUb[j], js.length); implicitLower[j] = true; });
        listed.push({ kind: 'dif', con: ci, addr: areaText(L), name: labelOf(se, ssh, L.r1, L.c1), js, formula: areaText(L) + '=AllDifferent' });
      } else {
        js.forEach((j, i) => {
          isInt[j] = true;
          if (con.op === 'bin') { intKind[j] = 'Binary'; implicitLb[j] = Math.max(implicitLb[j], 0); implicitUb[j] = Math.min(implicitUb[j], 1); implicitLower[j] = true; }
          else if (!intKind[j]) intKind[j] = 'Integer';
          listed.push({ kind: con.op, con: ci, addr: vars[j].addr, name: vars[j].name, j, formula: vars[j].addr + '=' + (con.op === 'bin' ? 'binary' : 'Integer') });
          void i;
        });
      }
      return;
    }
    const rtext = strip(con.r);
    if (!rtext) fail(`"${strip(con.l)}": the constraint has no right side.`);
    const isNum = NUM_RE.test(rtext);
    let R = null;
    if (!isNum) R = resolveArea(se, doc.names, rtext, ssh);
    let rhsAt;
    if (R) {
      const rc = cellsOf(R);
      if (rc.length !== 1 && rc.length !== cells.length) fail(`"${consText(con, A.LOCALE)}": the right side must be a number, a formula, a single cell or a range of the same size as the left side.`);
      rhsAt = (i) => {
        const [r, c] = rc.length === 1 ? rc[0] : rc[i];
        return { get: () => num(se.valueAt(R.sheet, r, c)), cell: [R.sheet, r, c], text: addr(R.sheet, r, c, ssh) };
      };
    } else {
      const f = '=' + rtext;
      const err = formulaSyntaxError(f);
      if (err) fail(`"${consText(con, A.LOCALE)}": ${err}`);
      rhsAt = () => ({ get: () => num(se.evalAt(f, ssh, 0, 0)), cell: null, text: rtext });
    }
    cells.forEach(([r, c], i) => {
      const rh = rhsAt(i);
      const j = varIndex.get(r + '|' + c);
      const lhsGet = () => num(se.valueAt(ssh, r, c));
      const entry = { kind: 'row', con: ci, rel: con.op, addr: addr(ssh, r, c), name: labelOf(se, ssh, r, c), formula: addr(ssh, r, c) + OP_TEXT[con.op] + rh.text, lhsGet, rhsGet: rh.get, r, c };
      // a bound: a variable cell compared with a number or with a constant cell that is not a variable
      let constVal = null;
      if (j !== undefined) {
        if (isNum) constVal = rh.get();
        else if (rh.cell) {
          const [s2, r2, c2] = rh.cell;
          if (!isFormula(rawAt(s2, r2, c2)) && !(s2 === ssh && varIndex.has(r2 + '|' + c2))) constVal = rh.get();
        }
      }
      if (constVal !== null && isFinite(constVal)) {
        entry.kind = 'bound';
        entry.j = j;
        entry.value = constVal;
        bounds.push(entry);
      } else {
        entry.row = rows.length;
        rows.push(entry);
      }
      listed.push(entry);
    });
  });
  if (rows.length > MAX_ROWS) fail(`${SV.MESSAGES.tooLarge} At most ${MAX_ROWS} constraint cells are supported.`);

  /** Bounds of the variables from implicit bounds (binary, AllDifferent), the given bound constraints and the
   * non-negativity option. */
  const boundsFor = (active) => {
    const lb = implicitLb.slice(), ub = implicitUb.slice(), lower = implicitLower.slice();
    for (const b of active) {
      if (b.rel !== 'ge') ub[b.j] = Math.min(ub[b.j], b.value);
      if (b.rel !== 'le') { lb[b.j] = Math.max(lb[b.j], b.value); lower[b.j] = true; }
    }
    if (model.nonneg) for (let j = 0; j < n; j++) if (!lower[j]) lb[j] = Math.max(lb[j], 0);
    return { lb, ub };
  };
  const { lb, ub } = boundsFor(bounds);

  const baseCells = vars.map((v) => ({ ...(ssh.cells[v.key] || {}) }));
  const setX = (x) => {
    for (let j = 0; j < n; j++) {
      const cell = { ...baseCells[j], v: numToText(tidy(x[j])) };
      ssh.cells[vars[j].key] = cell;
    }
    se.invalidateValues();
  };
  const ctype = rows.map((r) => r.rel);
  if (target !== null) ctype.push('eq');
  const evaluate = (x) => {
    setX(x);
    const objVal = obj ? num(se.valueAt(ssh, obj.r, obj.c)) : 0;
    const lhs = [], rhs = [], g = [];
    for (const row of rows) {
      const l = row.lhsGet(), rv = row.rhsGet();
      lhs.push(l);
      rhs.push(rv);
      g.push(l - rv);
    }
    if (target !== null) { lhs.push(objVal); rhs.push(target); g.push(objVal - target); }
    return { f: target !== null ? 0 : objVal, g, lhs, rhs, objCell: objVal };
  };
  const x0 = vars.map((v) => v.start);
  const P = { n, x0, lb, ub, isInt, difGroups, sense, ctype, evaluate };
  const original = evaluate(x0);
  return {
    P, snap, ssh, sheetId, model, obj, target, vars, rows, bounds, listed, intKind, boundsFor, original,
    hasInt: isInt.some(Boolean)
  };
}

/* ================================================================ dialogs: references and picking */

let picking = null;

/** An input for a reference with a button that lets the user select the cells on the sheet. */
function refField(A, value, opts = {}) {
  const input = h('input', { class: 'input mono', value: value || '', spellcheck: 'false', 'aria-label': opts.label || 'Reference', autocomplete: 'off' });
  const btn = h('button', {
    class: 'icon-btn pick-btn', type: 'button', title: 'Select cells on the sheet', 'aria-label': 'Select cells on the sheet',
    html: icon('pick'), onclick: () => pickRange(A, input, opts)
  });
  return { el: h('div', { class: 'ref-field' }, input, btn), input, btn };
}

/** Hides the open dialogs while the user selects cells; the selection is written into `input`. */
function pickRange(A, input, opts = {}) {
  if (picking) return;
  const backs = [...document.querySelectorAll('.modal-back')];
  backs.forEach((b) => { b.dataset.prevDisplay = b.style.display || ''; b.style.display = 'none'; });
  const startSheet = A.sheet().id;
  const baseId = opts.baseId || startSheet;
  const selText = () => {
    const sh = A.sheet();
    return (sh.id !== baseId ? quoteSheet(sh.name) + '!' : '') + M.rectText(A.selRect(), true);
  };
  const cur = h('input', { class: 'input mono', value: input.value, spellcheck: 'false', 'aria-label': 'Selected cells' });
  let last = selText();
  const finish = (ok) => {
    clearInterval(timer);
    document.removeEventListener('keydown', onKey, true);
    document.removeEventListener('mouseup', onUp, true);
    bar.remove();
    backs.forEach((b) => { b.style.display = b.dataset.prevDisplay; delete b.dataset.prevDisplay; });
    picking = null;
    if (A.sheet().id !== startSheet) A.goToSheet(startSheet);
    if (ok) {
      const v = cur.value.trim();
      input.value = opts.append && input.value.trim() && v && !splitList(input.value).includes(v) ? input.value.trim() + ', ' + v : v;
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
    input.focus();
  };
  const onKey = (e) => {
    if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); finish(true); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finish(false); }
  };
  // a click on the sheet takes the selection, even when it is the cell that was already selected
  const onUp = (e) => {
    if (bar.contains(e.target)) return;
    setTimeout(() => { if (picking) { last = selText(); cur.value = last; } }, 0);
  };
  const bar = h('div', { class: 'pick-bar', role: 'dialog', 'aria-label': opts.title || 'Select cells' },
    h('span', { class: 'pick-title' }, opts.title || 'Select cells on the sheet'),
    cur,
    h('button', { class: 'btn', type: 'button', onclick: () => finish(false) }, 'Cancel'),
    h('button', { class: 'btn primary', type: 'button', onclick: () => finish(true) }, 'Done'));
  document.body.appendChild(bar);
  document.addEventListener('keydown', onKey, true);
  document.addEventListener('mouseup', onUp, true);
  const timer = setInterval(() => {
    const t = selText();
    if (t !== last) { last = t; if (document.activeElement !== cur) cur.value = t; }
  }, 100);
  picking = { finish };
  A.focusGrid();
}

function radio(name, value, label, checked) {
  const input = h('input', { type: 'radio', name, value, checked });
  return { el: h('label', { class: 'check' }, input, h('span', null, label)), input };
}

/* ================================================================ Solver Parameters */

export function solverDialog(A, keep) {
  if (!A.canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  const sh = A.sheet();
  const sheetId = sh.id;
  const m = keep ? normModel(keep) : modelOf(sh) || defaultModel();
  if (!keep && !modelOf(sh)) {
    // like Excel: the active cell becomes the objective when it holds a formula
    const s = A.selRect();
    if (isFormula(A.engine.raw(sh, s.r1, s.c1))) m.obj = '=' + addr(sh, s.r1, s.c1);
  }
  const objF = refField(A, strip(m.obj), { title: 'Set objective: select one cell', label: 'Set objective' });
  const gname = 'goal-' + Core.randomId();
  const gMax = radio(gname, 'max', 'Max', m.goal === 'max'), gMin = radio(gname, 'min', 'Min', m.goal === 'min'), gVal = radio(gname, 'value', 'Value of:', m.goal === 'value');
  const valueIn = h('input', { class: 'input value-of', value: editNum(m.goal === 'value' ? parseNumber(m.value, 'en-US') ?? 0 : 0, A.LOCALE), 'aria-label': 'Value of' });
  // typing a value (not just tabbing through) chooses Value of
  valueIn.addEventListener('input', () => { gVal.input.checked = true; });
  valueIn.addEventListener('mousedown', () => { gVal.input.checked = true; });
  const varsF = refField(A, m.vars.map(strip).join(', '), { title: 'By changing variable cells', label: 'By changing variable cells', append: true });
  const list = h('div', { class: 'rule-list solver-cons', role: 'listbox', 'aria-label': 'Constraints', tabindex: '0' });
  let selected = -1;
  const nonneg = checkbox('Make unconstrained variables non-negative', m.nonneg);
  const methodSel = select(METHODS, m.method, { 'aria-label': 'Solving method' });

  const paint = () => {
    list.innerHTML = '';
    if (!m.cons.length) list.appendChild(h('div', { class: 'muted pad' }, 'No constraints.'));
    m.cons.forEach((c, i) => {
      const row = h('div', { class: 'rule cons-row' + (i === selected ? ' on' : ''), role: 'option', 'aria-selected': i === selected ? 'true' : 'false' },
        h('span', { class: 'mono rule-text' }, consText(c, A.LOCALE)));
      row.addEventListener('click', () => { selected = i; paint(); });
      row.addEventListener('dblclick', () => { selected = i; changeCons(); });
      list.appendChild(row);
    });
    btnChange.disabled = btnDelete.disabled = selected < 0;
  };
  const addCons = () => constraintDialog(A, sheetId, null, (c) => { m.cons.push(c); selected = m.cons.length - 1; paint(); });
  const changeCons = () => { if (selected >= 0) constraintDialog(A, sheetId, m.cons[selected], (c) => { m.cons[selected] = c; paint(); }); };
  const delCons = () => { if (selected >= 0) { m.cons.splice(selected, 1); selected = Math.min(selected, m.cons.length - 1); paint(); } };
  const btnAdd = h('button', { class: 'btn', type: 'button', onclick: addCons }, 'Add');
  const btnChange = h('button', { class: 'btn', type: 'button', onclick: changeCons }, 'Change');
  const btnDelete = h('button', { class: 'btn', type: 'button', onclick: delCons }, 'Delete');
  const btnReset = h('button', { class: 'btn', type: 'button', onclick: async () => {
    if (!(await confirmDialog('Reset all', 'Reset all Solver options and cell selections?', 'Reset'))) return;
    Object.assign(m, defaultModel());
    objF.input.value = ''; varsF.input.value = ''; gMax.input.checked = true; valueIn.value = '0';
    nonneg.input.checked = true; methodSel.value = 'grg'; selected = -1; paint();
  } }, 'Reset all');
  const btnLoad = h('button', { class: 'btn', type: 'button', onclick: () => { collect(); loadSaveDialog(A, sheetId, m, (loaded) => { Object.assign(m, loaded); fill(); paint(); }); } }, 'Load/save');
  list.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); selected = Math.max(0, Math.min(m.cons.length - 1, selected + (e.key === 'ArrowDown' ? 1 : -1))); paint(); }
    else if (e.key === 'Delete') { e.preventDefault(); delCons(); }
  });

  const fill = () => {
    objF.input.value = strip(m.obj);
    varsF.input.value = m.vars.map(strip).join(', ');
    ({ max: gMax, min: gMin, value: gVal })[m.goal].input.checked = true;
    valueIn.value = editNum(parseNumber(m.value, 'en-US') ?? 0, A.LOCALE);
    nonneg.input.checked = m.nonneg;
    methodSel.value = m.method;
  };
  const collect = () => {
    const o = absRefs(objF.input.value);
    m.obj = o ? '=' + o : '';
    m.goal = gMax.input.checked ? 'max' : gMin.input.checked ? 'min' : 'value';
    const v = parseNumber(valueIn.value, A.LOCALE);
    m.value = v === null ? valueIn.value.trim() : numToText(v);
    m.vars = splitList(varsF.input.value).map((t) => '=' + absRefs(t));
    m.nonneg = nonneg.input.checked;
    m.method = methodSel.value;
    return m;
  };
  /** Only checks that references can be read; the full check runs when solving. */
  const checkRefs = () => {
    const bad = [m.obj, ...m.vars].filter((t) => strip(t) && !resolveArea(A.engine, A.view().names, t, sheetById(A, sheetId)));
    if (bad.length) { toast(`"${strip(bad[0])}" is not a valid reference.`, 'error'); return false; }
    return true;
  };

  let done = false;
  // what the dialog showed when it opened: only fields the user changes are written back
  const orig = JSON.stringify(normModel(collect()));
  const changed = () => JSON.stringify(normModel(collect())) !== orig;
  const commitModel = () => { if (changed()) saveModel(A, sheetId, mergeModel(A, sheetId, orig, m)); };
  const body = h('div', { class: 'form solver' },
    field('Set objective', objF.el),
    h('div', { class: 'goal-row', role: 'radiogroup', 'aria-label': 'To' }, h('span', { class: 'goal-label' }, 'To:'), gMax.el, gMin.el, gVal.el, valueIn),
    field('By changing variable cells', varsF.el, 'Several ranges are separated by commas.'),
    h('div', { class: 'fld' }, h('span', null, 'Subject to the constraints'),
      h('div', { class: 'cons-wrap' }, list, h('div', { class: 'cons-btns' }, btnAdd, btnChange, btnDelete, btnReset, btnLoad))),
    nonneg.el,
    h('div', { class: 'row2 method-row' }, field('Select a solving method', methodSel),
      h('button', { class: 'btn', type: 'button', onclick: () => optionsDialog(A, m.opts, (o) => { m.opts = o; }) }, 'Options')),
    h('div', { class: 'method-help' }, h('b', null, 'Solving method'), h('p', null, METHOD_HELP)));
  dialog('Solver parameters', body, [
    { label: 'Close', action: () => { collect(); if (!checkRefs()) return false; done = true; commitModel(); } },
    { label: 'Solve', primary: true, action: () => {
      collect();
      if (!checkRefs()) return false;
      done = true;
      commitModel();
      const model = changed() ? mergeModel(A, sheetId, orig, m) : modelOf(sheetById(A, sheetId)) || normModel(m);
      setTimeout(() => runSolver(A, sheetId, model), 0);
    } }
  ], {
    wide: true, sticky: true, class: 'solver-dlg',
    // X or Esc keeps the edits too (references are checked when solving)
    onClose: () => { if (!done) commitModel(); }
  });
  paint();
}

function constraintDialog(A, sheetId, con, onSave) {
  const lf = refField(A, con ? strip(con.l) : '', { title: 'Cell reference', label: 'Cell reference', baseId: sheetId });
  const opSel = select(Object.entries(OP_TEXT), con ? con.op : 'le', { 'aria-label': 'Relation', class: 'input op-sel' });
  const rf = refField(A, con && !OP_WORD[con.op] ? rhsText(con.r, A.LOCALE) : '', { title: 'Constraint', label: 'Constraint', baseId: sheetId });
  const sync = () => {
    const word = OP_WORD[opSel.value];
    rf.input.disabled = rf.btn.disabled = !!word;
    if (word) rf.input.value = word;
    else if (Object.values(OP_WORD).includes(rf.input.value)) rf.input.value = '';
  };
  opSel.addEventListener('change', sync);
  sync();
  const read = () => {
    const sh = sheetById(A, sheetId);
    const l = strip(lf.input.value);
    const a = l && resolveArea(A.engine, A.view().names, l, sh);
    if (!a) { toast('Cell reference is empty or not valid.', 'error'); lf.input.focus(); return null; }
    if (a.sheet.id !== sheetId) { toast('Constraint cells must be on this sheet.', 'error'); return null; }
    const op = opSel.value;
    const la = absRefs(l);
    if (OP_WORD[op]) return { l: '=' + la, op, r: '' };
    const t = strip(rf.input.value);
    if (!t) { toast('Constraint must be a number, a cell reference or a formula.', 'error'); rf.input.focus(); return null; }
    const v = parseNumber(t, A.LOCALE);
    if (v !== null && !resolveArea(A.engine, A.view().names, t, sh)) return { l: '=' + la, op, r: '=' + numToText(v) };
    const err = formulaSyntaxError('=' + t);
    if (err) { toast(err, 'error'); return null; }
    return { l: '=' + la, op, r: '=' + absRefs(t) };
  };
  dialog(con ? 'Change constraint' : 'Add constraint', h('div', { class: 'cons-dlg' },
    field('Cell reference', lf.el), field('Relation', opSel), field('Constraint', rf.el)), [
    // the primary button comes first: OK, Add, Cancel
    { label: 'Add', action: () => {
      const c = read();
      if (!c) return false;
      onSave(c);
      if (con) return; // changing: Add saves this one and closes, like OK
      lf.input.value = ''; rf.input.value = ''; opSel.value = 'le'; sync(); lf.input.focus();
      return false;
    } },
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: () => { const c = read(); if (!c) return false; onSave(c); } }
  ], { sticky: true, class: 'cons-dialog' });
}

/* ================================================================ Solver Options */

function optionsDialog(A, current, onSave) {
  const o = { ...SV.DEFAULT_OPTIONS, ...current };
  const L = A.LOCALE;
  const shared = {}; // key → inputs on several tabs that stay in sync
  const numIn = (key, label, help, blank) => {
    const v = o[key];
    const input = h('input', { class: 'input num-in', value: blank && !v ? '' : editNum(v, L), 'data-key': key, inputmode: 'decimal', 'aria-label': label });
    (shared[key] = shared[key] || []).push(input);
    input.addEventListener('input', () => shared[key].forEach((x) => { if (x !== input) x.value = input.value; }));
    return field(label, input, help);
  };
  const chk = (key, label) => {
    const c = checkbox(label, !!o[key]);
    c.input.dataset.key = key;
    (shared[key] = shared[key] || []).push(c.input);
    c.input.addEventListener('change', () => shared[key].forEach((x) => { if (x !== c.input) x.checked = c.input.checked; }));
    return c.el;
  };
  const dname = 'drv-' + Core.randomId();
  const fwd = radio(dname, 'forward', 'Forward', o.derivatives !== 'central'), cen = radio(dname, 'central', 'Central', o.derivatives === 'central');
  const group = (title, ...kids) => h('fieldset', { class: 'opt-group' }, h('legend', null, title), ...kids);
  const panes = {
    all: h('div', { class: 'form' },
      numIn('precision', 'Constraint precision'),
      h('div', { class: 'checks cols' }, chk('scaling', 'Use automatic scaling'), chk('showIter', 'Show iteration results')),
      group('Solving with integer constraints', chk('ignoreInt', 'Ignore integer constraints'), numIn('intTolerance', 'Integer optimality (%)')),
      group('Solving limits', h('div', { class: 'row2' }, numIn('maxTime', 'Max time (seconds)', 'Empty: no limit', true), numIn('iterations', 'Iterations', 'Empty: no limit', true))),
      group('Evolutionary and integer constraints', h('div', { class: 'row2' }, numIn('maxSubproblems', 'Max subproblems', 'Empty: no limit', true), numIn('maxFeasible', 'Max feasible solutions', 'Empty: no limit', true)))),
    grg: h('div', { class: 'form' },
      numIn('convergence', 'Convergence'),
      group('Derivatives', h('div', { class: 'checks' }, fwd.el, cen.el)),
      group('Multistart', chk('multistart', 'Use multistart'), h('div', { class: 'row2' }, numIn('population', 'Population size'), numIn('seed', 'Random seed', '0: a different seed each time')), chk('requireBounds', 'Require bounds on variables'))),
    evo: h('div', { class: 'form' },
      h('div', { class: 'row2' }, numIn('convergence', 'Convergence'), numIn('mutation', 'Mutation rate')),
      h('div', { class: 'row2' }, numIn('population', 'Population size'), numIn('seed', 'Random seed', '0: a different seed each time')),
      numIn('maxNoImprove', 'Maximum time without improvement (seconds)'),
      chk('requireBounds', 'Require bounds on variables'))
  };
  const tabs = h('div', { class: 'opt-tabs', role: 'tablist' });
  const show = (k) => {
    Object.entries(panes).forEach(([key, p]) => { p.hidden = key !== k; });
    tabs.querySelectorAll('button').forEach((b) => { const on = b.dataset.tab === k; b.classList.toggle('on', on); b.setAttribute('aria-selected', on ? 'true' : 'false'); });
  };
  [['all', 'All methods'], ['grg', 'GRG Nonlinear'], ['evo', 'Evolutionary']].forEach(([k, label]) =>
    tabs.appendChild(h('button', { class: 'opt-tab', type: 'button', role: 'tab', 'data-tab': k, onclick: () => show(k) }, label)));
  show('all');
  const RULES = OPTION_RULES;
  const NAMES = { precision: 'Constraint precision', intTolerance: 'Integer optimality', maxTime: 'Max time', iterations: 'Iterations', maxSubproblems: 'Max subproblems', maxFeasible: 'Max feasible solutions', convergence: 'Convergence', population: 'Population size', seed: 'Random seed', mutation: 'Mutation rate', maxNoImprove: 'Maximum time without improvement' };
  dialog('Options', h('div', { class: 'solver-opts' }, tabs, panes.all, panes.grg, panes.evo), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: () => {
      const out = { ...o };
      for (const [key, inputs] of Object.entries(shared)) {
        const el = inputs[0];
        if (el.type === 'checkbox') { out[key] = el.checked; continue; }
        const [min, max, int] = RULES[key];
        const t = el.value.trim();
        const v = t === '' ? 0 : parseNumber(t, L);
        if (v === null || v < min || v > max || (int && !Number.isInteger(v)) || (t === '' && min > 0)) {
          toast(`${NAMES[key]} must be ${int ? 'a whole number' : 'a number'} between ${fmtNum(min, L)} and ${fmtNum(max, L)}.`, 'error');
          el.focus();
          return false;
        }
        out[key] = v;
      }
      out.derivatives = cen.input.checked ? 'central' : 'forward';
      onSave(out);
    } }
  ], { sticky: true, class: 'opts-dialog' });
}

/* ================================================================ Load / Save model */

function modelCells(m) {
  const o = strip(m.obj);
  const v = parseNumber(m.value, 'en-US') ?? 0;
  const cells = [!o ? '=0' : m.goal === 'max' ? `=MAX(${o})` : m.goal === 'min' ? `=MIN(${o})` : `=${o}=${numToText(v)}`];
  cells.push(`=COUNT(${m.vars.map(strip).join(',') || '0'})`);
  for (const c of m.cons) {
    const l = strip(c.l);
    if (c.op === 'int') cells.push(`=AND(INT(${l})=${l})`);
    else if (c.op === 'bin') cells.push(`=AND((${l}=0)+(${l}=1))`);
    else if (c.op === 'dif') cells.push(`=AND(SORT(TOCOL(${l}))=SEQUENCE(COUNT(${l})))`);
    else cells.push(`=AND(${l}${OP_TEXT[c.op]}${strip(c.r)})`);
  }
  const op = m.opts;
  cells.push('Solver options: ' + [
    'method=' + m.method, 'nonneg=' + (m.nonneg ? 1 : 0),
    ...Object.keys(SV.DEFAULT_OPTIONS).map((k) => k + '=' + (typeof op[k] === 'boolean' ? (op[k] ? 1 : 0) : op[k]))
  ].join('; '));
  return cells;
}
/** Splits "L<=R" at the first comparison outside parentheses and quotes. */
function splitComparison(t) {
  let depth = 0, q = null;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) { if (ch === q) q = null; continue; }
    if (ch === '"' || ch === "'") { q = ch; continue; }
    if (ch === '(' || ch === '{') depth++;
    else if (ch === ')' || ch === '}') depth--;
    else if (depth === 0 && (ch === '<' || ch === '>') && t[i + 1] === '=') return [t.slice(0, i), ch === '<' ? 'le' : 'ge', t.slice(i + 2)];
    else if (depth === 0 && ch === '=' && i > 0) return [t.slice(0, i), 'eq', t.slice(i + 1)];
  }
  return null;
}
function parseModelCells(texts) {
  const m = defaultModel();
  const fail = (msg) => { throw new InputError(msg); };
  if (texts.length < 2) fail('The selected cells do not contain a Solver model.');
  let t = texts[0], mm;
  if ((mm = /^=MAX\((.+)\)$/i.exec(t))) { m.obj = '=' + mm[1]; m.goal = 'max'; }
  else if ((mm = /^=MIN\((.+)\)$/i.exec(t))) { m.obj = '=' + mm[1]; m.goal = 'min'; }
  else if (/^=0$/.test(t)) m.obj = '';
  else {
    const p = t[0] === '=' && splitComparison(t.slice(1));
    if (!p || p[1] !== 'eq') fail('The first cell of a Solver model holds the objective, such as =MAX($B$10).');
    m.obj = '=' + p[0]; m.goal = 'value'; m.value = p[2];
  }
  if (!(mm = /^=COUNT\((.*)\)$/i.exec(texts[1]))) fail('The second cell of a Solver model holds the variable cells, such as =COUNT($B$2:$B$5).');
  m.vars = splitList(mm[1]).filter((x) => x !== '0').map((x) => '=' + x);
  for (const raw of texts.slice(2)) {
    t = String(raw || '');
    if (/^Solver options:/i.test(t)) {
      t.replace(/^Solver options:/i, '').split(';').forEach((kv) => {
        const [k, v] = kv.split('=').map((s) => s.trim());
        if (k === 'method' && METHODS[v]) m.method = v;
        else if (k === 'nonneg') m.nonneg = v !== '0';
        else if (k === 'derivatives') m.opts.derivatives = v === 'central' ? 'central' : 'forward';
        else if (k in SV.DEFAULT_OPTIONS) m.opts[k] = typeof SV.DEFAULT_OPTIONS[k] === 'boolean' ? v === '1' || /^true$/i.test(v) : (isFinite(+v) ? +v : SV.DEFAULT_OPTIONS[k]);
      });
      break;
    }
    if (!t) break;
    if ((mm = /^=AND\(INT\((.+)\)=\1\)$/i.exec(t))) m.cons.push({ l: '=' + mm[1], op: 'int', r: '' });
    else if ((mm = /^=AND\(\((.+)=0\)\+\(\1=1\)\)$/i.exec(t))) m.cons.push({ l: '=' + mm[1], op: 'bin', r: '' });
    else if ((mm = /^=AND\(SORT\(TOCOL\((.+)\)\)=SEQUENCE\(COUNT\(\1\)\)\)$/i.exec(t))) m.cons.push({ l: '=' + mm[1], op: 'dif', r: '' });
    else {
      let inner = t.replace(/^=/, '');
      if ((mm = /^AND\((.+)\)$/i.exec(inner))) inner = mm[1];
      const p = splitComparison(inner);
      if (!p) fail(`"${t}" is not a Solver constraint.`);
      m.cons.push({ l: '=' + p[0], op: p[1], r: '=' + p[2] });
    }
  }
  return m;
}

function loadSaveDialog(A, sheetId, m, onLoad) {
  const need = modelCells(m).length;
  const f = refField(A, M.rectText(A.selRect(), true), { title: 'Model area', label: 'Model area', baseId: sheetId });
  const readArea = () => {
    const sh = sheetById(A, sheetId);
    const a = resolveArea(A.engine, A.view().names, f.input.value, sh);
    if (!a) { toast('Select the model area, such as $H$2:$H$10.', 'error'); return null; }
    return a;
  };
  dialog('Load/Save Model', h('div', { class: 'form' },
    h('p', { class: 'dlg-text' }, `To save, select the first cell of an empty column range (this model needs ${need} cells). To load, select the cells of a saved model.`),
    field('Model area', f.el)), [
    { label: 'Cancel' },
    { label: 'Load', action: async () => {
      const a = readArea();
      if (!a) return false;
      const texts = [];
      for (let r = a.r1; r <= (a.r1 === a.r2 ? a.sheet.rows.length - 1 : a.r2); r++) {
        const raw = A.engine.raw(a.sheet, r, a.c1);
        if (raw === undefined || raw === '') break;
        texts.push(String(raw).replace(/^'/, ''));
      }
      let loaded;
      try { loaded = parseModelCells(texts); } catch (e) { toast(e.message, 'error'); return false; }
      const mode = await choiceDialog('Load model', 'Replace the current model or merge the loaded constraints into it?', [['merge', 'Merge'], ['replace', 'Replace']]);
      if (!mode) return false;
      if (mode === 'merge') { loaded = { ...m, cons: m.cons.concat(loaded.cons) }; }
      onLoad(loaded);
    } },
    { label: 'Save', primary: true, action: async () => {
      const a = readArea();
      if (!a) return false;
      if (a.sheet.id !== sheetId) { toast('Save the model on this sheet.', 'error'); return false; }
      const cells = modelCells(m);
      if (a.r1 + cells.length > a.sheet.rows.length) { toast('There are not enough rows below this cell.', 'error'); return false; }
      const used = cells.some((_, i) => { const raw = A.engine.raw(a.sheet, a.r1 + i, a.c1); return raw !== undefined && raw !== ''; });
      if (used && !(await confirmDialog('Save model', `The model needs ${cells.length} cells starting at ${addr(a.sheet, a.r1, a.c1)}. Replace the contents of these cells?`, 'Replace'))) return false;
      const ops = cells.map((v, i) => ({ t: 'set', s: sheetId, k: M.keyAt(a.sheet, a.r1 + i, a.c1), v }));
      if (A.commit(ops, { label: 'Save Solver model' })) toast(`Saved the model in ${addr(a.sheet, a.r1, a.c1)}:${addr(a.sheet, a.r1 + cells.length - 1, a.c1).replace(/^.*!/, '')}.`);
    } }
  ]);
}

function choiceDialog(title, text, choices) {
  return new Promise((resolve) => {
    let result = null;
    dialog(title, h('p', { class: 'dlg-text' }, text), [{ label: 'Cancel' },
      ...choices.map(([v, label], i) => ({ label, primary: i === choices.length - 1, action: () => { result = v; } }))],
    { onClose: () => resolve(result) });
  });
}

/* ================================================================ solving */

const RESULT_HELP = {
  found: 'When the GRG engine is used, Solver has found at least a local optimal solution. When Simplex LP is used, this means Solver has found a global optimal solution.',
  converged: 'Solver has performed 5 iterations for which the objective did not move significantly. Try a smaller convergence setting, or a different starting point.',
  noImprove: 'Solver could not find a better solution in the time without improvement set in Options. Try a larger population, or a longer time.',
  intTol: 'Solver found an integer solution within the Integer Optimality tolerance. To be sure of the best integer solution, set Integer Optimality (%) to 0 in Options.',
  global: 'Multistart found this solution from several starting points; it is very likely, but not certain, to be the global optimum.',
  infeasible: 'Solver could not find a point for which all Constraints are satisfied. Choose a Feasibility report to see which constraints conflict.',
  unbounded: 'Solver can make the Objective Cell as large (or as small, when minimizing) as it wants. Add constraints or bounds to the variable cells.',
  linearity: 'Use the GRG Nonlinear or Evolutionary engine, or choose the Linearity report to see which formulas are not linear.',
  stopped: 'You can keep the current values or restore the original values.',
  time: 'You can keep the current values or restore the original values, or raise Max Time in Options.',
  iterations: 'You can keep the current values or restore the original values, or raise Iterations in Options.',
  subproblems: 'Raise "Max subproblems" in Options to let Solver explore more of the problem.',
  feasibleLimit: 'Raise "Max feasible solutions" in Options to let Solver try more integer solutions.',
  error: 'Check the formulas of the objective and of the constraint cells for error values at the current values.',
  bounds: 'Add lower and upper bounds for every variable cell, or turn off "Require bounds on variables" in Options.'
};
const EVO_HELP = {
  converged: 'Nearly all members of the population have the same objective value. Try a smaller convergence setting, a larger population or a higher mutation rate to look further.',
  infeasible: 'No member of the population satisfied all Constraints. Check the constraints, or try a larger population or a longer time.'
};
const SOLVED = ['found', 'converged', 'noImprove', 'intTol', 'global'];

async function runSolver(A, sheetId, model) {
  let ctx;
  try { ctx = buildProblem(A, sheetId, model); } catch (e) {
    if (!(e instanceof InputError)) console.error(e);
    await infoDialog('Solver', e.message || String(e));
    solverDialog(A, model);
    return;
  }
  const opts = { ...SV.DEFAULT_OPTIONS, ...model.opts };
  if (!isFinite(ctx.original.f) || ctx.original.g.some((v) => !isFinite(v))) {
    await infoDialog('Solver', SV.MESSAGES.error + ' ' + RESULT_HELP.error);
    solverDialog(A, model);
    return;
  }
  let stop = false;
  const status = h('span', null, 'Setting up the problem…');
  const prog = dialog('Solving…', h('div', { class: 'form' }, h('div', { class: 'loading solve-status' }, h('span', { class: 'loader', 'aria-hidden': 'true' }), status), h('p', { class: 'help' }, 'Press Stop or Esc to stop.')), [
    { label: 'Stop', action: () => { stop = true; status.textContent = 'Stopping…'; return false; } }
  ], { sticky: true, class: 'progress-dialog', onClose: () => { stop = true; } });
  const hooks = {
    shouldStop: () => stop,
    onProgress: (i) => { status.textContent = `Iteration ${i.iteration}${i.subproblem ? ', subproblem ' + i.subproblem : ''}${i.objective !== undefined && ctx.obj && ctx.target === null ? ', objective ' + fmtNum(i.objective, A.LOCALE) : ''} · ${i.elapsed.toFixed(1)} s`; },
    onIteration: (i) => trialDialog(A, ctx, i, 'Solver paused, current solution values are shown below.'),
    onLimit: (lim, i) => trialDialog(A, ctx, i, lim === 'time' ? 'The maximum time limit was reached; continue anyway?' : 'The maximum iteration limit was reached; continue anyway?')
  };
  const engine = model.method === 'lp' ? SV.simplexLP : model.method === 'evo' ? SV.evolutionary : SV.grg;
  let res;
  try { res = await engine(ctx.P, opts, hooks); } catch (e) {
    console.error(e);
    res = { status: 'error', message: SV.MESSAGES.error, time: 0, iterations: 0, subproblems: 0 };
  }
  prog.close();
  if (res.x && !opts.ignoreInt) res.x = res.x.map((v, j) => (ctx.P.isInt[j] ? Math.round(v) : v));
  resultsDialog(A, ctx, opts, res);
}

function valueTable(A, ctx, x, before) {
  const L = A.LOCALE;
  // values in the number format of their cells (currency, percent, …)
  const fmt = (key, v) => {
    const st = (ctx.ssh.cells[key] || {}).st;
    return typeof v === 'number' && isFinite(v) ? formatValue(tidy(v), st && st.nf ? st : null, L) : fmtNum(v, L);
  };
  const rows = [];
  if (ctx.obj) {
    const e = ctx.P.evaluate(x);
    const k = M.keyAt(ctx.ssh, ctx.obj.r, ctx.obj.c);
    rows.push([ctx.obj.addr + (ctx.obj.name ? ' ' + ctx.obj.name : ''), before ? fmt(k, ctx.original.objCell) : null, fmt(k, e.objCell), true]);
  }
  ctx.vars.slice(0, 12).forEach((v, j) => rows.push([v.addr + (v.name ? ' ' + v.name : ''), before ? fmt(v.key, v.start) : null, fmt(v.key, x[j]), false]));
  const table = h('table', { class: 'solver-values' },
    h('thead', null, h('tr', null, h('th', null, 'Cell'), before ? h('th', null, 'Original') : null, h('th', null, before ? 'Solution' : 'Current'))),
    h('tbody', null, rows.map(([a, b, c, isObj]) => h('tr', { class: isObj ? 'obj' : '' }, h('td', null, a), before ? h('td', { class: 'num' }, b) : null, h('td', { class: 'num' }, c)))));
  const more = ctx.vars.length > 12 ? h('p', { class: 'help' }, `and ${ctx.vars.length - 12} more variable cells`) : null;
  return h('div', { class: 'solver-values-wrap' }, table, more);
}

/** "Show Trial Solution": pause with Continue / Stop / Save Scenario. Resolves true to continue. */
function trialDialog(A, ctx, info, message) {
  return new Promise((resolve) => {
    let go = false;
    const x = info && info.x;
    dialog('Show trial solution', h('div', { class: 'form' },
      h('p', { class: 'dlg-text' }, message),
      h('p', { class: 'help' }, `Iteration ${info.iteration}`),
      x ? valueTable(A, ctx, x, false) : null), [
      x ? { label: 'Save scenario…', action: async () => { await saveScenarioFrom(A, ctx, x); return false; } } : null,
      { label: 'Stop' },
      { label: 'Continue', primary: true, action: () => { go = true; } }
    ].filter(Boolean), { sticky: true, onClose: () => resolve(go) });
  });
}

function infoDialog(title, text) {
  return new Promise((resolve) => dialog(title, h('p', { class: 'dlg-text' }, text), [{ label: 'OK', primary: true }], { onClose: resolve }));
}

/* ================================================================ Solver Results */

function resultsDialog(A, ctx, opts, res) {
  const model = ctx.model;
  const hasX = !!res.x;
  const isLP = model.method === 'lp', isEvo = model.method === 'evo';
  const intOn = ctx.hasInt && !opts.ignoreInt;
  const reports = [];
  if (res.status === 'linearity') reports.push(['linearity', 'Linearity']);
  else if (res.status === 'infeasible' && !isEvo) reports.push(['feasibility', 'Feasibility'], ['feasibilityBounds', 'Feasibility-Bounds']);
  else if (hasX) {
    reports.push(['answer', 'Answer']);
    if (isEvo && res.population) reports.push(['population', 'Population']);
    else if (!isEvo && !intOn && SOLVED.includes(res.status)) reports.push(['sensitivity', 'Sensitivity'], ['limits', 'Limits']);
  }
  const keepName = 'keep-' + Core.randomId();
  const keep = radio(keepName, 'keep', 'Keep Solver solution', hasX);
  const restore = radio(keepName, 'restore', 'Restore original values', !hasX);
  keep.input.disabled = !hasX;
  const back = checkbox('Return to Solver parameters', false);
  const repChecks = reports.map(([k, label]) => { const c = checkbox(label, false); c.input.dataset.report = k; return c; });
  const ok = SOLVED.includes(res.status);
  const body = h('div', { class: 'form solver-results' },
    h('p', { class: 'dlg-text result-msg' + (ok ? ' ok' : ' warn') }, res.message || SV.MESSAGES[res.status] || res.status),
    h('div', { class: 'results-cols' },
      h('div', { class: 'checks cols' }, keep.el, restore.el),
      reports.length ? h('fieldset', { class: 'opt-group' }, h('legend', null, 'Reports'), h('div', { class: 'checks cols' }, repChecks.map((c) => c.el))) : null),
    back.el,
    h('p', { class: 'help' }, (isEvo && EVO_HELP[res.status]) || RESULT_HELP[res.status] || ''),
    hasX ? valueTable(A, ctx, res.x, true) : null,
    h('p', { class: 'help' }, `${METHODS[model.method]} · ${res.time.toFixed(3)} s · ${res.iterations} iterations${res.subproblems ? ' · ' + res.subproblems + ' subproblems' : ''}`));
  let reopen = false;
  dialog('Solver results', body, [
    hasX ? { label: 'Save scenario…', action: async () => { await saveScenarioFrom(A, ctx, res.x); return false; } } : null,
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: async () => {
      reopen = back.input.checked;
      const wanted = repChecks.filter((c) => c.input.checked).map((c) => c.input.dataset.report);
      const ops = [];
      const sh = sheetById(A, ctx.sheetId);
      if (!sh) { toast('The sheet of this model no longer exists.', 'error'); return; }
      if (keep.input.checked && hasX) {
        ctx.vars.forEach((v, j) => {
          const x = ctx.P.isInt[j] && !opts.ignoreInt ? Math.round(res.x[j]) : tidy(res.x[j]);
          ops.push({ t: 'set', s: ctx.sheetId, k: v.key, v: numToText(x) });
        });
      }
      let made = [];
      if (wanted.length) {
        try { made = await buildReports(A, ctx, opts, res, wanted); } catch (e) { console.error(e); toast('Could not create the reports: ' + (e.message || e), 'error'); }
        const room = Core.LIMITS.sheets - A.view().sheets.length;
        if (made.length > room) {
          toast(`A spreadsheet can have at most ${Core.LIMITS.sheets} sheets, so ${room <= 0 ? 'no reports were added' : `only the first ${room} reports were added`}.`, 'error');
          made = made.slice(0, Math.max(0, room));
        }
      }
      const valueOps = ops.slice();
      const at = A.view().sheets.findIndex((s) => s.id === ctx.sheetId);
      made.forEach((m, i) => ops.push({ t: 'addSheet', at: Math.max(0, at) + i, sheet: m }));
      if (!ops.length) return;
      if (A.commit(ops, { label: 'Solver' })) { if (made.length) toast(`Created ${made.map((m) => m.name).join(', ')}.`); }
      else if (made.length && valueOps.length) A.commit(valueOps, { label: 'Solver' }); // the reports failed: keep the solution at least
    } }
  ].filter(Boolean), { sticky: true, wide: true, class: 'results-dialog', onClose: () => { if (reopen) setTimeout(() => solverDialog(A, model), 0); } });
}

/* ================================================================ reports */

const B = (v) => ({ v, st: { b: 1 } });
const HEAD = (v) => ({ v, st: { b: 1, bg: '#e8ebf0' } });
function reportName(doc, base, taken) {
  for (let i = 1; ; i++) {
    const name = `${base} ${i}`;
    const low = name.toLowerCase();
    if (!doc.sheets.some((s) => s.name.toLowerCase() === low) && !taken.has(low)) { taken.add(low); return name; }
  }
}
/** A new sheet with the report rows (arrays of values or { v, st }); column B onwards holds the tables. */
function reportSheet(name, rows, widths) {
  const nR = Math.max(rows.length + 5, 30), nC = Math.max(10, ...rows.map((r) => r.length + 1));
  const sh = Core.newSheet(name, nR, nC);
  rows.forEach((row, r) => row.forEach((cell, c) => {
    if (cell === null || cell === undefined || cell === '') return;
    const o = typeof cell === 'object' && !isErr(cell) ? cell : { v: cell };
    const raw = rawOf(typeof o.v === 'number' ? (Math.abs(o.v) >= BIG ? (o.v > 0 ? BIG : -BIG) : tidy(o.v)) : o.v);
    if (!raw) return;
    const out = { v: raw };
    if (o.st) out.st = o.st;
    sh.cells[sh.rows[r] + ':' + sh.cols[c]] = out;
  }));
  (widths || [24, 90, 170, 120, 120, 120, 120, 120, 110]).forEach((w, c) => { if (sh.cols[c]) sh.colW[sh.cols[c]] = w; });
  return sh;
}
function reportHead(A, ctx, title, res) {
  return [
    [B(`Spreadsheets for YouTrack ${title}`)],
    [`Worksheet: ${ctx.ssh.name}`],
    [`Report Created: ${new Date().toLocaleString(A.LOCALE)}`],
    [`Result: ${res.message || SV.MESSAGES[res.status] || res.status}`]
  ];
}
const unlimited = (v, L) => (v > 0 ? fmtNum(v, L) : 'Unlimited');
function engineBlock(A, ctx, opts, res) {
  const L = A.LOCALE, m = ctx.model;
  const rows = [[B('Solver Engine')],
    ['', `Engine: ${METHODS[m.method]}`],
    ['', `Solution Time: ${fmtNum(+res.time.toFixed(3), L)} Seconds.`],
    ['', `Iterations: ${res.iterations} Subproblems: ${res.subproblems || 0}`],
    [B('Solver Options')],
    ['', `Max Time ${unlimited(opts.maxTime, L)}, Iterations ${unlimited(opts.iterations, L)}, Precision ${fmtNum(opts.precision, L)}${opts.scaling ? ', Use Automatic Scaling' : ''}${opts.showIter ? ', Show Iteration Results' : ''}`]];
  if (m.method === 'grg') rows.push(['', `Convergence ${fmtNum(opts.convergence, L)}, Derivatives ${opts.derivatives === 'central' ? 'Central' : 'Forward'}${opts.multistart ? `, Multistart, Population Size ${opts.population}, Random Seed ${opts.seed}${opts.requireBounds ? ', Require Bounds' : ''}` : ''}`]);
  if (m.method === 'evo') rows.push(['', `Convergence ${fmtNum(opts.convergence, L)}, Population Size ${opts.population}, Random Seed ${opts.seed}, Mutation Rate ${fmtNum(opts.mutation, L)}, Time w/o Improvement ${opts.maxNoImprove} sec${opts.requireBounds ? ', Require Bounds' : ''}`]);
  rows.push(['', `Max Subproblems ${unlimited(opts.maxSubproblems, L)}, Max Integer Sols ${unlimited(opts.maxFeasible, L)}, Integer Tolerance ${fmtNum(opts.intTolerance, L)}%${opts.ignoreInt ? ', Ignore Integer Constraints' : ''}${m.nonneg ? ', Assume NonNegative' : ''}`]);
  return rows;
}
const goalText = (m) => (m.goal === 'max' ? 'Max' : m.goal === 'min' ? 'Min' : 'Value Of');
function statusOf(e, tol) {
  const l = e.lhsGet(), r = e.rhsGet();
  const t = tol * Math.max(1, Math.abs(r));
  if (!isFinite(l) || !isFinite(r)) return { value: l, status: 'Violated', slack: '' };
  if (Math.abs(l - r) <= t) return { value: l, status: 'Binding', slack: 0 };
  const slack = e.rel === 'le' ? r - l : e.rel === 'ge' ? l - r : -Math.abs(l - r);
  return { value: l, status: e.rel === 'eq' || slack < 0 ? 'Violated' : 'Not Binding', slack };
}
function constraintRows(ctx, entries, x, tol) {
  const rows = [];
  for (const e of entries) {
    if (e.kind === 'row' || e.kind === 'bound') {
      const s = statusOf(e, tol);
      rows.push(['', e.addr, e.name, s.value, e.formula, s.status, s.slack]);
    } else if (e.kind === 'dif') {
      // the values of the group must be exactly 1..N
      const vals = x ? e.js.map((j) => Math.round(x[j])).sort((a, b) => a - b) : [];
      const ok = x && e.js.every((j) => Math.abs(x[j] - Math.round(x[j])) <= 1e-6) && vals.every((v, i) => v === i + 1);
      rows.push(['', e.addr, e.name, '', e.formula, ok ? 'Binding' : 'Violated', 0]);
    } else {
      const v = x ? x[e.j] : NaN;
      const whole = Math.abs(v - Math.round(v)) <= 1e-6;
      const ok = whole && (e.kind !== 'bin' || Math.round(v) === 0 || Math.round(v) === 1);
      rows.push(['', e.addr, e.name, x ? tidy(v) : '', e.formula, ok ? 'Binding' : 'Violated', 0]);
    }
  }
  return rows;
}

async function buildReports(A, ctx, opts, res, wanted) {
  const L = A.LOCALE, m = ctx.model, P = ctx.P;
  const taken = new Set();
  const doc = A.view();
  const out = [];
  const x = res.x || P.x0;
  const final = P.evaluate(x); // the private copy now holds the solution: entry getters read it
  const tol = Math.max(opts.precision, 1e-9);
  const kindOf = (j) => (ctx.intKind[j] && !opts.ignoreInt ? ctx.intKind[j] : 'Contin');
  const generalRows = ctx.rows; // the internal "Value Of" row is not a user constraint

  for (const w of wanted) {
    if (w === 'answer') {
      const rows = reportHead(A, ctx, 'Answer Report', res).concat(engineBlock(A, ctx, opts, res));
      rows.push([], [B(`Objective Cell (${goalText(m)})`)], ['', HEAD('Cell'), HEAD('Name'), HEAD('Original Value'), HEAD('Final Value')]);
      if (ctx.obj) rows.push(['', ctx.obj.addr, ctx.obj.name, ctx.original.objCell, final.objCell]);
      rows.push([], [B('Variable Cells')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Original Value'), HEAD('Final Value'), HEAD('Integer')]);
      ctx.vars.forEach((v, j) => rows.push(['', v.addr, v.name, v.start, x[j], kindOf(j)]));
      rows.push([], [B('Constraints')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Cell Value'), HEAD('Formula'), HEAD('Status'), HEAD('Slack')]);
      P.evaluate(x);
      rows.push(...constraintRows(ctx, ctx.listed.filter((e) => opts.ignoreInt ? !OP_WORD[e.kind] : true), x, tol));
      out.push(reportSheet(reportName(doc, 'Answer Report', taken), rows));
    }
    if (w === 'sensitivity') {
      const rows = reportHead(A, ctx, 'Sensitivity Report', res).slice(0, 3);
      const sens = res.sensitivity;
      P.evaluate(x);
      if (sens && sens.kind === 'lp') {
        const lin = res.linear;
        const sense = P.sense;
        rows.push([], [B('Variable Cells')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Final Value'), HEAD('Reduced Cost'), HEAD('Objective Coefficient'), HEAD('Allowable Increase'), HEAD('Allowable Decrease')]);
        ctx.vars.forEach((v, j) => {
          const [dec, inc] = sens.costRange[j];
          const coef = ctx.target === null ? lin.c[j] : 0;
          rows.push(['', v.addr, v.name, x[j], sense * sens.reduced[j], coef, sense > 0 ? inc : dec, sense > 0 ? dec : inc]);
        });
        rows.push([], [B('Constraints')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Final Value'), HEAD('Shadow Price'), HEAD('Constraint R.H. Side'), HEAD('Allowable Increase'), HEAD('Allowable Decrease')]);
        generalRows.forEach((e, k) => {
          const [dec, inc] = sens.gRange[k];
          rows.push(['', e.addr, e.name, e.lhsGet(), sense * sens.gDuals[k], e.rhsGet(), inc, dec]);
        });
      } else if (sens && sens.kind === 'nlp') {
        rows.push([], [B('Variable Cells')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Final Value'), HEAD('Reduced Gradient')]);
        ctx.vars.forEach((v, j) => rows.push(['', v.addr, v.name, x[j], Math.abs(sens.reducedGradient[j]) < 1e-7 ? 0 : sens.reducedGradient[j]]));
        rows.push([], [B('Constraints')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Final Value'), HEAD('Lagrange Multiplier')]);
        generalRows.forEach((e, k) => rows.push(['', e.addr, e.name, e.lhsGet(), Math.abs(sens.multipliers[k]) < 1e-7 ? 0 : sens.multipliers[k]]));
      } else rows.push([], ['Sensitivity information is not available for this solution.']);
      out.push(reportSheet(reportName(doc, 'Sensitivity Report', taken), rows));
    }
    if (w === 'limits') {
      const rows = reportHead(A, ctx, 'Limits Report', res).slice(0, 3);
      const lim = SV.variableLimits(P, x, m.method === 'lp' ? res.linear : null);
      P.evaluate(x);
      rows.push([], ['', HEAD('Objective')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Value')]);
      if (ctx.obj) rows.push(['', ctx.obj.addr, ctx.obj.name, final.objCell]);
      rows.push([], ['', HEAD('Variable'), '', '', HEAD('Lower'), HEAD('Objective'), HEAD('Upper'), HEAD('Objective')],
        ['', HEAD('Cell'), HEAD('Name'), HEAD('Value'), HEAD('Limit'), HEAD('Result'), HEAD('Limit'), HEAD('Result')]);
      ctx.vars.forEach((v, j) => {
        const l = lim[j];
        rows.push(['', v.addr, v.name, x[j], l.lower, l.objLower ?? '#N/A', l.upper, l.objUpper ?? '#N/A']);
      });
      P.evaluate(x);
      out.push(reportSheet(reportName(doc, 'Limits Report', taken), rows));
    }
    if (w === 'population' && res.population) {
      const rows = reportHead(A, ctx, 'Population Report', res).slice(0, 3);
      const cols = ['', HEAD('Cell'), HEAD('Name'), HEAD('Best Value'), HEAD('Mean Value'), HEAD('Standard Deviation'), HEAD('Maximum Value'), HEAD('Minimum Value')];
      rows.push([], [B('Variable Cells')], cols);
      ctx.vars.forEach((v, j) => { const s = res.population.vars[j]; rows.push(['', v.addr, v.name, s.best, s.mean, s.sd, s.max, s.min]); });
      rows.push([], [B('Constraints')], cols.slice());
      generalRows.forEach((e, k) => { const s = res.population.constraints[k]; rows.push(['', e.addr, e.name, s.best, s.mean, s.sd, s.max, s.min]); });
      out.push(reportSheet(reportName(doc, 'Population Report', taken), rows));
    }
    if (w === 'linearity') {
      const rows = reportHead(A, ctx, 'Linearity Report', res).slice(0, 4);
      const lr = SV.linearityReport(P, P.x0.map((v, j) => Math.min(P.ub[j], Math.max(P.lb[j], v))));
      P.evaluate(P.x0);
      const yn = (b) => (b ? 'Yes' : 'No');
      rows.push([], [B(`Objective Cell (${goalText(m)})`)], ['', HEAD('Cell'), HEAD('Name'), HEAD('Value'), HEAD('Linear Function')]);
      if (ctx.obj && lr) rows.push(['', ctx.obj.addr, ctx.obj.name, ctx.original.objCell, yn(ctx.target === null ? lr.objective : lr.constraints[lr.constraints.length - 1])]);
      rows.push([], [B('Variable Cells')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Value'), HEAD('Occurs Linearly')]);
      ctx.vars.forEach((v, j) => rows.push(['', v.addr, v.name, v.start, lr ? yn(lr.vars[j]) : '']));
      rows.push([], [B('Constraints')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Value'), HEAD('Formula'), HEAD('Linear Function')]);
      ctx.listed.forEach((e) => {
        if (e.kind === 'row') rows.push(['', e.addr, e.name, e.lhsGet(), e.formula, lr ? yn(lr.constraints[e.row]) : '']);
        else rows.push(['', e.addr, e.name, e.kind === 'bound' ? e.lhsGet() : '', e.formula, 'Yes']);
      });
      out.push(reportSheet(reportName(doc, 'Linearity Report', taken), rows));
    }
    if (w === 'feasibility' || w === 'feasibilityBounds') {
      const withBounds = w === 'feasibility';
      const rows = reportHead(A, ctx, withBounds ? 'Feasibility Report' : 'Feasibility-Bounds Report', res);
      ctx.lastX = x;
      const iis = await findConflict(ctx, opts, withBounds);
      P.evaluate(x);
      if (iis === null) rows.push([], ['The problem is feasible when the integer constraints are ignored; the integer constraints make it infeasible.']);
      else if (!iis.length) {
        rows.push([], [withBounds ? 'The bounds on the variable cells cannot be met together.'
          : 'The bounds on the variable cells (constraints that compare a variable cell with a constant) cannot be met together. Choose the Feasibility report to include them in the analysis.']);
      } else {
        rows.push([], [B('Constraints that make the problem infeasible')], ['', HEAD('Cell'), HEAD('Name'), HEAD('Cell Value'), HEAD('Formula'), HEAD('Status'), HEAD('Slack')]);
        rows.push(...constraintRows(ctx, iis, x, tol));
        if (!withBounds) rows.push([], ['All bounds on the variable cells were kept; only the other constraints were analyzed.']);
      }
      out.push(reportSheet(reportName(doc, withBounds ? 'Feasibility Report' : 'Feasibility-Bounds Report', taken), rows));
    }
  }
  return out;
}

/** A smallest set of constraints (and bounds) that cannot be satisfied together (deletion filter). */
async function findConflict(ctx, opts, withBounds) {
  const P = ctx.P;
  const elements = ctx.rows.map((e) => ({ type: 'row', e })).concat(withBounds ? ctx.bounds.map((e) => ({ type: 'bound', e })) : []);
  if (ctx.target !== null) {
    // "Value Of" is a condition like the constraints: objective cell = target
    const k = P.ctype.length - 1, o = ctx.obj;
    elements.push({ type: 'row', target: true, e: {
      kind: 'row', rel: 'eq', row: k, addr: o.addr, name: o.name, formula: `${o.addr}=${numToText(ctx.target)} (Value Of)`,
      lhsGet: () => { const e = P.evaluate(ctx.lastX || P.x0); return e.objCell; }, rhsGet: () => ctx.target
    } });
  }
  const quick = { ...opts, showIter: false, ignoreInt: true, multistart: false, maxTime: 20, iterations: 0 };
  const feasible = async (active) => {
    const rowIdx = active.filter((a) => a.type === 'row').map((a) => a.e.row);
    const { lb, ub } = ctx.boundsFor(withBounds ? active.filter((a) => a.type === 'bound').map((a) => a.e) : ctx.bounds);
    if (lb.some((v, j) => v > ub[j] + 1e-12)) return false;
    const sub = {
      ...P, lb, ub, sense: 1, ctype: rowIdx.map((k) => P.ctype[k]),
      evaluate: (x) => { const e = P.evaluate(x); return { f: 0, g: rowIdx.map((k) => e.g[k]), lhs: rowIdx.map((k) => e.lhs[k]), rhs: rowIdx.map((k) => e.rhs[k]), objCell: 0 }; }
    };
    const r = ctx.model.method === 'lp' ? await SV.simplexLP(sub, quick) : await SV.grg(sub, quick);
    return !!(r.x && r.feasible !== false && SOLVED.concat(['limit']).includes(r.status));
  };
  if (await feasible(elements)) return null;
  let active = elements.slice();
  for (const el of elements) {
    const without = active.filter((a) => a !== el);
    if (!(await feasible(without))) active = without;
  }
  return active.map((a) => a.e);
}

/* ================================================================ scenarios */

function scenariosOf(sheet) { return (sheet && sheet.scenarios || []).slice(); }

async function saveScenarioFrom(A, ctx, x) {
  const sh = sheetById(A, ctx.sheetId);
  if (!sh) return;
  const name = await promptDialog('Save scenario', 'Scenario name', '');
  if (name === null) return;
  const n = name.trim();
  if (!n) { toast('Enter a name for the scenario.', 'error'); return; }
  if (scenariosOf(sh).some((s) => s.name.toLowerCase() === n.toLowerCase())) { toast('A scenario with this name already exists.', 'error'); return; }
  if (scenariosOf(sh).length >= Core.LIMITS.objects) { toast('This sheet has too many scenarios.', 'error'); return; }
  if (ctx.vars.length > 32) { toast('A scenario can change at most 32 cells, and this model has ' + ctx.vars.length + ' variable cells.', 'error'); return; }
  const val = {
    name: n.slice(0, 255), cells: ctx.vars.map((v) => '=' + v.addr), values: ctx.vars.map((_, j) => tidy(x[j])),
    comment: `Created by ${A.store.me && A.store.me.name || 'Solver'} on ${new Date().toLocaleDateString(A.LOCALE)}`
  };
  if (A.commit([{ t: 'obj', s: ctx.sheetId, kind: 'scenarios', id: Core.randomId(), val }], { label: 'Scenario' })) toast(`Saved scenario "${n}".`);
}

/** Changing cells of a scenario as positions on its sheet; null entries for references that no longer work. */
function scenarioCells(A, sh, scn) {
  return scn.cells.map((t) => { const a = resolveArea(A.engine, A.view().names, t, sh); return a && a.sheet.id === sh.id ? { r: a.r1, c: a.c1, addr: addr(sh, a.r1, a.c1) } : null; });
}

export function scenarioManager(A) {
  if (!A.canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  const sheetId = A.sheet().id;
  let selected = 0;
  const list = h('div', { class: 'rule-list scn-list', role: 'listbox', 'aria-label': 'Scenarios', tabindex: '0' });
  const info = h('div', { class: 'scn-info' });
  const btn = (label, fn) => h('button', { class: 'btn', type: 'button', onclick: fn }, label);
  const btns = {
    add: btn('Add…', () => scenarioEdit(A, sheetId, null, () => { selected = scenariosOf(sheetById(A, sheetId)).length - 1; repaint(); })),
    del: btn('Delete', () => { const s = current(); if (s && A.commit([{ t: 'obj', s: sheetId, kind: 'scenarios', id: s.id, val: null }], { label: 'Delete scenario' })) setTimeout(repaint, 0); }),
    edit: btn('Edit…', () => { const s = current(); if (s) scenarioEdit(A, sheetId, s, repaint); }),
    merge: btn('Merge…', () => mergeScenarios(A, sheetId, repaint)),
    summary: btn('Summary…', () => scenarioSummary(A, sheetId, () => { if (dlg) dlg.close(); }))
  };
  const current = () => scenariosOf(sheetById(A, sheetId))[selected] || null;
  let showBtn = null;
  function repaint() {
    const sh = sheetById(A, sheetId);
    const all = scenariosOf(sh);
    selected = Math.min(Math.max(0, selected), all.length - 1);
    list.innerHTML = '';
    if (!all.length) list.appendChild(h('div', { class: 'muted pad' }, 'No scenarios defined. Choose Add to add scenarios.'));
    all.forEach((s, i) => {
      const row = h('div', { class: 'rule cons-row' + (i === selected ? ' on' : ''), role: 'option', 'aria-selected': i === selected ? 'true' : 'false' }, h('span', { class: 'rule-text' }, s.name));
      row.addEventListener('click', () => { selected = i; repaint(); });
      row.addEventListener('dblclick', () => { selected = i; showScenario(A, sheetId, s); });
      list.appendChild(row);
    });
    const s = current();
    info.innerHTML = '';
    if (s) info.append(field('Changing cells', h('div', { class: 'mono' }, s.cells.map(strip).join(', '))), field('Comment', h('div', { class: 'scn-comment' }, s.comment || '')));
    btns.del.disabled = btns.edit.disabled = !s;
    btns.summary.disabled = !all.length;
    if (showBtn) showBtn.disabled = !s;
  }
  list.addEventListener('keydown', (e) => {
    const n = scenariosOf(sheetById(A, sheetId)).length;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); selected = Math.max(0, Math.min(n - 1, selected + (e.key === 'ArrowDown' ? 1 : -1))); repaint(); }
  });
  const dlg = dialog('Scenario Manager', h('div', { class: 'scn' }, h('div', { class: 'cons-wrap' }, list, h('div', { class: 'cons-btns' }, Object.values(btns))), info), [
    { label: 'Show', primary: true, action: () => { const s = current(); if (s) showScenario(A, sheetId, s); return false; } },
    { label: 'Close' }
  ], { wide: true });
  showBtn = dlg.el.querySelector('.modal-foot .btn.primary');
  repaint();
}

function showScenario(A, sheetId, s) {
  const sh = sheetById(A, sheetId);
  if (!sh) return;
  const cells = scenarioCells(A, sh, s);
  const ops = [];
  cells.forEach((c, i) => { if (c) ops.push({ t: 'set', s: sheetId, k: M.keyAt(sh, c.r, c.c), v: rawOf(s.values[i]) }); });
  if (cells.some((c) => !c)) toast('Some changing cells of this scenario no longer exist.', 'error');
  if (A.commit(ops, { label: 'Show scenario' })) toast(`Showing scenario "${s.name}".`);
}

function scenarioEdit(A, sheetId, scn, onDone) {
  const sh = sheetById(A, sheetId);
  const nameIn = h('input', { class: 'input', value: scn ? scn.name : '', autofocus: true, maxlength: 255, 'aria-label': 'Scenario name' });
  const cellsF = refField(A, scn ? scn.cells.map(strip).join(', ') : M.rectText(A.selRect(), true), { title: 'Changing cells', label: 'Changing cells', baseId: sheetId, append: false });
  const who = A.store.me && A.store.me.name || '';
  const today = new Date().toLocaleDateString(A.LOCALE);
  const comment = h('textarea', { class: 'input', rows: 3, 'aria-label': 'Comment' });
  comment.value = scn ? (scn.comment || '') + (who ? `\nModified by ${who} on ${today}` : '') : (who ? `Created by ${who} on ${today}` : '');
  dialog(scn ? 'Edit scenario' : 'Add scenario', h('div', { class: 'form' },
    field('Scenario name', nameIn), field('Changing cells', cellsF.el, 'Up to 32 cells; several ranges are separated by commas.'), field('Comment', comment)), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: async () => {
      const name = nameIn.value.trim();
      if (!name) { toast('Enter a name for the scenario.', 'error'); return false; }
      const others = scenariosOf(sheetById(A, sheetId)).filter((s) => !scn || s.id !== scn.id);
      if (others.some((s) => s.name.toLowerCase() === name.toLowerCase())) { toast('A scenario with this name already exists.', 'error'); return false; }
      const cells = [];
      for (const part of splitList(cellsF.input.value)) {
        const a = resolveArea(A.engine, A.view().names, part, sh);
        if (!a) { toast(`"${part}" is not a valid reference.`, 'error'); return false; }
        if (a.sheet.id !== sheetId) { toast('Changing cells must be on this sheet.', 'error'); return false; }
        for (const [r, c] of cellsOf(a)) if (!cells.some((x) => x.r === r && x.c === c)) cells.push({ r, c, addr: addr(sh, r, c) });
      }
      if (!cells.length) { toast('Enter the changing cells.', 'error'); return false; }
      if (cells.length > 32) { toast('A scenario can change at most 32 cells.', 'error'); return false; }
      if (cells.some((c) => isFormula(A.engine.raw(sh, c.r, c.c))) &&
        !(await confirmDialog('Add scenario', 'One or more changing cells contain a formula. Showing the scenario replaces these formulas with values.', 'OK'))) return false;
      const prev = scn ? Object.fromEntries(scn.cells.map((t, i) => [strip(t), scn.values[i]])) : {};
      const values = await scenarioValues(A, sh, cells, cells.map((c) => (c.addr in prev ? prev[c.addr] : A.engine.valueAt(sh, c.r, c.c))));
      if (!values) return false;
      const val = { name: name.slice(0, 255), cells: cells.map((c) => '=' + c.addr), values, comment: comment.value.slice(0, 1000) };
      if (A.commit([{ t: 'obj', s: sheetId, kind: 'scenarios', id: scn ? scn.id : Core.randomId(), val }], { label: scn ? 'Edit scenario' : 'Add scenario' })) setTimeout(onDone, 0);
    } }
  ], { sticky: true });
}

function scenarioValues(A, sh, cells, defaults) {
  return new Promise((resolve) => {
    let result = null;
    const inputs = cells.map((c, i) => {
      const v = defaults[i];
      return h('input', { class: 'input', value: v === null || v === undefined || isErr(v) ? '' : typeof v === 'number' ? editNum(v, A.LOCALE) : String(v), 'aria-label': c.addr });
    });
    const grid = h('div', { class: 'scn-values' }, cells.map((c, i) => [h('label', { class: 'mono' }, `${i + 1}: ${c.addr}`), inputs[i]]));
    dialog('Scenario values', h('div', { class: 'form' }, h('p', { class: 'dlg-text' }, 'Enter values for each of the changing cells.'), grid), [
      { label: 'Cancel' },
      { label: 'OK', primary: true, action: () => {
        let converted = false;
        result = inputs.map((inp, i) => {
          const n = parseNumber(inp.value, A.LOCALE);
          if (n !== null) return n;
          const p = parseInput(inp.value, A.LOCALE);
          if (p.kind === 'empty') return null;
          if (p.kind === 'formula') { converted = true; const v = A.engine.evalAt(p.raw, sh, cells[i].r, cells[i].c); return isErr(v) ? null : v; }
          if (p.num !== undefined) return p.num;
          return literal(p.raw);
        });
        if (converted) toast('Formula results were converted to values.');
      } }
    ], { sticky: true, onClose: () => resolve(result) });
  });
}

function mergeScenarios(A, sheetId, onDone) {
  const sources = A.view().sheets.filter((s) => s.id !== sheetId && scenariosOf(s).length);
  if (!sources.length) { toast('No other sheet in this workbook has scenarios.'); return; }
  const sel = select(sources.map((s) => [s.id, `${s.name} (${scenariosOf(s).length} scenarios)`]), sources[0].id, { 'aria-label': 'Sheet' });
  dialog('Merge scenarios', h('div', { class: 'form' }, field('Merge scenarios from sheet', sel, 'Their changing cells are used at the same addresses on this sheet.')), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: () => {
      const src = sheetById(A, sel.value), dst = sheetById(A, sheetId);
      if (!src || !dst) return;
      const names = new Set(scenariosOf(dst).map((s) => s.name.toLowerCase()));
      const ops = scenariosOf(src).map((s) => {
        let name = s.name, k = 2;
        while (names.has(name.toLowerCase())) name = `${s.name} (${k++})`;
        names.add(name.toLowerCase());
        const { id, ...rest } = s;
        void id;
        return { t: 'obj', s: sheetId, kind: 'scenarios', id: Core.randomId(), val: { ...rest, name } };
      });
      if (A.commit(ops, { label: 'Merge scenarios' })) { toast(`Merged ${ops.length} scenarios.`); setTimeout(onDone, 0); }
    } }
  ]);
}

function scenarioSummary(A, sheetId, onDone) {
  const sh = sheetById(A, sheetId);
  const resF = refField(A, '', { title: 'Result cells', label: 'Result cells', baseId: sheetId });
  dialog('Scenario summary', h('div', { class: 'form' },
    h('p', { class: 'dlg-text' }, 'Creates a sheet that compares the scenarios of this sheet.'),
    field('Result cells', resF.el, 'Optional: cells whose values each scenario changes, such as totals. Several ranges are separated by commas.')), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: () => {
      const results = [];
      for (const part of splitList(resF.input.value)) {
        const a = resolveArea(A.engine, A.view().names, part, sh);
        if (!a) { toast(`"${part}" is not a valid reference.`, 'error'); return false; }
        for (const [r, c] of cellsOf(a)) results.push({ sheet: a.sheet, r, c, addr: addr(a.sheet, r, c, sh) });
      }
      if (results.length > 500) { toast('Choose at most 500 result cells.', 'error'); return false; }
      const sheetOp = summarySheet(A, sh, results);
      if (A.commit([sheetOp], { label: 'Scenario summary' })) { A.goToSheet(sheetOp.sheet.id); if (onDone) setTimeout(onDone, 0); }
    } }
  ]);
}

function summarySheet(A, sh, results) {
  const all = scenariosOf(sh);
  const snap = snapshot(A);
  const ssh = snap.sheet(sh.id);
  // the union of the changing cells, in the order they first appear
  const cells = [];
  const pos = all.map((s) => scenarioCells(A, sh, s).map((c) => {
    if (!c) return -1;
    let i = cells.findIndex((x) => x.r === c.r && x.c === c.c);
    if (i < 0) { i = cells.length; cells.push({ ...c, key: M.keyAt(ssh, c.r, c.c), raw: ssh.cells[M.keyAt(ssh, c.r, c.c)] ? ssh.cells[M.keyAt(ssh, c.r, c.c)].v : '' }); }
    return i;
  }));
  const current = cells.map((c) => snap.se.valueAt(ssh, c.r, c.c));
  const resultAt = () => results.map((r) => snap.se.valueAt(snap.sheet(r.sheet.id), r.r, r.c));
  const currentResults = resultAt();
  const perScenario = all.map((s, si) => {
    cells.forEach((c) => snap.setRaw(ssh, c.key, c.raw ?? ''));
    pos[si].forEach((ci, k) => { if (ci >= 0) snap.setRaw(ssh, cells[ci].key, rawOf(s.values[k])); });
    snap.se.invalidateValues();
    return resultAt();
  });
  const SHADE = { bg: '#e8ebf0' };
  const rows = [[], ['', B('Scenario Summary')], ['', '', '', HEAD('Current Values:'), ...all.map((s) => HEAD(s.name))], ['', B('Changing Cells:')]];
  cells.forEach((c, ci) => rows.push(['', '', c.addr, current[ci], ...all.map((s, si) => {
    const k = pos[si].indexOf(ci);
    return k >= 0 ? { v: s.values[k] ?? '', st: SHADE } : '';
  })]));
  if (results.length) {
    rows.push(['', B('Result Cells:')]);
    results.forEach((r, ri) => rows.push(['', '', r.addr, currentResults[ri], ...perScenario.map((vals) => vals[ri])]));
  }
  rows.push([], ['', 'Notes: The Current Values column shows the values of the changing cells when this summary was created.'],
    ['', 'Changing cells of each scenario are shaded.']);
  const doc = A.view();
  let name = 'Scenario Summary', k = 2;
  while (doc.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = `Scenario Summary ${k++}`;
  const sheet = reportSheet(name, rows, [24, 110, 120, ...all.map(() => 110), 110]);
  return { t: 'addSheet', at: Math.max(0, doc.sheets.findIndex((s) => s.id === sh.id)), sheet };
}

/* ================================================================ Goal Seek */

export function goalSeek(A) {
  if (!A.canEdit()) { toast('You can view this spreadsheet but not edit it.'); return; }
  const sh = A.sheet();
  const s = A.selRect();
  const setF = refField(A, addr(sh, s.r1, s.c1), { title: 'Set cell', label: 'Set cell' });
  const toIn = h('input', { class: 'input', value: '', 'aria-label': 'To value', inputmode: 'decimal' });
  const byF = refField(A, '', { title: 'By changing cell', label: 'By changing cell' });
  dialog('Goal Seek', h('div', { class: 'form' }, field('Set cell', setF.el), field('To value', toIn), field('By changing cell', byF.el)), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: async () => {
      const base = sheetById(A, sh.id);
      const a = resolveArea(A.engine, A.view().names, setF.input.value, base);
      if (!a || a.r1 !== a.r2 || a.c1 !== a.c2) { toast('Set cell must be a reference to a single cell.', 'error'); return false; }
      if (!isFormula(A.engine.raw(a.sheet, a.r1, a.c1))) { toast('Set cell must contain a formula.', 'error'); return false; }
      const target = parseNumber(toIn.value, A.LOCALE);
      if (target === null) { toast('To value must be a number.', 'error'); toIn.focus(); return false; }
      const b = resolveArea(A.engine, A.view().names, byF.input.value, base);
      if (!b || b.r1 !== b.r2 || b.c1 !== b.c2) { toast('By changing cell must be a reference to a single cell.', 'error'); return false; }
      if (isFormula(A.engine.raw(b.sheet, b.r1, b.c1))) { toast('By changing cell must contain a value, not a formula.', 'error'); return false; }
      setTimeout(() => runGoalSeek(A, a, target, b), 0);
    } }
  ], { sticky: true });
}

async function runGoalSeek(A, setA, target, byA) {
  const snap = snapshot(A);
  const ssh = snap.sheet(setA.sheet.id), bsh = snap.sheet(byA.sheet.id);
  const key = M.keyAt(bsh, byA.r1, byA.c1);
  const x0v = num(snap.se.valueAt(bsh, byA.r1, byA.c1));
  const x0 = isFinite(x0v) ? x0v : 0;
  const F = (x) => {
    snap.setRaw(bsh, key, numToText(tidy(x)));
    snap.se.invalidateValues();
    return num(snap.se.valueAt(ssh, setA.r1, setA.c1)) - target;
  };
  let stop = false;
  const status = h('p', { class: 'dlg-text' }, 'Goal seeking…');
  const prog = dialog('Goal Seek status', status, [{ label: 'Stop', action: () => { stop = true; return false; } }], { sticky: true, onClose: () => { stop = true; } });
  const best = await seekRoot(F, x0, () => stop);
  prog.close();
  const setName = addr(setA.sheet, setA.r1, setA.c1).replace(/\$/g, '');
  const tol = 0.001;
  const found = isFinite(best.f) && Math.abs(best.f) <= tol;
  F(best.x);
  const cur = snap.se.valueAt(ssh, setA.r1, setA.c1);
  const st = (setA.sheet.cells[M.keyAt(setA.sheet, setA.r1, setA.c1)] || {}).st;
  dialog('Goal Seek status', h('div', { class: 'form' },
    h('p', { class: 'dlg-text' }, `Goal seeking with cell ${setName} ${found ? 'found a solution.' : 'may not have found a solution.'}`),
    h('div', { class: 'gs-values' },
      h('span', null, 'Target value:'), h('span', { class: 'num' }, fmtNum(target, A.LOCALE)),
      h('span', null, 'Current value:'), h('span', { class: 'num' }, formatValue(typeof cur === 'number' ? tidy(cur) : cur, st, A.LOCALE)))), [
    { label: 'Cancel' },
    { label: 'OK', primary: true, action: () => {
      const live = sheetById(A, byA.sheet.id);
      if (!live) return;
      A.commit([{ t: 'set', s: live.id, k: key, v: numToText(tidy(best.x)) }], { label: 'Goal Seek' });
    } }
  ], { sticky: true });
}

/** Finds x with F(x) = 0: damped Newton steps, then a search for a sign change and bisection. */
export async function seekRoot(F, x0, shouldStop = () => false, maxIter = 100) {
  let best = { x: x0, f: F(x0) };
  const consider = (x, f) => { if (isFinite(f) && (!isFinite(best.f) || Math.abs(f) < Math.abs(best.f))) best = { x, f }; };
  const tiny = (f) => Math.abs(f) <= 1e-12;
  let x = x0, fx = best.f, last = Date.now();
  const yieldNow = async () => { if (Date.now() - last > 40) { await new Promise((r) => setTimeout(r, 0)); last = Date.now(); } };
  for (let i = 0; i < maxIter && isFinite(fx) && !tiny(fx) && !shouldStop(); i++) {
    const hstep = Math.max(1e-7, Math.abs(x) * 1e-7);
    const d = (F(x + hstep) - fx) / hstep;
    if (!isFinite(d) || d === 0) break;
    let step = -fx / d, xn = x + step, fn = F(xn), k = 0;
    while ((!isFinite(fn) || Math.abs(fn) >= Math.abs(fx)) && k < 40) { step /= 2; xn = x + step; fn = F(xn); k++; }
    consider(xn, fn);
    if (!isFinite(fn) || Math.abs(fn) >= Math.abs(fx)) break;
    x = xn; fx = fn;
    await yieldNow();
  }
  if (tiny(best.f) || shouldStop()) return best;
  // walk outwards from the start in growing steps until the sign changes, then bisect
  const f0 = F(x0);
  const prev = { 1: { x: x0, f: f0 }, '-1': { x: x0, f: f0 } };
  for (let k = 0; k < 64 && !shouldStop(); k++) {
    const step = Math.max(1, Math.abs(x0)) * Math.pow(2, k - 10);
    for (const dir of [1, -1]) {
      const xb = x0 + dir * step, fb = F(xb);
      consider(xb, fb);
      const p = prev[dir];
      if (isFinite(p.f) && isFinite(fb) && fb !== 0 && Math.sign(p.f) !== Math.sign(fb)) {
        let lo = p.x, hi = xb, flo = p.f;
        for (let it = 0; it < 200 && !shouldStop(); it++) {
          const mid = (lo + hi) / 2;
          const fm = F(mid);
          consider(mid, fm);
          if (!isFinite(fm) || tiny(fm) || Math.abs(hi - lo) <= 1e-15 * Math.max(1, Math.abs(mid))) break;
          if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else hi = mid;
        }
        return best;
      }
      if (isFinite(fb)) prev[dir] = { x: xb, f: fb };
    }
    await yieldNow();
  }
  return best;
}
