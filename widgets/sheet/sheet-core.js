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
/**
 * Spreadsheets for YouTrack – workbook model and change operations (utility module).
 *
 * Shared by the backend (YouTrack JavaScript runtime, CommonJS) and the widget (browser, window.SheetCore).
 * build.py copies this file to widgets/sheet/sheet-core.js; edit only this copy.
 *
 * Workbook
 *   { v, rev, sheets: [Sheet], names: { NAME: "=Sheet1!$A$1:$B$4" }, createdAt, createdBy, updatedAt, updatedBy, updatedByName,
 *     base, sv }   (base: first revision of this workbook; sv: version of the cleaning rules it went through)
 * Sheet
 *   { id, name, rows: [rowId], cols: [colId], cells: { "rowId:colId": Cell }, colW: { colId: px }, rowH: { rowId: px },
 *     hr: { rowId: 1 } (hidden rows), hc: { colId: 1 } (hidden columns), merges: [{ r1, c1, r2, c2 }],
 *     freeze: { r, c }, color, nogrid, tables: [], cf: [], dv: [], charts: [] }
 * Cell
 *   { v: "raw input or =formula", st: { b, i, u, x, fc, bg, ha, va, w, fs, ff, nf, dp, cur, pat, bd, bc }, n: "note" }
 * Objects (tables, conditional formats, validation rules, charts) cover a range given by row and column ids:
 *   { id, r1, c1, r2, c2, ... }; charts also have an anchor cell { ar, ac }.
 *
 * Rows and columns have stable ids, so a cell edit lands on the right cell even when someone inserts or deletes rows
 * at the same time, and ranges grow when rows are inserted inside them. Formulas use A1 references; the client
 * rewrites them when rows or columns move and sends those rewrites in the same batch.
 *
 * Operations (a batch is applied in order, atomically):
 *   { t: 'set', s, k, v?, st?, n? }                 value, style and/or note of one cell (null clears)
 *   { t: 'sheet', s, name?, colW?, rowH?, hr?, hc?, merges?, freeze?, color?, nogrid? }
 *   { t: 'rows' | 'cols', s, at, add: [ids] }       insert (at = -1 appends)
 *   { t: 'rows' | 'cols', s, del: [ids] }           delete, with the cells in them
 *   { t: 'obj', s, kind, id, val }                  add, replace (val) or delete (val null) a table, rule or chart
 *   { t: 'name', name, ref }                        defined name (ref null deletes)
 *   { t: 'addSheet', at, sheet }  { t: 'delSheet', s }  { t: 'moveSheet', s, to }
 * Operations that point at a sheet, row or column that no longer exists are skipped.
 *
 * Everything that another person can store is rebuilt from known fields before it is stored or shown (cleanOps,
 * cleanObject, cleanSheet, cleanDoc): unknown fields are dropped, strings are limited, numbers are finite and in range,
 * colors are #rrggbb, ids match ID_RE and enumerations are checked. The server cleans every batch it accepts and
 * logs the cleaned batch; the widget cleans every batch and workbook it receives as well.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SheetCore = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  const LIMITS = {
    sheets: 50,
    rows: 20000,
    cols: 200,
    cellChars: 32767,       // same as Excel
    noteChars: 4000,
    nameChars: 31,          // same as Excel
    opsPerBatch: 50000,
    structuralOps: 20,      // row and column inserts and deletes per batch (the widget sends at most 3)
    batchChars: 12000000,   // serialized batch: the undo of a large deletion carries all of its cells, up to docChars
                            // of data plus the wrapping of each change
    objects: 200,           // per kind and sheet
    objectChars: 20000,
    formulaChars: 8192,     // values and formulas of rules, as long as an Excel formula
    merges: 5000,           // merged ranges per sheet (as in 1.2.5)
    names: 500,
    docChars: 4000000       // serialized workbook
  };
  const CLEAN_VERSION = 2;  // raise when the cleaning rules change, so stored workbooks are cleaned again
  const DEFAULT_ROWS = 100;
  const DEFAULT_COLS = 26;
  const ID_RE = /^[a-z0-9]{1,16}$/;
  const NAME_RE = /^[A-Za-z_À-￿][\w.À-￿]{0,63}$/;
  const OBJ_KINDS = ['tables', 'cf', 'dv', 'charts', 'solver', 'scenarios'];
  const STYLE_KEYS = {
    b: 'bool', i: 'bool', u: 'bool', x: 'bool', w: 'bool',
    fc: 'color', bg: 'color', bc: 'color',
    ha: ['left', 'center', 'right'], va: ['top', 'middle', 'bottom'],
    nf: ['gen', 'num', 'cur', 'acc', 'pct', 'date', 'ldate', 'time', 'dt', 'dur', 'text', 'sci', 'custom'],
    ff: ['sans', 'serif', 'mono'],
    fs: 'size', dp: 'int', cur: 'str', pat: 'pattern', bd: 'border'
  };
  const COLOR_RE = /^#[0-9a-fA-F]{6}$/;
  // the values the widget knows for each kind of object (see model.js, charts.js and whatif.js)
  const TABLE_STYLES = ['blue', 'teal', 'green', 'orange', 'red', 'purple', 'gray', 'none'];
  const FILTER_OPS = ['contains', 'notContains', 'begins', 'ends', 'eq', 'ne', 'gt', 'lt', 'ge', 'le', 'blank', 'notBlank'];
  const CF_TYPES = ['gt', 'lt', 'ge', 'le', 'eq', 'ne', 'between', 'notBetween', 'contains', 'notContains', 'begins', 'ends',
    'blank', 'notBlank', 'dup', 'unique', 'top', 'bottom', 'aboveAvg', 'belowAvg', 'today', 'past', 'future', 'last7',
    'thisMonth', 'formula', 'scale2', 'scale3', 'bar', 'icons'];
  const DV_TYPES = ['list', 'checkbox', 'number', 'int', 'date', 'length', 'formula'];
  const DV_OPS = ['between', 'notBetween', 'eq', 'ne', 'gt', 'lt', 'ge', 'le'];
  const CHART_TYPES = ['column', 'stacked', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];
  const SOLVER_OPS = ['le', 'eq', 'ge', 'int', 'bin', 'dif'];
  const ERROR_CODES = ['#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#CIRC!', '#ERROR!'];
  // Solver options: [default, min, max, whole number] for numbers; booleans and choices by their default
  const SOLVER_OPTIONS = {
    precision: [1e-6, 1e-12, 0.999999, false], scaling: true, showIter: false, ignoreInt: false,
    intTolerance: [1, 0, 100, false], maxTime: [0, 0, 1e7, false], iterations: [0, 0, 1e9, true],
    maxSubproblems: [0, 0, 1e9, true], maxFeasible: [0, 0, 1e9, true], convergence: [1e-4, 1e-12, 0.999999, false],
    derivatives: ['forward', 'central'], multistart: false, population: [100, 10, 1000, true],
    seed: [0, 0, 2147483647, true], requireBounds: true, mutation: [0.075, 0.0001, 0.999, false],
    maxNoImprove: [30, 1, 1e7, false]
  };

  function randomId() {
    let s = '';
    while (s.length < 10) s += Math.random().toString(36).slice(2);
    return s.slice(0, 10);
  }
  function ids(n) {
    const out = [];
    for (let i = 0; i < n; i++) out.push(randomId());
    return out;
  }

  function colName(index) {
    let n = index + 1, s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function newSheet(name, rowCount, colCount) {
    return {
      id: randomId(), name: name || 'Sheet1',
      rows: ids(rowCount || DEFAULT_ROWS), cols: ids(colCount || DEFAULT_COLS),
      cells: {}, colW: {}, rowH: {}, hr: {}, hc: {}, merges: [], freeze: { r: 0, c: 0 }, color: null, nogrid: false,
      tables: [], cf: [], dv: [], charts: [], solver: [], scenarios: []
    };
  }

  function newWorkbook(user, sheets) {
    const now = new Date().toISOString();
    return {
      v: 1, rev: 0,
      sheets: sheets && sheets.length ? sheets : [newSheet('Sheet1')],
      names: {},
      createdAt: now, createdBy: user ? user.login : null,
      updatedAt: now, updatedBy: user ? user.login : null, updatedByName: user ? user.name : null,
      sv: CLEAN_VERSION
    };
  }

  /** Fills in properties added in later versions, so old workbooks keep working. */
  function upgrade(doc) {
    if (!doc) return doc;
    if (!doc.names || typeof doc.names !== 'object') doc.names = {};
    doc.sheets.forEach(function (s) {
      ['colW', 'rowH', 'hr', 'hc', 'cells'].forEach(function (k) { if (!s[k] || typeof s[k] !== 'object') s[k] = {}; });
      ['merges'].concat(OBJ_KINDS).forEach(function (k) { if (!Array.isArray(s[k])) s[k] = []; });
      if (!s.freeze) s.freeze = { r: 0, c: 0 };
    });
    return doc;
  }

  function findSheet(doc, id) {
    for (let i = 0; i < doc.sheets.length; i++) if (doc.sheets[i].id === id) return doc.sheets[i];
    return null;
  }

  function cellLabel(sheet, key) {
    const p = key.split(':');
    const r = sheet.rows.indexOf(p[0]), c = sheet.cols.indexOf(p[1]);
    if (r < 0 || c < 0) return '?';
    return colName(c) + (r + 1);
  }

  // ------------------------------------------------------------ validation helpers

  function fail(msg) { const e = new Error(msg); e.code = 400; throw e; }

  function own(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }
  function isObj(v) { return !!v && typeof v === 'object' && !Array.isArray(v); }
  /** A sheet, row, column or object id. "constructor" is refused: plain objects keyed by ids would find Object.prototype.constructor. */
  function isId(v) { return typeof v === 'string' && ID_RE.test(v) && v !== 'constructor'; }
  function rid(v) { return isId(v) ? v : undefined; }
  /** A defined name: it must match the name rule also in capitals, the form it is stored under ("ß" becomes "SS"). */
  function validName(n) { return typeof n === 'string' && NAME_RE.test(n) && NAME_RE.test(n.toUpperCase()); }
  function str(v, max) {
    if (typeof v === 'string') return v.length > max ? v.slice(0, max) : v;
    if (typeof v === 'number' && isFinite(v)) return String(v);
    return undefined;
  }
  /** Text of a field that the widget evaluates or compares (values, formulas, references): kept whole, or dropped when it
   * is longer than an Excel formula (cutting it would change its meaning). */
  function text(v) {
    const s = typeof v === 'string' ? v : typeof v === 'number' && isFinite(v) ? String(v) : undefined;
    return s !== undefined && s.length <= LIMITS.formulaChars ? s : undefined;
  }
  /** A stored or sent text, number or truth value as text; objects and arrays count as absent. They are never converted:
   * an object can carry its own toString, and converting it may fail. */
  function scalar(v) { return typeof v === 'string' ? v : (typeof v === 'number' && isFinite(v)) || typeof v === 'boolean' ? String(v) : undefined; }
  /** A number from a number or numeric text; anything else is NaN (objects are never converted, as in scalar). */
  function num(v) { return typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN; }
  function bool(v) { return v === true || v === 1 ? true : v === false || v === 0 ? false : undefined; }
  function oneOf(v, list) { return typeof v === 'string' && list.indexOf(v) >= 0 ? v : undefined; }
  function color(v) { return typeof v === 'string' && COLOR_RE.test(v) ? v.toLowerCase() : undefined; }
  function clampInt(v, lo, hi) {
    const n = typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN;
    return isFinite(n) ? Math.max(lo, Math.min(hi, Math.round(n))) : undefined;
  }
  function put(out, k, v) { if (v !== undefined) out[k] = v; }
  /** A map keyed by row or column ids; fn cleans a value (undefined drops the entry). */
  function idMap(m, fn, max) {
    const out = {};
    if (!isObj(m)) return out;
    let n = 0;
    const keys = Object.keys(m);
    for (let i = 0; i < keys.length && n < max; i++) {
      if (!isId(keys[i])) continue;
      const v = fn(m[keys[i]]);
      if (v !== undefined) { out[keys[i]] = v; n++; }
    }
    return out;
  }

  function cleanStyle(st) {
    if (st === null || st === undefined) return null;
    if (typeof st !== 'object') fail('Invalid style');
    const out = {};
    Object.keys(st).forEach(function (k) {
      if (!own(STYLE_KEYS, k)) return;
      const rule = STYLE_KEYS[k], v = st[k];
      if (v === null || v === undefined || v === false || v === '') return;
      if (rule === 'bool') { if (v) out[k] = 1; }
      else if (rule === 'color') { const c = color(v); if (c) out[k] = c; }
      else if (rule === 'int') { const n = Math.round(num(v)); if (n >= 0 && n <= 15) out[k] = n; }
      else if (rule === 'size') { const n = Math.round(num(v)); if (n >= 6 && n <= 72) out[k] = n; }
      // currency symbol: shown as text, and the toolbar puts it into a CSS custom property, so no quotes, brackets or ;
      else if (rule === 'str') { const s = (scalar(v) || '').replace(/[\u0000-\u001f\u007f"'`\\;:{}()<>[\]]/g, '').slice(0, 8); if (s) out[k] = s; }
      else if (rule === 'pattern') { if (typeof v === 'string') out[k] = v.slice(0, 80); }
      else if (rule === 'border') { const b = (scalar(v) || '').replace(/[^tlrbTLRB]/g, '').slice(0, 8); if (b) out[k] = b; }
      else if (Array.isArray(rule)) { if (rule.indexOf(v) >= 0) out[k] = v; }
    });
    return Object.keys(out).length ? out : null;
  }

  function checkId(id, what) {
    if (!isId(id)) fail('Invalid ' + what + ' id');
    return id;
  }

  function cellKey(k) {
    const p = typeof k === 'string' ? k.split(':') : [];
    if (p.length !== 2 || !isId(p[0]) || !isId(p[1])) fail('Invalid cell');
    return k;
  }

  function cleanName(name) {
    const n = String(typeof name === 'string' || typeof name === 'number' ? name : '')
      .replace(/[\u0000-\u001f\u007f\\\/?*\[\]:]/g, '').trim().slice(0, LIMITS.nameChars);
    if (!n) fail('Sheet name is empty');
    return n;
  }

  /** Row and column id lists: ids only, without duplicates. */
  function idList(list) {
    const out = [], seen = new Set();
    if (!Array.isArray(list)) return out;
    for (let i = 0; i < list.length; i++) if (isId(list[i]) && !seen.has(list[i])) { seen.add(list[i]); out.push(list[i]); }
    return out;
  }

  // ------------------------------------------------------------ objects (tables, rules, charts, Solver models, scenarios)

  const RANGE = ['r1', 'c1', 'r2', 'c2'];
  function rangeOf(out, v, keys) { keys.forEach(function (k) { put(out, k, rid(v[k])); }); }
  function formulaText(v) { return typeof v === 'string' && v.charAt(0) === '=' ? v : undefined; }

  function cleanFilter(spec) {
    if (!isObj(spec)) return undefined;
    const out = {};
    if (Array.isArray(spec.vals)) out.vals = spec.vals.filter(function (s) { return typeof s === 'string'; });
    if (isObj(spec.cond)) {
      const op = oneOf(spec.cond.op, FILTER_OPS);
      if (op) out.cond = { op: op, a: str(spec.cond.a, 1000) || '' };
    }
    return out.vals || out.cond ? out : undefined;
  }

  const CLEANERS = {
    tables: function (out, v) {
      const name = typeof v.name === 'string' ? v.name : '';
      if (!NAME_RE.test(name)) fail('Table names start with a letter and contain only letters, digits, "_" and "."');
      out.name = name;
      rangeOf(out, v, RANGE);
      ['header', 'totals', 'banded', 'filter'].forEach(function (k) { put(out, k, bool(v[k])); });
      put(out, 'style', oneOf(v.style, TABLE_STYLES));
      if (isObj(v.filters)) out.filters = idMap(v.filters, cleanFilter, LIMITS.cols);
    },
    cf: function (out, v) {
      const type = oneOf(v.type, CF_TYPES);
      if (!type) fail('Unknown conditional formatting rule');
      rangeOf(out, v, RANGE);
      out.type = type;
      put(out, 'a', text(v.a));
      put(out, 'b', text(v.b));
      put(out, 'pct', bool(v.pct));
      if (type === 'formula') put(out, 'formula', text(v.formula));
      if (v.style === null) out.style = null;
      else if (isObj(v.style)) {
        const st = {};
        put(st, 'bg', color(v.style.bg));
        put(st, 'fc', color(v.style.fc));
        ['b', 'i', 'x'].forEach(function (k) { if (v.style[k]) st[k] = 1; });
        out.style = st;
      }
      if (type === 'scale2' || type === 'scale3') {
        const n = type === 'scale2' ? 2 : 3;
        const cols = Array.isArray(v.colors) && v.colors.length === n ? v.colors.map(color) : [];
        if (cols.length === n && cols.every(function (c) { return !!c; })) out.colors = cols;
      }
      if (type === 'bar') put(out, 'color', color(v.color));
      if (v.stop === true) out.stop = true;
    },
    dv: function (out, v) {
      const type = oneOf(v.type, DV_TYPES);
      if (!type) fail('Unknown data validation rule');
      rangeOf(out, v, RANGE);
      out.type = type;
      put(out, 'strict', bool(v.strict));
      if (v.msg === null) out.msg = null; else put(out, 'msg', str(v.msg, 1000));
      if (type === 'list') {
        // the items of a list are only split, never translated: they may be as long as the object allows (a list
        // whose items were dropped would refuse every entry)
        put(out, 'source', scalar(v.source));
        if (v.arrows === 'always') out.arrows = 'always';
      } else if (type === 'formula') put(out, 'formula', text(v.formula));
      else if (type !== 'checkbox') {
        put(out, 'op', oneOf(v.op, DV_OPS));
        put(out, 'a', text(v.a));
        put(out, 'b', text(v.b));
      }
    },
    charts: function (out, v) {
      out.type = oneOf(v.type, CHART_TYPES) || 'column';
      put(out, 'title', str(v.title, 1000));
      put(out, 'legend', bool(v.legend));
      rangeOf(out, v, RANGE.concat(['ar', 'ac']));
      put(out, 'ox', clampInt(v.ox, 0, 10000));
      put(out, 'oy', clampInt(v.oy, 0, 10000));
      put(out, 'w', clampInt(v.w, 40, 10000));
      put(out, 'h', clampInt(v.h, 40, 10000));
    },
    solver: function (out, v, id) {
      if (id !== 'model') fail('Unknown Solver model');
      out.obj = typeof v.obj === 'string' && (v.obj === '' || v.obj.charAt(0) === '=') ? v.obj : '';
      out.goal = oneOf(v.goal, ['max', 'min', 'value']) || 'max';
      out.value = typeof v.value === 'string' ? v.value.slice(0, 255) : '0';
      out.vars = (Array.isArray(v.vars) ? v.vars : []).filter(function (x) { return !!formulaText(x); }).slice(0, 1000);
      out.cons = (Array.isArray(v.cons) ? v.cons : []).map(function (c) {
        if (!isObj(c)) return null;
        const op = oneOf(c.op, SOLVER_OPS), l = formulaText(c.l);
        if (!op || !l) return null;
        const r = op === 'int' || op === 'bin' || op === 'dif' ? '' : formulaText(c.r) || '';
        return { l: l, op: op, r: r };
      }).filter(Boolean).slice(0, 1000);
      out.nonneg = v.nonneg !== false;
      out.method = oneOf(v.method, ['grg', 'lp', 'evo']) || 'grg';
      if (isObj(v.opts)) {
        const o = {};
        Object.keys(SOLVER_OPTIONS).forEach(function (k) {
          const rule = SOLVER_OPTIONS[k], x = v.opts[k];
          if (typeof rule === 'boolean') o[k] = typeof x === 'boolean' ? x : rule;
          else if (typeof rule[0] === 'string') o[k] = rule.indexOf(x) >= 0 ? x : rule[0];
          else o[k] = typeof x === 'number' && isFinite(x) && x >= rule[1] && x <= rule[2] && (!rule[3] || Math.round(x) === x) ? x : rule[0];
        });
        out.opts = o;
      }
    },
    scenarios: function (out, v) {
      const cells = [], values = [];
      const vc = Array.isArray(v.cells) ? v.cells : [], vv = Array.isArray(v.values) ? v.values : [];
      for (let i = 0; i < vc.length && cells.length < 32; i++) {
        if (!formulaText(vc[i]) || vc[i].length > 255) continue;
        cells.push(vc[i]);
        const x = vv[i];
        values.push(typeof x === 'number' ? (isFinite(x) ? x : null)
          : typeof x === 'string' ? x.slice(0, LIMITS.cellChars)
          : typeof x === 'boolean' ? x
          : isObj(x) && ERROR_CODES.indexOf(x.code) >= 0 ? x.code : null);
      }
      if (!cells.length) fail('A scenario needs at least one cell');
      const name = typeof v.name === 'string' ? v.name.slice(0, 255) : '';
      out.name = name.trim() ? name : 'Scenario';
      out.cells = cells;
      out.values = values;
      out.comment = typeof v.comment === 'string' ? v.comment.slice(0, 1000) : '';
    }
  };

  /**
   * Rebuilds a stored object from the fields its kind has (fixed key order, so cleaning is idempotent). Fails only
   * when the object cannot be used at all; fields with bad values are dropped.
   */
  /** Size of an object without its row and column ids: deleting rows or columns moves its corners to other ids (longer
   * or shorter), and that must not push a stored object over the limit, which would drop it on the next cleaning. */
  const ID_FIELDS = ['r1', 'c1', 'r2', 'c2', 'ar', 'ac'];
  function contentSize(val) {
    // the object's own id is left out as well: a change sends the object without it, the cleaned copy has it
    let n = JSON.stringify(own(val, 'id') ? Object.assign({}, val, { id: undefined }) : val).length;
    ID_FIELDS.forEach(function (k) { if (typeof val[k] === 'string' && ID_RE.test(val[k])) n -= val[k].length; });
    return n;
  }

  function cleanObject(kind, id, val) {
    if (OBJ_KINDS.indexOf(kind) < 0) fail('Unknown object kind');
    checkId(id, 'object');
    if (val === null) return null;
    if (!isObj(val)) fail('Invalid object');
    if (contentSize(val) > LIMITS.objectChars) fail('The object is too large');
    const out = { id: id };
    CLEANERS[kind](out, val, id);
    return out;
  }

  /** A short label of an object for the change history. */
  function labelOf(o) {
    if (!o) return null;
    const s = typeof o.name === 'string' ? o.name : typeof o.title === 'string' ? o.title : '';
    return s ? s.slice(0, 100) : null;
  }

  /**
   * Merged ranges: only ranges whose ids exist on the sheet, larger than one cell and not overlapping an earlier one,
   * stored top-left to bottom-right.
   */
  function cleanMerges(sheet, list) {
    const rows = new Map(), cols = new Map(), kept = [], rects = [];
    sheet.rows.forEach(function (id, i) { rows.set(id, i); });
    sheet.cols.forEach(function (id, i) { cols.set(id, i); });
    for (let i = 0; i < list.length && kept.length < LIMITS.merges; i++) {
      const m = list[i];
      if (!isObj(m)) continue;
      const a = rows.get(m.r1), b = rows.get(m.r2), c = cols.get(m.c1), d = cols.get(m.c2);
      if (a === undefined || b === undefined || c === undefined || d === undefined) continue;
      const R = { r1: Math.min(a, b), r2: Math.max(a, b), c1: Math.min(c, d), c2: Math.max(c, d) };
      if (R.r1 === R.r2 && R.c1 === R.c2) continue;
      let overlaps = false;
      for (let j = 0; j < rects.length && !overlaps; j++) {
        const o = rects[j];
        overlaps = o.r1 <= R.r2 && R.r1 <= o.r2 && o.c1 <= R.c2 && R.c1 <= o.c2;
      }
      if (overlaps) continue;
      rects.push(R);
      kept.push({ r1: sheet.rows[R.r1], c1: sheet.cols[R.c1], r2: sheet.rows[R.r2], c2: sheet.cols[R.c2] });
    }
    return kept;
  }

  /** A sheet from an imported file, a copy or an undone deletion. Objects that cannot be used are left out. */
  function cleanSheet(sh) {
    if (!isObj(sh)) fail('Invalid sheet');
    const rows = idList(sh.rows), cols = idList(sh.cols);
    if (!rows.length || !cols.length) fail('A sheet needs rows and columns');
    if (rows.length > LIMITS.rows || cols.length > LIMITS.cols) fail('The sheet is too large (at most ' + LIMITS.rows + ' rows and ' + LIMITS.cols + ' columns)');
    // a name that is empty once cleaned gets a default one (cleanDoc numbers duplicates): the sheet's data is kept
    let name = 'Sheet';
    try { name = cleanName(sh.name); } catch (e) { /* keep the default */ }
    const out = newSheet(name, 1, 1);
    out.id = checkId(sh.id, 'sheet');
    out.rows = rows;
    out.cols = cols;
    const rowSet = new Set(rows), colSet = new Set(cols);
    const cells = isObj(sh.cells) ? sh.cells : {};
    Object.keys(cells).forEach(function (k) {
      const p = k.split(':');
      if (p.length !== 2 || !rowSet.has(p[0]) || !colSet.has(p[1])) return;
      const c = cells[k];
      if (!isObj(c)) return;
      // a value, style or note that is not valid is dropped on its own; the rest of the sheet is kept
      const cell = {};
      const v = scalar(c.v);
      if (v !== undefined && v !== '') cell.v = v.slice(0, LIMITS.cellChars);
      const st = isObj(c.st) ? cleanStyle(c.st) : null;
      if (st) cell.st = st;
      const n = scalar(c.n);
      if (n) cell.n = n.slice(0, LIMITS.noteChars);
      if (cell.v !== undefined || cell.st || cell.n) out.cells[k] = cell;
    });
    applySheetProps(out, { colW: sh.colW, rowH: sh.rowH, hr: sh.hr, hc: sh.hc, merges: sh.merges, freeze: sh.freeze, color: sh.color, nogrid: sh.nogrid }, { rows: rowSet, cols: colSet });
    OBJ_KINDS.forEach(function (kind) {
      const seen = new Set();
      (Array.isArray(sh[kind]) ? sh[kind] : []).forEach(function (o) {
        if (out[kind].length >= LIMITS.objects || !isObj(o) || !isId(o.id) || seen.has(o.id)) return;
        let clean = null;
        try { clean = cleanObject(kind, o.id, o); } catch (e) { clean = null; }
        if (clean) { seen.add(o.id); out[kind].push(clean); }
      });
    });
    return out;
  }

  /**
   * Applies sheet properties. Sizes and hidden flags are kept only for rows and columns the sheet has (`known`: Sets of
   * its row and column ids, built here when not given), so entries for other ids cannot fill the workbook.
   */
  function applySheetProps(sheet, op, known) {
    if (op.name !== undefined) sheet.name = cleanName(op.name);
    if (op.color !== undefined) sheet.color = color(op.color) || null;
    if (op.nogrid !== undefined) sheet.nogrid = !!op.nogrid;
    const ids = function (prop) {
      if (!known) known = { rows: new Set(sheet.rows), cols: new Set(sheet.cols) };
      return prop === 'colW' || prop === 'hc' ? known.cols : known.rows;
    };
    ['colW', 'rowH'].forEach(function (prop) {
      const patch = op[prop];
      if (!isObj(patch)) return;
      Object.keys(patch).forEach(function (id) {
        if (!isId(id)) return;
        const v = patch[id];
        if (v === null || v === undefined) delete sheet[prop][id];
        else if (ids(prop).has(id)) { const n = Math.round(num(v)); if (n >= 2 && n <= 2000) sheet[prop][id] = n; }
      });
    });
    ['hr', 'hc'].forEach(function (prop) {
      const patch = op[prop];
      if (!isObj(patch)) return;
      Object.keys(patch).forEach(function (id) {
        if (!isId(id)) return;
        if (!patch[id]) delete sheet[prop][id];
        else if (ids(prop).has(id)) sheet[prop][id] = 1;
      });
    });
    if (Array.isArray(op.merges)) sheet.merges = cleanMerges(sheet, op.merges);
    if (isObj(op.freeze)) {
      sheet.freeze = {
        r: Math.max(0, Math.min(50, Math.round(num(op.freeze.r) || 0))),
        c: Math.max(0, Math.min(26, Math.round(num(op.freeze.c) || 0)))
      };
    }
  }

  // ------------------------------------------------------------ cleaning of batches and workbooks

  /** A cleaned copy of one operation with only the fields its type has. Fails when the operation is not valid. */
  function cleanOp(op) {
    if (!isObj(op)) fail('Invalid change');
    switch (op.t) {
      case 'set': {
        const o = { t: 'set', s: checkId(op.s, 'sheet'), k: cellKey(op.k) };
        if (op.v !== undefined) {
          if (op.v === null || op.v === '') o.v = null;
          else { const v = scalar(op.v); if (v === undefined) fail('Invalid value'); o.v = v.slice(0, LIMITS.cellChars); }
        }
        if (op.st !== undefined) o.st = cleanStyle(op.st);
        if (op.n !== undefined) {
          if (!op.n) o.n = null;
          else { const n = scalar(op.n); if (n === undefined) fail('Invalid note'); o.n = n.slice(0, LIMITS.noteChars); }
        }
        return o;
      }
      case 'sheet': {
        const o = { t: 'sheet', s: checkId(op.s, 'sheet') };
        if (op.name !== undefined) o.name = cleanName(op.name);
        if (op.color !== undefined) o.color = color(op.color) || null;
        if (op.nogrid !== undefined) o.nogrid = !!op.nogrid;
        const size = function (v) {
          if (v === null) return null;
          const n = Math.round(num(v));
          return n >= 2 && n <= 2000 ? n : undefined;
        };
        if (op.colW !== undefined) o.colW = idMap(op.colW, size, LIMITS.cols * 2);
        if (op.rowH !== undefined) o.rowH = idMap(op.rowH, size, LIMITS.rows * 2);
        if (op.hc !== undefined) o.hc = idMap(op.hc, function (v) { return v ? 1 : 0; }, LIMITS.cols * 2);
        if (op.hr !== undefined) o.hr = idMap(op.hr, function (v) { return v ? 1 : 0; }, LIMITS.rows * 2);
        if (Array.isArray(op.merges)) {
          o.merges = op.merges.slice(0, LIMITS.merges * 2).filter(isObj).map(function (m) {
            return { r1: rid(m.r1), c1: rid(m.c1), r2: rid(m.r2), c2: rid(m.c2) };
          }).filter(function (m) { return m.r1 && m.c1 && m.r2 && m.c2; });
        }
        if (isObj(op.freeze)) o.freeze = { r: clampInt(op.freeze.r, 0, 50) || 0, c: clampInt(op.freeze.c, 0, 26) || 0 };
        return o;
      }
      case 'rows':
      case 'cols': {
        const o = { t: op.t, s: checkId(op.s, 'sheet') };
        const max = op.t === 'rows' ? LIMITS.rows : LIMITS.cols;
        const what = op.t === 'rows' ? ' rows' : ' columns';
        if (Array.isArray(op.add)) {
          if (op.add.length > max) fail('A sheet can have at most ' + max + what);
          o.add = idList(op.add);
          if (op.at !== undefined) o.at = op.at === -1 ? -1 : clampInt(op.at, 0, max) || 0;
        } else if (Array.isArray(op.del)) {
          if (op.del.length > max) fail('A sheet can have at most ' + max + what);
          o.del = idList(op.del);
        } else fail('Invalid change');
        return o;
      }
      case 'obj': {
        if (OBJ_KINDS.indexOf(op.kind) < 0) fail('Unknown object kind');
        const id = checkId(op.id, 'object');
        return { t: 'obj', s: checkId(op.s, 'sheet'), kind: op.kind, id: id, val: op.val === null ? null : cleanObject(op.kind, id, op.val) };
      }
      case 'name': {
        if (!validName(op.name)) fail('Names start with a letter and contain only letters, digits, "_" and "."');
        if (op.ref === null || op.ref === undefined || op.ref === '') return { t: 'name', name: op.name, ref: null };
        const ref = scalar(op.ref);
        if (ref === undefined) fail('Invalid reference');
        return { t: 'name', name: op.name, ref: ref.slice(0, 1000) };
      }
      case 'addSheet': {
        const o = { t: 'addSheet' };
        if (op.at !== undefined) o.at = clampInt(op.at, 0, LIMITS.sheets) || 0;
        o.sheet = cleanSheet(op.sheet);
        return o;
      }
      case 'delSheet': return { t: 'delSheet', s: checkId(op.s, 'sheet') };
      case 'moveSheet': return { t: 'moveSheet', s: checkId(op.s, 'sheet'), to: clampInt(op.to, 0, LIMITS.sheets) || 0 };
      default: fail('Unknown change: ' + (scalar(op.t) || '?').slice(0, 20));
    }
  }

  /**
   * Cleans a batch before it is applied or stored. lenient: leaves out operations that are not valid instead of
   * failing (for batches the server accepted earlier, possibly from an older version).
   */
  function cleanOps(ops, lenient) {
    if (!Array.isArray(ops)) { if (lenient) return []; fail('Invalid change'); }
    if (ops.length > LIMITS.opsPerBatch) { if (lenient) return []; fail('Too many changes at once'); }
    const out = [];
    let structural = 0;
    for (let i = 0; i < ops.length; i++) {
      let op = null;
      try { op = cleanOp(ops[i]); } catch (e) { if (!lenient) throw e; }
      if (!op) continue;
      if ((op.t === 'rows' || op.t === 'cols') && ++structural > LIMITS.structuralOps && !lenient) fail('Too many changes at once');
      out.push(op);
    }
    return out;
  }

  /** Fails when sheet ids, sheet names or table names are not unique in the workbook. */
  function checkWorkbook(doc) {
    const sheetIds = new Set(), names = new Set(), tables = new Set();
    doc.sheets.forEach(function (sh) {
      if (sheetIds.has(sh.id)) fail('Two sheets have the same id');
      sheetIds.add(sh.id);
      const name = scalar(sh.name) || '', n = name.toLowerCase();
      if (names.has(n)) fail('Another sheet is already named "' + name + '"');
      names.add(n);
      (Array.isArray(sh.tables) ? sh.tables : []).forEach(function (t) {
        const tname = scalar(t && t.name) || '', tn = tname.toLowerCase();
        if (tables.has(tn)) fail('Another table is already named "' + tname + '"');
        tables.add(tn);
      });
    });
  }

  /**
   * A cleaned copy of a whole workbook (stored by any version, or received from the server). Never fails: sheets and
   * objects that cannot be used are left out, and duplicate sheet and table names get a number.
   */
  function cleanDoc(doc) {
    if (!isObj(doc) || !Array.isArray(doc.sheets)) return null;
    const out = {
      v: 1, rev: typeof doc.rev === 'number' && isFinite(doc.rev) && doc.rev >= 0 ? doc.rev : 0,
      sheets: [], names: {},
      createdAt: str(doc.createdAt, 40) || null, createdBy: str(doc.createdBy, 200) || null,
      updatedAt: str(doc.updatedAt, 40) || null, updatedBy: str(doc.updatedBy, 200) || null,
      updatedByName: str(doc.updatedByName, 200) || null
    };
    // the revision the workbook was created with: never above its current revision (a broken one would make every
    // browser think the workbook was deleted)
    if (typeof doc.base === 'number' && isFinite(doc.base) && doc.base >= 0 && doc.base <= out.rev) out.base = doc.base;
    const sheetIds = new Set(), names = new Set(), tables = new Set();
    doc.sheets.forEach(function (sh) {
      if (out.sheets.length >= LIMITS.sheets) return;
      let c = null;
      try { c = cleanSheet(sh); } catch (e) { return; }
      if (sheetIds.has(c.id)) return;
      let name = c.name, n = 2;
      while (names.has(name.toLowerCase())) name = c.name.slice(0, LIMITS.nameChars - 5) + ' (' + (n++) + ')';
      c.name = name;
      c.tables.forEach(function (t) {
        let tn = t.name, k = 2;
        while (tables.has(tn.toLowerCase())) tn = t.name.slice(0, 58) + '_' + (k++);
        t.name = tn;
        tables.add(tn.toLowerCase());
      });
      sheetIds.add(c.id);
      names.add(name.toLowerCase());
      out.sheets.push(c);
    });
    if (!out.sheets.length) out.sheets.push(newSheet('Sheet1'));
    if (isObj(doc.names)) {
      let count = 0;
      Object.keys(doc.names).forEach(function (k) {
        if (count >= LIMITS.names || !validName(k) || typeof doc.names[k] !== 'string' || !doc.names[k]) return;
        const key = k.toUpperCase();
        if (own(out.names, key)) return;
        out.names[key] = doc.names[k].slice(0, 1000);
        count++;
      });
    }
    out.sv = CLEAN_VERSION;
    return out;
  }

  /**
   * After rows or columns were deleted, moves range corners that pointed at them to the nearest remaining row or
   * column inside the range, and drops ranges that lost all of their rows or columns.
   */
  function remapRanges(sheet, axis, oldList) {
    const list = axis === 'rows' ? sheet.rows : sheet.cols;
    const alive = new Map();
    list.forEach(function (id, i) { alive.set(id, i); });
    const oldIndex = new Map();
    oldList.forEach(function (id, i) { oldIndex.set(id, i); });
    const k1 = axis === 'rows' ? 'r1' : 'c1', k2 = axis === 'rows' ? 'r2' : 'c2', ka = axis === 'rows' ? 'ar' : 'ac';
    function fix(o) {
      if (!o) return false;
      if (o[k1] === undefined && o[ka] === undefined) return true;
      if (o[k1] !== undefined) {
        let a = oldIndex.get(o[k1]), b = oldIndex.get(o[k2]);
        if (a === undefined || b === undefined) return false;
        if (a > b) { const t = a; a = b; b = t; }
        while (a <= b && !alive.has(oldList[a])) a++;
        while (b >= a && !alive.has(oldList[b])) b--;
        if (a > b) return false;
        o[k1] = oldList[a];
        o[k2] = oldList[b];
      }
      if (o[ka] !== undefined && !alive.has(o[ka])) {
        let i = oldIndex.get(o[ka]);
        if (i === undefined) i = 0;
        while (i >= 0 && !alive.has(oldList[i])) i--;
        o[ka] = i >= 0 ? oldList[i] : list[0];
      }
      return true;
    }
    // a merge that shrank to one cell is no merge any more (cleanMerges drops it too, so server and browsers agree)
    sheet.merges = sheet.merges.filter(function (m) { return fix(m) && !(m.r1 === m.r2 && m.c1 === m.c2); });
    OBJ_KINDS.forEach(function (kind) { sheet[kind] = sheet[kind].filter(fix); });
  }

  /** True when the batch must be based on the latest revision (it moves rows or columns). */
  function isStrict(ops) {
    return ops.some(function (op) {
      return !!op && (op.t === 'rows' || op.t === 'cols') && (op.del || (op.add && op.at !== -1));
    });
  }

  // ------------------------------------------------------------ apply

  /** True when two cells hold the same value, format and note (compared field by field: a batch can touch many cells,
   * and turning each one into text would cost far more). Format keys without a value do not count. */
  function sameCell(a, b) {
    if (!a || !b) return !a && !b;
    if (a.v !== b.v || a.n !== b.n) return false;
    const x = a.st || null, y = b.st || null;
    if (x === y) return true;
    const kx = x ? Object.keys(x).filter(function (k) { return x[k] !== undefined; }) : [];
    const ky = y ? Object.keys(y).filter(function (k) { return y[k] !== undefined; }) : [];
    if (kx.length !== ky.length) return false;
    for (let i = 0; i < kx.length; i++) if (x[kx[i]] !== y[kx[i]]) return false;
    return true;
  }

  /**
   * Applies a batch to the workbook (in place). Returns the changes, for the history:
   *   { kind: 'cell', s, k, from, to } | { kind: 'rows'|'cols', s, added|deleted, at } | { kind: 'obj', … } | …
   * With validate off (batches that went through cleanOps already, or the widget's own pending batches), values are
   * taken as they are.
   */
  function applyOps(doc, ops, options) {
    const validate = !options || options.validate !== false;
    if (!Array.isArray(ops)) fail('Invalid change');
    if (ops.length > LIMITS.opsPerBatch) fail('Too many changes at once');
    upgrade(doc);
    const changes = [];
    // row and column lookups for this batch; dropped when a sheet's rows or columns change
    const index = new Map();
    function lookup(sheet) {
      let e = index.get(sheet);
      if (!e) { e = { rows: new Set(sheet.rows), cols: new Set(sheet.cols) }; index.set(sheet, e); }
      return e;
    }
    ops.forEach(function (op) {
      if (!isObj(op)) { if (validate) fail('Invalid change'); return; }
      const sheet = typeof op.s === 'string' ? findSheet(doc, op.s) : null;
      switch (op.t) {
        case 'set': {
          if (!sheet || typeof op.k !== 'string') return;
          const p = op.k.split(':');
          const e = lookup(sheet);
          if (p.length !== 2 || !e.rows.has(p[0]) || !e.cols.has(p[1])) return;
          const before = sheet.cells[op.k] || null;
          const cell = {};
          if (before) { if (before.v !== undefined) cell.v = before.v; if (before.st) cell.st = before.st; if (before.n) cell.n = before.n; }
          if (op.v !== undefined) {
            if (op.v === null || op.v === '') delete cell.v;
            else cell.v = String(op.v).slice(0, LIMITS.cellChars);
          }
          if (op.st !== undefined) {
            const st = validate ? cleanStyle(op.st) : op.st;
            if (st) cell.st = st; else delete cell.st;
          }
          if (op.n !== undefined) {
            if (op.n) cell.n = String(op.n).slice(0, LIMITS.noteChars); else delete cell.n;
          }
          const after = cell.v !== undefined || cell.st || cell.n ? cell : null;
          if (sameCell(before, after)) return;
          if (after) sheet.cells[op.k] = after; else delete sheet.cells[op.k];
          changes.push({ kind: 'cell', s: sheet.id, k: op.k, from: before, to: after });
          return;
        }
        case 'sheet': {
          if (!sheet) return;
          const was = sheet.name;
          applySheetProps(sheet, op, lookup(sheet));
          changes.push({ kind: 'sheet', s: sheet.id, renamed: was !== sheet.name ? { from: was, to: sheet.name } : null });
          return;
        }
        case 'rows':
        case 'cols': {
          if (!sheet) return;
          const list = op.t === 'rows' ? sheet.rows : sheet.cols;
          const max = op.t === 'rows' ? LIMITS.rows : LIMITS.cols;
          if (Array.isArray(op.add)) {
            if (op.add.length > max) { if (validate) fail('A sheet can have at most ' + max + (op.t === 'rows' ? ' rows' : ' columns')); return; }
            const have = new Set(list), add = [];
            op.add.forEach(function (id) { if (isId(id) && !have.has(id)) { have.add(id); add.push(id); } });
            if (!add.length) return;
            if (list.length + add.length > max) fail('A sheet can have at most ' + max + (op.t === 'rows' ? ' rows' : ' columns'));
            const at = op.at === -1 || typeof op.at !== 'number' || !isFinite(op.at) || op.at > list.length ? list.length : Math.max(0, Math.round(op.at));
            const end = at === list.length;
            Array.prototype.splice.apply(list, [at, 0].concat(add));
            index.delete(sheet);
            changes.push({ kind: op.t, s: sheet.id, added: add.length, at: at, end: end });
          } else if (Array.isArray(op.del)) {
            if (op.del.length > max) { if (validate) fail('A sheet can have at most ' + max + (op.t === 'rows' ? ' rows' : ' columns')); return; }
            const del = new Set();
            op.del.forEach(function (id) { if (isId(id)) del.add(id); });
            const kept = [];
            let first = -1;
            list.forEach(function (id, i) { if (del.has(id)) { if (first < 0) first = i; } else kept.push(id); });
            const count = list.length - kept.length;
            if (!count) return;
            if (!kept.length) fail('A sheet needs at least one ' + (op.t === 'rows' ? 'row' : 'column'));
            const oldList = list.slice();
            list.length = 0;
            Array.prototype.push.apply(list, kept);
            index.delete(sheet);
            const part = op.t === 'rows' ? 0 : 1;
            Object.keys(sheet.cells).forEach(function (k) { if (del.has(k.split(':')[part])) delete sheet.cells[k]; });
            const sizes = op.t === 'rows' ? sheet.rowH : sheet.colW;
            const hidden = op.t === 'rows' ? sheet.hr : sheet.hc;
            del.forEach(function (id) { delete sizes[id]; delete hidden[id]; });
            remapRanges(sheet, op.t, oldList);
            changes.push({ kind: op.t, s: sheet.id, deleted: count, at: first });
          }
          return;
        }
        case 'obj': {
          if (!sheet) return;
          if (OBJ_KINDS.indexOf(op.kind) < 0) { if (validate) fail('Unknown object kind'); return; }
          const val = validate ? cleanObject(op.kind, op.id, op.val) : op.val;
          const listO = sheet[op.kind];
          const idx = listO.findIndex(function (o) { return o.id === op.id; });
          if (val === null || val === undefined) {
            if (idx < 0) return;
            const gone = listO.splice(idx, 1)[0];
            changes.push({ kind: 'obj', s: sheet.id, objKind: op.kind, action: 'delete', name: labelOf(gone) });
            return;
          }
          if (idx >= 0) {
            if (JSON.stringify(listO[idx]) === JSON.stringify(val)) return;
            listO[idx] = val;
          } else {
            if (listO.length >= LIMITS.objects) fail('Too many items of this kind on one sheet');
            listO.push(val);
          }
          changes.push({ kind: 'obj', s: sheet.id, objKind: op.kind, action: idx >= 0 ? 'update' : 'add', name: labelOf(val) });
          return;
        }
        case 'name': {
          if (!validName(op.name)) { if (validate) fail('Names start with a letter and contain only letters, digits, "_" and "."'); return; }
          const key = op.name.toUpperCase();
          const had = own(doc.names, key);
          if (op.ref === null || op.ref === undefined || op.ref === '') {
            if (!had) return;
            delete doc.names[key];
          } else {
            if (!had && Object.keys(doc.names).length >= LIMITS.names) fail('Too many names');
            doc.names[key] = String(op.ref).slice(0, 1000);
          }
          changes.push({ kind: 'name', name: key, action: op.ref ? (had ? 'update' : 'add') : 'delete' });
          return;
        }
        case 'addSheet': {
          if (doc.sheets.length >= LIMITS.sheets) fail('A workbook can have at most ' + LIMITS.sheets + ' sheets');
          const sh = validate ? cleanSheet(op.sheet) : op.sheet;
          if (!isObj(sh) || !Array.isArray(sh.rows) || !Array.isArray(sh.cols) || typeof sh.name !== 'string') return;
          if (findSheet(doc, sh.id)) return;
          let name = sh.name, n = 2;
          while (doc.sheets.some(function (x) { return x.name.toLowerCase() === name.toLowerCase(); })) {
            name = sh.name.slice(0, LIMITS.nameChars - 5) + ' (' + (n++) + ')';
          }
          sh.name = name;
          const at = typeof op.at !== 'number' || !isFinite(op.at) || op.at < 0 || op.at > doc.sheets.length ? doc.sheets.length : Math.round(op.at);
          doc.sheets.splice(at, 0, sh);
          changes.push({ kind: 'addSheet', s: sh.id, name: sh.name });
          return;
        }
        case 'delSheet': {
          if (!sheet) return;
          if (doc.sheets.length <= 1) fail('A workbook needs at least one sheet');
          doc.sheets.splice(doc.sheets.indexOf(sheet), 1);
          index.delete(sheet);
          changes.push({ kind: 'delSheet', s: sheet.id, name: sheet.name });
          return;
        }
        case 'moveSheet': {
          if (!sheet) return;
          const from = doc.sheets.indexOf(sheet);
          const to = Math.max(0, Math.min(doc.sheets.length - 1, Math.round(Number(op.to) || 0)));
          if (from === to) return;
          doc.sheets.splice(from, 1);
          doc.sheets.splice(to, 0, sheet);
          changes.push({ kind: 'moveSheet', s: sheet.id, name: sheet.name });
          return;
        }
        default:
          if (validate) fail('Unknown change: ' + (scalar(op.t) || '?').slice(0, 20));
      }
    });
    if (validate) checkWorkbook(doc);
    return changes;
  }

  return {
    LIMITS: LIMITS,
    CLEAN_VERSION: CLEAN_VERSION,
    DEFAULT_ROWS: DEFAULT_ROWS,
    DEFAULT_COLS: DEFAULT_COLS,
    OBJ_KINDS: OBJ_KINDS,
    NAME_RE: NAME_RE,
    TABLE_STYLES: TABLE_STYLES,
    randomId: randomId,
    ids: ids,
    isId: isId,
    validName: validName,
    colName: colName,
    cellLabel: cellLabel,
    newSheet: newSheet,
    newWorkbook: newWorkbook,
    upgrade: upgrade,
    findSheet: findSheet,
    cleanSheet: cleanSheet,
    cleanStyle: cleanStyle,
    cleanObject: cleanObject,
    cleanOps: cleanOps,
    cleanDoc: cleanDoc,
    checkWorkbook: checkWorkbook,
    isStrict: isStrict,
    applyOps: applyOps
  };
});
