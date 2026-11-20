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
 *   { v, rev, sheets: [Sheet], names: { NAME: "=Sheet1!$A$1:$B$4" }, createdAt, createdBy, updatedAt, updatedBy, updatedByName }
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
    objects: 200,           // per kind and sheet
    objectChars: 20000,
    names: 500,
    docChars: 4000000       // serialized workbook
  };
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
      updatedAt: now, updatedBy: user ? user.login : null, updatedByName: user ? user.name : null
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

  // ------------------------------------------------------------ validation

  function fail(msg) { const e = new Error(msg); e.code = 400; throw e; }

  function cleanStyle(st) {
    if (st === null || st === undefined) return null;
    if (typeof st !== 'object') fail('Invalid style');
    const out = {};
    Object.keys(st).forEach(function (k) {
      const rule = STYLE_KEYS[k], v = st[k];
      if (!rule || v === null || v === undefined || v === false || v === '') return;
      if (rule === 'bool') { if (v) out[k] = 1; }
      else if (rule === 'color') { if (COLOR_RE.test(v)) out[k] = String(v).toLowerCase(); }
      else if (rule === 'int') { const n = Math.round(Number(v)); if (n >= 0 && n <= 15) out[k] = n; }
      else if (rule === 'size') { const n = Math.round(Number(v)); if (n >= 6 && n <= 72) out[k] = n; }
      else if (rule === 'str') { out[k] = String(v).slice(0, 8); }
      else if (rule === 'pattern') { out[k] = String(v).slice(0, 80); }
      else if (rule === 'border') { const b = String(v).replace(/[^tlrbTLRB]/g, ''); if (b) out[k] = b; }
      else if (Array.isArray(rule)) { if (rule.indexOf(v) >= 0) out[k] = v; }
    });
    return Object.keys(out).length ? out : null;
  }

  function checkId(id, what) {
    if (typeof id !== 'string' || !ID_RE.test(id)) fail('Invalid ' + what + ' id');
    return id;
  }

  function cleanName(name) {
    const n = String(name || '').replace(/[\\/?*\[\]:]/g, '').trim().slice(0, LIMITS.nameChars);
    if (!n) fail('Sheet name is empty');
    return n;
  }

  function cleanObject(kind, id, val) {
    if (OBJ_KINDS.indexOf(kind) < 0) fail('Unknown object kind');
    checkId(id, 'object');
    if (val === null) return null;
    if (!val || typeof val !== 'object' || Array.isArray(val)) fail('Invalid object');
    const raw = JSON.stringify(val);
    if (raw.length > LIMITS.objectChars) fail('The object is too large');
    const out = JSON.parse(raw);
    out.id = id;
    if (kind === 'tables') {
      if (!NAME_RE.test(String(out.name || ''))) fail('Table names start with a letter and contain only letters, digits, "_" and "."');
    }
    return out;
  }

  function cleanSheet(sh) {
    if (!sh || typeof sh !== 'object') fail('Invalid sheet');
    const rows = Array.isArray(sh.rows) ? sh.rows.map(function (id) { return checkId(id, 'row'); }) : [];
    const cols = Array.isArray(sh.cols) ? sh.cols.map(function (id) { return checkId(id, 'column'); }) : [];
    if (!rows.length || !cols.length) fail('A sheet needs rows and columns');
    if (rows.length > LIMITS.rows || cols.length > LIMITS.cols) fail('The sheet is too large (at most ' + LIMITS.rows + ' rows and ' + LIMITS.cols + ' columns)');
    const out = newSheet(cleanName(sh.name), 1, 1);
    out.id = checkId(sh.id, 'sheet');
    out.rows = rows;
    out.cols = cols;
    const rowSet = {}, colSet = {};
    rows.forEach(function (r) { rowSet[r] = 1; });
    cols.forEach(function (c) { colSet[c] = 1; });
    Object.keys(sh.cells || {}).forEach(function (k) {
      const p = k.split(':');
      if (!rowSet[p[0]] || !colSet[p[1]]) return;
      const c = sh.cells[k] || {};
      const cell = {};
      if (c.v !== null && c.v !== undefined && c.v !== '') cell.v = String(c.v).slice(0, LIMITS.cellChars);
      const st = cleanStyle(c.st);
      if (st) cell.st = st;
      if (c.n) cell.n = String(c.n).slice(0, LIMITS.noteChars);
      if (cell.v !== undefined || cell.st || cell.n) out.cells[k] = cell;
    });
    applySheetProps(out, { colW: sh.colW, rowH: sh.rowH, hr: sh.hr, hc: sh.hc, merges: sh.merges, freeze: sh.freeze, color: sh.color, nogrid: sh.nogrid });
    OBJ_KINDS.forEach(function (kind) {
      (Array.isArray(sh[kind]) ? sh[kind] : []).slice(0, LIMITS.objects).forEach(function (o) {
        if (o && ID_RE.test(o.id)) out[kind].push(cleanObject(kind, o.id, o));
      });
    });
    return out;
  }

  function applySheetProps(sheet, op) {
    if (op.name !== undefined) sheet.name = cleanName(op.name);
    if (op.color !== undefined) sheet.color = op.color && COLOR_RE.test(op.color) ? op.color.toLowerCase() : null;
    if (op.nogrid !== undefined) sheet.nogrid = !!op.nogrid;
    ['colW', 'rowH'].forEach(function (prop) {
      const patch = op[prop];
      if (!patch || typeof patch !== 'object') return;
      Object.keys(patch).forEach(function (id) {
        const v = patch[id];
        if (v === null || v === undefined) delete sheet[prop][id];
        else { const n = Math.round(Number(v)); if (n >= 2 && n <= 2000) sheet[prop][id] = n; }
      });
    });
    ['hr', 'hc'].forEach(function (prop) {
      const patch = op[prop];
      if (!patch || typeof patch !== 'object') return;
      Object.keys(patch).forEach(function (id) {
        if (patch[id]) sheet[prop][id] = 1; else delete sheet[prop][id];
      });
    });
    if (Array.isArray(op.merges)) {
      sheet.merges = op.merges.filter(function (m) {
        return m && ID_RE.test(m.r1) && ID_RE.test(m.c1) && ID_RE.test(m.r2) && ID_RE.test(m.c2);
      }).slice(0, 5000).map(function (m) { return { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 }; });
    }
    if (op.freeze && typeof op.freeze === 'object') {
      sheet.freeze = {
        r: Math.max(0, Math.min(50, Math.round(Number(op.freeze.r) || 0))),
        c: Math.max(0, Math.min(26, Math.round(Number(op.freeze.c) || 0)))
      };
    }
  }

  /**
   * After rows or columns were deleted, moves range corners that pointed at them to the nearest remaining row or
   * column inside the range, and drops ranges that lost all of their rows or columns.
   */
  function remapRanges(sheet, axis, oldList) {
    const list = axis === 'rows' ? sheet.rows : sheet.cols;
    const alive = {};
    list.forEach(function (id, i) { alive[id] = i; });
    const oldIndex = {};
    oldList.forEach(function (id, i) { oldIndex[id] = i; });
    const k1 = axis === 'rows' ? 'r1' : 'c1', k2 = axis === 'rows' ? 'r2' : 'c2', ka = axis === 'rows' ? 'ar' : 'ac';
    function fix(o) {
      if (!o) return false;
      if (o[k1] === undefined && o[ka] === undefined) return true;
      if (o[k1] !== undefined) {
        let a = oldIndex[o[k1]], b = oldIndex[o[k2]];
        if (a === undefined || b === undefined) return false;
        if (a > b) { const t = a; a = b; b = t; }
        while (a <= b && alive[oldList[a]] === undefined) a++;
        while (b >= a && alive[oldList[b]] === undefined) b--;
        if (a > b) return false;
        o[k1] = oldList[a];
        o[k2] = oldList[b];
      }
      if (o[ka] !== undefined && alive[o[ka]] === undefined) {
        let i = oldIndex[o[ka]];
        if (i === undefined) i = 0;
        while (i >= 0 && alive[oldList[i]] === undefined) i--;
        o[ka] = i >= 0 ? oldList[i] : list[0];
      }
      return true;
    }
    sheet.merges = sheet.merges.filter(fix);
    OBJ_KINDS.forEach(function (kind) { sheet[kind] = sheet[kind].filter(fix); });
  }

  /** True when the batch must be based on the latest revision (it moves rows or columns). */
  function isStrict(ops) {
    return ops.some(function (op) {
      return (op.t === 'rows' || op.t === 'cols') && (op.del || (op.add && op.at !== -1));
    });
  }

  // ------------------------------------------------------------ apply

  function sameCell(a, b) { return JSON.stringify(a || null) === JSON.stringify(b || null); }

  /**
   * Applies a batch to the workbook (in place). Returns the changes, for the history:
   *   { kind: 'cell', s, k, from, to } | { kind: 'rows'|'cols', s, added|deleted, at } | { kind: 'obj', … } | …
   */
  function applyOps(doc, ops, options) {
    const validate = !options || options.validate !== false;
    if (!Array.isArray(ops)) fail('Invalid change');
    if (ops.length > LIMITS.opsPerBatch) fail('Too many changes at once');
    upgrade(doc);
    const changes = [];
    ops.forEach(function (op) {
      if (!op || typeof op !== 'object') { if (validate) fail('Invalid change'); return; }
      const sheet = op.s ? findSheet(doc, op.s) : null;
      switch (op.t) {
        case 'set': {
          if (!sheet || typeof op.k !== 'string') return;
          const p = op.k.split(':');
          if (sheet.rows.indexOf(p[0]) < 0 || sheet.cols.indexOf(p[1]) < 0) return;
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
          applySheetProps(sheet, op);
          changes.push({ kind: 'sheet', s: sheet.id, renamed: was !== sheet.name ? { from: was, to: sheet.name } : null });
          return;
        }
        case 'rows':
        case 'cols': {
          if (!sheet) return;
          const list = op.t === 'rows' ? sheet.rows : sheet.cols;
          if (Array.isArray(op.add)) {
            const add = op.add.filter(function (id) { return ID_RE.test(id) && list.indexOf(id) < 0; });
            if (!add.length) return;
            const max = op.t === 'rows' ? LIMITS.rows : LIMITS.cols;
            if (list.length + add.length > max) fail('A sheet can have at most ' + max + (op.t === 'rows' ? ' rows' : ' columns'));
            const at = op.at === -1 || op.at === undefined || op.at > list.length ? list.length : Math.max(0, Math.round(op.at));
            const end = at === list.length;
            Array.prototype.splice.apply(list, [at, 0].concat(add));
            changes.push({ kind: op.t, s: sheet.id, added: add.length, at: at, end: end });
          } else if (Array.isArray(op.del)) {
            const del = {};
            let first = -1, count = 0;
            op.del.forEach(function (id) {
              const i = list.indexOf(id);
              if (i >= 0 && !del[id]) { del[id] = 1; count++; if (first < 0 || i < first) first = i; }
            });
            if (!count) return;
            if (count >= list.length) fail('A sheet needs at least one ' + (op.t === 'rows' ? 'row' : 'column'));
            const oldList = list.slice();
            const kept = list.filter(function (id) { return !del[id]; });
            list.length = 0;
            Array.prototype.push.apply(list, kept);
            const part = op.t === 'rows' ? 0 : 1;
            Object.keys(sheet.cells).forEach(function (k) { if (del[k.split(':')[part]]) delete sheet.cells[k]; });
            const sizes = op.t === 'rows' ? sheet.rowH : sheet.colW;
            const hidden = op.t === 'rows' ? sheet.hr : sheet.hc;
            Object.keys(del).forEach(function (id) { delete sizes[id]; delete hidden[id]; });
            remapRanges(sheet, op.t, oldList);
            changes.push({ kind: op.t, s: sheet.id, deleted: count, at: first });
          }
          return;
        }
        case 'obj': {
          if (!sheet) return;
          const val = validate ? cleanObject(op.kind, op.id, op.val) : op.val;
          const listO = sheet[op.kind];
          if (!listO) return;
          const idx = listO.findIndex(function (o) { return o.id === op.id; });
          if (val === null) {
            if (idx < 0) return;
            const gone = listO.splice(idx, 1)[0];
            changes.push({ kind: 'obj', s: sheet.id, objKind: op.kind, action: 'delete', name: gone.name || gone.title || null });
            return;
          }
          if (idx >= 0) {
            if (JSON.stringify(listO[idx]) === JSON.stringify(val)) return;
            listO[idx] = val;
          } else {
            if (listO.length >= LIMITS.objects) fail('Too many items of this kind on one sheet');
            listO.push(val);
          }
          changes.push({ kind: 'obj', s: sheet.id, objKind: op.kind, action: idx >= 0 ? 'update' : 'add', name: val.name || val.title || null });
          return;
        }
        case 'name': {
          if (!NAME_RE.test(String(op.name || ''))) { if (validate) fail('Names start with a letter and contain only letters, digits, "_" and "."'); return; }
          const key = String(op.name).toUpperCase();
          const had = Object.prototype.hasOwnProperty.call(doc.names, key);
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
          if (findSheet(doc, sh.id)) return;
          let name = sh.name, n = 2;
          while (doc.sheets.some(function (x) { return x.name.toLowerCase() === name.toLowerCase(); })) {
            name = sh.name.slice(0, LIMITS.nameChars - 5) + ' (' + (n++) + ')';
          }
          sh.name = name;
          const at = op.at === undefined || op.at < 0 || op.at > doc.sheets.length ? doc.sheets.length : op.at;
          doc.sheets.splice(at, 0, sh);
          changes.push({ kind: 'addSheet', s: sh.id, name: sh.name });
          return;
        }
        case 'delSheet': {
          if (!sheet) return;
          if (doc.sheets.length <= 1) fail('A workbook needs at least one sheet');
          doc.sheets.splice(doc.sheets.indexOf(sheet), 1);
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
          if (validate) fail('Unknown change: ' + op.t);
      }
    });
    if (validate) {
      const names = {}, tables = {};
      doc.sheets.forEach(function (sh) {
        const n = sh.name.toLowerCase();
        if (names[n]) fail('Another sheet is already named "' + sh.name + '"');
        names[n] = 1;
        sh.tables.forEach(function (t) {
          const tn = String(t.name).toLowerCase();
          if (tables[tn]) fail('Another table is already named "' + t.name + '"');
          tables[tn] = 1;
        });
      });
    }
    return changes;
  }

  return {
    LIMITS: LIMITS,
    DEFAULT_ROWS: DEFAULT_ROWS,
    DEFAULT_COLS: DEFAULT_COLS,
    OBJ_KINDS: OBJ_KINDS,
    NAME_RE: NAME_RE,
    randomId: randomId,
    ids: ids,
    colName: colName,
    cellLabel: cellLabel,
    newSheet: newSheet,
    newWorkbook: newWorkbook,
    upgrade: upgrade,
    findSheet: findSheet,
    cleanSheet: cleanSheet,
    cleanStyle: cleanStyle,
    isStrict: isStrict,
    applyOps: applyOps
  };
});
