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
/* Spreadsheets for YouTrack – formula engine.
 *
 * Excel-compatible formulas: =SUM(A1:B3)*2, $A$1, Sheet2!A1, 'My sheet'!A1:B2, A:A, 1:3, "text" & A1, 10%,
 * comparisons, ^, array arithmetic (=SUM(A1:A3*B1:B3)), dynamic arrays that spill (=SORT(UNIQUE(A2:A50)),
 * A1# for a spilled range), structured table references (=SUM(Sales[Amount]), [@Price]*[@Qty]), named
 * ranges, LET, and more than 300 functions. Arguments are separated with "," or ";".
 *
 * Values: number, string, boolean, null (empty), FormulaError, 2-D arrays (rows of values) and Range.
 * Dates are Excel serial numbers (days since 1899-12-30).
 */

export class FormulaError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
const ERR = {};
['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!', '#CIRC!', '#ERROR!']
  .forEach((c) => { ERR[c] = new FormulaError(c); });
export const ERRORS = ERR;
export const isErr = (v) => v instanceof FormulaError;
/* #CALC! of a formula that the work budget stopped (not an empty result): a separate value with the same code, so the
 * widget can tell the two apart and say why. */
const SLOW = new FormulaError('#CALC!');
SLOW.slow = true;
export const isSlowCalc = (v) => v === SLOW;

/* Work budget. One burst of evaluation (a render, a recalculation) may take this many steps; after that formulas
 * give #CALC! instead of freezing the page. The budget refills when the burst is over and when values are cleared. */
const BUDGET = 3e7, BUDGET_MS = 1500;
// byClock: the budget of this burst ran out because the clock did (not the steps); burstNo counts the bursts that are over
let budget = BUDGET, burst = false, clock = 0, calls = 0, deadline = 0, byClock = false, burstNo = 0;
function refill() { budget = BUDGET; deadline = 0; byClock = false; }
function tick(n) {
  if (!burst) { burst = true; Promise.resolve().then(() => { burst = false; burstNo++; refill(); }); }
  budget -= n;
  if (budget < 0) throw SLOW;
  // steps are not equally slow (one array function can work on a million values), so the clock is checked now and
  // then too: after many steps, and after every few calls (each formula part that is evaluated makes one)
  if ((clock += n) >= 50000 || ++calls >= 16) {
    clock = 0; calls = 0;
    const now = Date.now();
    if (!deadline) deadline = now + BUDGET_MS;
    else if (now > deadline) { budget = -1; byClock = true; throw SLOW; }
  }
}
const spent = () => budget < 0;
/** True while the current burst has used up the work budget: values read now may be #CALC! only for lack of time, so
 * results built from them must not be kept (the widget draws again to compute the rest). */
export const budgetSpent = spent;
/** Counts n steps of work done outside formulas (statistics of conditional formatting rules) against the budget:
 * throws the budget's #CALC! (isSlowCalc) when it is used up. */
export function charge(n) { tick(n); }
/* Whether all of the burst's budget is left, but for a few steps (the formula that reads the one starting now): an
 * evaluation that starts then and is still stopped cannot be finished in any burst. */
const wholeBudget = () => budget >= BUDGET - 1e4 && (!deadline || deadline - Date.now() >= BUDGET_MS - 15);
/* A page in the background runs slower (and may be throttled): there the clock does not prove a formula too heavy. */
const pageHidden = () => typeof document !== 'undefined' && document.hidden === true;
// kept results of formulas outside cells (conditional formatting, see evalAt): at most this many, in two generations
// (the older one is dropped when the newer one is half of it), so that results drawn recently stay
const KEEP_MAX = 300000;
const clockNow = () => (typeof performance !== 'undefined' ? performance.now() : Date.now());
const fixedFormulas = new Map();
/**
 * Whether a formula gives the same wherever it is evaluated: no relative references (copying it elsewhere does not
 * change it), nothing that reads the position (ROW(), COLUMN(), this-row table references, columns of the cell's own
 * table, INDIRECT, OFFSET, CELL, ADDRESS), and only defined names (names: the workbook's) of which the same is true.
 * Such a formula too heavy to compute in one cell is too heavy in all of them. Kept until Engine.invalidate.
 */
export function cellIndependent(text, names, depth = 0) {
  const t = String(text ?? '');
  let f = fixedFormulas.get(t);
  if (f !== undefined) return f;
  const withEq = t[0] === '=' ? t : '=' + t;
  f = !/\b(ROW|COLUMN|INDIRECT|OFFSET|CELL|ADDRESS)\s*\(|\[\s*@|#this\s+row/i.test(t) && translateFormula(withEq, 1, 1) === withEq;
  if (f) {
    let toks = [];
    try { toks = tokenize(withEq.slice(1)); } catch (e) { f = false; }
    for (const k of toks) {
      if (k.t === 'sref' && !k.table) { f = false; break; }
      if (k.t !== 'name' || !names) continue;
      const def = names[k.v] ?? Object.entries(names).find(([n]) => n.toUpperCase() === k.v)?.[1];
      if (def !== undefined && (depth >= 4 || !cellIndependent(def, names, depth + 1))) { f = false; break; }
    }
  }
  if (fixedFormulas.size >= 2000) fixedFormulas.clear();
  fixedFormulas.set(t, f);
  return f;
}

/* ================================================================ addresses */

export function colName(i) {
  let n = i + 1, s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function colIndex(name) {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
const MAX_COL = 16383, MAX_ROW = 1048575;
// text as long as Excel allows in a cell, arrays no bigger than a sheet (20000 rows x 200 columns)
const MAX_TEXT = 32767, MAX_CELLS = 20000 * 200;
const longText = (s) => typeof s === 'string' && s.length > MAX_TEXT;
/** Throws #VALUE! when an array of rows x cols would be bigger than a sheet; returns the cell count. */
function arraySize(rows, cols) {
  if (!(rows >= 1 && cols >= 1 && rows * cols <= MAX_CELLS)) throw ERR['#VALUE!'];
  tick(rows * cols);
  return rows * cols;
}

/** "B3" → {r: 2, c: 1}; also accepts "$B$3". */
export function parseAddr(text) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(text).trim());
  if (!m) return null;
  const c = colIndex(m[1]), r = +m[2] - 1;
  if (r < 0 || r > MAX_ROW || c > MAX_COL) return null;
  return { r, c };
}
export const addrName = (r, c) => colName(c) + (r + 1);

export function quoteSheet(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/* ================================================================ tokenizer */

const SHEET_RE = /^(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!/;
const CELL_RANGE_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?/;
const COL_RANGE_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
const ROW_RANGE_RE = /^(\$?)(\d+):(\$?)(\d+)(?![\w.(])/;
const NUM_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const NAME_RE = /^[A-Za-z_\\À-￿][\w.À-￿]*/;
const ERR_RE = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|CIRC!|ERROR!)/i;

/** Reads a reference at the start of `s`. Returns {ref, len} or null. */
function readRef(s) {
  let sheet = null, off = 0;
  const sm = SHEET_RE.exec(s);
  if (sm) { sheet = sm[1] !== undefined ? sm[1].replace(/''/g, "'") : sm[2]; off = sm[0].length; }
  const rest = s.slice(off);
  let m = COL_RANGE_RE.exec(rest);
  if (m) {
    const c1 = colIndex(m[2]), c2 = colIndex(m[4]);
    if (c1 <= MAX_COL && c2 <= MAX_COL) {
      return { len: off + m[0].length, ref: { kind: 'cols', sheet, r1: null, r2: null, c1, c2, ca1: !!m[1], ca2: !!m[3], ra1: false, ra2: false } };
    }
  }
  m = ROW_RANGE_RE.exec(rest);
  if (m) {
    const r1 = +m[2] - 1, r2 = +m[4] - 1;
    if (r1 >= 0 && r2 >= 0 && r1 <= MAX_ROW && r2 <= MAX_ROW) {
      return { len: off + m[0].length, ref: { kind: 'rows', sheet, c1: null, c2: null, r1, r2, ra1: !!m[1], ra2: !!m[3], ca1: false, ca2: false } };
    }
  }
  m = CELL_RANGE_RE.exec(rest);
  if (m) {
    let len = m[0].length;
    const after = rest.slice(len);
    if (/^[\w.(\[]/.test(after)) return null; // a function or name such as LOG10( or A1B
    const c1 = colIndex(m[2]), r1 = +m[4] - 1;
    if (c1 > MAX_COL || r1 < 0 || r1 > MAX_ROW) return null;
    const ref = { kind: 'cell', sheet, r1, c1, r2: r1, c2: c1, ca1: !!m[1], ra1: !!m[3], ca2: !!m[1], ra2: !!m[3] };
    if (m[6]) {
      const c2 = colIndex(m[6]), r2 = +m[8] - 1;
      if (c2 > MAX_COL || r2 < 0 || r2 > MAX_ROW) return null;
      Object.assign(ref, { kind: 'range', c2, r2, ca2: !!m[5], ra2: !!m[7] });
    } else if (after[0] === '#') {
      ref.spill = true;
      len += 1;
    }
    return { len: off + len, ref };
  }
  return null;
}

/** Reads a bracketed structured-reference spec starting at s[0] === '['. Returns its length or -1. */
function readBrackets(s) {
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'" && i + 1 < s.length) { i++; continue; } // escaped char inside a column name
    if (ch === '[') depth++;
    else if (ch === ']') { depth--; if (depth === 0) return i + 1; }
  }
  return -1;
}

/** Splits a formula (without the leading "=") into tokens with their source positions. */
export function tokenize(src) {
  const out = [];
  let i = 0;
  let prev = null;
  while (i < src.length) {
    const ch = src[i];
    const s = src.slice(i);
    let tok = null;
    if (/\s/.test(ch)) {
      tok = { t: 'ws', text: /^\s+/.exec(s)[0] };
    } else if (ch === '"') {
      let j = i + 1, val = '';
      for (;;) {
        if (j >= src.length) { tok = { t: 'bad', text: src.slice(i) }; break; }
        if (src[j] === '"') {
          if (src[j + 1] === '"') { val += '"'; j += 2; continue; }
          tok = { t: 'str', text: src.slice(i, j + 1), v: val };
          break;
        }
        val += src[j++];
      }
    } else if (ch === '{') {
      const j = src.indexOf('}', i);
      tok = j < 0 ? { t: 'bad', text: src.slice(i) } : { t: 'arr', text: src.slice(i, j + 1) };
    } else if (ERR_RE.test(s)) {
      const m = ERR_RE.exec(s);
      tok = { t: 'err', text: m[0], v: m[0].toUpperCase() };
    } else if (ch === '[') {
      const len = readBrackets(s);
      tok = len < 0 ? { t: 'bad', text: s } : { t: 'sref', text: s.slice(0, len), table: null, spec: s.slice(0, len) };
    } else {
      const r = !prev || prev.t !== 'ref' ? readRef(s) : null;
      if (r) {
        tok = { t: 'ref', text: s.slice(0, r.len), ref: r.ref };
      } else if (NUM_RE.test(s) && !/^\d+:/.test(s)) {
        const m = NUM_RE.exec(s);
        tok = { t: 'num', text: m[0], v: parseFloat(m[0]) };
      } else if (NAME_RE.test(s)) {
        const m = NAME_RE.exec(s);
        const rest = src.slice(i + m[0].length);
        const up = m[0].toUpperCase();
        if (rest[0] === '[') {
          const len = readBrackets(rest);
          tok = len < 0 ? { t: 'bad', text: s } : { t: 'sref', text: m[0] + rest.slice(0, len), table: m[0], spec: rest.slice(0, len) };
        } else if (/^\s*\(/.test(rest)) tok = { t: 'func', text: m[0], v: up.replace(/^_XLFN\.|^_XLWS\./, '') };
        else if (up === 'TRUE' || up === 'FALSE') tok = { t: 'bool', text: m[0], v: up === 'TRUE' };
        else tok = { t: 'name', text: m[0], v: up };
      } else if (/^(<=|>=|<>)/.test(s)) {
        tok = { t: 'op', text: s.slice(0, 2) };
      } else if ('+-*/^&=<>%'.includes(ch)) {
        tok = { t: 'op', text: ch };
      } else if (ch === '(') tok = { t: 'lp', text: ch };
      else if (ch === ')') tok = { t: 'rp', text: ch };
      else if (ch === ',' || ch === ';') tok = { t: 'sep', text: ch };
      else if (ch === ':') tok = { t: 'op', text: ':' };
      // a quote that does not start 'Sheet name'! makes the formula invalid: take the rest in one token, so that
      // runs of quotes are not scanned again from every position
      else tok = { t: 'bad', text: ch === "'" ? s : ch };
    }
    tok.s = i;
    tok.e = i + tok.text.length;
    out.push(tok);
    if (tok.t !== 'ws') prev = tok;
    i = tok.e;
  }
  return out;
}

/* ================================================================ parser */

class ParseError extends Error {}
const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** Splits a row of an array constant at the commas outside "text" (with an even number of quotes after them). */
function splitRow(row) {
  const cuts = [];
  let quotes = 0;
  for (let i = row.length - 1; i >= 0; i--) {
    if (row[i] === '"') quotes++;
    else if (row[i] === ',' && quotes % 2 === 0) cuts.push(i);
  }
  const out = [];
  let from = 0;
  for (let k = cuts.length - 1; k >= 0; k--) { out.push(row.slice(from, cuts[k])); from = cuts[k] + 1; }
  out.push(row.slice(from));
  return out;
}
function parseArrayLiteral(text) {
  const inner = text.slice(1, -1);
  return inner.split(';').map((row) => splitRow(row).map((cell) => {
    const t = cell.trim();
    if (/^".*"$/.test(t)) return t.slice(1, -1).replace(/""/g, '"');
    if (/^(TRUE|FALSE)$/i.test(t)) return t.toUpperCase() === 'TRUE';
    if (ERR[t.toUpperCase()]) return ERR[t.toUpperCase()];
    const n = Number(t);
    return t !== '' && isFinite(n) ? n : t;
  }));
}

function parseTokens(tokens) {
  const toks = tokens.filter((t) => t.t !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];

  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || !(t.text in BIN_PREC)) break;
      const prec = BIN_PREC[t.text];
      if (prec < minPrec) break;
      next();
      const right = expr(prec + 1); // all binary operators are left-associative in Excel (2^3^2 = 64)
      left = { k: 'bin', op: t.text, a: left, b: right };
    }
    return left;
  }
  function unary() {
    const t = peek();
    if (t && t.t === 'op' && (t.text === '-' || t.text === '+')) {
      next();
      const v = unary();
      return t.text === '-' ? { k: 'neg', a: v } : v;
    }
    return postfix(primary());
  }
  function postfix(node) {
    let n = node;
    for (;;) {
      const t = peek();
      if (t && t.t === 'op' && t.text === '%') { next(); n = { k: 'pct', a: n }; continue; }
      if (t && t.t === 'op' && t.text === ':') { next(); n = { k: 'span', a: n, b: postfix(primary()) }; continue; }
      return n;
    }
  }
  function primary() {
    const t = next();
    if (!t) throw new ParseError('Unexpected end');
    switch (t.t) {
      case 'num': return { k: 'lit', v: t.v };
      case 'str': return { k: 'lit', v: t.v };
      case 'bool': return { k: 'lit', v: t.v };
      case 'err': return { k: 'lit', v: ERR[t.v] || ERR['#ERROR!'] };
      case 'arr': return { k: 'lit', v: parseArrayLiteral(t.text) };
      case 'ref': return { k: 'ref', ref: t.ref };
      case 'sref': return { k: 'sref', table: t.table, spec: t.spec };
      case 'name': return { k: 'name', v: t.v };
      case 'lp': {
        const e = expr(1);
        const c = next();
        if (!c || c.t !== 'rp') throw new ParseError('Missing )');
        return e;
      }
      case 'func': {
        const lp = next();
        if (!lp || lp.t !== 'lp') throw new ParseError('Missing (');
        const args = [];
        if (peek() && peek().t === 'rp') { next(); return { k: 'fn', name: t.v, args }; }
        for (;;) {
          if (peek() && (peek().t === 'sep' || peek().t === 'rp')) args.push({ k: 'empty' });
          else args.push(expr(1));
          const sep = next();
          if (!sep) throw new ParseError('Missing )');
          if (sep.t === 'rp') break;
          if (sep.t !== 'sep') throw new ParseError('Unexpected ' + sep.text);
          if (peek() && peek().t === 'rp') { args.push({ k: 'empty' }); next(); break; }
        }
        return { k: 'fn', name: t.v, args };
      }
      default:
        throw new ParseError('Unexpected ' + t.text);
    }
  }
  const node = expr(1);
  if (p < toks.length) throw new ParseError('Unexpected ' + toks[p].text);
  return node;
}

const astCache = new Map();
/** Parses formula text (without "="). Bad formulas give {k: 'bad'}. */
export function parseFormula(text) {
  let node = astCache.get(text);
  if (node) return node;
  try { node = parseTokens(tokenize(text)); } catch (e) { node = { k: 'bad', msg: e.message }; }
  if (astCache.size > 5000) astCache.clear();
  astCache.set(text, node);
  return node;
}
/** Error message for a formula the user typed, or null when it parses. */
export function formulaSyntaxError(formula) {
  const n = parseFormula(String(formula).replace(/^=/, ''));
  return n.k === 'bad' ? n.msg : null;
}

/* ================================================================ reference rewriting */

export function refText(ref) {
  const sh = ref.sheet != null ? quoteSheet(ref.sheet) + '!' : '';
  const d = (abs) => (abs ? '$' : '');
  if (ref.kind === 'cols') return sh + d(ref.ca1) + colName(ref.c1) + ':' + d(ref.ca2) + colName(ref.c2);
  if (ref.kind === 'rows') return sh + d(ref.ra1) + (ref.r1 + 1) + ':' + d(ref.ra2) + (ref.r2 + 1);
  const a = d(ref.ca1) + colName(ref.c1) + d(ref.ra1) + (ref.r1 + 1);
  if (ref.kind === 'cell') return sh + a + (ref.spill ? '#' : '');
  return sh + a + ':' + d(ref.ca2) + colName(ref.c2) + d(ref.ra2) + (ref.r2 + 1);
}

/** Rewrites every reference in a formula ("=..."). fn(ref) returns a ref (changed or not) or null (#REF!). */
export function rewriteRefs(formula, fn) {
  if (typeof formula !== 'string' || formula[0] !== '=') return formula;
  const toks = tokenize(formula.slice(1));
  let out = '', changed = false;
  for (const t of toks) {
    if (t.t === 'ref') {
      const next = fn({ ...t.ref });
      if (next === null) { out += '#REF!'; changed = true; continue; }
      const txt = refText(next);
      if (txt !== refText(t.ref)) { changed = true; out += txt; } else out += t.text;
    } else out += t.text;
  }
  return changed ? '=' + out : formula;
}

/** Shifts relative references by (dr, dc), as Excel does when a formula is copied. */
export function translateFormula(formula, dr, dc) {
  if (!dr && !dc) return formula;
  return rewriteRefs(formula, (ref) => {
    if (ref.kind !== 'cols') {
      if (!ref.ra1) ref.r1 += dr;
      if (!ref.ra2) ref.r2 += dr;
    }
    if (ref.kind !== 'rows') {
      if (!ref.ca1) ref.c1 += dc;
      if (!ref.ca2) ref.c2 += dc;
    }
    if ((ref.r1 !== null && (ref.r1 < 0 || ref.r2 < 0)) || (ref.c1 !== null && (ref.c1 < 0 || ref.c2 < 0))) return null;
    return ref;
  });
}

/**
 * Adjusts references after rows or columns were inserted or deleted.
 *   axis: 'r' | 'c'; at: first index; count > 0 inserted, < 0 deleted; isTarget(ref): points at that sheet
 */
export function shiftFormula(formula, axis, at, count, isTarget) {
  return rewriteRefs(formula, (ref) => shiftRef(ref, axis, at, count, isTarget));
}
export function shiftRef(ref, axis, at, count, isTarget) {
  if (isTarget && !isTarget(ref)) return ref;
  const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
  if (ref[k1] === null) return ref;
  if (count > 0) {
    if (ref[k1] >= at) ref[k1] += count;
    if (ref[k2] >= at) ref[k2] += count;
    return ref;
  }
  const n = -count, end = at + n;
  const a = ref[k1], b = ref[k2];
  const na = a < at ? a : (a >= end ? a - n : at);
  const nb = b < at ? b : (b >= end ? b - n : at - 1);
  if (nb < na) return null;
  ref[k1] = na; ref[k2] = nb;
  return ref;
}

/** Points references from one sheet name to another (rename) or to #REF! (newName null: deleted sheet). */
export function renameSheetInFormula(formula, oldName, newName) {
  const low = oldName.toLowerCase();
  return rewriteRefs(formula, (ref) => {
    if (ref.sheet == null || ref.sheet.toLowerCase() !== low) return ref;
    if (newName === null) return null;
    ref.sheet = newName;
    return ref;
  });
}

/** References in a formula with their positions in the text, for highlighting while editing. */
export function formulaRefs(formula) {
  if (typeof formula !== 'string' || formula[0] !== '=') return [];
  return tokenize(formula.slice(1)).filter((t) => t.t === 'ref').map((t) => ({ ...t.ref, s: t.s + 1, e: t.e + 1 }));
}

/** Parses "Sheet1!$A$1:$B$4" or "A1" into a ref, or null. */
export function parseRefText(text) {
  const t = String(text).trim().replace(/^=/, '');
  const r = readRef(t);
  return r && r.len === t.length ? r.ref : null;
}

/* ================================================================ dates */

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
export const dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
export function serialToParts(serial) {
  const dt = new Date(Math.round(serial * DAY_MS) + EPOCH);
  return {
    y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(),
    H: dt.getUTCHours(), M: dt.getUTCMinutes(), S: dt.getUTCSeconds(), dow: dt.getUTCDay()
  };
}
function nowSerial() {
  const n = new Date();
  return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()) + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
}
function todaySerial() { const n = new Date(); return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()); }

/* ================================================================ values and coercion */

export class Range {
  constructor(engine, sheet, r1, c1, r2, c2) { Object.assign(this, { engine, sheet, r1, c1, r2, c2 }); }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  get(i, j) { return this.engine.valueAt(this.sheet, this.r1 + i, this.c1 + j); }
  *values() { for (let i = 0; i < this.rows; i++) for (let j = 0; j < this.cols; j++) yield this.get(i, j); }
  matrix() {
    const m = [];
    for (let i = 0; i < this.rows; i++) { const row = []; for (let j = 0; j < this.cols; j++) row.push(this.get(i, j)); m.push(row); }
    return m;
  }
}
const isMatrix = (v) => Array.isArray(v);
const isMulti = (v) => (v instanceof Range && (v.rows > 1 || v.cols > 1)) || (isMatrix(v) && (v.length > 1 || (v[0] && v[0].length > 1)));
function toMatrix(v) {
  if (v instanceof Range) return v.matrix(); // each cell read counts against the work budget
  if (isMatrix(v)) return v;
  return [[v]];
}
function scalar(v) {
  if (v instanceof Range) return v.rows === 1 && v.cols === 1 ? v.get(0, 0) : ERR['#VALUE!'];
  if (isMatrix(v)) return v.length && v[0].length ? v[0][0] : ERR['#VALUE!'];
  return v;
}
function mapMatrix(m, fn) { return m.map((row) => row.map(fn)); }
/** Element-wise operation with Excel broadcasting (1×n, n×1 and 1×1 stretch). */
function broadcast(a, b, fn) {
  const A = toMatrix(a), B = toMatrix(b);
  const rows = Math.max(A.length, B.length), cols = Math.max(A[0].length, B[0].length);
  arraySize(rows, cols);
  const pick = (M, i, j) => {
    const ii = M.length === 1 ? 0 : i, jj = M[0].length === 1 ? 0 : j;
    return ii < M.length && jj < M[0].length ? M[ii][jj] : ERR['#N/A'];
  };
  const out = [];
  for (let i = 0; i < rows; i++) { const row = []; for (let j = 0; j < cols; j++) row.push(fn(pick(A, i, j), pick(B, i, j))); out.push(row); }
  return out;
}

/** Text → number the way Excel does in arithmetic ("12", "1e3", "50%", " 3.5 ", "2026-01-31"). */
function textToNumber(s) {
  const t = String(s).trim();
  if (!t) return 0;
  let m = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(%?)$/i.exec(t);
  if (m) return parseFloat(m[1]) / (m[2] ? 100 : 1);
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) return dateToSerial(+m[1], +m[2], +m[3]);
  return null;
}
function toNum(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === null || v === undefined) return 0;
  const n = textToNumber(v);
  return n === null ? ERR['#VALUE!'] : n;
}
function toStr(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numToText(v);
  return String(v);
}
function toBool(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  const u = String(v).toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return ERR['#VALUE!'];
}
export function numToText(n) {
  if (!isFinite(n)) return '#NUM!';
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(+n.toPrecision(15));
}

/** All values of the arguments, flattened (ranges and arrays included). */
function* flat(args) {
  for (const a of args) {
    if (a instanceof Range) { for (const v of a.values()) yield { v, ref: true }; }
    else if (isMatrix(a)) { for (const row of a) { tick(row.length); for (const v of row) yield { v, ref: true }; } }
    else yield { v: a, ref: false };
  }
}
/** Numbers of the arguments, Excel style: in ranges only numbers count; direct arguments are converted. */
function numbers(args, opts = {}) {
  const out = [];
  for (const { v, ref } of flat(args)) {
    if (isErr(v)) throw v;
    if (typeof v === 'number') out.push(v);
    else if (!ref) {
      if (v === null) continue;
      const n = toNum(v);
      if (isErr(n)) throw n;
      out.push(n);
    } else if (opts.countLogical && typeof v === 'boolean') out.push(v ? 1 : 0);
    else if (opts.countLogical && typeof v === 'string') out.push(0);
  }
  return out;
}
function numArg(v, def) {
  if (v === undefined || v === null && def !== undefined) return def;
  const n = toNum(v);
  if (isErr(n)) throw n;
  return n;
}
function strArg(v, def) {
  if (v === undefined) return def;
  const s = toStr(v);
  if (isErr(s)) throw s;
  return s;
}
function boolArg(v, def) {
  if (v === undefined || v === null) return def;
  const b = toBool(v);
  if (isErr(b)) throw b;
  return b;
}
const finite = (x) => (typeof x === 'number' && !isFinite(x) ? ERR['#NUM!'] : x);

/* ================================================================ criteria (COUNTIF, SUMIF, …) */

// Wildcards (* any text, ? one character, ~ escapes the next one) are matched without regular expressions, so that
// a pattern such as *a*a*a*b costs at most text length x pattern length steps instead of backtracking.
const W_STAR = 1, W_ANY = 2;
/** A character in the case used to compare without case (as a regular expression with the i flag does). */
function canon(ch) {
  const u = ch.toUpperCase();
  return u.length !== 1 || (ch.charCodeAt(0) >= 128 && u.charCodeAt(0) < 128) ? ch : u;
}
function wildcardTokens(pattern) {
  const t = [];
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) t.push(canon(pattern[++i]));
    else if (ch === '*') { if (t[t.length - 1] !== W_STAR) t.push(W_STAR); }
    else t.push(ch === '?' ? W_ANY : canon(ch));
  }
  return t;
}
/** Whether the whole text matches the wildcard tokens. */
function wildcardMatch(tok, s) {
  let i = 0, p = 0, star = -1, back = 0, steps = 0;
  while (i < s.length) {
    steps++;
    const t = tok[p];
    if (p < tok.length && t !== W_STAR && (t === W_ANY || t === canon(s[i]))) { i++; p++; }
    else if (t === W_STAR) { star = p++; back = i; }
    else if (star >= 0) { p = star + 1; i = ++back; }
    else break;
  }
  tick(steps);
  if (i < s.length) return false;
  while (tok[p] === W_STAR) p++;
  return p === tok.length;
}
/** Start of the first match of the wildcard tokens in the text at or after `from`, or -1 (for SEARCH). */
function wildcardFind(tok, s, from) {
  // the part before the first * starts the match; the parts after each * follow it in order
  const parts = [[]];
  for (const t of tok) { if (t === W_STAR) parts.push([]); else parts[parts.length - 1].push(t); }
  const find = (part, q) => {
    for (; q + part.length <= s.length; q++) {
      tick(part.length + 1);
      let k = 0;
      while (k < part.length && (part[k] === W_ANY || part[k] === canon(s[q + k]))) k++;
      if (k === part.length) return q;
    }
    return -1;
  };
  const start = find(parts[0], Math.max(0, from));
  if (start < 0) return -1;
  let q = start + parts[0].length;
  for (let k = 1; k < parts.length; k++) { const f = find(parts[k], q); if (f < 0) return -1; q = f + parts[k].length; }
  return start;
}
function makeCriteria(crit) {
  crit = scalar(crit);
  if (isErr(crit)) throw crit;
  if (typeof crit === 'number' || typeof crit === 'boolean') {
    return (v) => v === crit || (typeof crit === 'number' && typeof v === 'string' && textToNumber(v) === crit && v.trim() !== '');
  }
  let s = crit === null ? '' : String(crit);
  let op = '=';
  const m = /^(<=|>=|<>|=|<|>)/.exec(s);
  if (m) { op = m[1]; s = s.slice(m[1].length); }
  const n = s.trim() === '' ? null : textToNumber(s);
  if (n !== null && n !== undefined) {
    return (v) => {
      const x = typeof v === 'number' ? v : (typeof v === 'string' && v.trim() !== '' ? textToNumber(v) : null);
      if (x === null || x === undefined || typeof v === 'boolean') return op === '<>';
      switch (op) {
        case '=': return x === n; case '<>': return x !== n; case '<': return x < n;
        case '>': return x > n; case '<=': return x <= n; default: return x >= n;
      }
    };
  }
  if (s === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => !(v === null || v === '');
  }
  const up = s.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') { const b = up === 'TRUE'; return (v) => (op === '<>' ? v !== b : v === b); }
  const tok = /[*?~]/.test(s) ? wildcardTokens(s) : null;
  if (tok && s.length > 255) throw ERR['#VALUE!']; // Excel's limit for criteria
  const low = s.toLowerCase();
  return (v) => {
    if (isErr(v)) return false;
    const str = v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'number' ? numToText(v) : String(v);
    if (op === '=' || op === '<>') {
      const eq = tok ? wildcardMatch(tok, str) : str.toLowerCase() === low;
      return op === '=' ? eq : !eq;
    }
    if (typeof v !== 'string') return false;
    const c = str.toLowerCase().localeCompare(low);
    switch (op) { case '<': return c < 0; case '>': return c > 0; case '<=': return c <= 0; default: return c >= 0; }
  };
}
function rangeSize(r) { return r instanceof Range ? [r.rows, r.cols] : isMatrix(r) ? [r.length, r[0].length] : [1, 1]; }
function rangeGet(r, i, j) { return r instanceof Range ? r.get(i, j) : isMatrix(r) ? (r[i] ? r[i][j] : undefined) : r; }
function sameShapeGet(r, i, j) {
  // SUMIF's sum range is read with the criteria range's shape even when it is smaller
  if (r instanceof Range) return r.engine.valueAt(r.sheet, r.r1 + i, r.c1 + j);
  return rangeGet(r, i, j);
}
function ifsAggregate(target, pairs) {
  const [rows, cols] = rangeSize(target);
  const tests = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const [rr, cc] = rangeSize(pairs[i]);
    if (rr !== rows || cc !== cols) throw ERR['#VALUE!'];
    tests.push([pairs[i], makeCriteria(pairs[i + 1])]);
  }
  const out = [];
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) if (tests.every(([r, f]) => f(rangeGet(r, i, j)))) out.push(rangeGet(target, i, j));
  }
  return out;
}

/* ================================================================ comparison */

export function compare(a, b) {
  // Excel order: numbers < text < logical; empty equals 0 / "" / FALSE
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (isErr(a)) return 0;
  return a < b ? -1 : a > b ? 1 : 0;
}

/* ================================================================ TEXT() and custom formats */

export function formatWithPattern(value, pattern, locale) {
  if (typeof value !== 'number') return toStr(value);
  if (String(pattern).length > 255) return numToText(value); // longer than Excel allows for a format
  const sections = String(pattern).split(';');
  let pat = sections[0];
  let v = value;
  if (v < 0 && sections[1] !== undefined) { pat = sections[1]; v = -v; }
  else if (v === 0 && sections[2] !== undefined) pat = sections[2];
  const bare = pat.replace(/"[^"]*"|\\./g, '').replace(/\[[^\]]*\]/g, '');
  pat = pat.replace(/\[[^\]]*\]/g, '');
  if (/^general$/i.test(bare.trim())) return numToText(v);
  if (/[dmyhs]/i.test(bare) && !/^[#0.,%\s]+$/.test(bare)) return formatDatePattern(v, pat, locale);
  const m = /([#0?,]*)(?:\.([0#?]+))?\s*(%?)/.exec(bare);
  const decimals = m && m[2] ? m[2].replace(/#/g, '').length : 0;
  const maxDecimals = m && m[2] ? m[2].length : 0;
  const grouping = /[#0],[#0]/.test(bare);
  const pct = bare.includes('%');
  const num = pct ? v * 100 : v;
  const intPart = m && m[1] ? m[1].replace(/,/g, '') : '';
  const minInt = (intPart.match(/0/g) || []).length;
  let txt = Math.abs(num).toFixed(Math.min(maxDecimals, 100)); // toFixed allows 100 decimals; more are padded below
  let [ip, dp] = txt.split('.');
  if (dp) { dp = dp.replace(/0+$/, ''); if (dp.length < decimals) dp = dp.padEnd(decimals, '0'); }
  ip = ip.padStart(minInt, '0');
  if (minInt === 0 && ip === '0' && dp) ip = '';
  if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  txt = ip + (dp ? '.' + dp : '');
  const literalText = (part) => part.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const first = pat.search(/[#0?]/), lastHash = Math.max(pat.lastIndexOf('#'), pat.lastIndexOf('0'), pat.lastIndexOf('?'));
  const prefix = first > 0 ? literalText(pat.slice(0, first)) : '';
  const suffix = lastHash >= 0 ? literalText(pat.slice(lastHash + 1)).replace('%', '') : '';
  const sign = num < 0 && sections[1] === undefined ? '-' : '';
  return sign + prefix + txt + (pct ? '%' : '') + suffix.replace(/^\.?[0#?]+/, '');
}

function formatDatePattern(serial, pat, locale) {
  const p = serialToParts(serial);
  if (!Number.isFinite(p.y)) return '########'; // not a date (as Excel shows it)
  const ampm = /AM\/PM|A\/P/i.test(pat);
  let out = '';
  const re = /"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|./gi;
  let lastWasHour = false;
  const loc = locale || 'en-US';
  const fmt = (o) => {
    try { return new Intl.DateTimeFormat(loc, { ...o, timeZone: 'UTC' }).format(Date.UTC(p.y, p.m - 1, p.d)); } catch (e) { return '#'; }
  };
  let m;
  while ((m = re.exec(pat))) {
    const tok = m[0];
    const low = tok.toLowerCase();
    if (tok[0] === '"') { out += tok.slice(1, -1); continue; }
    if (tok[0] === '\\') { out += tok[1]; continue; }
    const nextIsSec = /^[:\s]*s/i.test(pat.slice(re.lastIndex));
    switch (low) {
      case 'yyyy': out += String(p.y).padStart(4, '0'); break;
      case 'yy': out += String(p.y % 100).padStart(2, '0'); break;
      case 'mmmmm': out += fmt({ month: 'narrow' }); break;
      case 'mmmm': out += fmt({ month: 'long' }); break;
      case 'mmm': out += fmt({ month: 'short' }); break;
      case 'mm': out += lastWasHour || nextIsSec ? String(p.M).padStart(2, '0') : String(p.m).padStart(2, '0'); break;
      case 'm': out += lastWasHour || nextIsSec ? String(p.M) : String(p.m); break;
      case 'dddd': out += fmt({ weekday: 'long' }); break;
      case 'ddd': out += fmt({ weekday: 'short' }); break;
      case 'dd': out += String(p.d).padStart(2, '0'); break;
      case 'd': out += String(p.d); break;
      case 'hh': out += String(ampm ? ((p.H + 11) % 12) + 1 : p.H).padStart(2, '0'); break;
      case 'h': out += String(ampm ? ((p.H + 11) % 12) + 1 : p.H); break;
      case 'ss': out += String(p.S).padStart(2, '0'); break;
      case 's': out += String(p.S); break;
      case 'am/pm': out += p.H < 12 ? 'AM' : 'PM'; break;
      case 'a/p': out += p.H < 12 ? 'A' : 'P'; break;
      default: out += tok;
    }
    if (low === 'hh' || low === 'h') lastWasHour = true;
    else if (!/^[:\s]$/.test(tok)) lastWasHour = false;
  }
  return out;
}

/* ================================================================ statistics helpers */

function erf(x) {
  // Abramowitz-Stegun 7.1.26 with higher precision series for small |x|
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x);
  if (x < 2) {
    let sum = x, term = x;
    for (let n = 1; n < 60; n++) { term *= -x * x / n; sum += term / (2 * n + 1); }
    return sign * 2 / Math.sqrt(Math.PI) * sum;
  }
  const t = 1 / (1 + 0.5 * x);
  const y = 1 - t * Math.exp(-x * x - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 +
    t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))));
  return sign * y;
}
const normCdf = (z) => 0.5 * (1 + erf(z / Math.SQRT2));
const normPdf = (z) => Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI);
function normInv(p) {
  if (p <= 0 || p >= 1) return NaN;
  const a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
  const b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
  const c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
  const d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];
  const lo = 0.02425, hi = 1 - lo;
  let q, r, x;
  if (p < lo) { q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  else if (p <= hi) { q = p - 0.5; r = q * q; x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
  else { q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  // one Newton step for full precision
  const e = normCdf(x) - p;
  return x - e / normPdf(x);
}
function lnGamma(z) {
  const g = 7, coef = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61503916999185,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  if (z < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * z)) - lnGamma(1 - z);
  z -= 1;
  let x = coef[0];
  for (let i = 1; i < g + 2; i++) x += coef[i] / (z + i);
  const t = z + g + 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(x);
}
function combin(n, k) {
  if (k < 0 || k > n) return 0;
  return Math.round(Math.exp(lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1)));
}
function pairs(xs, ys) {
  const X = [...flat([xs])].map((x) => x.v), Y = [...flat([ys])].map((x) => x.v);
  if (X.length !== Y.length) throw ERR['#N/A'];
  const out = [];
  for (let i = 0; i < X.length; i++) if (typeof X[i] === 'number' && typeof Y[i] === 'number') out.push([X[i], Y[i]]);
  return out;
}
function linreg(ys, xs) {
  const p = pairs(xs, ys);
  if (p.length < 2) throw ERR['#DIV/0!'];
  const mx = p.reduce((s, [x]) => s + x, 0) / p.length, my = p.reduce((s, [, y]) => s + y, 0) / p.length;
  let sxy = 0, sxx = 0, syy = 0;
  for (const [x, y] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; }
  if (sxx === 0) throw ERR['#DIV/0!'];
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, r: sxy / Math.sqrt(sxx * syy), n: p.length, mx, my, sxx, sxy, syy };
}
function percentileInc(arr, k) {
  const n = numbers([arr]).sort((a, b) => a - b);
  if (!n.length || k < 0 || k > 1) return ERR['#NUM!'];
  const pos = k * (n.length - 1), lo = Math.floor(pos);
  return n[lo] + (pos - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]);
}
function percentileExc(arr, k) {
  const n = numbers([arr]).sort((a, b) => a - b);
  const pos = k * (n.length + 1) - 1;
  if (!n.length || pos < 0 || pos > n.length - 1) return ERR['#NUM!'];
  const lo = Math.floor(pos);
  return n[lo] + (pos - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]);
}

/* ================================================================ financial helpers */

function fv(rate, nper, pmt, pv, type) {
  if (rate === 0) return -(pv + pmt * nper);
  const f = Math.pow(1 + rate, nper);
  return -(pv * f + pmt * (1 + rate * type) * (f - 1) / rate);
}
function pmtCalc(rate, nper, pv, fvv, type) {
  if (rate === 0) return -(pv + fvv) / nper;
  const f = Math.pow(1 + rate, nper);
  return -(rate * (fvv + pv * f)) / ((1 + rate * type) * (f - 1));
}
function ipmt(rate, per, nper, pv, fvv, type) {
  const payment = pmtCalc(rate, nper, pv, fvv, type);
  let interest;
  if (per === 1) interest = type === 1 ? 0 : -pv;
  else interest = type === 1 ? fv(rate, per - 2, payment, pv, 1) - payment : fv(rate, per - 1, payment, pv, 0);
  return interest * rate;
}
function solve(f, guess) {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x), h = 1e-6 * Math.max(1, Math.abs(x));
    const d = (f(x + h) - y) / h;
    if (!isFinite(d) || d === 0) break;
    const nx = x - y / d;
    if (Math.abs(nx - x) < 1e-10) return nx;
    x = nx;
  }
  return Math.abs(f(x)) < 1e-6 ? x : NaN;
}

/* ================================================================ functions */

const F = Object.create(null);
// Receive (argNodes, ev, ctx, engine) instead of values
const LAZY = new Set(['IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'CHOOSE', 'AND', 'OR', 'ROW', 'COLUMN', 'ISREF', 'LET',
  'HYPERLINK', 'OFFSET', 'INDIRECT', 'FORMULATEXT', 'ISFORMULA', 'SHEET', 'SHEETS', 'CELL']);
// Receive errors as values instead of failing straight away
const RAW = new Set(['ISBLANK', 'ISERROR', 'ISERR', 'ISNA', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'COUNT', 'COUNTA',
  'COUNTBLANK', 'ERROR.TYPE', 'N', 'T', 'TYPE', 'AGGREGATE']);

/** Lifts a function over its first argument when that is an array or a multi-cell range. */
function L(fn) {
  return (a, ...rest) => {
    if (isMulti(a)) {
      return mapMatrix(toMatrix(a), (x) => {
        let v;
        try { v = fn(x, ...rest); } catch (e) { if (isErr(e)) return e; throw e; }
        // text made for every element counts against the work budget, and is not longer than a cell can hold
        if (typeof v === 'string') { tick(v.length); if (v.length > MAX_TEXT) return ERR['#VALUE!']; }
        return v;
      });
    }
    return fn(a, ...rest);
  };
}
const N1 = (fn) => L((a) => { const x = toNum(a); return isErr(x) ? x : finite(fn(x)); });
const S1 = (fn) => L((a) => { const x = toStr(a); return isErr(x) ? x : fn(x); });
function roundTo(x, digits, mode) {
  const f = Math.pow(10, digits);
  const y = x * f;
  let r;
  if (mode === 'up') r = y < 0 ? -Math.ceil(-y - 1e-9) : Math.ceil(y - 1e-9);
  else if (mode === 'down') r = y < 0 ? -Math.floor(-y + 1e-9) : Math.floor(y + 1e-9);
  else r = Math.sign(y) * Math.round(Math.abs(y) + 1e-9);
  return r / f;
}
const sum = (a) => a.reduce((s, x) => s + x, 0);
const mean = (a) => sum(a) / a.length;

// ---- math
F.SUM = (...a) => sum(numbers(a));
F.SUMSQ = (...a) => numbers(a).reduce((s, x) => s + x * x, 0);
F.PRODUCT = (...a) => numbers(a).reduce((s, x) => s * x, 1);
F.ROUND = L((x, d) => roundTo(numArg(x), numArg(d, 0), 'half'));
F.ROUNDUP = L((x, d) => roundTo(numArg(x), numArg(d, 0), 'up'));
F.ROUNDDOWN = L((x, d) => roundTo(numArg(x), numArg(d, 0), 'down'));
F.TRUNC = L((x, d) => roundTo(numArg(x), numArg(d, 0), 'down'));
F.INT = N1(Math.floor);
F.ABS = N1(Math.abs);
F.SIGN = N1(Math.sign);
F.SQRT = N1((x) => (x < 0 ? ERR['#NUM!'] : Math.sqrt(x)));
F.SQRTPI = N1((x) => (x < 0 ? ERR['#NUM!'] : Math.sqrt(x * Math.PI)));
F.POWER = L((a, b) => finite(Math.pow(numArg(a), numArg(b))));
F.EXP = N1(Math.exp);
F.LN = N1((x) => (x <= 0 ? ERR['#NUM!'] : Math.log(x)));
F.LOG10 = N1((x) => (x <= 0 ? ERR['#NUM!'] : Math.log10(x)));
F.LOG = L((x, b) => { const v = numArg(x), base = numArg(b, 10); return v <= 0 || base <= 0 || base === 1 ? ERR['#NUM!'] : Math.log(v) / Math.log(base); });
F.MOD = L((a, b) => { const x = numArg(a), y = numArg(b); if (y === 0) return ERR['#DIV/0!']; const r = x - y * Math.floor(x / y); return Math.abs(r) < 1e-12 ? 0 : r; });
F.QUOTIENT = (a, b) => { const y = numArg(b); return y === 0 ? ERR['#DIV/0!'] : Math.trunc(numArg(a) / y); };
F.PI = () => Math.PI;
F.RAND = () => Math.random();
F.RANDBETWEEN = (a, b) => { const lo = Math.ceil(numArg(a)), hi = Math.floor(numArg(b)); return lo > hi ? ERR['#NUM!'] : lo + Math.floor(Math.random() * (hi - lo + 1)); };
F.RANDARRAY = (rows, cols, min, max, whole) => {
  const r = numArg(rows, 1), c = numArg(cols, 1), lo = numArg(min, 0), hi = numArg(max, 1), int = boolArg(whole, false);
  if (r < 1 || c < 1 || hi < lo) return ERR['#VALUE!'];
  arraySize(Math.floor(r), Math.floor(c));
  return Array.from({ length: r }, () => Array.from({ length: c }, () => (int ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo))));
};
F.SEQUENCE = (rows, cols, start, step) => {
  const r = Math.floor(numArg(rows, 1)), c = Math.floor(numArg(cols, 1)), s = numArg(start, 1), st = numArg(step, 1);
  if (r < 1 || c < 1 || r * c > 1e6) return ERR['#VALUE!'];
  tick(r * c);
  return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (__, j) => s + (i * c + j) * st));
};
F.CEILING = (x, s) => { const v = numArg(x), sig = numArg(s, 1); return sig === 0 ? 0 : Math.ceil(v / sig - 1e-12) * sig; };
F['CEILING.MATH'] = (x, s, mode) => { const v = numArg(x), sig = Math.abs(numArg(s, 1)); if (!sig) return 0; return v < 0 && numArg(mode, 0) ? Math.floor(v / sig) * sig : Math.ceil(v / sig - 1e-12) * sig; };
F['CEILING.PRECISE'] = F['CEILING.MATH'];
F.FLOOR = (x, s) => { const v = numArg(x), sig = numArg(s, 1); return sig === 0 ? ERR['#DIV/0!'] : Math.floor(v / sig + 1e-12) * sig; };
F['FLOOR.MATH'] = (x, s, mode) => { const v = numArg(x), sig = Math.abs(numArg(s, 1)); if (!sig) return 0; return v < 0 && numArg(mode, 0) ? Math.ceil(v / sig) * sig : Math.floor(v / sig + 1e-12) * sig; };
F['FLOOR.PRECISE'] = F['FLOOR.MATH'];
F.MROUND = (x, m) => { const v = numArg(x), mul = numArg(m); if (mul === 0) return 0; if (Math.sign(v) * Math.sign(mul) < 0) return ERR['#NUM!']; return roundTo(v / mul, 0, 'half') * mul; };
F.EVEN = N1((x) => { const c = Math.ceil(Math.abs(x)); return Math.sign(x || 1) * (c % 2 ? c + 1 : c); });
F.ODD = N1((x) => { const c = Math.ceil(Math.abs(x)); return Math.sign(x || 1) * (c % 2 ? c : c + 1); });
// 170! and 300!! are the largest that fit in a number; larger arguments give #NUM! before looping (the whole part
// counts, as in Excel: FACT(170.5) is 170!)
F.FACT = N1((x) => { if (x < 0 || Math.floor(x) > 170) return ERR['#NUM!']; let r = 1; for (let i = 2; i <= Math.floor(x); i++) r *= i; return r; });
F.FACTDOUBLE = N1((x) => { if (x < -1 || Math.floor(x) > 300) return ERR['#NUM!']; let r = 1; for (let i = Math.floor(x); i > 1; i -= 2) r *= i; return r; });
F.COMBIN = (n, k) => { const a = Math.floor(numArg(n)), b = Math.floor(numArg(k)); return a < 0 || b < 0 || b > a ? ERR['#NUM!'] : combin(a, b); };
F.COMBINA = (n, k) => { const a = Math.floor(numArg(n)), b = Math.floor(numArg(k)); return a < 0 || b < 0 ? ERR['#NUM!'] : combin(a + b - 1, b); };
F.PERMUT = (n, k) => { const a = Math.floor(numArg(n)), b = Math.floor(numArg(k)); if (a < 0 || b < 0 || b > a) return ERR['#NUM!']; let r = 1; for (let i = 0; i < b && isFinite(r); i++) r *= a - i; return r; };
F.GCD = (...a) => numbers(a).map(Math.floor).reduce((x, y) => { while (y) [x, y] = [y, x % y]; return x; }, 0);
F.LCM = (...a) => numbers(a).map(Math.floor).reduce((x, y) => { if (!x || !y) return 0; let g = x, h = y; while (h) [g, h] = [h, g % h]; return (x * y) / g; }, 1);
F.SIN = N1(Math.sin); F.COS = N1(Math.cos); F.TAN = N1(Math.tan);
F.ASIN = N1(Math.asin); F.ACOS = N1(Math.acos); F.ATAN = N1(Math.atan);
F.SINH = N1(Math.sinh); F.COSH = N1(Math.cosh); F.TANH = N1(Math.tanh);
F.ASINH = N1(Math.asinh); F.ACOSH = N1(Math.acosh); F.ATANH = N1(Math.atanh);
F.COT = N1((x) => 1 / Math.tan(x)); F.SEC = N1((x) => 1 / Math.cos(x)); F.CSC = N1((x) => 1 / Math.sin(x));
F.ATAN2 = (x, y) => Math.atan2(numArg(y), numArg(x));
F.DEGREES = N1((x) => (x * 180) / Math.PI);
F.RADIANS = N1((x) => (x * Math.PI) / 180);
F.SUMPRODUCT = (...a) => {
  const [rows, cols] = rangeSize(a[0]);
  for (const r of a) { const [rr, cc] = rangeSize(r); if (rr !== rows || cc !== cols) return ERR['#VALUE!']; }
  let s = 0;
  for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
    let p = 1;
    for (const r of a) { const v = rangeGet(r, i, j); if (isErr(v)) throw v; p *= typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0; }
    s += p;
  }
  return s;
};
F.SUMX2MY2 = (x, y) => pairs(x, y).reduce((s, [a, b]) => s + a * a - b * b, 0);
F.SUMX2PY2 = (x, y) => pairs(x, y).reduce((s, [a, b]) => s + a * a + b * b, 0);
F.SUMXMY2 = (x, y) => pairs(x, y).reduce((s, [a, b]) => s + (a - b) ** 2, 0);
F.MMULT = (a, b) => {
  const A = toMatrix(a), B = toMatrix(b);
  if (A[0].length !== B.length) return ERR['#VALUE!'];
  arraySize(A.length, B[0].length);
  tick(A.length * B.length * B[0].length / 4); // multiply-adds are cheaper than the steps the budget counts
  return A.map((row) => B[0].map((_, j) => row.reduce((s, x, k) => s + toNum(x) * toNum(B[k][j]), 0)));
};
F.MDETERM = (a) => {
  const M = toMatrix(a).map((r) => r.map((x) => toNum(x)));
  const n = M.length; if (n !== M[0].length) return ERR['#VALUE!'];
  tick(n * n * n / 4);
  let det = 1;
  for (let i = 0; i < n; i++) {
    let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(M[k][i]) > Math.abs(M[p][i])) p = k;
    if (M[p][i] === 0) return 0;
    if (p !== i) { [M[p], M[i]] = [M[i], M[p]]; det = -det; }
    det *= M[i][i];
    for (let k = i + 1; k < n; k++) { const f = M[k][i] / M[i][i]; for (let j = i; j < n; j++) M[k][j] -= f * M[i][j]; }
  }
  return det;
};
F.MINVERSE = (a) => {
  const M = toMatrix(a).map((r) => r.map((x) => toNum(x)));
  const n = M.length; if (n !== M[0].length) return ERR['#VALUE!'];
  tick(n * n * n / 2);
  const I = M.map((r, i) => r.concat(Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
  for (let i = 0; i < n; i++) {
    let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(I[k][i]) > Math.abs(I[p][i])) p = k;
    if (Math.abs(I[p][i]) < 1e-15) return ERR['#NUM!'];
    [I[p], I[i]] = [I[i], I[p]];
    const d = I[i][i]; for (let j = 0; j < 2 * n; j++) I[i][j] /= d;
    for (let k = 0; k < n; k++) if (k !== i) { const f = I[k][i]; for (let j = 0; j < 2 * n; j++) I[k][j] -= f * I[i][j]; }
  }
  return I.map((r) => r.slice(n));
};
F.MUNIT = (n) => { const k = Math.floor(numArg(n)); arraySize(k, k); return Array.from({ length: k }, (_, i) => Array.from({ length: k }, (__, j) => (i === j ? 1 : 0))); };
F.BASE = (n, radix, len) => {
  const v = Math.floor(numArg(n)), r = numArg(radix), k = numArg(len, 0);
  if (v < 0 || r < 2 || r > 36 || !(k >= 0 && k <= 255)) return ERR['#NUM!']; // Excel allows at most 255 digits
  return v.toString(r).toUpperCase().padStart(k, '0');
};
F.DECIMAL = (t, radix) => { const r = numArg(radix); const v = parseInt(strArg(t), r); return isNaN(v) ? ERR['#NUM!'] : v; };
F.DEC2BIN = (n, p) => F.BASE(n, 2, p); F.DEC2OCT = (n, p) => F.BASE(n, 8, p); F.DEC2HEX = (n, p) => F.BASE(n, 16, p);
F.BIN2DEC = (t) => F.DECIMAL(t, 2); F.OCT2DEC = (t) => F.DECIMAL(t, 8); F.HEX2DEC = (t) => F.DECIMAL(t, 16);
F.ROMAN = (n) => {
  let v = Math.floor(numArg(n)); if (v < 0 || v > 3999) return ERR['#VALUE!'];
  const map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = ''; for (const [k, r] of map) while (v >= k) { s += r; v -= k; } return s;
};
F.ARABIC = (t) => {
  const s = strArg(t).toUpperCase().trim(); const val = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0; for (let i = 0; i < s.length; i++) { const a = val[s[i]], b = val[s[i + 1]] || 0; if (!a) return ERR['#VALUE!']; total += a < b ? -a : a; } return total;
};
/**
 * The cells of the range arguments without the rows that SUBTOTAL/AGGREGATE leave out, as Excel does: rows hidden
 * by a filter, manually hidden rows (when asked), and cells that are themselves SUBTOTAL or AGGREGATE results.
 */
function visibleCells(args, opts) {
  return args.map((x) => {
    if (!(x instanceof Range)) return x;
    const book = x.engine.book;
    const out = [];
    for (let i = 0; i < x.rows; i++) {
      const r = x.r1 + i;
      const state = book && book.rowHidden ? book.rowHidden(x.sheet, r) : null;
      if ((state === 'filter' && opts.filtered) || (state === 'manual' && opts.manual)) continue;
      // a row whose filter value is too heavy to compute: whether it is hidden is not known (see Engine.slowRead)
      if (state === 'unknown' && opts.filtered) x.engine.slowRead();
      const row = [];
      for (let j = 0; j < x.cols; j++) {
        const raw = x.engine.raw(x.sheet, r, x.c1 + j);
        if (opts.nested && typeof raw === 'string' && raw[0] === '=' && /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(raw)) { row.push(null); continue; }
        row.push(x.get(i, j));
      }
      out.push(row);
    }
    return out.length ? out : [[null]];
  });
}
F.SUBTOTAL = (fnum, ...a) => {
  const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
  const n = Math.trunc(numArg(fnum));
  const name = map[n % 100];
  if (!name || (n > 11 && (n < 101 || n > 111))) return ERR['#VALUE!'];
  // 1-11 leave out rows hidden by a filter; 101-111 also manually hidden rows
  return F[name](...visibleCells(a, { filtered: true, manual: n > 100, nested: true }));
};
F.AGGREGATE = (fnum, options, ...a) => {
  const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' };
  const name = map[numArg(fnum)];
  if (!name) return ERR['#VALUE!'];
  const opt = numArg(options, 0);
  const skipHidden = [1, 3, 5, 7].includes(opt), skipErrors = [2, 3, 6, 7].includes(opt), skipNested = opt <= 3;
  // functions 14-19 take the array first and k second
  const data = name === 'LARGE' || name === 'SMALL' || name.startsWith('PERCENTILE') || name.startsWith('QUARTILE') ? a.slice(0, 1) : a;
  const rest = data === a ? [] : a.slice(1);
  let cleaned = visibleCells(data, { filtered: skipHidden, manual: skipHidden, nested: skipNested });
  if (skipErrors) cleaned = cleaned.map((x) => (isMatrix(x) ? mapMatrix(x, (v) => (isErr(v) ? null : v)) : x));
  return F[name](...cleaned, ...rest);
};

// ---- statistics
F.AVERAGE = (...a) => { const n = numbers(a); return n.length ? mean(n) : ERR['#DIV/0!']; };
F.AVERAGEA = (...a) => { const n = numbers(a, { countLogical: true }); return n.length ? mean(n) : ERR['#DIV/0!']; };
F.MIN = (...a) => { const n = numbers(a); return n.length ? Math.min(...n) : 0; };
F.MAX = (...a) => { const n = numbers(a); return n.length ? Math.max(...n) : 0; };
F.MINA = (...a) => { const n = numbers(a, { countLogical: true }); return n.length ? Math.min(...n) : 0; };
F.MAXA = (...a) => { const n = numbers(a, { countLogical: true }); return n.length ? Math.max(...n) : 0; };
F.MEDIAN = (...a) => {
  const n = numbers(a).sort((x, y) => x - y);
  if (!n.length) return ERR['#NUM!'];
  const h = n.length >> 1;
  return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2;
};
F['MODE.SNGL'] = (...a) => {
  const counts = new Map(); let best = null, bc = 1;
  for (const x of numbers(a)) { const c = (counts.get(x) || 0) + 1; counts.set(x, c); if (c > bc) { bc = c; best = x; } }
  return best === null ? ERR['#N/A'] : best;
};
F.MODE = F['MODE.SNGL'];
F['MODE.MULT'] = (...a) => {
  const counts = new Map(); for (const x of numbers(a)) counts.set(x, (counts.get(x) || 0) + 1);
  const max = Math.max(...counts.values()); if (max < 2) return ERR['#N/A'];
  return [...counts].filter(([, c]) => c === max).map(([x]) => [x]);
};
function variance(a, sample, opts) {
  const n = numbers(a, opts);
  if (n.length < (sample ? 2 : 1)) return ERR['#DIV/0!'];
  const m = mean(n);
  return n.reduce((s, x) => s + (x - m) * (x - m), 0) / (n.length - (sample ? 1 : 0));
}
F['VAR.S'] = (...a) => variance(a, true); F.VAR = F['VAR.S'];
F['VAR.P'] = (...a) => variance(a, false); F.VARP = F['VAR.P'];
F.VARA = (...a) => variance(a, true, { countLogical: true }); F.VARPA = (...a) => variance(a, false, { countLogical: true });
F['STDEV.S'] = (...a) => { const v = variance(a, true); return isErr(v) ? v : Math.sqrt(v); }; F.STDEV = F['STDEV.S'];
F['STDEV.P'] = (...a) => { const v = variance(a, false); return isErr(v) ? v : Math.sqrt(v); }; F.STDEVP = F['STDEV.P'];
F.STDEVA = (...a) => { const v = variance(a, true, { countLogical: true }); return isErr(v) ? v : Math.sqrt(v); };
F.STDEVPA = (...a) => { const v = variance(a, false, { countLogical: true }); return isErr(v) ? v : Math.sqrt(v); };
F.AVEDEV = (...a) => { const n = numbers(a); if (!n.length) return ERR['#NUM!']; const m = mean(n); return mean(n.map((x) => Math.abs(x - m))); };
F.DEVSQ = (...a) => { const n = numbers(a); const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0); };
F.GEOMEAN = (...a) => { const n = numbers(a); if (!n.length || n.some((x) => x <= 0)) return ERR['#NUM!']; return Math.exp(mean(n.map(Math.log))); };
F.HARMEAN = (...a) => { const n = numbers(a); if (!n.length || n.some((x) => x <= 0)) return ERR['#NUM!']; return n.length / sum(n.map((x) => 1 / x)); };
F.SKEW = (...a) => { const n = numbers(a); const k = n.length; if (k < 3) return ERR['#DIV/0!']; const m = mean(n); const s = Math.sqrt(variance(n, true)); return k / ((k - 1) * (k - 2)) * sum(n.map((x) => ((x - m) / s) ** 3)); };
F.KURT = (...a) => { const n = numbers(a); const k = n.length; if (k < 4) return ERR['#DIV/0!']; const m = mean(n); const s = Math.sqrt(variance(n, true)); return k * (k + 1) / ((k - 1) * (k - 2) * (k - 3)) * sum(n.map((x) => ((x - m) / s) ** 4)) - 3 * (k - 1) ** 2 / ((k - 2) * (k - 3)); };
F.LARGE = (r, k) => { const n = numbers([r]).sort((x, y) => y - x); const i = Math.ceil(numArg(k)) - 1; return i >= 0 && i < n.length ? n[i] : ERR['#NUM!']; };
F.SMALL = (r, k) => { const n = numbers([r]).sort((x, y) => x - y); const i = Math.ceil(numArg(k)) - 1; return i >= 0 && i < n.length ? n[i] : ERR['#NUM!']; };
F['RANK.EQ'] = (v, r, order) => {
  const x = numArg(v); const n = numbers([r]);
  if (!n.includes(x)) return ERR['#N/A'];
  const asc = numArg(order, 0) !== 0;
  return n.filter((y) => (asc ? y < x : y > x)).length + 1;
};
F.RANK = F['RANK.EQ'];
F['RANK.AVG'] = (v, r, order) => {
  const x = numArg(v); const n = numbers([r]);
  const same = n.filter((y) => y === x).length; if (!same) return ERR['#N/A'];
  const asc = numArg(order, 0) !== 0;
  return n.filter((y) => (asc ? y < x : y > x)).length + (same + 1) / 2;
};
F['PERCENTILE.INC'] = (r, k) => percentileInc(r, numArg(k)); F.PERCENTILE = F['PERCENTILE.INC'];
F['PERCENTILE.EXC'] = (r, k) => percentileExc(r, numArg(k));
F['QUARTILE.INC'] = (r, q) => { const k = Math.floor(numArg(q)); return k < 0 || k > 4 ? ERR['#NUM!'] : percentileInc(r, k / 4); }; F.QUARTILE = F['QUARTILE.INC'];
F['QUARTILE.EXC'] = (r, q) => { const k = Math.floor(numArg(q)); return k <= 0 || k >= 4 ? ERR['#NUM!'] : percentileExc(r, k / 4); };
F['PERCENTRANK.INC'] = (r, x, sig) => {
  const n = numbers([r]).sort((a, b) => a - b); const v = numArg(x), d = numArg(sig, 3);
  if (!n.length || v < n[0] || v > n[n.length - 1]) return ERR['#N/A'];
  let i = 0; while (i < n.length - 1 && n[i + 1] <= v) i++;
  let rank = i / (n.length - 1);
  if (n[i] !== v) rank += ((v - n[i]) / (n[i + 1] - n[i])) / (n.length - 1);
  return Math.floor(rank * 10 ** d) / 10 ** d;
};
F.PERCENTRANK = F['PERCENTRANK.INC'];
F.COUNT = (...a) => { let c = 0; for (const { v, ref } of flat(a)) { if (typeof v === 'number' || (!ref && (typeof v === 'boolean' || (typeof v === 'string' && textToNumber(v) !== null && v.trim() !== '')))) c++; } return c; };
F.COUNTA = (...a) => { let c = 0; for (const { v } of flat(a)) if (v !== null && v !== undefined) c++; return c; };
F.COUNTBLANK = (...a) => { let c = 0; for (const { v } of flat(a)) if (v === null || v === '') c++; return c; };
F.COUNTIF = (r, crit) => { const f = makeCriteria(crit); let c = 0; for (const { v } of flat([r])) if (f(v)) c++; return c; };
F.COUNTIFS = (...a) => ifsAggregate(a[0], a).length;
F.SUMIF = (r, crit, sumRange) => {
  const f = makeCriteria(crit);
  const [rows, cols] = rangeSize(r);
  let s = 0;
  for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
    if (!f(rangeGet(r, i, j))) continue;
    const v = sumRange !== undefined && sumRange !== null ? sameShapeGet(sumRange, i, j) : rangeGet(r, i, j);
    if (isErr(v)) throw v;
    if (typeof v === 'number') s += v;
  }
  return s;
};
F.SUMIFS = (sumRange, ...p) => ifsAggregate(sumRange, p).reduce((s, v) => { if (isErr(v)) throw v; return typeof v === 'number' ? s + v : s; }, 0);
F.AVERAGEIF = (r, crit, avgRange) => {
  const f = makeCriteria(crit);
  const [rows, cols] = rangeSize(r);
  let s = 0, c = 0;
  for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
    if (!f(rangeGet(r, i, j))) continue;
    const v = avgRange !== undefined && avgRange !== null ? sameShapeGet(avgRange, i, j) : rangeGet(r, i, j);
    if (typeof v === 'number') { s += v; c++; }
  }
  return c ? s / c : ERR['#DIV/0!'];
};
F.AVERAGEIFS = (r, ...p) => { const n = ifsAggregate(r, p).filter((v) => typeof v === 'number'); return n.length ? mean(n) : ERR['#DIV/0!']; };
F.MAXIFS = (r, ...p) => { const n = ifsAggregate(r, p).filter((v) => typeof v === 'number'); return n.length ? Math.max(...n) : 0; };
F.MINIFS = (r, ...p) => { const n = ifsAggregate(r, p).filter((v) => typeof v === 'number'); return n.length ? Math.min(...n) : 0; };
F.CORREL = (a, b) => linreg(a, b).r; F.PEARSON = F.CORREL;
F.RSQ = (y, x) => linreg(y, x).r ** 2;
F.SLOPE = (y, x) => linreg(y, x).slope;
F.INTERCEPT = (y, x) => linreg(y, x).intercept;
F['FORECAST.LINEAR'] = (x, y, xs) => { const l = linreg(y, xs); return l.intercept + l.slope * numArg(x); }; F.FORECAST = F['FORECAST.LINEAR'];
F.STEYX = (y, x) => { const l = linreg(y, x); return l.n < 3 ? ERR['#DIV/0!'] : Math.sqrt((l.syy - l.sxy ** 2 / l.sxx) / (l.n - 2)); };
F.TREND = (ys, xs, newXs) => { const l = linreg(ys, xs); return mapMatrix(toMatrix(newXs === undefined ? xs : newXs), (x) => l.intercept + l.slope * toNum(x)); };
F['COVARIANCE.P'] = (a, b) => { const p = pairs(a, b); if (!p.length) return ERR['#DIV/0!']; const mx = mean(p.map((x) => x[0])), my = mean(p.map((x) => x[1])); return sum(p.map(([x, y]) => (x - mx) * (y - my))) / p.length; };
F.COVAR = F['COVARIANCE.P'];
F['COVARIANCE.S'] = (a, b) => { const p = pairs(a, b); if (p.length < 2) return ERR['#DIV/0!']; const mx = mean(p.map((x) => x[0])), my = mean(p.map((x) => x[1])); return sum(p.map(([x, y]) => (x - mx) * (y - my))) / (p.length - 1); };
F.STANDARDIZE = (x, m, s) => { const sd = numArg(s); return sd <= 0 ? ERR['#NUM!'] : (numArg(x) - numArg(m)) / sd; };
F['NORM.DIST'] = (x, m, s, cum) => { const sd = numArg(s); if (sd <= 0) return ERR['#NUM!']; const z = (numArg(x) - numArg(m)) / sd; return boolArg(cum, true) ? normCdf(z) : normPdf(z) / sd; };
F.NORMDIST = F['NORM.DIST'];
F['NORM.S.DIST'] = (z, cum) => (boolArg(cum, true) ? normCdf(numArg(z)) : normPdf(numArg(z)));
F.NORMSDIST = (z) => normCdf(numArg(z));
F['NORM.INV'] = (p, m, s) => { const v = numArg(p), sd = numArg(s); if (v <= 0 || v >= 1 || sd <= 0) return ERR['#NUM!']; return numArg(m) + sd * normInv(v); };
F.NORMINV = F['NORM.INV'];
F['NORM.S.INV'] = (p) => { const v = numArg(p); return v <= 0 || v >= 1 ? ERR['#NUM!'] : normInv(v); };
F.NORMSINV = F['NORM.S.INV'];
F['BINOM.DIST'] = (k, n, p, cum) => {
  const x = Math.floor(numArg(k)), t = Math.floor(numArg(n)), pr = numArg(p);
  if (x < 0 || x > t || pr < 0 || pr > 1) return ERR['#NUM!'];
  const pmf = (i) => combin(t, i) * pr ** i * (1 - pr) ** (t - i);
  if (!boolArg(cum, false)) return pmf(x);
  if (x > 1e6) return ERR['#NUM!'];
  tick(x + 1);
  let s = 0; for (let i = 0; i <= x; i++) s += pmf(i); return s;
};
F.BINOMDIST = F['BINOM.DIST'];
F['POISSON.DIST'] = (x, m, cum) => {
  const k = Math.floor(numArg(x)), lambda = numArg(m);
  if (k < 0 || lambda < 0) return ERR['#NUM!'];
  const pmf = (i) => Math.exp(-lambda + i * Math.log(lambda || 1e-300) - lnGamma(i + 1));
  if (!boolArg(cum, false)) return pmf(k);
  if (k > 1e6) return ERR['#NUM!'];
  tick(k + 1);
  let s = 0; for (let i = 0; i <= k; i++) s += pmf(i); return s;
};
F.POISSON = F['POISSON.DIST'];
F['EXPON.DIST'] = (x, l, cum) => { const v = numArg(x), lambda = numArg(l); if (v < 0 || lambda <= 0) return ERR['#NUM!']; return boolArg(cum, true) ? 1 - Math.exp(-lambda * v) : lambda * Math.exp(-lambda * v); };
F.EXPONDIST = F['EXPON.DIST'];
F.GAMMALN = N1((x) => (x <= 0 ? ERR['#NUM!'] : lnGamma(x))); F['GAMMALN.PRECISE'] = F.GAMMALN;
F.GAMMA = N1((x) => (x <= 0 && Number.isInteger(x) ? ERR['#NUM!'] : Math.exp(lnGamma(x)) * (x < 0 && Math.floor(-x) % 2 === 0 ? -1 : 1)));
F.FISHER = N1((x) => (Math.abs(x) >= 1 ? ERR['#NUM!'] : 0.5 * Math.log((1 + x) / (1 - x))));
F.FISHERINV = N1((y) => (Math.exp(2 * y) - 1) / (Math.exp(2 * y) + 1));
F['CONFIDENCE.NORM'] = (alpha, s, n) => normInv(1 - numArg(alpha) / 2) * numArg(s) / Math.sqrt(numArg(n)); F.CONFIDENCE = F['CONFIDENCE.NORM'];
F.FREQUENCY = (data, bins) => {
  const d = numbers([data]), b = numbers([bins]).sort((x, y) => x - y);
  const out = b.map(() => 0).concat([0]);
  tick(d.length + b.length);
  for (const x of d) {
    // the first bin that is not smaller than x (binary search over the sorted bins)
    let lo = 0, hi = b.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (x > b[mid]) lo = mid + 1; else hi = mid; }
    out[lo]++;
  }
  return out.map((x) => [x]);
};
F.PROB = (xs, ps, lo, hi) => { const a = numArg(lo), b = hi === undefined ? a : numArg(hi); return pairs(xs, ps).reduce((s, [x, p]) => (x >= a && x <= b ? s + p : s), 0); };

// ---- financial
F.PMT = (r, n, pv, fvv, type) => pmtCalc(numArg(r), numArg(n), numArg(pv), numArg(fvv, 0), numArg(type, 0) ? 1 : 0);
F.FV = (r, n, pmt, pv, type) => fv(numArg(r), numArg(n), numArg(pmt, 0), numArg(pv, 0), numArg(type, 0) ? 1 : 0);
F.PV = (r, n, pmt, fvv, type) => {
  const rate = numArg(r), nper = numArg(n), p = numArg(pmt, 0), f = numArg(fvv, 0), t = numArg(type, 0) ? 1 : 0;
  if (rate === 0) return -(f + p * nper);
  const g = Math.pow(1 + rate, nper);
  return -(f + p * (1 + rate * t) * (g - 1) / rate) / g;
};
F.NPER = (r, pmt, pv, fvv, type) => {
  const rate = numArg(r), p = numArg(pmt), v = numArg(pv), f = numArg(fvv, 0), t = numArg(type, 0) ? 1 : 0;
  if (rate === 0) return p === 0 ? ERR['#NUM!'] : -(v + f) / p;
  const num = p * (1 + rate * t) - f * rate, den = v * rate + p * (1 + rate * t);
  return finite(Math.log(num / den) / Math.log(1 + rate));
};
F.RATE = (n, pmt, pv, fvv, type, guess) => {
  const nper = numArg(n), p = numArg(pmt), v = numArg(pv), f = numArg(fvv, 0), t = numArg(type, 0) ? 1 : 0;
  const r = solve((rate) => (Math.abs(rate) < 1e-12 ? v + p * nper + f : v * Math.pow(1 + rate, nper) + p * (1 + rate * t) * (Math.pow(1 + rate, nper) - 1) / rate + f), numArg(guess, 0.1));
  return isNaN(r) ? ERR['#NUM!'] : r;
};
F.IPMT = (r, per, n, pv, fvv, type) => ipmt(numArg(r), numArg(per), numArg(n), numArg(pv), numArg(fvv, 0), numArg(type, 0) ? 1 : 0);
F.PPMT = (r, per, n, pv, fvv, type) => {
  const t = numArg(type, 0) ? 1 : 0;
  return pmtCalc(numArg(r), numArg(n), numArg(pv), numArg(fvv, 0), t) - ipmt(numArg(r), numArg(per), numArg(n), numArg(pv), numArg(fvv, 0), t);
};
/** The whole periods start..end of CUMIPMT/CUMPRINC (Excel truncates them), or null when they are not valid. */
function periods(s, e) {
  const a = Math.trunc(numArg(s)), b = Math.trunc(numArg(e));
  if (!(a >= 1 && b >= a && b <= 1e9)) return null;
  tick(b - a + 1);
  return [a, b];
}
F.CUMIPMT = (r, n, pv, s, e, type) => {
  const se = periods(s, e); if (!se) return ERR['#NUM!'];
  let t = 0; for (let i = se[0]; i <= se[1]; i++) t += ipmt(numArg(r), i, numArg(n), numArg(pv), 0, numArg(type)); return t;
};
F.CUMPRINC = (r, n, pv, s, e, type) => {
  const se = periods(s, e); if (!se) return ERR['#NUM!'];
  let t = 0; const ty = numArg(type);
  const p = pmtCalc(numArg(r), numArg(n), numArg(pv), 0, ty);
  for (let i = se[0]; i <= se[1]; i++) t += p - ipmt(numArg(r), i, numArg(n), numArg(pv), 0, ty);
  return t;
};
F.NPV = (r, ...vals) => { const rate = numArg(r); return numbers(vals).reduce((s, v, i) => s + v / Math.pow(1 + rate, i + 1), 0); };
F.IRR = (vals, guess) => {
  const cf = numbers([vals]);
  const r = solve((rate) => cf.reduce((s, v, i) => s + v / Math.pow(1 + rate, i), 0), numArg(guess, 0.1));
  return isNaN(r) ? ERR['#NUM!'] : r;
};
F.XNPV = (r, vals, dates) => { const rate = numArg(r); const p = pairs(vals, dates); return p.reduce((s, [v, d]) => s + v / Math.pow(1 + rate, (d - p[0][1]) / 365), 0); };
F.XIRR = (vals, dates, guess) => {
  const p = pairs(vals, dates);
  const r = solve((rate) => p.reduce((s, [v, d]) => s + v / Math.pow(1 + rate, (d - p[0][1]) / 365), 0), numArg(guess, 0.1));
  return isNaN(r) ? ERR['#NUM!'] : r;
};
F.MIRR = (vals, fr, rr) => {
  const cf = numbers([vals]), f = numArg(fr), r = numArg(rr), n = cf.length;
  const pos = cf.reduce((s, v, i) => (v > 0 ? s + v * Math.pow(1 + r, n - 1 - i) : s), 0);
  const neg = cf.reduce((s, v, i) => (v < 0 ? s + v / Math.pow(1 + f, i) : s), 0);
  return neg === 0 || pos === 0 ? ERR['#DIV/0!'] : Math.pow(-pos / neg, 1 / (n - 1)) - 1;
};
F.SLN = (c, s, l) => { const life = numArg(l); return life === 0 ? ERR['#DIV/0!'] : (numArg(c) - numArg(s)) / life; };
F.SYD = (c, s, l, per) => { const life = numArg(l); return (numArg(c) - numArg(s)) * (life - numArg(per) + 1) * 2 / (life * (life + 1)); };
F.DDB = (c, s, l, per, factor) => {
  const cost = numArg(c), salvage = numArg(s), life = numArg(l), p = numArg(per), f = numArg(factor, 2);
  if (!(life > 0 && p > 0 && p <= life && p <= 1e9)) return ERR['#NUM!'];
  tick(p);
  let book = cost, dep = 0;
  for (let i = 1; i <= p; i++) { dep = Math.min(book * f / life, Math.max(0, book - salvage)); book -= dep; }
  return dep;
};
F.DB = (c, s, l, per, month) => {
  const cost = numArg(c), salvage = numArg(s), life = numArg(l), p = numArg(per), m = numArg(month, 12);
  if (!(life > 0 && p >= 1 && p <= life + 1 && p <= 1e9 && m >= 1 && m <= 12)) return ERR['#NUM!'];
  tick(p);
  const rate = roundTo(1 - Math.pow(salvage / cost, 1 / life), 3, 'half');
  let total = 0, dep = cost * rate * m / 12;
  if (p === 1) return dep;
  total = dep;
  for (let i = 2; i <= p; i++) { dep = i === life + 1 ? (cost - total) * rate * (12 - m) / 12 : (cost - total) * rate; total += dep; }
  return dep;
};
F.EFFECT = (n, p) => { const r = numArg(n), k = Math.floor(numArg(p)); return r <= 0 || k < 1 ? ERR['#NUM!'] : Math.pow(1 + r / k, k) - 1; };
F.NOMINAL = (e, p) => { const r = numArg(e), k = Math.floor(numArg(p)); return r <= 0 || k < 1 ? ERR['#NUM!'] : k * (Math.pow(1 + r, 1 / k) - 1); };
F.RRI = (n, pv, fvv) => Math.pow(numArg(fvv) / numArg(pv), 1 / numArg(n)) - 1;
F.PDURATION = (r, pv, fvv) => (Math.log(numArg(fvv)) - Math.log(numArg(pv))) / Math.log(1 + numArg(r));

// ---- logic
function lazyIf(cond, yes, no) {
  if (isMulti(cond)) {
    return broadcast(cond, broadcast(yes, no, (a, b) => [a, b]), (c, pair) => { const b = toBool(c); return isErr(b) ? b : b ? pair[0] : pair[1]; });
  }
  const c = toBool(cond);
  if (isErr(c)) return c;
  return c ? yes : no;
}
F.IF = (args, ev) => {
  if (args.length < 2) return ERR['#VALUE!'];
  const cond = ev(args[0]);
  const thenV = () => (args[1].k === 'empty' ? 0 : ev(args[1]));
  const elseV = () => (args.length > 2 ? (args[2].k === 'empty' ? 0 : ev(args[2])) : false);
  if (isMulti(cond)) return lazyIf(cond, toMatrix(thenV()), toMatrix(elseV()));
  const c = toBool(cond);
  if (isErr(c)) return c;
  return c ? thenV() : elseV();
};
F.IFS = (args, ev) => {
  for (let i = 0; i + 1 < args.length; i += 2) {
    const c = toBool(ev(args[i]));
    if (isErr(c)) return c;
    if (c) return ev(args[i + 1]);
  }
  return ERR['#N/A'];
};
F.IFERROR = (args, ev) => {
  const v = ev(args[0]);
  if (isMulti(v)) { const alt = ev(args[1]); return mapMatrix(toMatrix(v), (x) => (isErr(x) ? scalar(alt) : x)); }
  const s = scalar(v);
  return isErr(s) ? ev(args[1]) : v;
};
F.IFNA = (args, ev) => {
  const v = ev(args[0]);
  if (isMulti(v)) { const alt = ev(args[1]); return mapMatrix(toMatrix(v), (x) => (x === ERR['#N/A'] ? scalar(alt) : x)); }
  return scalar(v) === ERR['#N/A'] ? ev(args[1]) : v;
};
F.SWITCH = (args, ev) => {
  const v = scalar(ev(args[0]));
  if (isErr(v)) return v;
  let i = 1;
  for (; i + 1 < args.length; i += 2) if (compare(v, scalar(ev(args[i]))) === 0) return ev(args[i + 1]);
  return i < args.length ? ev(args[i]) : ERR['#N/A'];
};
F.CHOOSE = (args, ev) => {
  const i = toNum(ev(args[0]));
  if (isErr(i)) return i;
  const k = Math.floor(i);
  return k >= 1 && k < args.length ? ev(args[k]) : ERR['#VALUE!'];
};
function logicals(args, ev) {
  const out = [];
  for (const node of args) {
    const v = ev(node);
    if (v instanceof Range || isMatrix(v)) {
      for (const { v: x } of flat([v])) { if (isErr(x)) throw x; if (typeof x === 'boolean') out.push(x); else if (typeof x === 'number') out.push(x !== 0); }
    } else {
      const b = toBool(v);
      if (isErr(b)) throw b;
      out.push(b);
    }
  }
  if (!out.length) throw ERR['#VALUE!'];
  return out;
}
F.AND = (args, ev) => logicals(args, ev).every(Boolean);
F.OR = (args, ev) => logicals(args, ev).some(Boolean);
F.XOR = (...a) => { let c = 0; for (const { v } of flat(a)) { const b = toBool(v); if (isErr(b)) return b; if (b) c++; } return c % 2 === 1; };
F.NOT = L((a) => { const b = toBool(a); return isErr(b) ? b : !b; });
F.TRUE = () => true;
F.FALSE = () => false;
F.LET = (args, ev, ctx, engine) => {
  if (args.length < 3 || args.length % 2 === 0) return ERR['#VALUE!'];
  const vars = Object.create(ctx.vars || null);
  const inner = { ...ctx, vars };
  for (let i = 0; i + 1 < args.length; i += 2) {
    if (args[i].k !== 'name') return ERR['#NAME?'];
    vars[args[i].v] = engine.ev(args[i + 1], inner);
  }
  return engine.ev(args[args.length - 1], inner);
};

// ---- text
F.CONCAT = (...a) => { let s = ''; for (const { v } of flat(a)) { const t = toStr(v); if (isErr(t)) return t; s += t; if (longText(s)) return ERR['#VALUE!']; } return s; };
F.CONCATENATE = (...a) => { let s = ''; for (const x of a) { const t = toStr(x); if (isErr(t)) return t; s += t; if (longText(s)) return ERR['#VALUE!']; } return s; };
F.TEXTJOIN = (delim, ignoreEmpty, ...a) => {
  const d = strArg(delim), skip = boolArg(ignoreEmpty, true);
  const parts = [];
  let len = 0;
  for (const { v } of flat(a)) {
    const t = toStr(v); if (isErr(t)) return t; if (skip && t === '') continue;
    len += t.length + (parts.length ? d.length : 0);
    if (len > MAX_TEXT) return ERR['#VALUE!'];
    parts.push(t);
  }
  return parts.join(d);
};
F.LEN = S1((s) => s.length);
F.LEFT = L((t, n) => { const s = strArg(t), k = numArg(n, 1); return k < 0 ? ERR['#VALUE!'] : s.slice(0, k); });
F.RIGHT = L((t, n) => { const s = strArg(t), k = numArg(n, 1); return k < 0 ? ERR['#VALUE!'] : (k === 0 ? '' : s.slice(-k)); });
F.MID = L((t, start, n) => { const s = strArg(t), a = numArg(start), k = numArg(n); return a < 1 || k < 0 ? ERR['#VALUE!'] : s.substr(a - 1, k); });
F.UPPER = S1((s) => s.toLocaleUpperCase(F._locale));
F.LOWER = S1((s) => s.toLocaleLowerCase(F._locale));
F.PROPER = S1((s) => s.toLocaleLowerCase(F._locale).replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase(F._locale)));
F.TRIM = S1((s) => s.trim().replace(/ {2,}/g, ' '));
F.CLEAN = S1((s) => s.replace(/[\x00-\x1f]/g, ''));
F.SUBSTITUTE = L((t, from, to, nth) => {
  const s = strArg(t), a = strArg(from), b = strArg(to);
  if (!a) return s;
  if (nth === undefined || nth === null) {
    const parts = s.split(a);
    return s.length + (parts.length - 1) * (b.length - a.length) > MAX_TEXT ? ERR['#VALUE!'] : parts.join(b);
  }
  const k = numArg(nth); let idx = -1;
  for (let i = 0; i < k; i++) { idx = s.indexOf(a, idx + 1); if (idx < 0) return s; }
  return s.slice(0, idx) + b + s.slice(idx + a.length);
});
F.REPLACE = (t, start, n, by) => { const s = strArg(t), a = numArg(start), k = numArg(n); return s.slice(0, a - 1) + strArg(by) + s.slice(a - 1 + k); };
F.FIND = L((needle, hay, start) => { const i = strArg(hay).indexOf(strArg(needle), numArg(start, 1) - 1); return i < 0 ? ERR['#VALUE!'] : i + 1; });
F.SEARCH = L((needle, hay, start) => {
  const n = strArg(needle), h = strArg(hay), from = numArg(start, 1) - 1;
  if (/[*?~]/.test(n)) {
    if (n.length > 255) return ERR['#VALUE!']; // Excel's limit for find_text
    const i = wildcardFind(wildcardTokens(n), h, from);
    return i < 0 ? ERR['#VALUE!'] : i + 1;
  }
  const i = h.toLowerCase().indexOf(n.toLowerCase(), from);
  return i < 0 ? ERR['#VALUE!'] : i + 1;
});
F.REPT = (t, n) => {
  const s = strArg(t), k = Math.floor(numArg(n));
  // the result may not be longer than Excel's text limit (checked before building it)
  if (!(k >= 0) || (k > 0 && s.length * k > MAX_TEXT)) return ERR['#VALUE!'];
  return s === '' ? '' : s.repeat(k);
};
F.EXACT = (a, b) => strArg(a) === strArg(b);
F.CHAR = N1((x) => String.fromCharCode(x));
F.UNICHAR = N1((x) => String.fromCodePoint(x));
F.CODE = S1((s) => (s.length ? s.charCodeAt(0) : ERR['#VALUE!']));
F.UNICODE = S1((s) => (s.length ? s.codePointAt(0) : ERR['#VALUE!']));
F.VALUE = L((a) => { const v = scalar(a); if (typeof v === 'number') return v; const n = textToNumber(toStr(v)); return n === null ? ERR['#VALUE!'] : n; });
F.NUMBERVALUE = (t, dec, grp) => {
  const d = strArg(dec, '.'), g = strArg(grp, ',');
  let s = strArg(t).replace(/\s/g, '').split(g).join('');
  if (d !== '.') s = s.replace(d, '.');
  const pct = (s.match(/%/g) || []).length; s = s.replace(/%/g, '');
  const n = Number(s);
  return s === '' ? 0 : isFinite(n) ? n / 100 ** pct : ERR['#VALUE!'];
};
F.TEXT = L((v, fmt) => {
  const s = scalar(v);
  const n = typeof s === 'number' ? s : typeof s === 'string' ? textToNumber(s) : null;
  const pat = strArg(fmt);
  if (pat.length > 255) return ERR['#VALUE!']; // Excel's limit for a format
  return formatWithPattern(n === null || n === undefined ? toStr(s) : n, pat, F._locale);
});
F.DOLLAR = (x, d) => { const k = numArg(d, 2); if (!(k <= 127)) return ERR['#VALUE!']; return formatWithPattern(roundTo(numArg(x), k, 'half'), '$#,##0' + (k > 0 ? '.' + '0'.repeat(k) : '') + ';($#,##0' + (k > 0 ? '.' + '0'.repeat(k) : '') + ')', 'en-US'); };
F.FIXED = (x, d, noCommas) => { const k = numArg(d, 2); if (!(k <= 127)) return ERR['#VALUE!']; return formatWithPattern(roundTo(numArg(x), k, 'half'), (boolArg(noCommas, false) ? '0' : '#,##0') + (k > 0 ? '.' + '0'.repeat(k) : ''), 'en-US'); };
F.T = (a) => { const v = scalar(a); return typeof v === 'string' ? v : ''; };
F.N = (a) => { const v = scalar(a); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : isErr(v) ? v : 0; };
// the instance can be at most one more than the length of the text (an empty delimiter is found after every character)
const badInstance = (n, s) => n === 0 || !(Math.abs(n) <= s.length + 1);
F.TEXTBEFORE = (t, d, inst) => { const s = strArg(t), dl = strArg(d), n = numArg(inst, 1); if (badInstance(n, s)) return ERR['#VALUE!']; let i = -1; for (let k = 0; k < n; k++) { i = s.indexOf(dl, i + 1); if (i < 0) return ERR['#N/A']; } return s.slice(0, i); };
F.TEXTAFTER = (t, d, inst) => { const s = strArg(t), dl = strArg(d), n = numArg(inst, 1); if (badInstance(n, s)) return ERR['#VALUE!']; let i = -1; for (let k = 0; k < n; k++) { i = s.indexOf(dl, i + 1); if (i < 0) return ERR['#N/A']; } return s.slice(i + dl.length); };
F.TEXTSPLIT = (t, colDelim, rowDelim, ignoreEmpty) => {
  const s = strArg(t), cd = colDelim === null || colDelim === undefined ? null : strArg(colDelim), rd = rowDelim === undefined || rowDelim === null ? null : strArg(rowDelim);
  const skip = boolArg(ignoreEmpty, false);
  const rows = (rd ? s.split(rd) : [s]).filter((r) => !skip || r !== '');
  const m = rows.map((r) => (cd ? r.split(cd) : [r]).filter((x) => !skip || x !== ''));
  const w = m.reduce((x, r) => Math.max(x, r.length), 0);
  if (!w) return ERR['#CALC!'];
  arraySize(m.length, w);
  return m.map((r) => r.concat(Array(w - r.length).fill(ERR['#N/A'])));
};
F.VALUETOTEXT = (v) => toStr(v);
F.ARRAYTOTEXT = (a, format) => {
  const m = toMatrix(a), strict = numArg(format, 0);
  let len = 0;
  const text = (x) => {
    const s = strict && typeof x === 'string' ? `"${x}"` : toStr(x);
    if (isErr(s)) return s.code;
    if ((len += s.length + 2) > MAX_TEXT) throw ERR['#VALUE!'];
    return s;
  };
  if (strict) return '{' + m.map((r) => r.map(text).join(',')).join(';') + '}';
  return m.map((r) => r.map(text).join(', ')).join(', ');
};
/*
 * Regular expressions typed in formulas (REGEXTEST, REGEXEXTRACT, REGEXREPLACE) run on a small matcher that follows
 * all ways through the pattern at once (a Pike VM), not on the browser's backtracking engine, so that no pattern can
 * take exponential time: the cost is at most text length x pattern size, and it counts against the work budget.
 * It reads patterns as JavaScript does and finds the same matches: classes, groups, alternatives, greedy and lazy
 * quantifiers, anchors, \b, and groups that switch the i, m and s flags, such as (?i:abc) and (?-i:a). Back references
 * and lookarounds, which need backtracking, give #VALUE!, like patterns over 1000 characters and patterns that take
 * more than 20000 steps to match one character (such as a{30000}).
 */
const RX_MAX = 1000, RX_PROG = 20000;
const rxCache = new Map();
const isWordCh = (c) => c !== undefined && /[A-Za-z0-9_]/.test(c);
/** Whether a character ends a line for ".", "^" and "$" (the line terminators of JavaScript). */
const isLineEnd = (c) => { const k = c === undefined ? -1 : c.charCodeAt(0); return k === 10 || k === 13 || k === 0x2028 || k === 0x2029; };
const ANCHORS = new Set(['bol', 'eol', 'mbol', 'meol', 'wb']);
/* The characters that are equal without case (canon gives them the same value), for each character that has such
 * others: a case-insensitive class matches a character when it holds one of them, as in JavaScript (so "ς" is in [σ],
 * and "ı" is not in [A-Z], because "ı" and "I" are not equal without case). Made once, when first needed. */
let caseGroups = null;
function caseGroup(ch) {
  if (!caseGroups) {
    const byCanon = new Map();
    for (let k = 0; k < 0x10000; k++) {
      const x = String.fromCharCode(k), key = canon(x), g = byCanon.get(key);
      if (g) g.push(x); else byCanon.set(key, [x]);
    }
    caseGroups = new Map();
    for (const g of byCanon.values()) if (g.length > 1) for (const x of g) caseGroups.set(x, g);
  }
  return caseGroups.get(ch) || null;
}
function rxCompile(p, icase) {
  // icase, dotAll and multi are the flags at the current place in the pattern: (?ims-ims:…) groups change them
  let i = 0, ncap = 0, dotAll = false, multi = false;
  const names = Object.create(null);
  const fail = () => { throw ERR['#VALUE!']; };
  const lit = (ch) => { const cc = canon(ch); return { t: 'char', ch, test: icase ? (c) => canon(c) === cc : (c) => c === ch }; };
  const set = (test) => ({ t: 'char', test });
  const quantAt = (k) => /^\{(\d+)(,(\d*))?\}/.exec(p.slice(k, k + 24));
  function escape(inClass) {
    const ch = p[i++];
    switch (ch) {
      case undefined: return fail();
      case 'd': return set((c) => c >= '0' && c <= '9');
      case 'D': return set((c) => !(c >= '0' && c <= '9'));
      case 'w': return set(isWordCh);
      case 'W': return set((c) => !isWordCh(c));
      case 's': return set((c) => /\s/.test(c));
      case 'S': return set((c) => !/\s/.test(c));
      case 'b': return inClass ? lit('\b') : { t: 'wb', neg: false };
      case 'B': return inClass ? lit('B') : { t: 'wb', neg: true };
      case 't': return lit('\t');
      case 'n': return lit('\n');
      case 'r': return lit('\r');
      case 'f': return lit('\f');
      case 'v': return lit('\v');
      case '0': return /\d/.test(p[i] || '') ? fail() : lit('\0');
      case 'x': case 'u': {
        const len = ch === 'x' ? 2 : 4, hex = p.slice(i, i + len);
        if (!new RegExp('^[0-9a-fA-F]{' + len + '}$').test(hex)) return lit(ch);
        i += len;
        return lit(String.fromCharCode(parseInt(hex, 16)));
      }
      case 'c': return /[A-Za-z]/.test(p[i] || '') ? lit(String.fromCharCode(p[i++].charCodeAt(0) % 32)) : fail();
      case 'k': return fail(); // a back reference to a named group
      default: return /[1-9]/.test(ch) ? fail() : lit(ch); // back references are not supported
    }
  }
  function cls() {
    let neg = false;
    if (p[i] === '^') { neg = true; i++; }
    const tests = [];
    const one = () => {
      const ch = p[i++];
      if (ch !== '\\') return { ch, test: (c) => c === ch };
      const e = escape(true);
      return e.ch !== undefined ? { ch: e.ch, test: (c) => c === e.ch } : { test: e.test };
    };
    while (i < p.length && p[i] !== ']') {
      const a = one();
      if (p[i] === '-' && i + 1 < p.length && p[i + 1] !== ']') {
        i++;
        const b = one();
        if (a.ch === undefined || b.ch === undefined) tests.push(a.test, (c) => c === '-', b.test); // [\d-z]: "-" stands for itself
        else { if (a.ch > b.ch) fail(); tests.push((c) => c >= a.ch && c <= b.ch); }
      } else tests.push(a.test);
    }
    if (p[i++] !== ']') fail();
    const any = (c) => tests.some((t) => t(c));
    // without case: the class holds the character or one that is equal to it without case
    const test = icase ? (c) => { const g = caseGroup(c); return g ? g.some(any) : any(c); } : any;
    return set(neg ? (c) => !test(c) : test);
  }
  function atom() {
    const ch = p[i++];
    switch (ch) {
      case '(': {
        let cap = 0;
        const flags = [icase, dotAll, multi];
        if (p[i] === '?') {
          // (?:...), (?<name>...), or a group that switches flags: (?i:...), (?-i:...), (?ms-i:...)
          const m = /^\?(?::|<([A-Za-z_$][\w$]*)>|([ims]*)(?:-([ims]*))?:)/.exec(p.slice(i, i + 40));
          if (!m) fail(); // lookarounds
          i += m[0].length;
          if (m[1]) { cap = ++ncap; names[m[1]] = cap; }
          if (m[2] !== undefined) {
            const on = m[2], all = on + (m[3] || '');
            // as in JavaScript: each flag at most once, and at least one flag
            if (!all || new Set(all).size !== all.length) fail();
            for (const f of all) { if (f === 'i') icase = on.includes(f); else if (f === 's') dotAll = on.includes(f); else multi = on.includes(f); }
          }
        } else cap = ++ncap;
        const node = alt();
        [icase, dotAll, multi] = flags; // the flags of a group end with it
        if (p[i++] !== ')') fail();
        return { t: 'group', node, cap };
      }
      case '[': return cls();
      case '.': return set(dotAll ? () => true : (c) => !isLineEnd(c));
      case '^': return { t: multi ? 'mbol' : 'bol' };
      case '$': return { t: multi ? 'meol' : 'eol' };
      case '\\': return escape(false);
      case '*': case '+': case '?': case ')': return fail();
      case '{': return quantAt(i - 1) ? fail() : lit(ch);
      default: return lit(ch);
    }
  }
  function quant(node) {
    let min, max;
    const ch = p[i];
    if (ch === '*') { min = 0; max = Infinity; i++; }
    else if (ch === '+') { min = 1; max = Infinity; i++; }
    else if (ch === '?') { min = 0; max = 1; i++; }
    else if (ch === '{' && quantAt(i)) {
      const m = quantAt(i);
      min = +m[1]; max = !m[2] ? min : m[3] === '' ? Infinity : +m[3];
      // any count, as in JavaScript: the size of the compiled pattern (RX_PROG) limits what a count can cost
      if (min > max) fail();
      i += m[0].length;
    } else return node;
    let greedy = true;
    if (p[i] === '?') { greedy = false; i++; }
    if (p[i] === '*' || p[i] === '+' || p[i] === '?' || (p[i] === '{' && quantAt(i))) fail(); // nothing to repeat
    if (ANCHORS.has(node.t)) fail(); // nothing to repeat, as in JavaScript
    return { t: 'rep', node, min, max, greedy };
  }
  function cat() {
    const items = [];
    while (i < p.length && p[i] !== '|' && p[i] !== ')') items.push(quant(atom()));
    return { t: 'cat', items };
  }
  function alt() {
    const alts = [cat()];
    while (p[i] === '|') { i++; alts.push(cat()); }
    return alts.length === 1 ? alts[0] : { t: 'alt', alts };
  }
  const root = alt();
  if (i < p.length) fail(); // an unmatched ")"
  // compile to instructions: c (one character), p (split, x first), j (jump), s (save a position), a (assert), m (match)
  const prog = [];
  const emit = (ins) => { if (prog.length >= RX_PROG) fail(); prog.push(ins); return ins; };
  // whether a part compiles to no instructions (an empty group, a{0}): repeated any number of times it adds nothing,
  // so a count such as (?:){1000000000} costs no time
  const empty = (n) => {
    if (n.e === undefined) n.e = n.t === 'cat' ? n.items.every(empty) : n.t === 'group' ? !n.cap && empty(n.node) : n.t === 'rep' ? n.max === 0 || empty(n.node) : false;
    return n.e;
  };
  function gen(n) {
    switch (n.t) {
      case 'char': emit({ op: 'c', test: n.test }); break;
      case 'bol': case 'eol': case 'mbol': case 'meol': emit({ op: 'a', kind: n.t }); break;
      case 'wb': emit({ op: 'a', kind: n.neg ? 'nwb' : 'wb' }); break;
      case 'cat': n.items.forEach(gen); break;
      case 'group':
        if (n.cap) emit({ op: 's', n: 2 * n.cap });
        gen(n.node);
        if (n.cap) emit({ op: 's', n: 2 * n.cap + 1 });
        break;
      case 'alt': {
        const jumps = [];
        n.alts.forEach((a, k) => {
          if (k === n.alts.length - 1) { gen(a); return; }
          const sp = emit({ op: 'p' });
          sp.x = prog.length; gen(a);
          jumps.push(emit({ op: 'j' }));
          sp.y = prog.length;
        });
        jumps.forEach((j) => { j.x = prog.length; });
        break;
      }
      default: { // rep
        if (empty(n.node)) break;
        for (let k = 0; k < n.min; k++) gen(n.node);
        const loop = n.max === Infinity;
        const splits = [];
        for (let k = n.min; k < (loop ? n.min + 1 : n.max); k++) {
          const sp = emit({ op: 'p' }), at = prog.length - 1;
          splits.push([sp, prog.length]);
          gen(n.node);
          if (loop) emit({ op: 'j', x: at });
        }
        splits.forEach(([sp, body]) => { sp.x = n.greedy ? body : prog.length; sp.y = n.greedy ? prog.length : body; });
      }
    }
  }
  emit({ op: 's', n: 0 }); gen(root); emit({ op: 's', n: 1 }); emit({ op: 'm' });
  return { prog, ncap, names };
}
function rxAssert(kind, s, pos) {
  if (kind === 'bol') return pos === 0;
  if (kind === 'eol') return pos === s.length;
  // with the m flag, also at the start and the end of each line
  if (kind === 'mbol') return pos === 0 || isLineEnd(s[pos - 1]);
  if (kind === 'meol') return pos === s.length || isLineEnd(s[pos]);
  const b = isWordCh(s[pos - 1]) !== isWordCh(s[pos]);
  return kind === 'wb' ? b : !b;
}
/** The first match at or after `from`: [start, end, group 1 start, group 1 end, …] (-1 for no group), or null. */
function rxExec(rx, s, from) {
  const prog = rx.prog, mark = new Int32Array(prog.length), init = new Array(2 * rx.ncap + 2).fill(-1);
  let stamp = 1, match = null;
  // adds the threads reachable from pc without reading a character, in priority order
  const add = (list, pc0, caps0, pos) => {
    const stack = [pc0, caps0];
    while (stack.length) {
      const caps = stack.pop(), pc = stack.pop();
      if (mark[pc] === stamp) continue;
      mark[pc] = stamp;
      const ins = prog[pc];
      if (ins.op === 'j') stack.push(ins.x, caps);
      else if (ins.op === 'p') stack.push(ins.y, caps, ins.x, caps);
      else if (ins.op === 's') { const c = caps.slice(); c[ins.n] = pos; stack.push(pc + 1, c); }
      else if (ins.op === 'a') { if (rxAssert(ins.kind, s, pos)) stack.push(pc + 1, caps); }
      else list.push(pc, caps);
    }
  };
  let list = [];
  add(list, 0, init, from);
  for (let pos = from; pos <= s.length; pos++) {
    tick(list.length / 2 + 1);
    stamp++;
    const next = [], c = s[pos];
    for (let k = 0; k < list.length; k += 2) {
      const ins = prog[list[k]];
      if (ins.op === 'm') { match = list[k + 1]; break; } // threads after this one have lower priority
      if (c !== undefined && ins.test(c)) add(next, list[k] + 1, list[k + 1], pos + 1);
    }
    if (!match && pos < s.length) add(next, 0, init, pos + 1); // a match starting one character later
    if (match && !next.length) break;
    list = next;
  }
  return match;
}
function userRegex(pattern, icase) {
  const p = strArg(pattern);
  if (p.length > RX_MAX) throw ERR['#VALUE!'];
  const key = (icase ? 'i' : 'c') + p;
  let rx = rxCache.get(key);
  if (!rx) {
    rx = rxCompile(p, icase);
    if (rxCache.size > 200) rxCache.clear();
    rxCache.set(key, rx);
  }
  return rx;
}
function regexText(t) {
  const s = strArg(t);
  if (longText(s)) throw ERR['#VALUE!'];
  return s;
}
/** All matches, each as its capture positions (empty matches move on by one character, as in JavaScript). */
function rxAll(rx, s, limit) {
  const out = [];
  for (let pos = 0; pos <= s.length && out.length < limit;) {
    const m = rxExec(rx, s, pos);
    if (!m) break;
    out.push(m);
    pos = m[1] > m[0] ? m[1] : m[1] + 1;
  }
  return out;
}
/** A replacement text with $&, $1…$99, $<name>, $`, $' and $$ filled in, as String.replace does. */
function rxExpand(repl, rx, s, m) {
  const group = (k) => (m[2 * k] < 0 ? '' : s.slice(m[2 * k], m[2 * k + 1]));
  let out = '';
  for (let i = 0; i < repl.length; i++) {
    const ch = repl[i], nx = repl[i + 1];
    if (ch !== '$' || nx === undefined) { out += ch; continue; }
    if (nx === '$') { out += '$'; i++; }
    else if (nx === '&') { out += group(0); i++; }
    else if (nx === '`') { out += s.slice(0, m[0]); i++; }
    else if (nx === "'") { out += s.slice(m[1]); i++; }
    else if (/\d/.test(nx)) {
      const two = +(nx + (repl[i + 2] || 'x'));
      if (two >= 1 && two <= rx.ncap) { out += group(two); i += 2; }
      else if (+nx >= 1 && +nx <= rx.ncap) { out += group(+nx); i++; }
      else out += ch;
    } else if (nx === '<' && Object.keys(rx.names).length) {
      const end = repl.indexOf('>', i + 2);
      if (end < 0) { out += ch; continue; }
      const k = rx.names[repl.slice(i + 2, end)];
      if (k) out += group(k);
      i = end;
    } else out += ch;
    if (longText(out)) throw ERR['#VALUE!'];
  }
  return out;
}
F.REGEXTEST = L((t, p, cs) => rxExec(userRegex(p, !!numArg(cs, 0)), regexText(t), 0) !== null);
F.REGEXEXTRACT = (t, p, mode, cs) => {
  const s = regexText(t), md = numArg(mode, 0);
  const rx = userRegex(p, !!numArg(cs, 0));
  if (md === 1) {
    const all = rxAll(rx, s, MAX_CELLS);
    return all.length ? all.map((m) => [s.slice(m[0], m[1])]) : ERR['#N/A'];
  }
  const m = rxExec(rx, s, 0);
  if (!m) return ERR['#N/A'];
  if (md !== 2) return s.slice(m[0], m[1]);
  const groups = [];
  for (let k = 1; k <= rx.ncap; k++) groups.push(m[2 * k] < 0 ? undefined : s.slice(m[2 * k], m[2 * k + 1]));
  return [groups];
};
F.REGEXREPLACE = (t, p, r, occ, cs) => {
  const s = regexText(t), n = numArg(occ, 0), repl = strArg(r);
  const rx = userRegex(p, !!numArg(cs, 0));
  const all = rxAll(rx, s, n > 0 ? n : MAX_CELLS);
  let out = '', last = 0;
  all.forEach((m, k) => {
    if (n && k + 1 !== n) return;
    out += s.slice(last, m[0]) + rxExpand(repl, rx, s, m);
    last = m[1];
    if (longText(out)) throw ERR['#VALUE!'];
  });
  return out + s.slice(last);
};

// ---- date & time
F.TODAY = () => todaySerial();
F.NOW = () => nowSerial();
F.DATE = (y, m, d) => {
  let yy = Math.floor(numArg(y)); if (yy < 1900) yy += 1900;
  const dt = new Date(Date.UTC(yy, Math.floor(numArg(m)) - 1, Math.floor(numArg(d))));
  return dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
};
F.TIME = (h, m, s) => { const v = (numArg(h) * 3600 + numArg(m) * 60 + numArg(s)) / 86400; return v - Math.floor(v); };
F.YEAR = N1((x) => serialToParts(x).y);
F.MONTH = N1((x) => serialToParts(x).m);
F.DAY = N1((x) => serialToParts(x).d);
F.HOUR = N1((x) => serialToParts(x).H);
F.MINUTE = N1((x) => serialToParts(x).M);
F.SECOND = N1((x) => serialToParts(x).S);
F.WEEKDAY = L((x, type) => {
  const dow = serialToParts(numArg(x)).dow; const t = numArg(type, 1);
  if (t === 1 || t === 17) return dow + 1;
  if (t === 2 || t === 11) return ((dow + 6) % 7) + 1;
  if (t === 3) return (dow + 6) % 7;
  if (t >= 12 && t <= 16) return ((dow - (t - 10) + 7) % 7) + 1;
  return ERR['#NUM!'];
});
F.WEEKNUM = L((x, type) => {
  const s = Math.floor(numArg(x)); const p = serialToParts(s); const t = numArg(type, 1);
  if (t === 21) return F.ISOWEEKNUM(s);
  const start = t === 1 || t === 17 ? 0 : t === 2 || t === 11 ? 1 : (t - 10) % 7;
  const jan1 = dateToSerial(p.y, 1, 1);
  const off = (serialToParts(jan1).dow - start + 7) % 7;
  return Math.floor((s - jan1 + off) / 7) + 1;
});
F.ISOWEEKNUM = (x) => {
  const p = serialToParts(Math.floor(numArg(x)));
  const dt = new Date(Date.UTC(p.y, p.m - 1, p.d));
  const day = (dt.getUTCDay() + 6) % 7;
  dt.setUTCDate(dt.getUTCDate() - day + 3);
  const firstThu = new Date(Date.UTC(dt.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((dt - firstThu) / DAY_MS - 3 + ((firstThu.getUTCDay() + 6) % 7)) / 7);
};
F.EDATE = L((x, months) => { const p = serialToParts(numArg(x)); const dt = new Date(Date.UTC(p.y, p.m - 1 + Math.trunc(numArg(months)), 1)); const last = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate(); return dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, Math.min(p.d, last)); });
F.EOMONTH = L((x, months) => { const p = serialToParts(numArg(x)); const dt = new Date(Date.UTC(p.y, p.m + Math.trunc(numArg(months)), 0)); return dateToSerial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()); });
F.DAYS = (end, start) => Math.floor(numArg(end)) - Math.floor(numArg(start));
F.DAYS360 = (start, end, method) => {
  const a = serialToParts(numArg(start)), b = serialToParts(numArg(end)), eu = boolArg(method, false);
  let d1 = a.d, d2 = b.d;
  if (eu) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
  else {
    const lastFeb = (p) => p.m === 2 && p.d === new Date(Date.UTC(p.y, 2, 0)).getUTCDate();
    if (d1 === 31 || lastFeb(a)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
};
F.YEARFRAC = (start, end, basis) => {
  let s = Math.floor(numArg(start)), e = Math.floor(numArg(end)); const b = numArg(basis, 0);
  if (s > e) [s, e] = [e, s];
  switch (b) {
    case 0: return F.DAYS360(s, e, false) / 360;
    case 1: { const ya = serialToParts(s).y, yb = serialToParts(e).y; let days = 0; for (let y = ya; y <= yb; y++) days += (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0)) ? 366 : 365; return (e - s) / (days / (yb - ya + 1)); }
    case 2: return (e - s) / 360;
    case 3: return (e - s) / 365;
    case 4: return F.DAYS360(s, e, true) / 360;
    default: return ERR['#NUM!'];
  }
};
F.DATEDIF = (start, end, unit) => {
  const s = numArg(start), e = numArg(end);
  const a = serialToParts(s), b = serialToParts(e), u = strArg(unit).toUpperCase();
  if (s > e) return ERR['#NUM!'];
  const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
  switch (u) {
    case 'D': return Math.floor(e) - Math.floor(s);
    case 'M': return months;
    case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { const prev = new Date(Date.UTC(b.y, b.m - 1, 0)).getUTCDate(); return b.d >= a.d ? b.d - a.d : prev - a.d + b.d; }
    case 'YD': { let d = dateToSerial(b.y, a.m, a.d); if (d > e) d = dateToSerial(b.y - 1, a.m, a.d); return Math.floor(e - d); }
    default: return ERR['#NUM!'];
  }
};
function holidaySet(h) { const s = new Set(); if (h !== undefined && h !== null) for (const { v } of flat([h])) if (typeof v === 'number') s.add(Math.floor(v)); return s; }
function weekendMask(w) {
  if (w === undefined || w === null) return [0, 6];
  const v = scalar(w);
  if (typeof v === 'string' && /^[01]{7}$/.test(v)) return [...v].map((c, i) => (c === '1' ? (i + 1) % 7 : -1)).filter((x) => x >= 0);
  const map = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6], 11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6] };
  return map[toNum(v)] || [6, 0];
}
// dates are serials 0..2958465 (31 Dec 9999); serial 0 is a Saturday
const MAX_DATE = 2958465;
const validDate = (s) => s >= 0 && s <= MAX_DATE;
const weekday = (s) => (((s + 6) % 7) + 7) % 7;
function networkdays(start, end, weekend, holidays) {
  let a = Math.floor(numArg(start)), b = Math.floor(numArg(end)), sign = 1;
  if (!validDate(a) || !validDate(b)) return ERR['#NUM!'];
  if (a > b) { [a, b] = [b, a]; sign = -1; }
  const hol = holidaySet(holidays), we = weekendMask(weekend);
  tick(b - a + 1);
  let c = 0;
  for (let s = a; s <= b; s++) if (!we.includes(weekday(s)) && !hol.has(s)) c++;
  return sign * c;
}
function workday(start, days, weekend, holidays) {
  let s = Math.floor(numArg(start)), n = Math.trunc(numArg(days));
  if (!validDate(s) || !(Math.abs(n) <= MAX_DATE)) return ERR['#NUM!'];
  const hol = holidaySet(holidays), we = weekendMask(weekend), step = n < 0 ? -1 : 1;
  if (we.length >= 7) return ERR['#VALUE!']; // every day is a weekend day
  while (n !== 0) {
    s += step;
    if (!validDate(s)) return ERR['#NUM!'];
    tick(1);
    if (!we.includes(weekday(s)) && !hol.has(s)) n -= step;
  }
  return s;
}
F.NETWORKDAYS = (s, e, h) => networkdays(s, e, undefined, h);
F['NETWORKDAYS.INTL'] = (s, e, w, h) => networkdays(s, e, w, h);
F.WORKDAY = (s, d, h) => workday(s, d, undefined, h);
F['WORKDAY.INTL'] = (s, d, w, h) => workday(s, d, w, h);
F.DATEVALUE = (t) => { const n = parseInput(strArg(t), F._locale); return n.kind === 'date' || n.kind === 'datetime' ? Math.floor(n.num) : ERR['#VALUE!']; };
F.TIMEVALUE = (t) => { const n = parseInput(strArg(t), F._locale); return n.kind === 'time' || n.kind === 'datetime' ? n.num - Math.floor(n.num) : ERR['#VALUE!']; };

// ---- lookup & reference
function exactMatch(a, b) {
  if (a === null || a === undefined) return false;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return a === b;
}
function lookupIndex(value, vec, exact, approxMode) {
  const v = scalar(value);
  if (isErr(v)) throw v;
  tick(vec.length);
  if (exact) {
    const f = typeof v === 'string' && /[*?~]/.test(v) ? makeCriteria(v) : null;
    for (let i = 0; i < vec.length; i++) if (f ? f(vec[i]) : exactMatch(vec[i], v)) return i;
    return -1;
  }
  // approximate match on sorted data (binary search like Excel)
  let lo = 0, hi = vec.length - 1, found = -1;
  if (approxMode === -1) {
    for (let i = 0; i < vec.length; i++) if (vec[i] !== null && compare(vec[i], v) >= 0) found = i; else break;
    return found;
  }
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    while (m >= lo && (vec[m] === null || (typeof vec[m] === 'string') !== (typeof v === 'string'))) m--;
    if (m < lo) { lo = mid + 1; continue; }
    if (compare(vec[m], v) <= 0) { found = m; lo = mid + 1; } else hi = m - 1;
  }
  return found;
}
function columnOf(r, j) { const [rows] = rangeSize(r); tick(rows); const out = []; for (let i = 0; i < rows; i++) out.push(rangeGet(r, i, j)); return out; }
function rowOf(r, i) { const [, cols] = rangeSize(r); tick(cols); const out = []; for (let j = 0; j < cols; j++) out.push(rangeGet(r, i, j)); return out; }
F.VLOOKUP = L((value, table, colNum, approx) => {
  const k = Math.floor(numArg(colNum)); const [rows, cols] = rangeSize(table);
  if (k < 1 || k > cols || !rows) return ERR['#REF!'];
  const exact = approx !== undefined && approx !== null && boolArg(approx, true) === false;
  const i = lookupIndex(value, columnOf(table, 0), exact, 1);
  return i < 0 ? ERR['#N/A'] : rangeGet(table, i, k - 1);
});
F.HLOOKUP = L((value, table, rowNum, approx) => {
  const k = Math.floor(numArg(rowNum)); const [rows] = rangeSize(table);
  if (k < 1 || k > rows) return ERR['#REF!'];
  const exact = approx !== undefined && approx !== null && boolArg(approx, true) === false;
  const j = lookupIndex(value, rowOf(table, 0), exact, 1);
  return j < 0 ? ERR['#N/A'] : rangeGet(table, k - 1, j);
});
F.MATCH = L((value, vec, type) => {
  const t = numArg(type, 1);
  const [rows, cols] = rangeSize(vec);
  if (rows > 1 && cols > 1) return ERR['#N/A'];
  const list = rows === 1 ? rowOf(vec, 0) : columnOf(vec, 0);
  const i = lookupIndex(value, list, t === 0, t < 0 ? -1 : 1);
  return i < 0 ? ERR['#N/A'] : i + 1;
});
F.XMATCH = L((value, vec, mode, searchMode) => {
  const [rows] = rangeSize(vec);
  const list = rows === 1 ? rowOf(vec, 0) : columnOf(vec, 0);
  const i = xIndex(value, list, numArg(mode, 0), numArg(searchMode, 1));
  return i < 0 ? ERR['#N/A'] : i + 1;
});
function xIndex(value, list, mode, searchMode) {
  const v = scalar(value);
  tick(list.length);
  const order = searchMode === -1 || searchMode === -2 ? list.map((_, i) => list.length - 1 - i) : list.map((_, i) => i);
  if (mode === 0 || mode === 2) {
    const f = mode === 2 && typeof v === 'string' ? makeCriteria(v) : null;
    for (const i of order) if (f ? f(list[i]) : exactMatch(list[i], v)) return i;
    return -1;
  }
  let best = -1;
  for (const i of order) {
    const x = list[i];
    if (x === null || (typeof x === 'string') !== (typeof v === 'string')) continue;
    const c = compare(x, v);
    if (c === 0) return i;
    if (mode === -1 && c < 0 && (best < 0 || compare(x, list[best]) > 0)) best = i;
    if (mode === 1 && c > 0 && (best < 0 || compare(x, list[best]) < 0)) best = i;
  }
  return best;
}
F.XLOOKUP = (value, lookup, ret, notFound, matchMode, searchMode) => {
  const [rows, cols] = rangeSize(lookup);
  const vertical = cols === 1 && rows >= 1;
  // the lookup vector is read once, also when many values are looked up
  const list = vertical ? columnOf(lookup, 0) : rowOf(lookup, 0);
  const mode = numArg(matchMode, 0), search = numArg(searchMode, 1);
  const one = (v) => xlookupResult(ret, vertical, xIndex(v, list, mode, search), notFound);
  if (!isMulti(value)) return one(value);
  if (mode !== 0) return mapMatrix(toMatrix(value), (v) => scalar(one(v)));
  // many exact lookups: index the lookup vector once (the first match, or the last one when searching backwards)
  const index = new Map();
  list.forEach((x, i) => { const k = exactKey(x); if (k !== null && (search < 0 || !index.has(k))) index.set(k, i); });
  return mapMatrix(toMatrix(value), (v) => {
    const k = exactKey(scalar(v));
    return scalar(xlookupResult(ret, vertical, k !== null && index.has(k) ? index.get(k) : -1, notFound));
  });
};
/** A key under which exactMatch finds a value (text without case), or null for values it never matches. */
function exactKey(x) {
  if (x === null || x === undefined || (typeof x === 'number' && x !== x)) return null;
  return typeof x === 'string' ? 's' + x.toLowerCase() : isErr(x) ? 'e' + x.code : typeof x + ':' + x;
}
function xlookupResult(ret, vertical, i, notFound) {
  if (i < 0) return notFound !== undefined && notFound !== null ? notFound : ERR['#N/A'];
  if (ret instanceof Range) {
    return vertical ? new Range(ret.engine, ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2)
      : new Range(ret.engine, ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i);
  }
  const m = toMatrix(ret);
  return vertical ? [m[i]] : m.map((r) => [r[i]]);
};
F.LOOKUP = (value, vec, result) => {
  const [rows, cols] = rangeSize(vec);
  const horizontal = cols > rows;
  const list = horizontal ? rowOf(vec, 0) : columnOf(vec, 0);
  const i = lookupIndex(value, list, false, 1);
  if (i < 0) return ERR['#N/A'];
  if (result !== undefined) { const [rr] = rangeSize(result); return rr === 1 ? rangeGet(result, 0, i) : rangeGet(result, i, 0); }
  return horizontal ? rangeGet(vec, rows - 1, i) : rangeGet(vec, i, cols - 1);
};
F.INDEX = (r, rowNum, colNum) => {
  const [rows, cols] = rangeSize(r);
  let i = Math.floor(numArg(rowNum, 0)), j = Math.floor(numArg(colNum, 0));
  if (rows === 1 && (colNum === undefined || colNum === null)) { j = i; i = 1; }
  if (cols === 1 && (colNum === undefined || colNum === null)) j = 1;
  if (i < 0 || j < 0 || i > rows || j > cols) return ERR['#REF!'];
  if (r instanceof Range) {
    if (i === 0 && j === 0) return r;
    if (i === 0) return new Range(r.engine, r.sheet, r.r1, r.c1 + j - 1, r.r2, r.c1 + j - 1);
    if (j === 0) return new Range(r.engine, r.sheet, r.r1 + i - 1, r.c1, r.r1 + i - 1, r.c2);
    return new Range(r.engine, r.sheet, r.r1 + i - 1, r.c1 + j - 1, r.r1 + i - 1, r.c1 + j - 1);
  }
  const m = toMatrix(r);
  if (i === 0) return m.map((row) => [row[j - 1]]);
  if (j === 0) return [m[i - 1]];
  return m[i - 1][j - 1];
};
F.ROW = (args, ev, ctx) => {
  if (!args.length || args[0].k === 'empty') return ctx.r + 1;
  const v = ev(args[0]);
  if (!(v instanceof Range)) return ERR['#VALUE!'];
  return v.rows > 1 ? Array.from({ length: v.rows }, (_, i) => [v.r1 + i + 1]) : v.r1 + 1;
};
F.COLUMN = (args, ev, ctx) => {
  if (!args.length || args[0].k === 'empty') return ctx.c + 1;
  const v = ev(args[0]);
  if (!(v instanceof Range)) return ERR['#VALUE!'];
  return v.cols > 1 ? [Array.from({ length: v.cols }, (_, j) => v.c1 + j + 1)] : v.c1 + 1;
};
F.ROWS = (r) => rangeSize(r)[0];
F.COLUMNS = (r) => rangeSize(r)[1];
F.AREAS = () => 1;
F.ISREF = (args, ev) => ev(args[0]) instanceof Range;
F.OFFSET = (args, ev) => {
  const base = ev(args[0]);
  if (!(base instanceof Range)) return ERR['#VALUE!'];
  const num = (k, d) => (args[k] && args[k].k !== 'empty' ? toNum(ev(args[k])) : d);
  const dr = num(1, 0), dc = num(2, 0), h = num(3, base.rows), w = num(4, base.cols);
  if ([dr, dc, h, w].some(isErr)) return ERR['#VALUE!'];
  const r1 = base.r1 + Math.trunc(dr), c1 = base.c1 + Math.trunc(dc);
  if (r1 < 0 || c1 < 0 || h < 1 || w < 1) return ERR['#REF!'];
  const r2 = r1 + Math.trunc(h) - 1, c2 = c1 + Math.trunc(w) - 1;
  // past the last row or column of a worksheet is #REF!, as in Excel; the rest is cut to the sheet like A:A
  if (!(r2 <= MAX_ROW && c2 <= MAX_COL)) return ERR['#REF!'];
  return base.engine.clip(base.sheet, r1, c1, r2, c2);
};
F.INDIRECT = (args, ev, ctx, engine) => {
  const t = toStr(ev(args[0]));
  if (isErr(t)) return t;
  const a1 = args.length < 2 || args[1].k === 'empty' || toBool(ev(args[1])) !== false;
  let ref = null;
  if (a1) ref = parseRefText(t);
  else {
    const m = /^(?:(.+)!)?R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(t.trim());
    if (m) ref = { kind: m[4] ? 'range' : 'cell', sheet: m[1] ? m[1].replace(/^'|'$/g, '') : null, r1: +m[2] - 1, c1: +m[3] - 1, r2: +(m[4] || m[2]) - 1, c2: +(m[5] || m[3]) - 1 };
  }
  if (!ref) {
    const named = engine.resolveName(t.toUpperCase(), ctx);
    return named === undefined ? ERR['#REF!'] : named;
  }
  try { return engine.rangeFor(ref, ctx); } catch (e) { return isErr(e) ? e : ERR['#REF!']; }
};
F.ADDRESS = (row, col, absNum, a1, sheet) => {
  const r = Math.floor(numArg(row)), c = Math.floor(numArg(col)), abs = numArg(absNum, 1), style = boolArg(a1, true);
  if (!(r >= 1 && c >= 1 && r <= MAX_ROW + 1 && c <= MAX_COL + 1)) return ERR['#VALUE!']; // inside a worksheet, as in Excel
  const sh = sheet !== undefined && sheet !== null ? quoteSheet(strArg(sheet)) + '!' : '';
  if (!style) return sh + 'R' + (abs === 1 || abs === 2 ? r : `[${r}]`) + 'C' + (abs === 1 || abs === 3 ? c : `[${c}]`);
  return sh + (abs === 1 || abs === 3 ? '$' : '') + colName(c - 1) + (abs === 1 || abs === 2 ? '$' : '') + r;
};
F.HYPERLINK = (args, ev, ctx, engine) => {
  const url = toStr(ev(args[0]));
  if (isErr(url)) return url;
  const label = args.length > 1 ? scalar(ev(args[1])) : url;
  // only web and mail links open from a cell (no javascript: or data: URLs), and only the cell's own formula sets
  // its link: not conditional formatting, validation or a defined name evaluated for that cell
  if (ctx.cell && ctx.sheet && /^(https?:\/\/|mailto:)/i.test(String(url).trim())) engine.links.set(ctx.sheet.id + '|' + ctx.r + '|' + ctx.c, String(url).trim());
  return label;
};
F.FORMULATEXT = (args, ev, ctx, engine) => {
  const v = ev(args[0]);
  if (!(v instanceof Range)) return ERR['#N/A'];
  const raw = engine.raw(v.sheet, v.r1, v.c1);
  return raw && raw[0] === '=' ? raw : ERR['#N/A'];
};
F.ISFORMULA = (args, ev, ctx, engine) => {
  const v = ev(args[0]);
  if (!(v instanceof Range)) return ERR['#VALUE!'];
  const raw = engine.raw(v.sheet, v.r1, v.c1);
  return !!raw && raw[0] === '=';
};
F.SHEET = (args, ev, ctx, engine) => {
  if (!args.length) return engine.book.sheets.indexOf(ctx.sheet) + 1;
  const v = ev(args[0]);
  if (v instanceof Range) return engine.book.sheets.indexOf(v.sheet) + 1;
  const s = engine.sheetByName(toStr(v));
  return s ? engine.book.sheets.indexOf(s) + 1 : ERR['#N/A'];
};
F.SHEETS = (args, ev, ctx, engine) => engine.book.sheets.length;
F.CELL = (args, ev, ctx, engine) => {
  const info = toStr(ev(args[0])).toLowerCase();
  const v = args.length > 1 ? ev(args[1]) : new Range(engine, ctx.sheet, ctx.r, ctx.c, ctx.r, ctx.c);
  if (!(v instanceof Range)) return ERR['#VALUE!'];
  switch (info) {
    case 'address': return '$' + colName(v.c1) + '$' + (v.r1 + 1);
    case 'row': return v.r1 + 1;
    case 'col': return v.c1 + 1;
    case 'contents': return v.get(0, 0);
    case 'type': { const x = v.get(0, 0); return x === null ? 'b' : typeof x === 'string' ? 'l' : 'v'; }
    case 'filename': return v.sheet.name;
    default: return ERR['#VALUE!'];
  }
};

// ---- dynamic arrays
const rowsOf = (a) => toMatrix(a);
F.TRANSPOSE = (a) => { const m = toMatrix(a); return m[0].map((_, j) => m.map((r) => r[j])); };
F.UNIQUE = (a, byCol, once) => {
  let m = toMatrix(a);
  const cols = boolArg(byCol, false);
  if (cols) m = F.TRANSPOSE(m);
  const key = (r) => JSON.stringify(r.map((x) => (typeof x === 'string' ? x.toLowerCase() : isErr(x) ? x.code : x)));
  const counts = new Map();
  m.forEach((r) => counts.set(key(r), (counts.get(key(r)) || 0) + 1));
  const seen = new Set();
  const out = m.filter((r) => { const k = key(r); if (boolArg(once, false) && counts.get(k) > 1) return false; if (seen.has(k)) return false; seen.add(k); return true; });
  if (!out.length) return ERR['#CALC!'];
  return cols ? F.TRANSPOSE(out) : out;
};
function sortRows(m, keys) {
  tick(m.length * Math.ceil(Math.log2(m.length + 1)) * keys.length);
  return m.map((r, i) => ({ r, i })).sort((x, y) => {
    for (const { get, order } of keys) {
      const c = compare(get(x), get(y));
      if (c) return c * order;
    }
    return x.i - y.i;
  }).map((x) => x.r);
}
F.SORT = (a, idx, order, byCol) => {
  let m = toMatrix(a);
  const cols = boolArg(byCol, false);
  if (cols) m = F.TRANSPOSE(m);
  const k = Math.floor(numArg(idx, 1)) - 1, o = numArg(order, 1) < 0 ? -1 : 1;
  if (k < 0 || k >= m[0].length) return ERR['#VALUE!'];
  const out = sortRows(m, [{ get: (x) => x.r[k], order: o }]);
  return cols ? F.TRANSPOSE(out) : out;
};
F.SORTBY = (a, ...spec) => {
  const m = toMatrix(a);
  const keys = [];
  for (let i = 0; i < spec.length; i += 2) {
    const by = toMatrix(spec[i]).map((r) => r[0]);
    if (by.length !== m.length) return ERR['#VALUE!'];
    const o = numArg(spec[i + 1], 1) < 0 ? -1 : 1;
    keys.push({ get: (x) => by[x.i], order: o });
  }
  return sortRows(m, keys);
};
F.FILTER = (a, include, ifEmpty) => {
  const m = toMatrix(a), inc = toMatrix(include);
  let out;
  if (inc.length === m.length && inc[0].length === 1) out = m.filter((_, i) => { const b = toBool(inc[i][0]); if (isErr(b)) throw b; return b; });
  else if (inc.length === 1 && inc[0].length === m[0].length) {
    const keep = inc[0].map((x) => toBool(x) === true);
    out = m.map((r) => r.filter((_, j) => keep[j]));
    if (!out[0].length) out = [];
  } else return ERR['#VALUE!'];
  if (!out.length) return ifEmpty !== undefined && ifEmpty !== null ? ifEmpty : ERR['#CALC!'];
  return out;
};
F.VSTACK = (...a) => { const ms = a.map(toMatrix); const w = Math.max(...ms.map((m) => m[0].length)); arraySize(ms.reduce((n, m) => n + m.length, 0), w); return ms.flatMap((m) => m.map((r) => r.concat(Array(w - r.length).fill(ERR['#N/A'])))); };
F.HSTACK = (...a) => { const ms = a.map(toMatrix); const h = Math.max(...ms.map((m) => m.length)); arraySize(h, ms.reduce((n, m) => n + m[0].length, 0)); return Array.from({ length: h }, (_, i) => ms.flatMap((m) => m[i] || Array(m[0].length).fill(ERR['#N/A']))); };
F.TAKE = (a, rows, cols) => {
  let m = toMatrix(a);
  if (rows !== undefined && rows !== null) { const n = Math.trunc(numArg(rows)); m = n >= 0 ? m.slice(0, n) : m.slice(n); }
  if (cols !== undefined && cols !== null) { const n = Math.trunc(numArg(cols)); m = m.map((r) => (n >= 0 ? r.slice(0, n) : r.slice(n))); }
  return m.length && m[0].length ? m : ERR['#CALC!'];
};
F.DROP = (a, rows, cols) => {
  let m = toMatrix(a);
  if (rows !== undefined && rows !== null) { const n = Math.trunc(numArg(rows)); m = n >= 0 ? m.slice(n) : m.slice(0, n); }
  if (cols !== undefined && cols !== null) { const n = Math.trunc(numArg(cols)); m = m.map((r) => (n >= 0 ? r.slice(n) : r.slice(0, n))); }
  return m.length && m[0].length ? m : ERR['#CALC!'];
};
F.CHOOSECOLS = (a, ...idx) => { const m = toMatrix(a); const w = m[0].length; const ks = numbers(idx).map((k) => (k < 0 ? w + k : k - 1)); if (ks.some((k) => k < 0 || k >= w)) return ERR['#VALUE!']; arraySize(m.length, ks.length); return m.map((r) => ks.map((k) => r[k])); };
F.CHOOSEROWS = (a, ...idx) => { const m = toMatrix(a); const h = m.length; const ks = numbers(idx).map((k) => (k < 0 ? h + k : k - 1)); if (ks.some((k) => k < 0 || k >= h)) return ERR['#VALUE!']; arraySize(ks.length, m[0].length); return ks.map((k) => m[k]); };
F.TOCOL = (a, ignore, byCol) => { let m = toMatrix(a); if (boolArg(byCol, false)) m = F.TRANSPOSE(m); const ig = numArg(ignore, 0); return m.flat().filter((x) => !((ig === 1 || ig === 3) && x === null) && !((ig === 2 || ig === 3) && isErr(x))).map((x) => [x]); };
F.TOROW = (a, ignore, byCol) => [F.TOCOL(a, ignore, byCol).map((r) => r[0])];
F.WRAPROWS = (v, n, pad) => { const items = toMatrix(v).flat(); const w = Math.floor(numArg(n)); if (!(w >= 1)) return ERR['#NUM!']; arraySize(Math.ceil(items.length / w), w); const out = []; for (let i = 0; i < items.length; i += w) { const r = items.slice(i, i + w); while (r.length < w) r.push(pad === undefined ? ERR['#N/A'] : scalar(pad)); out.push(r); } return out; };
F.WRAPCOLS = (v, n, pad) => F.TRANSPOSE(F.WRAPROWS(v, n, pad));
F.EXPAND = (a, rows, cols, pad) => { const m = toMatrix(a); const h = numArg(rows, m.length), w = numArg(cols, m[0].length); arraySize(Math.floor(h), Math.floor(w)); const p = pad === undefined ? ERR['#N/A'] : scalar(pad); return Array.from({ length: h }, (_, i) => Array.from({ length: w }, (__, j) => (m[i] && j < m[i].length ? m[i][j] : p))); };

// ---- information
F.ISBLANK = L((a) => scalar(a) === null);
F.ISNUMBER = L((a) => typeof scalar(a) === 'number');
F.ISTEXT = L((a) => typeof scalar(a) === 'string');
F.ISNONTEXT = L((a) => typeof scalar(a) !== 'string');
F.ISLOGICAL = L((a) => typeof scalar(a) === 'boolean');
F.ISERROR = L((a) => isErr(scalar(a)));
F.ISERR = L((a) => { const v = scalar(a); return isErr(v) && v !== ERR['#N/A']; });
F.ISNA = L((a) => scalar(a) === ERR['#N/A']);
F.ISEVEN = N1((x) => Math.floor(Math.abs(x)) % 2 === 0);
F.ISODD = N1((x) => Math.floor(Math.abs(x)) % 2 === 1);
F.NA = () => ERR['#N/A'];
F.TYPE = (a) => { if (isMulti(a)) return 64; const v = scalar(a); return typeof v === 'number' || v === null ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : isErr(v) ? 16 : 64; };
F['ERROR.TYPE'] = (a) => { const v = scalar(a); const codes = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!']; return isErr(v) ? (codes.indexOf(v.code) + 1 || ERR['#N/A']) : ERR['#N/A']; };

export const FUNCTION_NAMES = Object.keys(F).filter((k) => !k.startsWith('_')).sort();

/** Short help shown in the formula autocomplete. */
export const FUNCTION_HELP = {
  SUM: 'SUM(number1, [number2], …)', AVERAGE: 'AVERAGE(number1, [number2], …)', COUNT: 'COUNT(value1, …)',
  COUNTA: 'COUNTA(value1, …)', COUNTIF: 'COUNTIF(range, criteria)', COUNTIFS: 'COUNTIFS(range1, criteria1, …)',
  SUMIF: 'SUMIF(range, criteria, [sum_range])', SUMIFS: 'SUMIFS(sum_range, range1, criteria1, …)',
  AVERAGEIF: 'AVERAGEIF(range, criteria, [average_range])', MIN: 'MIN(number1, …)', MAX: 'MAX(number1, …)',
  IF: 'IF(condition, value_if_true, [value_if_false])', IFS: 'IFS(condition1, value1, …)', IFERROR: 'IFERROR(value, value_if_error)',
  AND: 'AND(logical1, …)', OR: 'OR(logical1, …)', NOT: 'NOT(logical)', SWITCH: 'SWITCH(expression, value1, result1, …, [default])',
  VLOOKUP: 'VLOOKUP(lookup_value, table, col_index, [approximate])', HLOOKUP: 'HLOOKUP(lookup_value, table, row_index, [approximate])',
  XLOOKUP: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
  INDEX: 'INDEX(array, row, [column])', MATCH: 'MATCH(lookup_value, lookup_array, [match_type])', XMATCH: 'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])',
  ROUND: 'ROUND(number, digits)', ROUNDUP: 'ROUNDUP(number, digits)', ROUNDDOWN: 'ROUNDDOWN(number, digits)',
  CONCAT: 'CONCAT(text1, …)', TEXTJOIN: 'TEXTJOIN(delimiter, ignore_empty, text1, …)', LEFT: 'LEFT(text, [count])',
  RIGHT: 'RIGHT(text, [count])', MID: 'MID(text, start, count)', LEN: 'LEN(text)', TRIM: 'TRIM(text)', TEXT: 'TEXT(value, format)',
  SUBSTITUTE: 'SUBSTITUTE(text, old, new, [instance])', FIND: 'FIND(find_text, within_text, [start])', SEARCH: 'SEARCH(find_text, within_text, [start])',
  TODAY: 'TODAY()', NOW: 'NOW()', DATE: 'DATE(year, month, day)', DATEDIF: 'DATEDIF(start, end, unit)', EDATE: 'EDATE(start, months)',
  EOMONTH: 'EOMONTH(start, months)', NETWORKDAYS: 'NETWORKDAYS(start, end, [holidays])', WORKDAY: 'WORKDAY(start, days, [holidays])',
  WEEKDAY: 'WEEKDAY(date, [type])', YEAR: 'YEAR(date)', MONTH: 'MONTH(date)', DAY: 'DAY(date)',
  UNIQUE: 'UNIQUE(array, [by_col], [exactly_once])', SORT: 'SORT(array, [sort_index], [order], [by_col])',
  SORTBY: 'SORTBY(array, by_array1, [order1], …)', FILTER: 'FILTER(array, include, [if_empty])', SEQUENCE: 'SEQUENCE(rows, [columns], [start], [step])',
  LET: 'LET(name1, value1, …, calculation)', PMT: 'PMT(rate, nper, pv, [fv], [type])', FV: 'FV(rate, nper, pmt, [pv], [type])',
  PV: 'PV(rate, nper, pmt, [fv], [type])', NPV: 'NPV(rate, value1, …)', IRR: 'IRR(values, [guess])', RATE: 'RATE(nper, pmt, pv, [fv], [type], [guess])',
  SUMPRODUCT: 'SUMPRODUCT(array1, [array2], …)', SUBTOTAL: 'SUBTOTAL(function_num, ref1, …)', OFFSET: 'OFFSET(reference, rows, cols, [height], [width])',
  INDIRECT: 'INDIRECT(ref_text, [a1])', HYPERLINK: 'HYPERLINK(url, [friendly_name])', TRANSPOSE: 'TRANSPOSE(array)',
  MEDIAN: 'MEDIAN(number1, …)', STDEV: 'STDEV(number1, …)', LARGE: 'LARGE(array, k)', SMALL: 'SMALL(array, k)', RANK: 'RANK(number, ref, [order])',
  PERCENTILE: 'PERCENTILE(array, k)', TEXTSPLIT: 'TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty])',
  MAXIFS: 'MAXIFS(max_range, range1, criteria1, …)', MINIFS: 'MINIFS(min_range, range1, criteria1, …)', ABS: 'ABS(number)', MOD: 'MOD(number, divisor)',
  POWER: 'POWER(number, power)', SQRT: 'SQRT(number)', INT: 'INT(number)', UPPER: 'UPPER(text)', LOWER: 'LOWER(text)', PROPER: 'PROPER(text)',
  VALUE: 'VALUE(text)', ISBLANK: 'ISBLANK(value)', ISNUMBER: 'ISNUMBER(value)', ISERROR: 'ISERROR(value)', CHOOSE: 'CHOOSE(index, value1, …)'
};

/* ================================================================ engine */

/** Functions whose result may be an array (formulas using them can spill). */
const ARRAY_FNS = new Set(['TRANSPOSE', 'UNIQUE', 'SORT', 'SORTBY', 'FILTER', 'SEQUENCE', 'RANDARRAY', 'VSTACK', 'HSTACK', 'TAKE',
  'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'TEXTSPLIT', 'MMULT', 'MINVERSE', 'MUNIT',
  'FREQUENCY', 'MODE.MULT', 'TREND', 'OFFSET', 'INDIRECT', 'INDEX', 'XLOOKUP', 'REGEXEXTRACT', 'ROW', 'COLUMN']);
const REDUCERS = new Set(['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTIF',
  'COUNTIFS', 'COUNTBLANK', 'MIN', 'MAX', 'MINIFS', 'MAXIFS', 'MEDIAN', 'PRODUCT', 'STDEV', 'VAR',
  'LOOKUP', 'CONCAT', 'TEXTJOIN', 'AND', 'OR', 'ROWS', 'COLUMNS', 'LARGE', 'SMALL', 'RANK', 'NPV', 'IRR', 'SUBTOTAL', 'AGGREGATE']);
// they give an array for an array of lookup values (their first argument), one value otherwise
const LOOKUP_FNS = new Set(['VLOOKUP', 'HLOOKUP', 'MATCH']);
function mayArray(node) {
  if (!node) return false;
  switch (node.k) {
    case 'ref': return node.ref.kind !== 'cell' || !!node.ref.spill;
    case 'sref': return true;
    case 'lit': return isMatrix(node.v);
    case 'name': return true;
    case 'fn':
      if (ARRAY_FNS.has(node.name)) return true;
      if (REDUCERS.has(node.name)) return false;
      if (LOOKUP_FNS.has(node.name)) return mayArray(node.args[0]);
      return node.args.some(mayArray);
    case 'bin': case 'span': return mayArray(node.a) || mayArray(node.b);
    case 'neg': case 'pct': return mayArray(node.a);
    default: return false;
  }
}
/**
 * Like mayArray, but stricter, for formulas kept as #CALC! (see Engine.noteSlowSpill): this-row table references and
 * names of a single cell give one value, so a formula made of them does not spill.
 */
// functions that give one value whatever their arguments (spillsArray), and ones that give an array only for an array
// of their later arguments (the k of PERCENTILE)
const SCALAR_FNS = new Set(['MDETERM', 'STDEV.S', 'STDEV.P', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR.S', 'VAR.P', 'VARA', 'VARP',
  'VARPA', 'CORREL', 'PEARSON', 'COVARIANCE.P', 'COVARIANCE.S', 'COVAR', 'SLOPE', 'INTERCEPT', 'RSQ', 'STEYX', 'GEOMEAN',
  'HARMEAN', 'KURT', 'SKEW', 'SKEW.P', 'DEVSQ', 'AVEDEV', 'SUMSQ', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'XNPV', 'XIRR', 'MIRR',
  'AVERAGEA', 'MAXA', 'MINA', 'MODE', 'MODE.SNGL']);
const K_FNS = new Set(['PERCENTILE', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE', 'QUARTILE.INC', 'QUARTILE.EXC',
  'PERCENTRANK', 'PERCENTRANK.INC', 'PERCENTRANK.EXC']);
function spillsArray(node, names, depth) {
  if (!node) return false;
  switch (node.k) {
    case 'ref': return node.ref.kind !== 'cell' || !!node.ref.spill;
    case 'sref': { const inner = String(node.spec || '').slice(1, -1).trim(); return !(inner[0] === '@' || /#this\s+row/i.test(inner)); }
    case 'lit': return isMatrix(node.v);
    case 'name': {
      const up = String(node.v).toUpperCase();
      const def = names[node.v] ?? Object.entries(names).find(([k]) => k.toUpperCase() === up)?.[1];
      if (def === undefined || depth > 4) return true; // a table's name, or names of names of names
      return spillsArray(parseFormula(String(def).replace(/^=/, '')), names, depth + 1);
    }
    case 'fn':
      // ROW() and COLUMN() of one cell give one number
      if (node.name === 'ROW' || node.name === 'COLUMN') return node.args.length > 0 && spillsArray(node.args[0], names, depth);
      if (ARRAY_FNS.has(node.name)) return true;
      if (REDUCERS.has(node.name) || SCALAR_FNS.has(node.name)) return false;
      if (LOOKUP_FNS.has(node.name)) return spillsArray(node.args[0], names, depth);
      if (K_FNS.has(node.name)) return node.args.slice(1).some((a) => spillsArray(a, names, depth));
      return node.args.some((a) => spillsArray(a, names, depth));
    case 'bin': case 'span': return spillsArray(node.a, names, depth) || spillsArray(node.b, names, depth);
    case 'neg': case 'pct': return spillsArray(node.a, names, depth);
    default: return false;
  }
}

/**
 * The formulas of a sheet that may spill (sorted by row, then column) that a recalculation has not evaluated yet. Each
 * node of a tree over the list keeps the smallest column below it, so the first one above and left of a cell is found
 * in log n steps, however many there are.
 */
const NO_COL = 0x7fffffff;
class SpillScan {
  constructor(cands) {
    let size = 1;
    while (size < cands.length) size *= 2;
    const min = new Int32Array(2 * size).fill(NO_COL);
    for (let i = 0; i < cands.length; i++) min[size + i] = cands[i][1];
    for (let k = size - 1; k >= 1; k--) min[k] = Math.min(min[2 * k], min[2 * k + 1]);
    this.cands = cands; this.size = size; this.min = min;
  }
  /** The first one in node k (list positions lo .. lo + span - 1) with row <= r and column <= c, or -1. */
  first(k, lo, span, r, c) {
    if (lo >= this.cands.length || this.min[k] > c || this.cands[lo][0] > r) return -1;
    if (span === 1) return lo;
    const half = span >> 1, i = this.first(2 * k, lo, half, r, c);
    return i >= 0 ? i : this.first(2 * k + 1, lo + half, half, r, c);
  }
  /** Takes the first one with row <= r and column <= c (it is not returned again): its list position, or -1. */
  take(r, c) {
    const i = this.first(1, 0, this.size, r, c);
    if (i < 0) return -1;
    let k = this.size + i;
    this.min[k] = NO_COL;
    for (k >>= 1; k >= 1; k >>= 1) this.min[k] = Math.min(this.min[2 * k], this.min[2 * k + 1]);
    return i;
  }
}

/**
 * Evaluates the workbook. `book` is the live view:
 *   book.sheets: [{ id, name, rows, cols, cells }]
 *   book.names: { NAME: "=Sheet1!$A$1:$B$4" }
 *   book.findTable(name) / book.tableAt(sheet, r, c) → { name, sheet, r1, c1, r2, c2, header, totals, columns: [names] }
 * Results are cached until invalidate().
 */
export class Engine {
  constructor(book, locale) {
    this.book = book;
    this.locale = locale || 'en-US';
    this.invalidate();
  }
  invalidate() {
    this.cache = new Map();    // "sid|r|c" → value; "@sid|r|c|formula" → value of a formula outside cells (evalAt)
    this.stack = new Set();
    this.spills = new Map();   // anchor "sid|r|c" → { rows, cols, m }
    this.cover = new Map();    // covered "sid|r|c" → anchor key
    this.links = new Map();    // "sid|r|c" → url (HYPERLINK)
    this.indexes = new Map();  // sheet id → { candidates }
    this.names = new Map();    // "NAME|sid|r|c" → value of a defined name where it is used
    this.scans = new Map();    // sheet id → { left: SpillScan, busy }: spill candidates not evaluated yet
    this.filled = new Map();   // sheet id → Map(row id → Set of column ids that have a cell)
    fixedFormulas.clear();     // defined names may have changed (see cellIndependent)
    this.forgetHeavy();
    refill();
  }
  /** Resets what is known about formulas too heavy to compute (see stopped), with the computed values. */
  forgetHeavy() {
    this.done = 0;               // formulas and names computed (or found too heavy) so far
    this.slowReads = 0;          // reads that gave the #CALC! of a formula too heavy to compute (see hit)
    this.heavyGroups = new Set(); // groups of formulas outside cells found too heavy (a conditional formatting rule)
    this.heavyNames = new Set(); // defined names found too heavy wherever they are used
    this.slowSpills = new Map(); // sheet id → { keys, rows, big }: where formulas kept as #CALC! might spill (noteSlowSpill)
    this.atCache = new Map();    // results of formulas outside cells (evalAt), and the generation before (see keep)
    this.atOld = new Map();
    this.atHeavy = new Set();    // keys of formulas outside cells found too heavy for their cell
    this.groupProofs = new Map(); // group → cells of it found too heavy (see proofs)
    this.nameProofs = new Map(); // name and sheet → the same for positions of a defined name
    this.computed = new Set();   // groups, and names on a sheet, that were computed for some cell (see proofs)
    this.fullAt = -1;            // the burst in which the kept results of evalAt reached KEEP_MAX / 2 (see keep)
    this.usefulMs = this.usefulMs || 0; // time of evaluations that computed something (see waste), never reset
    this.inner = null;           // the innermost evaluation that a stop of the budget went through (see stopped)
    this.pending = null;         // the evaluation to start the next burst with (see settle)
    // never reset, so that callers can tell whether a burst got further (see progress)
    this.work = this.work || 0;
    this.heavy = this.heavy || 0;
    this.wastedMs = this.wastedMs || 0;
  }
  /** Time spent on evaluations that the work budget stopped before they computed anything (in ms, never reset): the
   * widget limits it per change. */
  wasted() { return this.wastedMs; }
  /** Adds to wasted() the time of an evaluation that started at t0 and was just stopped: what it computed before
   * (usefulMs since u0: the formulas it read and computed) is not lost. A stop by the clock on a page in the background
   * is not counted: it proves nothing (see
   * provenHeavy), and the evaluation is done again when the page is in front. */
  waste(t0, u0) {
    if (byClock && pageHidden()) return;
    this.wastedMs += Math.max(0, clockNow() - t0 - (this.usefulMs - u0));
  }
  /** Counts one more cell of a group, or position of a name, found too heavy: true from the third on when none of its
   * cells was computed since values were cleared, so that it is then too heavy everywhere (one burst for each of its
   * cells could keep the page busy for a long time). One that was computed somewhere goes on cell by cell. */
  proofs(map, k) {
    if (this.computed.has(k)) return false;
    const n = (map.get(k) || 0) + 1;
    map.set(k, n);
    return n >= 3;
  }
  /** Counts work done for the widget with values of this engine (statistics of a conditional formatting rule) as
   * progress (see progress). */
  noteProgress() { this.work++; }
  /** Grows whenever something is computed or found too heavy to compute: a burst that leaves it as it was computed
   * nothing new. */
  progress() { return this.work; }
  /** Grows whenever a formula, a defined name or a rule is found too heavy to compute (see stopped). */
  heavyFound() { return this.heavy; }
  /** Whether the value kept for a cell is the #CALC! of a formula too heavy to compute (or of one built from it), and
   * not only a value that is not computed yet. */
  keptSlow(sheet, r, c) { return this.cache.get(sheet.id + '|' + r + '|' + c) === SLOW; }
  /** Counts a read of something that is not known because of a formula too heavy to compute, such as a row that a table
   * filter cannot tell is hidden: the formula being evaluated becomes #CALC! (see tainted). */
  slowRead() { this.slowReads++; }
  /** Forgets computed values but keeps the per-sheet indexes: for repeated recalculation when only constants
   * changed (Solver, Goal Seek, scenario summaries). */
  invalidateValues() {
    this.cache = new Map();
    this.stack = new Set();
    this.spills = new Map();
    this.cover = new Map();
    this.links = new Map();
    this.names = new Map();
    this.scans = new Map();
    this.filled = new Map();
    this.forgetHeavy();
    refill();
  }
  sheetByName(name) {
    const low = String(name).toLowerCase();
    return this.book.sheets.find((s) => s.name.toLowerCase() === low) || null;
  }
  raw(sheet, r, c) {
    const rid = sheet.rows[r], cid = sheet.cols[c];
    if (rid === undefined || cid === undefined) return undefined;
    // most cells of a big range are empty: find that out without building the "row:col" key
    let filled = this.filled.get(sheet.id);
    if (!filled) {
      filled = new Map();
      for (const k in sheet.cells) {
        const i = k.indexOf(':'), row = k.slice(0, i);
        if (!filled.has(row)) filled.set(row, new Set());
        filled.get(row).add(k.slice(i + 1));
      }
      this.filled.set(sheet.id, filled);
    }
    const cols = filled.get(rid);
    if (!cols || !cols.has(cid)) return undefined;
    const cell = sheet.cells[rid + ':' + cid];
    // stored values are text; anything else (damaged or foreign data) counts as an empty cell
    return cell && typeof cell.v === 'string' ? cell.v : undefined;
  }
  index(sheet) {
    let ix = this.indexes.get(sheet.id);
    if (ix) return ix;
    const rows = new Map(sheet.rows.map((id, i) => [id, i]));
    const cols = new Map(sheet.cols.map((id, i) => [id, i]));
    const candidates = [];
    for (const k in sheet.cells) {
      const cell = sheet.cells[k], v = cell && cell.v;
      if (typeof v !== 'string' || v[0] !== '=') continue;
      if (!mayArray(parseFormula(v.slice(1)))) continue;
      const p = k.split(':');
      const r = rows.get(p[0]), c = cols.get(p[1]);
      if (r !== undefined && c !== undefined) candidates.push([r, c]);
    }
    candidates.sort((a, b) => a[0] - b[0] || a[1] - b[1]); // row by row, the order the grid is drawn in
    // keys: to tell whether a cell kept as #CALC! is one of them (see noteSlowSpill)
    ix = { candidates, keys: new Set(candidates.map((p) => sheet.id + '|' + p[0] + '|' + p[1])) };
    this.indexes.set(sheet.id, ix);
    return ix;
  }
  /** Computed value of a cell (number, string, boolean, null or FormulaError). */
  valueAt(sheet, r, c) {
    if (r < 0 || c < 0 || r >= sheet.rows.length || c >= sheet.cols.length) return null;
    // reads made by a formula count against the work budget
    const top = this.stack.size === 0;
    if (!top) tick(1);
    else if (this.pending) this.settle();
    const raw = this.raw(sheet, r, c);
    let v;
    if (raw === undefined || raw === null || raw === '') {
      // part of a spilled array? (empty cells are not kept, so that reading a whole sheet does not fill the cache)
      if (!this.index(sheet).candidates.length) return null;
      const unsure = top && spent() ? false : this.scanSpills(sheet, r, c);
      if (this.cover.size) {
        const key = sheet.id + '|' + r + '|' + c;
        if (this.cache.has(key)) return this.hit(this.cache.get(key));
        const anchor = this.cover.get(key);
        if (anchor) {
          const sp = this.spills.get(anchor);
          const [, ar, ac] = anchor.split('|').map(Number);
          const row = sp && sp.m[r - ar];
          if (row && c - ac < row.length) {
            v = row[c - ac];
            if (v === null || v === undefined) v = 0;
            if (!unsure && !spent()) { this.cache.set(key, v); this.work++; }
            return v;
          }
          this.cover.delete(key); // not covered any more
        }
      }
      // an empty cell that a formula too heavy to compute might spill over (see noteSlowSpill): whether it is covered
      // is not known, so a formula that reads it is not computed either (see tainted). Drawn, the cell is empty.
      if (!top && this.slowSpillOver(sheet, r, c)) this.slowReads++;
      return null;
    }
    const key = sheet.id + '|' + r + '|' + c;
    if (this.cache.has(key)) return this.hit(this.cache.get(key));
    if (raw[0] === '=' && raw.length > 1) {
      // once the budget is spent, a formula that is not computed yet gives #CALC! (not kept, so it is computed when
      // there is budget again); typed values are read as usual
      if (top && spent()) return SLOW;
      if (this.stack.has(key)) return ERR['#CIRC!'];
      const whole = wholeBudget(), done = this.done, reads = this.slowReads, t0 = clockNow(), u0 = this.usefulMs;
      if (top) this.inner = null;
      this.stack.add(key);
      try { v = this.evalCell(raw.slice(1), sheet, r, c, key); } finally { this.stack.delete(key); }
      if (spent() || this.slowReads !== reads) {
        // stopped by the budget, or built from a formula too heavy to compute: the result may be partial or wrong, so
        // nothing of it is kept, not even the array it spilled
        this.dropSpill(sheet, r, c, key);
        this.links.delete(key);
        if (!spent()) { this.usefulMs = u0 + (clockNow() - t0); return this.tainted(key, sheet, r, c); }
        // the innermost evaluation that the stop went through is the first to get here
        if (!this.inner) this.inner = { key, sheet, r, c };
        if (top) this.waste(t0, u0);
        return this.stopped(key, whole, done, top, sheet, r, c);
      }
      this.done++;
      // its time is useful (it replaces the time of the formulas it read, which is part of it)
      this.usefulMs = u0 + (clockNow() - t0);
    } else v = literal(raw);
    this.cache.set(key, v);
    this.work++;
    return v;
  }
  /** A kept value. Reading the #CALC! of a formula too heavy to compute (see stopped) is counted, so that what is built
   * from it is #CALC! as well (see tainted). */
  hit(v) {
    if (v === SLOW) this.slowReads++;
    return v;
  }
  /**
   * Whether an evaluation that the work budget stopped is too heavy to compute in any burst: it started with the whole
   * budget (see wholeBudget) and computed nothing on the way. The clock does not prove it on a page in the background,
   * which runs slower (the widget draws again when the page is in front). Counted as found.
   */
  provenHeavy(whole, done) {
    if (this.done !== done || !whole || (byClock && pageHidden())) return false;
    this.done++;
    this.work++;
    this.heavy++;
    return true;
  }
  /**
   * After an evaluation of a cell that the work budget stopped: gives #CALC! (SLOW).
   * - The stop may have come from a formula or name that it reads (this.inner, the innermost evaluation the stop went
   *   through). When that one is not proven too heavy (it did not start with the whole budget, because this one read a
   *   lot first), it starts the next burst (see settle), and this one is not kept as heavy: otherwise every formula
   *   that reads a heavy one would be found heavy on its own, one burst each.
   * - One proven too heavy keeps SLOW until values are cleared, so that later bursts get to the formulas after it.
   * - A cell of the drawing (top) that is not proven (it did not start with the whole budget) starts the next burst.
   */
  stopped(key, whole, done, top, sheet, r, c) {
    const inner = this.inner;
    if (inner && inner.key !== key && !inner.proven) {
      if (top) { if (!inner.at) this.pending = inner; this.inner = null; }
      return SLOW;
    }
    const proven = this.provenHeavy(whole, done);
    if (proven) { this.cache.set(key, SLOW); this.noteSlowSpill(sheet, r, c, key); }
    if (inner && inner.key === key) inner.proven = proven;
    if (top) {
      if (!proven && (!inner || inner.key === key)) this.pending = { key, sheet, r, c };
      this.inner = null;
    }
    return SLOW;
  }
  /** Starts a burst with the evaluation that stopped a formula reading it in an earlier burst (see stopped), when the
   * whole budget is left: it is then computed, or proven too heavy, once. */
  settle() {
    if (!wholeBudget()) return;
    const p = this.pending;
    this.pending = null;
    if (p.name) {
      const t0 = clockNow(), u0 = this.usefulMs;
      try { this.resolveName(p.name, { sheet: p.sheet, r: p.r, c: p.c, cell: false }); } catch (e) { if (e !== SLOW) throw e; }
      if (spent()) this.waste(t0, u0);
      this.inner = null;
    } else if (!this.cache.has(p.key) && p.r < p.sheet.rows.length && p.c < p.sheet.cols.length) this.valueAt(p.sheet, p.r, p.c);
  }
  /**
   * Notes a formula kept as #CALC! that may give an array (see index and spillsArray): whether it covers the empty cells
   * right of it and below it is not known (see valueAt). Only cells it could cover count: up to the first cell with a
   * value below it and right of it (an array that would cover a value gives #SPILL! instead).
   */
  noteSlowSpill(sheet, r, c, key) {
    if (!sheet || !this.index(sheet).keys.has(key)) return;
    let s = this.slowSpills.get(sheet.id);
    if (!s) { s = { keys: new Set(), rows: new Map(), big: [] }; this.slowSpills.set(sheet.id, s); }
    if (s.keys.has(key)) return;
    s.keys.add(key);
    const raw = this.raw(sheet, r, c);
    if (typeof raw !== 'string' || raw[0] !== '=' || !spillsArray(parseFormula(raw.slice(1)), this.book.names || {}, 0)) return;
    const empty = (rr, cc) => { const x = this.raw(sheet, rr, cc); return x === undefined || x === ''; };
    let r2 = r + 1, c2 = c + 1;
    while (r2 < sheet.rows.length && r2 - r <= 4096 && empty(r2, c)) r2++;
    while (c2 < sheet.cols.length && empty(r, c2)) c2++;
    if (r2 - r > 4096) r2 = sheet.rows.length; // looked at far enough: to the end of the sheet
    // small areas by row (a read looks at its own row only), big ones in a list
    if ((r2 - r) * (c2 - c) <= 4096) {
      for (let i = r; i < r2; i++) { let l = s.rows.get(i); if (!l) s.rows.set(i, (l = [])); l.push(c, c2); }
    } else s.big.push(r, r2, c, c2);
  }
  /** Whether a formula kept as #CALC! might spill over (r, c) of the sheet (see noteSlowSpill). */
  slowSpillOver(sheet, r, c) {
    const s = this.slowSpills.get(sheet.id);
    if (!s) return false;
    const l = s.rows.get(r);
    if (l) for (let i = 0; i < l.length; i += 2) if (c >= l[i] && c < l[i + 1]) return true;
    const b = s.big;
    for (let i = 0; i < b.length; i += 4) if (r >= b[i] && r < b[i + 1] && c >= b[i + 2] && c < b[i + 3]) return true;
    return false;
  }
  /**
   * The #CALC! of an evaluation that read the #CALC! of a formula too heavy to compute, kept like that one: the value
   * read is not known, so neither is what is built from it (IFERROR, ISERROR or COUNT would otherwise turn it into a
   * real value, which commands would then sort, total or export).
   */
  tainted(key, sheet, r, c) {
    this.cache.set(key, SLOW);
    this.done++;
    this.work++;
    this.noteSlowSpill(sheet, r, c, key);
    return SLOW;
  }
  /** Forgets the array spilled from a cell and the cells it covers. */
  dropSpill(sheet, r, c, key) {
    const sp = this.spills.get(key);
    if (!sp) return;
    this.spills.delete(key);
    for (let i = 0; i < sp.rows; i++) {
      for (let j = 0; j < sp.cols; j++) {
        const ck = sheet.id + '|' + (r + i) + '|' + (c + j);
        if ((i || j) && this.cover.get(ck) === key) { this.cover.delete(ck); this.cache.delete(ck); }
      }
    }
  }
  /**
   * Evaluates the formulas that may spill over (r, c), the ones above and left of it (each once per recalculation),
   * so that an empty cell knows whether an array covers it. Formulas right of the cell or below it cannot cover it and
   * are not evaluated here: one that uses the formula being computed would get a false #CIRC!. True when one that is
   * still being evaluated could cover (r, c).
   */
  scanSpills(sheet, r, c) {
    const cands = this.index(sheet).candidates, at = (p) => sheet.id + '|' + p[0] + '|' + p[1];
    let st = this.scans.get(sheet.id);
    if (!st) { st = { left: new SpillScan(cands), busy: [] }; this.scans.set(sheet.id, st); }
    // the formulas evaluated here are not read by the one that reads (r, c): one too heavy to compute does not make
    // that one #CALC! (see hit). Like any error, its #CALC! does not spill, so (r, c) is empty as far as it goes.
    const reads = this.slowReads;
    try {
      for (let i; !spent() && (i = st.left.take(r, c)) >= 0;) {
        const p = cands[i];
        st.busy.push(p);
        if (!this.stack.has(at(p))) this.valueAt(sheet, p[0], p[1]);
      }
    } finally {
      this.slowReads = reads;
      if (spent()) this.scans.delete(sheet.id); // cut short by the work budget: scan again next time
    }
    if (!st.busy.length) return false;
    st.busy = st.busy.filter((p) => this.stack.has(at(p)));
    return st.busy.some((p) => p[0] <= r && p[1] <= c);
  }
  evalCell(text, sheet, r, c, key) {
    const v = this.evalFormula(text, { sheet, r, c, cell: true });
    if (v instanceof Range || isMatrix(v)) {
      // a range that cannot spill is not read cell by cell first
      if (v instanceof Range && (v.rows > 1 || v.cols > 1) && (r + v.rows > sheet.rows.length || c + v.cols > sheet.cols.length)) return ERR['#SPILL!'];
      let m;
      try { m = toMatrix(v); } catch (e) { return isErr(e) ? e : ERR['#VALUE!']; }
      if (!m.length || !m[0]) return ERR['#CALC!']; // an empty array
      if (m.length === 1 && m[0].length === 1) return fix(m[0][0]);
      // spill into the cells below and to the right
      if (r + m.length > sheet.rows.length || c + m[0].length > sheet.cols.length) return ERR['#SPILL!'];
      for (let i = 0; i < m.length; i++) {
        for (let j = 0; j < m[0].length; j++) {
          if (!i && !j) continue;
          const raw = this.raw(sheet, r + i, c + j);
          const ck = sheet.id + '|' + (r + i) + '|' + (c + j);
          if ((raw !== undefined && raw !== '') || (this.cover.has(ck) && this.cover.get(ck) !== key)) return ERR['#SPILL!'];
        }
      }
      this.spills.set(key, { rows: m.length, cols: m[0].length, m });
      for (let i = 0; i < m.length; i++) for (let j = 0; j < m[0].length; j++) if (i || j) this.cover.set(sheet.id + '|' + (r + i) + '|' + (c + j), key);
      return fix(m[0][0]);
    }
    return v;
    function fix(x) { return x === null || x === undefined ? 0 : x; }
  }
  /** Size of the array spilled from a cell, or null. */
  spillAt(sheet, r, c) {
    this.valueAt(sheet, r, c);
    return this.spills.get(sheet.id + '|' + r + '|' + c) || null;
  }
  /** The anchor cell of the spill covering a cell, or null. */
  spillAnchorOf(sheet, r, c) {
    const k = this.cover.get(sheet.id + '|' + r + '|' + c);
    if (!k) return null;
    const [, ar, ac] = k.split('|').map(Number);
    return { r: ar, c: ac };
  }
  linkAt(sheet, r, c) { return this.links.get(sheet.id + '|' + r + '|' + c) || null; }
  evalFormula(text, ctx) {
    F._locale = this.locale;
    const node = parseFormula(text);
    if (node.k === 'bad') return ERR['#ERROR!'];
    let v;
    try {
      v = this.ev(node, ctx);
      if (v instanceof Range) { if (v.rows === 1 && v.cols === 1) v = v.get(0, 0); else return v; }
    } catch (e) {
      if (isErr(e)) return e;
      console.warn('Formula failed', text, e);
      return ERR['#VALUE!'];
    }
    if (isMatrix(v)) return v;
    if (typeof v === 'number' && !isFinite(v)) return ERR['#NUM!'];
    if (v === null || v === undefined) return 0; // =A1 with an empty A1 shows 0, like Excel
    if (longText(v)) return ERR['#VALUE!']; // longer than a cell can hold
    return v;
  }
  /**
   * Evaluates a formula typed somewhere other than a cell (conditional formatting, validation). group names what the
   * formula belongs to, for all the cells it is evaluated for (a conditional formatting rule): its results are kept
   * until values are cleared (at most KEEP_MAX), so that drawing again does not compute them again and a redraw gets
   * further than the one before; and once it is too heavy for one cell, it is #CALC! for all of them, instead of
   * costing a whole burst for each. Without a group only a formula too heavy is kept (as #CALC!, see stopped). Keys are
   * ones that no cell has: cell keys start with a sheet id.
   */
  evalAt(formula, sheet, r, c, group, shared) {
    const key = '@' + sheet.id + '|' + r + '|' + c + '|' + (group || formula);
    // a result computed before stays, also once its group is found too heavy elsewhere
    const kept = this.keptValue(key);
    if (kept !== undefined) return kept;
    if (group && this.heavyGroups.has(group)) { this.slowReads++; return SLOW; }
    if (this.pending && this.stack.size === 0) this.settle();
    if (spent()) return SLOW;
    const whole = wholeBudget(), done = this.done, reads = this.slowReads, t0 = clockNow(), u0 = this.usefulMs;
    this.inner = null;
    // on the stack while it is evaluated: the cells it reads are then reads of a formula, which count against the work
    // budget (see valueAt), so that a formula outside cells is stopped like one in a cell
    this.stack.add(key);
    let v;
    try { v = scalar(this.evalFormula(String(formula).replace(/^=/, ''), { sheet, r, c })); } finally { this.stack.delete(key); }
    if (spent()) {
      const inner = this.inner || { key };
      this.inner = null;
      this.waste(t0, u0);
      // stopped in a formula or name that it reads, not proven too heavy: that one starts the next burst (see stopped)
      if (inner.key !== key && !inner.proven) { if (!inner.at) this.pending = inner; return SLOW; }
      // too heavy: for every cell when the formula is the same everywhere (shared), or once it was too heavy in a few
      // cells in a row (see proofs); else for this cell
      if (inner.key === key && this.provenHeavy(whole, done)) {
        if (group && (shared || this.proofs(this.groupProofs, group))) this.heavyGroups.add(group);
        else this.atHeavy.add(key);
      }
      return SLOW;
    }
    this.usefulMs = u0 + (clockNow() - t0);
    if (group) this.computed.add(group);
    if (this.slowReads !== reads) { this.keep(key, SLOW); this.done++; this.work++; return SLOW; } // see tainted
    if (group) { this.keep(key, v); this.done++; this.work++; }
    return v;
  }
  /** The kept result of evalAt for a key (a read of it, see hit), or undefined. */
  keptValue(key) {
    if (this.atHeavy.has(key)) { this.slowReads++; return SLOW; }
    if (this.atCache.has(key)) return this.hit(this.atCache.get(key));
    if (this.atOld.has(key)) { const v = this.atOld.get(key); this.keep(key, v); return this.hit(v); }
    return undefined;
  }
  /** Keeps a result of evalAt. Results go to a new generation when the current one holds KEEP_MAX / 2, and the
   * generation before is then dropped (its results are computed again when they are asked for), so that drawing a big
   * sheet with many rules cannot fill the memory while the results of the last drawings stay. */
  keep(key, v) {
    if (this.atCache.size >= KEEP_MAX / 2) {
      // a new generation, but not in the middle of a burst: the results of the drawing in progress stay
      if (this.fullAt < 0) this.fullAt = burstNo;
      else if (this.fullAt !== burstNo) { this.atOld = this.atCache; this.atCache = new Map(); this.fullAt = -1; }
    }
    this.atCache.set(key, v);
  }
  /** The kept result of a group's formula for a cell (see evalAt), or undefined: the caller then builds the formula for
   * the cell (a conditional formatting rule's formula, translated for each cell) and evaluates it. */
  keptAt(sheet, r, c, group) {
    const v = this.keptValue('@' + sheet.id + '|' + r + '|' + c + '|' + group);
    if (v !== undefined) return v;
    if (this.heavyGroups.has(group)) { this.slowReads++; return SLOW; }
    return undefined;
  }
  rangeFor(ref, ctx) {
    const sheet = ref.sheet != null ? this.sheetByName(ref.sheet) : ctx.sheet;
    if (!sheet) throw ERR['#REF!'];
    if (ref.spill) {
      const sp = this.spillAt(sheet, ref.r1, ref.c1);
      if (!sp) {
        const raw = this.raw(sheet, ref.r1, ref.c1);
        if (raw && raw[0] === '=') return new Range(this, sheet, ref.r1, ref.c1, ref.r1, ref.c1);
        throw ERR['#REF!'];
      }
      return new Range(this, sheet, ref.r1, ref.c1, ref.r1 + sp.rows - 1, ref.c1 + sp.cols - 1);
    }
    let { r1, r2, c1, c2 } = ref;
    if (ref.kind === 'cols') { r1 = 0; r2 = sheet.rows.length - 1; }
    if (ref.kind === 'rows') { c1 = 0; c2 = sheet.cols.length - 1; }
    if (r1 > r2) [r1, r2] = [r2, r1];
    if (c1 > c2) [c1, c2] = [c2, c1];
    return this.clip(sheet, r1, c1, r2, c2);
  }
  /** A range cut to the rows and columns the sheet has (but at least one cell). */
  clip(sheet, r1, c1, r2, c2) {
    return new Range(this, sheet, r1, c1, Math.min(r2, Math.max(r1, sheet.rows.length - 1)), Math.min(c2, Math.max(c1, sheet.cols.length - 1)));
  }
  resolveName(name, ctx) {
    if (ctx.vars && name in ctx.vars) return ctx.vars[name];
    const names = this.book.names || {};
    const def = names[name] ?? Object.entries(names).find(([k]) => k.toUpperCase() === name)?.[1];
    if (def !== undefined) {
      // a name is evaluated once for each cell that uses it (it may depend on that cell, as in ROW() or [@Col]);
      // evaluating it again for every reference made names that use names cost exponential time
      const at = name + '|' + (ctx.sheet ? ctx.sheet.id : '') + '|' + ctx.r + '|' + ctx.c;
      if (this.names.has(at)) return this.hit(this.names.get(at));
      // a name too heavy to compute is #CALC! wherever it is used on the sheet (references without a sheet and SHEET()
      // are the using cell's): finding that out costs a whole burst, once
      const onSheet = name + '|' + (ctx.sheet ? ctx.sheet.id : '');
      if (this.heavyNames.has(onSheet)) { this.slowReads++; return SLOW; }
      const key = 'name|' + name;
      if (this.stack.has(key)) return ERR['#CIRC!'];
      const whole = wholeBudget(), done = this.done, reads = this.slowReads, t0 = clockNow(), u0 = this.usefulMs;
      this.stack.add(key);
      let v, thrown = false;
      try {
        const node = parseFormula(String(def).replace(/^=/, ''));
        // cell: false, so that a HYPERLINK in a name does not set the link of the cells that use it
        v = node.k === 'bad' ? ERR['#NAME?'] : this.ev(node, { ...ctx, vars: null, cell: false });
      } catch (e) {
        // the work budget's stop is thrown through the name: it is noted here, then goes on up as before
        if (e !== SLOW || !spent()) throw e;
        thrown = true;
      } finally { this.stack.delete(key); }
      if (spent()) { // stopped by the work budget: the value may be partial
        const nk = key + '|' + at;
        if (!this.inner) this.inner = { key: nk, name, sheet: ctx.sheet, r: ctx.r, c: ctx.c };
        if (this.inner.key === nk) {
          // too heavy: wherever it is used when it gives the same everywhere, else where it was used
          const proven = this.inner.proven = this.provenHeavy(whole, done);
          if (proven) {
            if (cellIndependent(def, this.book.names) || this.proofs(this.nameProofs, onSheet)) this.heavyNames.add(onSheet);
            else this.names.set(at, SLOW);
          }
        }
        if (thrown) throw SLOW;
        return SLOW;
      }
      if (this.slowReads !== reads) v = SLOW; // built from a formula too heavy to compute (see tainted)
      this.names.set(at, v);
      this.computed.add(onSheet);
      this.usefulMs = u0 + (clockNow() - t0);
      this.done++;
      this.work++;
      return v;
    }
    const table = this.book.findTable ? this.book.findTable(name) : null;
    if (table) return this.tableRange(table, '#Data');
    return undefined;
  }
  tableRange(t, part, colFrom, colTo, thisRow) {
    const sheet = t.sheet;
    const dataTop = t.r1 + (t.header ? 1 : 0), dataBottom = t.r2 - (t.totals ? 1 : 0);
    let r1 = dataTop, r2 = dataBottom;
    if (part === '#All') { r1 = t.r1; r2 = t.r2; }
    else if (part === '#Headers') { if (!t.header) throw ERR['#REF!']; r1 = r2 = t.r1; }
    else if (part === '#Totals') { if (!t.totals) throw ERR['#REF!']; r1 = r2 = t.r2; }
    else if (part === '#Headers,#Data') { r1 = t.r1; }
    else if (part === '#Data,#Totals') { r2 = t.r2; }
    if (thisRow !== undefined) { if (thisRow < dataTop || thisRow > dataBottom) throw ERR['#VALUE!']; r1 = r2 = thisRow; }
    let c1 = t.c1, c2 = t.c2;
    if (colFrom !== undefined) {
      const find = (name) => { const i = t.columns.findIndex((x) => x.toLowerCase() === name.toLowerCase()); if (i < 0) throw ERR['#REF!']; return t.c1 + i; };
      c1 = find(colFrom); c2 = colTo !== undefined ? find(colTo) : c1;
      if (c2 < c1) [c1, c2] = [c2, c1];
    }
    if (r2 < r1) return [[null]];
    return new Range(this, sheet, r1, c1, r2, c2);
  }
  structured(node, ctx) {
    const t = node.table ? this.book.findTable && this.book.findTable(node.table) : this.book.tableAt && this.book.tableAt(ctx.sheet, ctx.r, ctx.c);
    if (!t) throw ERR['#REF!'];
    // spec: [Col] | [#Data] | [@Col] | [@[Col name]] | [[#Headers],[Col]] | [[Col1]:[Col2]] | [#This Row],[Col]
    let inner = node.spec.slice(1, -1).trim();
    if (!inner) return this.tableRange(t, '#Data');
    const unesc = (s) => s.replace(/'(.)/g, '$1');
    let thisRow;
    if (inner[0] === '@') { thisRow = ctx.r; inner = inner.slice(1).trim(); if (!inner) return this.tableRange(t, '#Data', undefined, undefined, thisRow); }
    const items = [];
    if (inner[0] === '[') {
      let depth = 0, cur = '';
      for (let i = 0; i < inner.length; i++) {
        const ch = inner[i];
        if (ch === "'" && i + 1 < inner.length) { cur += ch + inner[++i]; continue; }
        if (ch === '[') { depth++; if (depth === 1) { cur = ''; continue; } }
        if (ch === ']') { depth--; if (depth === 0) { items.push(cur); continue; } }
        if (depth === 0) { if (ch === ':') items.push(':'); continue; }
        cur += ch;
      }
    } else items.push(inner);
    let part = '#Data';
    const specials = items.filter((x) => x[0] === '#');
    if (specials.some((x) => /^#this row$/i.test(x))) thisRow = ctx.r;
    const parts = specials.filter((x) => !/^#this row$/i.test(x)).map((x) => '#' + x.slice(1, 2).toUpperCase() + x.slice(2).toLowerCase());
    if (parts.length) part = parts.join(',');
    const cols = items.filter((x) => x[0] !== '#' && x !== ':').map(unesc);
    return this.tableRange(t, part, cols[0], cols[1], thisRow);
  }
  ev(node, ctx) {
    tick(1);
    switch (node.k) {
      case 'lit': return node.v;
      case 'empty': return null;
      case 'ref': return this.rangeFor(node.ref, ctx);
      case 'sref': return this.structured(node, ctx);
      case 'name': { const v = this.resolveName(node.v, ctx); return v === undefined ? ERR['#NAME?'] : v; }
      case 'neg': return this.unary(this.ev(node.a, ctx), (x) => -x);
      case 'pct': return this.unary(this.ev(node.a, ctx), (x) => x / 100);
      case 'span': {
        const a = this.ev(node.a, ctx), b = this.ev(node.b, ctx);
        if (!(a instanceof Range) || !(b instanceof Range) || a.sheet !== b.sheet) return ERR['#REF!'];
        return this.clip(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      }
      case 'bin': return this.binary(node, ctx);
      case 'fn': {
        const fn = F[node.name];
        if (!fn) return ERR['#NAME?'];
        if (LAZY.has(node.name)) return fn(node.args, (n) => this.ev(n, ctx), ctx, this);
        const args = [];
        for (const a of node.args) {
          if (a.k === 'empty') { args.push(null); continue; }
          const v = this.ev(a, ctx);
          if (!RAW.has(node.name) && isErr(v)) return v;
          args.push(v);
        }
        const out = fn(...args);
        return longText(out) ? ERR['#VALUE!'] : out; // text longer than a cell can hold
      }
      default: return ERR['#ERROR!'];
    }
  }
  unary(v, fn) {
    if (isMulti(v)) return mapMatrix(toMatrix(v), (x) => { const n = toNum(x); return isErr(n) ? n : fn(n); });
    const n = toNum(v);
    return isErr(n) ? n : fn(n);
  }
  binary(node, ctx) {
    const A = this.ev(node.a, ctx), B = this.ev(node.b, ctx);
    if (isMulti(A) || isMulti(B)) return broadcast(A, B, (a, b) => op(node.op, a, b));
    return op(node.op, scalar(A), scalar(B));
  }
}

function op(o, a, b) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (o) {
    case '&': {
      const x = toStr(a), y = toStr(b);
      if (isErr(x)) return x;
      if (isErr(y)) return y;
      return x.length + y.length > MAX_TEXT ? ERR['#VALUE!'] : x + y; // not longer than a cell can hold
    }
    case '=': return compare(a, b) === 0;
    case '<>': return compare(a, b) !== 0;
    case '<': return compare(a, b) < 0;
    case '>': return compare(a, b) > 0;
    case '<=': return compare(a, b) <= 0;
    case '>=': return compare(a, b) >= 0;
    default: {
      const x = toNum(a); if (isErr(x)) return x;
      const y = toNum(b); if (isErr(y)) return y;
      switch (o) {
        case '+': return x + y;
        case '-': return x - y;
        case '*': return x * y;
        case '/': return y === 0 ? ERR['#DIV/0!'] : x / y;
        case '^': return finite(Math.pow(x, y));
        default: return ERR['#VALUE!'];
      }
    }
  }
}

/** Stored (non-formula) value → typed value. Stored numbers are canonical ("1234.5"). */
export function literal(raw) {
  if (typeof raw !== 'string') return null;
  if (raw[0] === "'") return raw.slice(1);
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw)) return parseFloat(raw);
  const up = raw.toUpperCase();
  if (up === 'TRUE') return true;
  if (up === 'FALSE') return false;
  if (ERR[up]) return ERR[up];
  return raw;
}

/* ================================================================ user input */

const infoCache = new Map();
export function localeInfo(locale) {
  if (infoCache.has(locale)) return infoCache.get(locale);
  let info;
  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
    const dec = (parts.find((p) => p.type === 'decimal') || {}).value || '.';
    const grp = (parts.find((p) => p.type === 'group') || {}).value || ',';
    const dparts = new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'UTC' }).formatToParts(Date.UTC(2026, 10, 25, 12));
    const order = dparts.filter((p) => ['year', 'month', 'day'].includes(p.type)).map((p) => p.type[0]).join('');
    info = { dec, grp, dateOrder: order || 'mdy' };
  } catch (e) {
    info = { dec: '.', grp: ',', dateOrder: 'mdy' };
  }
  infoCache.set(locale, info);
  return info;
}

// The number part starts and ends with a digit or separator, so that a run of spaces can be read in one way only
// (overlapping \s* around [\d\s]+ made these patterns take cubic time on long runs of spaces).
const CURRENCY_RE = /^([₺$€£¥₽₹]|TL|USD|EUR|TRY|GBP)?\s*((?:[-+]\s*)?[\d.,'](?:[\d.,\s']*[\d.,'])?)\s*([₺$€£¥₽₹]|TL|USD|EUR|TRY|GBP)?$/i;
const PERCENT_RE = /^%?\s*((?:[-+]\s*)?[\d.,](?:[\d.,\s]*[\d.,])?)\s*%$|^%\s*((?:[-+]\s*)?[\d.,](?:[\d.,\s]*[\d.,])?)$/;

function parseLocaleNumber(s, info) {
  let t = s.replace(/[\s ']/g, '');
  if (!/^[-+]?[\d.,]+$/.test(t) || !/\d/.test(t)) return null;
  const hasDot = t.includes('.'), hasComma = t.includes(',');
  if (hasDot && hasComma) {
    const dec = t.lastIndexOf('.') > t.lastIndexOf(',') ? '.' : ',';
    t = t.split(dec === '.' ? ',' : '.').join('').replace(dec, '.');
  } else if (hasComma || hasDot) {
    const sep = hasComma ? ',' : '.';
    const parts = t.split(sep);
    const head = parts[0].replace(/^[-+]/, '');
    const groupsOk = parts.length > 1 && parts.slice(1).every((p) => p.length === 3) && head.length >= 1 && head.length <= 3 && head !== '0';
    if (parts.length > 2) { if (!groupsOk) return null; t = parts.join(''); }
    else if (sep === info.dec) t = parts.join('.');
    else if (groupsOk) t = parts.join('');
    else t = parts.join('.');
  }
  const n = Number(t);
  return isFinite(n) ? n : null;
}
function decimalsOf(numText, info) {
  const t = numText.trim();
  const i = t.lastIndexOf(info.dec);
  if (i < 0) return 0;
  const tail = t.slice(i + 1).replace(/\D/g, '');
  return tail.length <= 4 ? tail.length : 2;
}

/**
 * What the user typed → { kind, raw, num?, fmt? }. `raw` is what is stored: formulas as typed, numbers and
 * dates in canonical form (so every user sees the same value whatever their locale), `fmt` a format hint.
 */
export function parseInput(text, locale) {
  const info = localeInfo(locale);
  const s = String(text);
  const t = s.trim();
  if (s === '') return { kind: 'empty', raw: '' };
  if (s[0] === '=' && s.length > 1) return { kind: 'formula', raw: s };
  if (s[0] === '+' && s.length > 1 && /[A-Za-z(]/.test(s)) return { kind: 'formula', raw: '=' + s.slice(1) };
  if (s[0] === "'") return { kind: 'text', raw: s };
  if (t.length > 1000) return { kind: 'text', raw: s }; // far too long for a number or a date
  const up = t.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE' || up === 'DOĞRU' || up === 'YANLIŞ') return { kind: 'bool', raw: up === 'TRUE' || up === 'DOĞRU' ? 'TRUE' : 'FALSE' };
  let m = PERCENT_RE.exec(t);
  if (m) {
    const numPart = m[1] || m[2];
    const n = parseLocaleNumber(numPart, info);
    if (n !== null) return { kind: 'number', raw: numToText(+(n / 100).toPrecision(15)), num: n / 100, fmt: { nf: 'pct', dp: decimalsOf(numPart, info) } };
  }
  if (/^[-+]?[\d.,\s ']+$/.test(t)) {
    const n = parseLocaleNumber(t, info);
    if (n !== null) {
      const grouped = new RegExp('\\d\\' + info.grp + '\\d{3}').test(t) && Math.abs(n) >= 1000;
      return { kind: 'number', raw: numToText(n), num: n, fmt: grouped ? { nf: 'num', dp: decimalsOf(t, info) } : null };
    }
  }
  if (/^[-+]?\d+(?:\.\d+)?e[-+]?\d+$/i.test(t)) { const n = Number(t); return { kind: 'number', raw: numToText(n), num: n, fmt: { nf: 'sci', dp: 2 } }; }
  m = CURRENCY_RE.exec(t);
  if (m && (m[1] || m[3]) && !(m[1] && m[3])) {
    const n = parseLocaleNumber(m[2], info);
    if (n !== null) {
      const sym = (m[1] || m[3]).toUpperCase();
      const cur = { TL: '₺', TRY: '₺', USD: '$', EUR: '€', GBP: '£' }[sym] || sym;
      return { kind: 'number', raw: numToText(n), num: n, fmt: { nf: 'cur', cur, dp: decimalsOf(m[2], info) || 2 } };
    }
  }
  const dt = parseDateTime(t, info);
  if (dt) return dt;
  return { kind: 'text', raw: s };
}

function parseDateTime(t, info) {
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  let y, mo, d, time = null;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; if (m[4] !== undefined) time = [+m[4], +m[5], +(m[6] || 0)]; }
  else {
    m = /^(\d{1,2})([./-])(\d{1,2})\2(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
    if (m) {
      const a = +m[1], b = +m[3];
      y = +m[4]; if (y < 100) y += 2000;
      if (info.dateOrder.startsWith('m')) { mo = a; d = b; } else { d = a; mo = b; }
      if (mo > 12 && d <= 12) [mo, d] = [d, mo];
      if (m[5] !== undefined) time = [+m[5], +m[6], +(m[7] || 0)];
    } else {
      m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(t);
      if (m && +m[1] < 24 && +m[2] < 60) {
        let h = +m[1];
        if (m[4]) { if (h === 12) h = 0; if (/pm/i.test(m[4])) h += 12; }
        const num = (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
        return { kind: 'time', raw: numToText(+num.toFixed(10)), num, fmt: { nf: 'time' } };
      }
      return null;
    }
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return null;
  if (new Date(Date.UTC(y, mo - 1, d)).getUTCDate() !== d) return null;
  let num = dateToSerial(y, mo, d);
  if (time) {
    if (time[0] > 23 || time[1] > 59 || time[2] > 59) return null;
    num += (time[0] * 3600 + time[1] * 60 + time[2]) / 86400;
    return { kind: 'datetime', raw: numToText(+num.toFixed(10)), num, fmt: { nf: 'dt' } };
  }
  return { kind: 'date', raw: numToText(num), num, fmt: { nf: 'date' } };
}

/* ================================================================ display */

const nfCache = new Map();
function numberFormatter(locale, opts) {
  const key = locale + JSON.stringify(opts);
  let f = nfCache.get(key);
  if (!f) {
    try { f = new Intl.NumberFormat(locale, opts); } catch (e) { f = new Intl.NumberFormat('en-US', opts); }
    nfCache.set(key, f);
  }
  return f;
}

const EURO = ['DE', 'FR', 'ES', 'IT', 'NL', 'BE', 'AT', 'PT', 'FI', 'IE', 'GR', 'SK', 'SI', 'EE', 'LV', 'LT', 'LU', 'MT', 'CY', 'HR'];
const CURRENCY_BY_REGION = { TR: '₺', US: '$', CA: '$', AU: '$', NZ: '$', GB: '£', JP: '¥', CN: '¥', IN: '₹', RU: '₽', KR: '₩', BR: 'R$', CH: 'CHF', PL: 'zł', CZ: 'Kč', SE: 'kr', NO: 'kr', DK: 'kr', UA: '₴', IL: '₪' };
const CURRENCY_BY_LANGUAGE = { tr: '₺', en: '$', de: '€', fr: '€', es: '€', it: '€', nl: '€', pt: '€', fi: '€', el: '€', ja: '¥', zh: '¥', ru: '₽', ko: '₩', pl: 'zł', cs: 'Kč', sv: 'kr', nb: 'kr', da: 'kr', uk: '₴', he: '₪', hi: '₹' };
/** Currency symbol for a locale ("tr-TR" → ₺, "en-GB" → £, "de" → €); $ when unknown. */
export function defaultCurrency(locale) {
  const parts = String(locale || '').replace('_', '-').split('-');
  const region = (parts.find((p, i) => i > 0 && /^[A-Za-z]{2}$/.test(p)) || '').toUpperCase();
  if (region) return EURO.includes(region) ? '€' : CURRENCY_BY_REGION[region] || '$';
  return CURRENCY_BY_LANGUAGE[(parts[0] || '').toLowerCase()] || '$';
}

/** How a computed value is shown in a cell with the given style. */
export function formatValue(v, st, locale) {
  if (v === null || v === undefined) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (isMatrix(v)) return formatValue(v[0] && v[0][0], st, locale);
  if (typeof v !== 'number') return '';
  // a damaged style (out-of-range decimals, a pattern no formatter accepts) shows the plain number, never throws
  try { return formatNumber(v, st, locale); } catch (e) { return numToText(v); }
}
function formatNumber(v, st, locale) {
  const nf = (st && st.nf) || 'gen';
  const dp = st && st.dp !== undefined ? st.dp : null;
  switch (nf) {
    case 'num': return numberFormatter(locale, { minimumFractionDigits: dp ?? 2, maximumFractionDigits: dp ?? 2, useGrouping: true }).format(v);
    case 'cur': {
      const txt = numberFormatter(locale, { minimumFractionDigits: dp ?? 2, maximumFractionDigits: dp ?? 2, useGrouping: true }).format(Math.abs(v));
      const sym = (st && st.cur) || defaultCurrency(locale);
      const after = /^(tr|de|fr|es|it|nl|pl|ru|pt|cs|sv|fi|da|nb)/i.test(locale) && sym !== '$' && sym !== '£';
      return (v < 0 ? '-' : '') + (after ? txt + ' ' + sym : sym + txt);
    }
    case 'acc': {
      const txt = numberFormatter(locale, { minimumFractionDigits: dp ?? 2, maximumFractionDigits: dp ?? 2, useGrouping: true }).format(Math.abs(v));
      return (v < 0 ? '(' + txt + ')' : txt);
    }
    case 'pct': return numberFormatter(locale, { style: 'percent', minimumFractionDigits: dp ?? 0, maximumFractionDigits: dp ?? 0 }).format(v);
    case 'sci': {
      // past 15 significant digits toExponential shows the binary error (9.95 → 9.949999999999999E+0): zeros instead,
      // as other spreadsheet apps show them
      const d = dp ?? 2;
      let t = v.toExponential(Math.min(d, 14));
      if (d > 14) t = t.replace('e', '0'.repeat(d - 14) + 'e');
      return t.toUpperCase().replace('.', localeInfo(locale).dec);
    }
    case 'date': return formatSerialDate(v, locale, false);
    case 'ldate': return formatSerialDate(v, locale, false, true);
    case 'dt': return formatSerialDate(v, locale, true);
    case 'time': { const p = serialToParts(v); return String(p.H).padStart(2, '0') + ':' + String(p.M).padStart(2, '0') + (p.S ? ':' + String(p.S).padStart(2, '0') : ''); }
    case 'dur': { const total = Math.round(v * 86400); const h = Math.floor(total / 3600), mm = Math.floor((total % 3600) / 60), ss = total % 60; return h + ':' + String(mm).padStart(2, '0') + ':' + String(ss).padStart(2, '0'); }
    case 'text': return numToText(v);
    case 'custom': return st.pat ? formatWithPattern(v, st.pat, locale) : numToText(v);
    default: {
      if (dp !== null) return numberFormatter(locale, { minimumFractionDigits: dp, maximumFractionDigits: dp, useGrouping: false }).format(v);
      if (Math.abs(v) >= 1e11 || (Math.abs(v) < 1e-9 && v !== 0)) return v.toExponential(5).replace(/\.?0+e/, 'e').replace('e', 'E').replace('.', localeInfo(locale).dec);
      if (Number.isInteger(v)) return numberFormatter(locale, { maximumFractionDigits: 0, useGrouping: false }).format(v); // whole numbers exactly
      return numberFormatter(locale, { maximumSignificantDigits: 10, useGrouping: false }).format(v);
    }
  }
}

export function formatSerialDate(v, locale, withTime, long) {
  const p = serialToParts(v);
  const opts = long ? { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long', timeZone: 'UTC' } : { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'UTC' };
  if (withTime) Object.assign(opts, { hour: '2-digit', minute: '2-digit', hourCycle: 'h23' });
  try {
    return new Intl.DateTimeFormat(locale, opts).format(Date.UTC(p.y, p.m - 1, p.d, p.H, p.M, p.S));
  } catch (e) {
    return `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`;
  }
}

/** Text shown in the formula bar and the cell editor for a stored value (dates as dates, locale decimals). */
/*
 * Formula separators. Formulas are stored in one form: "," between arguments and "." in numbers (=ROUND(1.5,0)), so
 * that people with different languages share them and .xlsx files get them as the file format expects. People whose
 * style is "semicolon" (the default for languages that write 1,5) see and type ";" between arguments and "," in numbers
 * (=ROUND(1,5;0)), like spreadsheet apps set to such a language. Array constants ({1,2;3,4}) keep their stored form.
 */

/** 'comma' or 'semicolon': the person's choice ('comma' | 'semicolon'), or for 'auto' the one that fits the language. */
export function separatorStyle(pref, locale) {
  if (pref === 'comma' || pref === 'semicolon') return pref;
  return localeInfo(locale).dec === ',' ? 'semicolon' : 'comma';
}

/** A stored formula as the person sees it in their separator style. Other text is returned unchanged. */
export function formulaToLocal(raw, style) {
  if (typeof raw !== 'string' || raw[0] !== '=') return raw;
  const toks = tokenize(raw.slice(1));
  if (toks.some((t) => t.t === 'bad')) return raw;
  const semi = style === 'semicolon';
  return '=' + toks.map((t) => (t.t === 'sep' ? (semi ? ';' : ',') : semi && t.t === 'num' ? t.text.replace('.', ',') : t.text)).join('');
}

/** A formula typed in the person's separator style, in the stored form. Other text is returned unchanged. */
export function formulaFromLocal(text, style) {
  if (style !== 'semicolon' || typeof text !== 'string' || text[0] !== '=') return text;
  let out = '';
  let i = 0;
  const s = text;
  // the end of a quoted part ("text" or 'sheet name'), where a doubled quote stands for one quote
  const quoted = (q, from) => { let j = from + 1; while (j < s.length) { if (s[j] === q) { if (s[j + 1] === q) { j += 2; continue; } return j; } j++; } return s.length - 1; };
  const closing = (open, close, from) => { let d = 0; for (let j = from; j < s.length; j++) { if (s[j] === open) d++; else if (s[j] === close && --d === 0) return j; } return s.length - 1; };
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"' || ch === "'") { const j = quoted(ch, i); out += s.slice(i, j + 1); i = j + 1; continue; }
    if (ch === '[') { const j = closing('[', ']', i); out += s.slice(i, j + 1); i = j + 1; continue; }
    if (ch === '{') { const j = closing('{', '}', i); out += s.slice(i, j + 1); i = j + 1; continue; }
    if (ch === ';') { out += ','; i++; continue; }
    if (ch === ',' && /\d/.test(s[i + 1] || '')) {
      // a decimal comma follows the digits of a number, not of a reference (A1,5) or a name
      const m = /(\d+)$/.exec(out);
      const before = m ? out[out.length - m[1].length - 1] : null;
      if (m && !/[A-Za-z0-9_$.\]]/.test(before || '')) { out += '.'; i++; continue; }
    }
    out += ch;
    i++;
  }
  return out;
}

/** The signature help of a function (SUM(number1, [number2], …)) in the person's separator style. */
export function helpToLocal(text, style) {
  return style === 'semicolon' && text ? text.replace(/, /g, '; ') : text;
}

export function editText(raw, st, locale) {
  if (typeof raw !== 'string') return '';
  if (raw[0] === '=' || raw[0] === "'") return raw;
  const v = literal(raw);
  if (typeof v !== 'number') return raw;
  const nf = st && st.nf;
  const dec = localeInfo(locale).dec;
  if (nf === 'date' || nf === 'ldate' || nf === 'dt') return formatSerialDate(v, locale, nf === 'dt' || v % 1 !== 0);
  if (nf === 'time') return formatValue(v, st, locale);
  if (nf === 'pct') return numToText(+(v * 100).toPrecision(15)).replace('.', dec) + '%';
  return numToText(v).replace('.', dec);
}
