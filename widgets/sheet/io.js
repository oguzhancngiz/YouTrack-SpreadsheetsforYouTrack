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
/* Spreadsheets for YouTrack – import and export: XLSX (Excel) and CSV. No external libraries: ZIP is read and written here,
 * DEFLATE uses the browser's CompressionStream / DecompressionStream. */
import { colName, colIndex, parseInput, formatValue, isErr, numToText, parseRefText, tokenize, quoteSheet, rewriteRefs, literal } from './formula.js';
import { DEFAULT_OPTIONS as SOLVER_DEFAULTS } from './solver.js';
import { cellAt, indexOf, objRect } from './model.js';

const Core = window.SheetCore;

/* ================================================================ ZIP */

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
async function streamBytes(bytes, stream) {
  const res = new Response(new Blob([bytes]).stream().pipeThrough(stream));
  return new Uint8Array(await res.arrayBuffer());
}
const canDeflate = typeof CompressionStream === 'function';

export async function unzip(buffer) {
  const data = new Uint8Array(buffer);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let eocd = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65557); i--) if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('This is not an .xlsx file.');
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const files = {};
  const dec = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== 0x02014b50) break;
    const method = view.getUint16(p + 10, true);
    const csize = view.getUint32(p + 20, true);
    const nlen = view.getUint16(p + 28, true), elen = view.getUint16(p + 30, true), clen = view.getUint16(p + 32, true);
    const off = view.getUint32(p + 42, true);
    const name = dec.decode(data.subarray(p + 46, p + 46 + nlen));
    p += 46 + nlen + elen + clen;
    const lnlen = view.getUint16(off + 26, true), lelen = view.getUint16(off + 28, true);
    const start = off + 30 + lnlen + lelen;
    files[name] = { method, raw: data.subarray(start, start + csize) };
  }
  return {
    names: Object.keys(files),
    async text(name) {
      const f = files[name] || files[Object.keys(files).find((k) => k.toLowerCase() === name.toLowerCase())];
      if (!f) return null;
      const bytes = f.method === 0 ? f.raw : await streamBytes(f.raw, new DecompressionStream('deflate-raw'));
      return dec.decode(bytes);
    }
  };
}

export async function zip(entries) {
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  for (const { name, text } of entries) {
    const raw = enc.encode(text);
    const nameBytes = enc.encode(name);
    const crc = crc32(raw);
    let body = raw, method = 0;
    if (canDeflate && raw.length > 256) { body = await streamBytes(raw, new CompressionStream('deflate-raw')); method = 8; }
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
    local.setUint16(8, method, true); local.setUint32(14, crc, true);
    local.setUint32(18, body.length, true); local.setUint32(22, raw.length, true); local.setUint16(26, nameBytes.length, true);
    parts.push(new Uint8Array(local.buffer), nameBytes, body);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true);
    cd.setUint16(10, method, true); cd.setUint32(16, crc, true); cd.setUint32(20, body.length, true); cd.setUint32(24, raw.length, true);
    cd.setUint16(28, nameBytes.length, true); cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), nameBytes);
    offset += 30 + nameBytes.length + body.length;
  }
  const cdSize = central.reduce((s, b) => s + b.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ================================================================ XLSX import */

const BUILTIN_FMT = {
  0: null, 1: { nf: 'gen', dp: 0 }, 2: { nf: 'gen', dp: 2 }, 3: { nf: 'num', dp: 0 }, 4: { nf: 'num', dp: 2 },
  9: { nf: 'pct', dp: 0 }, 10: { nf: 'pct', dp: 2 }, 11: { nf: 'sci', dp: 2 }, 14: { nf: 'date' }, 15: { nf: 'custom', pat: 'd-mmm-yy' },
  16: { nf: 'custom', pat: 'd-mmm' }, 17: { nf: 'custom', pat: 'mmm-yy' }, 18: { nf: 'custom', pat: 'h:mm AM/PM' },
  19: { nf: 'custom', pat: 'h:mm:ss AM/PM' }, 20: { nf: 'time' }, 21: { nf: 'time' }, 22: { nf: 'dt' }, 37: { nf: 'acc', dp: 0 },
  38: { nf: 'acc', dp: 0 }, 39: { nf: 'acc', dp: 2 }, 40: { nf: 'acc', dp: 2 }, 45: { nf: 'custom', pat: 'mm:ss' },
  46: { nf: 'dur' }, 47: { nf: 'custom', pat: 'mm:ss.0' }, 48: { nf: 'sci', dp: 1 }, 49: { nf: 'text' }
};
function fmtFromCode(code) {
  const c = code.replace(/\[\$-[0-9A-F]+\]/gi, '').trim();
  if (/^general$/i.test(c)) return null;
  if (c === '@') return { nf: 'text' };
  const bare = c.replace(/"[^"]*"|\\./g, '');
  const dpm = /\.(0+)/.exec(bare);
  const dp = dpm ? dpm[1].length : 0;
  if (/^0(\.0+)?%$/.test(bare)) return { nf: 'pct', dp };
  if (/^#,##0(\.0+)?$/.test(bare)) return { nf: 'num', dp };
  if (/^0(\.0+)?$/.test(bare)) return { nf: 'gen', dp };
  if (/^\[h\]:mm(:ss)?$/i.test(bare)) return { nf: 'dur' };
  const cur = /(₺|\$|€|£|TL)/.exec(c);
  if (cur && /#,##0/.test(bare) && !/[dmyhs]/i.test(bare.replace(/\[[^\]]*\]/g, ''))) return { nf: 'cur', cur: cur[1] === 'TL' ? '₺' : cur[1], dp };
  if (/^(d{1,2}|m{1,2})[./-](d{1,2}|m{1,2})[./-]y{2,4}$/i.test(bare) || /^y{4}-m{2}-d{2}$/i.test(bare)) return { nf: 'date' };
  return { nf: 'custom', pat: c.slice(0, 80) };
}
const xmlDoc = (text) => new DOMParser().parseFromString(text, 'application/xml');
const kids = (el, tag) => (el ? Array.from(el.getElementsByTagName(tag)) : []);
const first = (el, tag) => (el ? el.getElementsByTagName(tag)[0] || null : null);
function argb(el) {
  if (!el) return null;
  const rgb = el.getAttribute('rgb');
  if (rgb && /^[0-9A-F]{8}$/i.test(rgb)) return '#' + rgb.slice(2).toLowerCase();
  if (rgb && /^[0-9A-F]{6}$/i.test(rgb)) return '#' + rgb.toLowerCase();
  const theme = el.getAttribute('theme');
  if (theme !== null) {
    const themeColors = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
    return themeColors[+theme] || null;
  }
  const idx = el.getAttribute('indexed');
  if (idx !== null) {
    const map = { 8: '#000000', 9: '#ffffff', 10: '#ff0000', 11: '#00ff00', 12: '#0000ff', 13: '#ffff00', 14: '#ff00ff', 15: '#00ffff' };
    return map[+idx] || null;
  }
  return null;
}
function cellRef(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref);
  return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null;
}
function rangeRef(ref) {
  const [a, b] = ref.split(':');
  const A = cellRef(a), B = b ? cellRef(b) : A;
  return A && B ? { r1: Math.min(A.r, B.r), c1: Math.min(A.c, B.c), r2: Math.max(A.r, B.r), c2: Math.max(A.c, B.c) } : null;
}
const cleanFormula = (f) => '=' + f.replace(/_xlfn\.ANCHORARRAY\(([^()]+)\)/gi, '$1#').replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./gi, '').replace(/_xlpm\./gi, '');

/** Reads an .xlsx file into sheets (in our format) and defined names. */
export async function readXlsx(buffer, { translate } = {}) {
  const z = await unzip(buffer);
  const wb = xmlDoc(await z.text('xl/workbook.xml') || '');
  if (!first(wb, 'workbook')) throw new Error('This is not a valid .xlsx workbook.');
  const rels = xmlDoc(await z.text('xl/_rels/workbook.xml.rels') || '<Relationships/>');
  const relMap = {};
  kids(rels, 'Relationship').forEach((r) => { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
  const resolve = (base, target) => {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
    return parts.join('/');
  };

  // shared strings
  const shared = [];
  const sst = await z.text('xl/sharedStrings.xml');
  if (sst) kids(xmlDoc(sst), 'si').forEach((si) => shared.push(kids(si, 't').map((t) => t.textContent).join('')));

  // styles
  const styles = [];
  const stx = await z.text('xl/styles.xml');
  if (stx) {
    const sd = xmlDoc(stx);
    const numFmts = {};
    kids(first(sd, 'numFmts'), 'numFmt').forEach((n) => { numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode'); });
    const fonts = kids(first(sd, 'fonts'), 'font').map((f) => {
      const st = {};
      if (first(f, 'b') && first(f, 'b').getAttribute('val') !== '0') st.b = 1;
      if (first(f, 'i') && first(f, 'i').getAttribute('val') !== '0') st.i = 1;
      if (first(f, 'u') && first(f, 'u').getAttribute('val') !== 'none') st.u = 1;
      if (first(f, 'strike') && first(f, 'strike').getAttribute('val') !== '0') st.x = 1;
      const color = argb(first(f, 'color'));
      if (color && color !== '#000000') st.fc = color;
      const sz = first(f, 'sz');
      if (sz && Math.round(+sz.getAttribute('val')) !== 11) st.fs = Math.round(+sz.getAttribute('val'));
      const name = (first(f, 'name') || { getAttribute: () => '' }).getAttribute('val') || '';
      if (/times|georgia|serif|garamond|cambria/i.test(name)) st.ff = 'serif';
      else if (/consolas|courier|mono/i.test(name)) st.ff = 'mono';
      return st;
    });
    const fills = kids(first(sd, 'fills'), 'fill').map((f) => {
      const pf = first(f, 'patternFill');
      if (!pf || pf.getAttribute('patternType') !== 'solid') return null;
      return argb(first(pf, 'fgColor'));
    });
    const borders = kids(first(sd, 'borders'), 'border').map((b) => {
      let s = '', color = null;
      for (const [tag, ch] of [['top', 't'], ['left', 'l'], ['right', 'r'], ['bottom', 'b']]) {
        const e = first(b, tag);
        const style = e && e.getAttribute('style');
        if (style && style !== 'none') { s += /medium|thick|double/.test(style) ? ch.toUpperCase() : ch; color = color || argb(first(e, 'color')); }
      }
      return s ? { bd: s, bc: color && color !== '#000000' ? color : null } : null;
    });
    kids(first(sd, 'cellXfs'), 'xf').forEach((xf) => {
      const st = { ...(fonts[+xf.getAttribute('fontId') || 0] || {}) };
      const fill = fills[+xf.getAttribute('fillId') || 0];
      if (fill && fill !== '#ffffff') st.bg = fill;
      const bd = borders[+xf.getAttribute('borderId') || 0];
      if (bd) { st.bd = bd.bd; if (bd.bc) st.bc = bd.bc; }
      const fid = +xf.getAttribute('numFmtId') || 0;
      const fmt = numFmts[fid] !== undefined ? fmtFromCode(numFmts[fid]) : BUILTIN_FMT[fid];
      if (fmt) Object.assign(st, fmt);
      const al = first(xf, 'alignment');
      if (al) {
        const h = al.getAttribute('horizontal'), v = al.getAttribute('vertical');
        if (['left', 'center', 'right'].includes(h)) st.ha = h;
        if (v === 'top') st.va = 'top'; else if (v === 'center') st.va = 'middle';
        if (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true') st.w = 1;
      }
      styles.push(Object.keys(st).length ? st : null);
    });
  }

  const sheets = [];
  const nameOf = {};
  const sheetEls = kids(first(wb, 'sheets'), 'sheet');
  for (const s of sheetEls) {
    const name = s.getAttribute('name');
    const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const path = resolve('xl/workbook.xml', relMap[rid] || '');
    const xml = await z.text(path);
    if (!xml) continue;
    const sd = xmlDoc(xml);
    const cells = [];
    let maxR = 0, maxC = 0;
    const sharedF = {};
    const rowH = {}, hiddenRows = [];
    kids(first(sd, 'sheetData'), 'row').forEach((row) => {
      const r = +row.getAttribute('r') - 1;
      if (row.getAttribute('customHeight') === '1' && row.getAttribute('ht')) rowH[r] = Math.round(+row.getAttribute('ht') * 4 / 3);
      if (row.getAttribute('hidden') === '1') hiddenRows.push(r);
      kids(row, 'c').forEach((c) => {
        const pos = cellRef(c.getAttribute('r'));
        if (!pos) return;
        const t = c.getAttribute('t');
        const fEl = first(c, 'f');
        const vEl = first(c, 'v');
        let v = null;
        if (fEl) {
          const si = fEl.getAttribute('si');
          if (fEl.getAttribute('t') === 'shared' && si !== null) {
            if (fEl.textContent) { sharedF[si] = { f: cleanFormula(fEl.textContent), r: pos.r, c: pos.c }; v = sharedF[si].f; }
            else if (sharedF[si] && translate) v = translate(sharedF[si].f, pos.r - sharedF[si].r, pos.c - sharedF[si].c);
          } else if (fEl.textContent) v = cleanFormula(fEl.textContent);
        }
        if (v === null) {
          if (t === 's') v = vEl ? shared[+vEl.textContent] ?? '' : '';
          else if (t === 'inlineStr') v = kids(c, 't').map((x) => x.textContent).join('');
          else if (t === 'str') v = vEl ? vEl.textContent : '';
          else if (t === 'b') v = vEl && vEl.textContent === '1' ? 'TRUE' : 'FALSE';
          else if (t === 'e') v = vEl ? vEl.textContent : '#N/A';
          else if (vEl) v = numToText(+vEl.textContent);
          if (typeof v === 'string' && (t === 's' || t === 'inlineStr' || t === 'str') && /^(=|'|[-+]?\d|TRUE$|FALSE$|#)/i.test(v)) v = "'" + v;
        }
        const st = styles[+c.getAttribute('s') || 0] || null;
        if ((v === null || v === '') && !st) return;
        cells.push({ r: pos.r, c: pos.c, v: v === '' ? null : v, st });
        maxR = Math.max(maxR, pos.r); maxC = Math.max(maxC, pos.c);
      });
    });
    const colW = {}, hiddenCols = [];
    kids(first(sd, 'cols'), 'col').forEach((col) => {
      const a = +col.getAttribute('min') - 1, b = Math.min(+col.getAttribute('max') - 1, a + 200);
      for (let c = a; c <= b; c++) {
        if (col.getAttribute('width') && col.getAttribute('customWidth') !== '0') colW[c] = Math.round(+col.getAttribute('width') * 7 + 5);
        if (col.getAttribute('hidden') === '1') hiddenCols.push(c);
      }
    });
    const merges = kids(first(sd, 'mergeCells'), 'mergeCell').map((m) => rangeRef(m.getAttribute('ref'))).filter(Boolean);
    merges.forEach((m) => { maxR = Math.max(maxR, m.r2); maxC = Math.max(maxC, m.c2); });
    const pane = first(sd, 'pane');
    const freeze = pane && (pane.getAttribute('state') === 'frozen' || pane.getAttribute('state') === 'frozenSplit')
      ? { r: Math.min(50, +pane.getAttribute('ySplit') || 0), c: Math.min(26, +pane.getAttribute('xSplit') || 0) } : { r: 0, c: 0 };
    const nogrid = (first(sd, 'sheetView') || { getAttribute: () => null }).getAttribute('showGridLines') === '0';
    // data validation (lists)
    const dvs = [];
    kids(first(sd, 'dataValidations'), 'dataValidation').forEach((dv) => {
      const type = dv.getAttribute('type');
      const sqref = (dv.getAttribute('sqref') || '').split(/\s+/)[0];
      const rect = sqref && rangeRef(sqref);
      if (!rect) return;
      const f1 = (first(dv, 'formula1') || {}).textContent || '';
      if (type === 'list') dvs.push({ rect, type: 'list', source: /^".*"$/.test(f1) ? f1.slice(1, -1) : '=' + f1, strict: dv.getAttribute('errorStyle') !== 'warning' && dv.getAttribute('errorStyle') !== 'information' });
      else if (['whole', 'decimal', 'date', 'textLength'].includes(type)) {
        const f2 = (first(dv, 'formula2') || {}).textContent || '';
        const opMap = { between: 'between', notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' };
        dvs.push({ rect, type: { whole: 'int', decimal: 'number', date: 'date', textLength: 'length' }[type], op: opMap[dv.getAttribute('operator') || 'between'] || 'between', a: f1, b: f2, strict: true });
      }
    });
    // tables
    const tables = [];
    const sheetRels = await z.text(path.replace(/([^/]+)$/, '_rels/$1.rels'));
    if (sheetRels) {
      const rm = {};
      kids(xmlDoc(sheetRels), 'Relationship').forEach((r) => { rm[r.getAttribute('Id')] = r.getAttribute('Target'); });
      for (const tp of kids(first(sd, 'tableParts'), 'tablePart')) {
        const id = tp.getAttribute('r:id') || tp.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        const txml = rm[id] && await z.text(resolve(path, rm[id]));
        if (!txml) continue;
        const td = first(xmlDoc(txml), 'table');
        const rect = rangeRef(td.getAttribute('ref') || '');
        if (!rect) continue;
        const styleName = ((first(td, 'tableStyleInfo') || { getAttribute: () => '' }).getAttribute('name') || '');
        const tone = /Medium(2|9|16|23)\b|Light(9|16)\b/.test(styleName) ? 'blue' : /Medium(7|14|21|28)\b/.test(styleName) ? 'green' : /Medium(3|10|17|24)\b/.test(styleName) ? 'orange' : /Medium(1|8|15|22)\b/.test(styleName) ? 'gray' : 'blue';
        tables.push({ rect, name: (td.getAttribute('displayName') || td.getAttribute('name') || 'Table').replace(/[^\w.]/g, '_'), header: td.getAttribute('headerRowCount') !== '0', totals: td.getAttribute('totalsRowCount') === '1' || td.getAttribute('totalsRowShown') === '1' && !!td.getAttribute('totalsRowCount'), style: tone, banded: true, filter: true });
      }
    }
    const scenarios = kids(first(sd, 'scenarios'), 'scenario').map((sc) => ({
      name: sc.getAttribute('name') || 'Scenario', comment: sc.getAttribute('comment') || '',
      cells: kids(sc, 'inputCells').map((ic) => ic.getAttribute('r') || ''), values: kids(sc, 'inputCells').map((ic) => ic.getAttribute('val') ?? '')
    })).filter((sc) => sc.cells.length);
    sheets.push({ name, cells, maxR, maxC, colW, rowH, merges, freeze, nogrid, hiddenRows, hiddenCols, dvs, tables, scenarios, xmlIndex: sheetEls.indexOf(s) });
    nameOf[sheets.length - 1] = name;
  }

  const names = {};
  kids(first(wb, 'definedNames'), 'definedName').forEach((dn) => {
    const n = dn.getAttribute('name');
    const local = dn.getAttribute('localSheetId');
    if (n && local !== null && /^solver_/i.test(n)) {
      // Excel keeps each sheet's Solver model in hidden sheet-level names
      const sh = sheets.find((x) => x.xmlIndex === +local);
      if (sh) (sh.solver = sh.solver || {})[n.toLowerCase()] = dn.textContent;
      return;
    }
    if (!n || n.startsWith('_xlnm.') || local !== null) return;
    names[n.toUpperCase()] = cleanFormula(dn.textContent);
  });
  return { sheets, names };
}

/** Converts what readXlsx / readCsv returned into sheets in our format. */
export function toSheets(imported) {
  const L = Core.LIMITS;
  return imported.sheets.map((s) => {
    const rows = Math.min(L.rows, Math.max(Core.DEFAULT_ROWS, s.maxR + 21));
    const cols = Math.min(L.cols, Math.max(Core.DEFAULT_COLS, s.maxC + 3));
    const sheet = Core.newSheet(String(s.name || 'Sheet1').slice(0, 31), rows, cols);
    for (const cell of s.cells) {
      if (cell.r >= rows || cell.c >= cols) continue;
      const k = sheet.rows[cell.r] + ':' + sheet.cols[cell.c];
      const out = {};
      if (cell.v !== null && cell.v !== undefined) out.v = String(cell.v).slice(0, L.cellChars);
      if (cell.st) out.st = cell.st;
      if (out.v !== undefined || out.st) sheet.cells[k] = out;
    }
    for (const c in s.colW || {}) if (+c < cols) sheet.colW[sheet.cols[+c]] = Math.max(16, Math.min(1000, s.colW[c]));
    for (const r in s.rowH || {}) if (+r < rows) sheet.rowH[sheet.rows[+r]] = Math.max(8, Math.min(600, s.rowH[r]));
    (s.hiddenRows || []).forEach((r) => { if (r < rows) sheet.hr[sheet.rows[r]] = 1; });
    (s.hiddenCols || []).forEach((c) => { if (c < cols) sheet.hc[sheet.cols[c]] = 1; });
    const ids = (rc) => ({ r1: sheet.rows[rc.r1], c1: sheet.cols[rc.c1], r2: sheet.rows[Math.min(rc.r2, rows - 1)], c2: sheet.cols[Math.min(rc.c2, cols - 1)] });
    sheet.merges = (s.merges || []).filter((m) => m.r1 < rows && m.c1 < cols).map(ids);
    sheet.freeze = s.freeze || { r: 0, c: 0 };
    sheet.nogrid = !!s.nogrid;
    sheet.dv = (s.dvs || []).filter((d) => d.rect.r1 < rows && d.rect.c1 < cols).map((d) => ({ id: Core.randomId(), ...ids(d.rect), type: d.type, source: d.source, op: d.op, a: d.a, b: d.b, strict: d.strict }));
    sheet.tables = (s.tables || []).filter((t) => t.rect.r1 < rows && t.rect.c1 < cols).map((t) => ({ id: Core.randomId(), ...ids(t.rect), name: t.name, header: t.header, totals: t.totals, style: t.style, banded: true, filter: true, filters: {} }));
    if (s.solver) {
      const m = solverFromNames(s.solver, sheet.name);
      // a model larger than an object may be is left out rather than failing the whole import
      if (m && JSON.stringify(m).length < L.objectChars - 200) sheet.solver = [{ id: 'model', ...m }];
    }
    sheet.scenarios = (s.scenarios || []).slice(0, L.objects).map((sc) => ({
      id: Core.randomId(), name: String(sc.name).slice(0, 255), comment: String(sc.comment || '').slice(0, 1000),
      cells: sc.cells.map((r) => '=' + String(r).replace(/^\$?([A-Za-z]+)\$?(\d+)$/, '$$$1$$$2')),
      values: sc.values.map((v) => { const x = v === '' ? null : literal(String(v)); return isErr(x) ? null : x; })
    }));
    return sheet;
  });
}

/* ================================================================ Solver models in .xlsx (sheet-level solver_* names) */

const XL_UNLIMITED = 2147483647;
const REL_CODES = { le: 1, eq: 2, ge: 3, int: 4, bin: 5, dif: 6 };
const REL_WORDS = { int: 'integer', bin: 'binary', dif: 'AllDifferent' };
function splitUnion(text) {
  const out = [];
  let cur = '', q = false, depth = 0;
  for (const ch of String(text || '')) {
    if (ch === "'") q = !q;
    if (!q && ch === '(') depth++;
    if (!q && ch === ')') depth--;
    if (!q && !depth && ch === ',') { if (cur.trim()) out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}
/** Excel name text → our formula ("=…"), without the sheet name where it is the model's own sheet. */
function localFormula(text, sheetName) {
  const f = cleanFormula(String(text || '').trim().replace(/^=/, ''));
  return rewriteRefs(f, (ref) => { if (ref.sheet != null && ref.sheet.toLowerCase() === sheetName.toLowerCase()) ref.sheet = null; return ref; });
}
function solverFromNames(nm, sheetName) {
  const get = (k) => (nm[k] === undefined ? undefined : String(nm[k]).trim());
  const numv = (k, d) => { const v = get(k); return v === undefined || v === '' || !isFinite(+v) ? d : +v; };
  const flag = (k, d) => { const v = get(k); return v === undefined ? d : v === '1' || /^true$/i.test(v); };
  const limit = (k, d) => { const v = numv(k, d); return v >= XL_UNLIMITED || v <= 0 ? 0 : v; };
  if (get('solver_opt') === undefined && get('solver_adj') === undefined) return null;
  const typ = numv('solver_typ', 1);
  const eng = numv('solver_eng', get('solver_lin') === '1' ? 2 : 1);
  const cons = [];
  for (let i = 1; i <= Math.min(1000, numv('solver_num', 0)); i++) {
    const l = get('solver_lhs' + i), rel = numv('solver_rel' + i, 1), r = get('solver_rhs' + i);
    const op = Object.keys(REL_CODES).find((k) => REL_CODES[k] === rel);
    if (!l || !op) continue;
    cons.push({ l: localFormula(l, sheetName), op, r: REL_WORDS[op] ? '' : localFormula(r || '0', sheetName) });
  }
  const D = SOLVER_DEFAULTS;
  return {
    obj: get('solver_opt') ? localFormula(get('solver_opt'), sheetName) : '',
    goal: typ === 2 ? 'min' : typ === 3 ? 'value' : 'max',
    value: numToText(numv('solver_val', 0)),
    vars: splitUnion(get('solver_adj')).map((t) => localFormula(t, sheetName)),
    cons,
    nonneg: flag('solver_neg', true),
    method: eng === 2 ? 'lp' : eng === 3 ? 'evo' : 'grg',
    opts: {
      ...D,
      precision: numv('solver_pre', D.precision), scaling: flag('solver_scl', D.scaling), showIter: flag('solver_sho', D.showIter),
      ignoreInt: flag('solver_rlx', D.ignoreInt), intTolerance: numv('solver_tol', D.intTolerance / 100) * 100,
      maxTime: limit('solver_tim', 0), iterations: limit('solver_itr', 0), maxSubproblems: limit('solver_nod', 0), maxFeasible: limit('solver_mip', 0),
      convergence: numv('solver_cvg', D.convergence), derivatives: numv('solver_drv', 1) === 2 ? 'central' : 'forward',
      multistart: flag('solver_msl', D.multistart), population: numv('solver_ssz', D.population), seed: numv('solver_rsd', D.seed),
      requireBounds: flag('solver_rbv', D.requireBounds), mutation: numv('solver_mrt', D.mutation), maxNoImprove: numv('solver_mni', D.maxNoImprove)
    }
  };
}
/** The sheet's Solver model as Excel's hidden sheet-level names. */
function solverNamesXml(engine, sheet, si) {
  const o = (sheet.solver || []).find((x) => x.id === 'model');
  if (!o) return [];
  const opts = { ...SOLVER_DEFAULTS, ...(o.opts || {}) };
  // Solver names hold sheet-qualified absolute references ('Sheet 1'!$B$3), as the names written by other apps do
  const q = (f) => {
    const text = String(f || '').trim();
    if (!text) return '';
    const full = rewriteRefs(text[0] === '=' ? text : '=' + text, (ref) => {
      if (ref.sheet == null) ref.sheet = sheet.name;
      if (ref.kind !== 'rows') { ref.ca1 = true; ref.ca2 = true; }
      if (ref.kind !== 'cols') { ref.ra1 = true; ref.ra2 = true; }
      return ref;
    });
    return exportFormula(engine, full.slice(1), sheet, 0, 0);
  };
  const vars = (o.vars || []).map(q).filter((t) => t && !/#REF!/i.test(t));
  if (!vars.length && !String(o.obj || '').replace(/^=/, '').trim()) return []; // nothing to save
  const out = [];
  // a defined name with an empty formula makes the file invalid: such values are left out
  const add = (n, v) => { if (String(v) !== '') out.push(`<definedName name="${n}" localSheetId="${si}" hidden="1">${xmlEsc(v)}</definedName>`); };
  const b = (v) => (v ? 1 : 2);
  const lim = (v) => (v > 0 ? v : XL_UNLIMITED);
  const cons = Array.isArray(o.cons) ? o.cons : [];
  add('solver_adj', vars.join(','));
  add('solver_cvg', opts.convergence);
  add('solver_drv', opts.derivatives === 'central' ? 2 : 1);
  add('solver_eng', o.method === 'lp' ? 2 : o.method === 'evo' ? 3 : 1);
  add('solver_est', 1);
  add('solver_itr', lim(opts.iterations));
  // constraints whose references no longer work (#REF!) are left out, and the count follows
  const works = (t) => !!t && !/#REF!/i.test(t);
  const usable = cons.filter((c) => works(q(c.l)) && (REL_WORDS[c.op] || works(q(c.r))));
  usable.forEach((c, i) => {
    add('solver_lhs' + (i + 1), q(c.l));
    add('solver_rel' + (i + 1), REL_CODES[c.op] || 1);
    add('solver_rhs' + (i + 1), REL_WORDS[c.op] || q(c.r));
  });
  add('solver_lin', o.method === 'lp' ? 1 : 2);
  add('solver_mip', lim(opts.maxFeasible));
  add('solver_mni', opts.maxNoImprove);
  add('solver_mrt', opts.mutation);
  add('solver_msl', b(opts.multistart));
  add('solver_neg', b(o.nonneg !== false));
  add('solver_nod', lim(opts.maxSubproblems));
  add('solver_num', usable.length);
  add('solver_nwt', 1);
  if (String(o.obj || '').trim()) add('solver_opt', q(o.obj));
  add('solver_pre', opts.precision);
  add('solver_rbv', b(opts.requireBounds));
  add('solver_rlx', b(opts.ignoreInt));
  add('solver_rsd', opts.seed);
  add('solver_scl', b(opts.scaling));
  add('solver_sho', b(opts.showIter));
  add('solver_ssz', opts.population);
  add('solver_tim', lim(opts.maxTime));
  add('solver_tol', opts.intTolerance / 100);
  add('solver_typ', o.goal === 'min' ? 2 : o.goal === 'value' ? 3 : 1);
  add('solver_val', o.value || '0');
  add('solver_ver', 3);
  return out;
}
/** The sheet's scenarios as Excel's <scenarios> element. */
function scenariosXml(sheet) {
  const valText = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? numToText(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
  const attr = (v) => xmlEsc(v).replace(/\n/g, '&#10;');
  const list = (sheet.scenarios || []).map((sc) => {
    const cells = (sc.cells || []).map((t, i) => {
      const a = String(t).replace(/^=/, '').replace(/\$/g, '');
      return /^[A-Z]{1,3}\d+$/i.test(a) ? `<inputCells r="${a.toUpperCase()}" val="${attr(valText((sc.values || [])[i]))}"/>` : '';
    }).filter(Boolean);
    return cells.length ? `<scenario name="${attr(String(sc.name).slice(0, 255))}" locked="1" count="${cells.length}"${sc.comment ? ` comment="${attr(String(sc.comment).slice(0, 255))}"` : ''}>${cells.join('')}</scenario>` : '';
  }).filter(Boolean);
  return list.length ? `<scenarios current="0" show="0">${list.join('')}</scenarios>` : '';
}

/* ================================================================ XLSX export */

const xmlEsc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
function fmtCode(st) {
  if (!st || !st.nf) return st && st.dp !== undefined ? (st.dp ? '0.' + '0'.repeat(st.dp) : '0') : null;
  const dp = st.dp ?? 2, d = dp ? '.' + '0'.repeat(dp) : '';
  switch (st.nf) {
    case 'gen': return st.dp !== undefined ? '0' + d : null;
    case 'num': return '#,##0' + d;
    case 'cur': return st.cur === '$' || st.cur === '£' ? `"${st.cur}"#,##0${d}` : `#,##0${d} "${st.cur || '₺'}"`;
    case 'acc': return `#,##0${d};(#,##0${d})`;
    case 'pct': return '0' + (st.dp ? '.' + '0'.repeat(st.dp) : '') + '%';
    case 'date': return 'yyyy-mm-dd';
    case 'ldate': return 'dddd, mmmm d, yyyy';
    case 'dt': return 'yyyy-mm-dd hh:mm';
    case 'time': return 'hh:mm';
    case 'dur': return '[h]:mm:ss';
    case 'text': return '@';
    case 'sci': return '0' + d + 'E+00';
    case 'custom': return st.pat || null;
    default: return null;
  }
}

/* Functions added to Excel after 2007 are stored with a prefix in .xlsx files; without it Excel shows #NAME?. */
const XLFN = new Set(('AGGREGATE ARABIC ARRAYTOTEXT BASE BINOM.DIST CEILING.MATH CEILING.PRECISE CHOOSECOLS CHOOSEROWS COMBINA ' +
  'CONCAT CONFIDENCE.NORM COT COVARIANCE.P COVARIANCE.S CSC DAYS DECIMAL DROP EXPAND EXPON.DIST FILTER FLOOR.MATH FLOOR.PRECISE ' +
  'FORECAST.LINEAR FORMULATEXT GAMMA GAMMALN.PRECISE HSTACK IFNA IFS ISFORMULA ISOWEEKNUM LET MAXIFS MINIFS MODE.MULT MODE.SNGL ' +
  'NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC ' +
  'PERCENTRANK.INC POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SEC ' +
  'SEQUENCE SHEET SHEETS SORT SORTBY STDEV.P STDEV.S SWITCH TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR ' +
  'UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR MUNIT').split(' '));
const XLWS = new Set(['FILTER', 'SORT']); // these two are stored as _xlfn._xlws.NAME

/**
 * Formula text as Excel stores it: newer functions get their _xlfn. prefix, LET variables their _xlpm. prefix, and table
 * references (Budget[Total], [@Qty]) become cell ranges, because tables are not written to the file.
 */
export function exportFormula(engine, text, sheet, r, c) {
  const toks = tokenize(text);
  const letVars = new Set();
  toks.forEach((t, i) => {
    if (t.t !== 'func' || t.v !== 'LET') return;
    // LET(name1, value1, name2, value2, …, calculation): names follow "(" or a separator at the LET's own depth
    let depth = 0;
    for (let j = i + 1; j < toks.length; j++) {
      const x = toks[j];
      if (x.t === 'lp') { depth++; if (depth === 1) { const n = toks.slice(j + 1).find((y) => y.t !== 'ws'); if (n && n.t === 'name') letVars.add(n.v); } continue; }
      if (x.t === 'rp') { depth--; if (depth === 0) break; continue; }
      if (x.t === 'sep' && depth === 1) {
        const n = toks.slice(j + 1).find((y) => y.t !== 'ws');
        if (n && n.t === 'name') letVars.add(n.v);
      }
    }
  });
  const ctx = { sheet, r, c };
  const rangeText = (rng) => {
    const pre = rng.sheet !== sheet ? quoteSheet(rng.sheet.name) + '!' : '';
    const a = colName(rng.c1) + (rng.r1 + 1), b = colName(rng.c2) + (rng.r2 + 1);
    if (a === b) return pre + a;
    return pre + '$' + colName(rng.c1) + '$' + (rng.r1 + 1) + ':$' + colName(rng.c2) + '$' + (rng.r2 + 1);
  };
  let out = '';
  for (const t of toks) {
    if (t.t === 'func' && XLFN.has(t.v)) { out += (XLWS.has(t.v) ? '_xlfn._xlws.' : '_xlfn.') + t.v; continue; }
    // the file format separates arguments with commas, also when the formula was typed with semicolons
    if (t.t === 'sep') { out += ','; continue; }
    if (t.t === 'name' && letVars.has(t.v)) { out += '_xlpm.' + t.text; continue; }
    // a spilled range (A1#) is stored as ANCHORARRAY(A1)
    if (t.t === 'ref' && t.ref.spill) { out += '_xlfn.ANCHORARRAY(' + t.text.slice(0, -1) + ')'; continue; }
    const isTableName = t.t === 'name' && !(engine.book.names || {})[t.v] && engine.book.findTable && engine.book.findTable(t.text);
    if (t.t === 'sref' || isTableName) {
      try {
        const rng = t.t === 'sref' ? engine.structured({ k: 'sref', table: t.table, spec: t.spec }, ctx) : engine.resolveName(t.v, ctx);
        out += rng && rng.sheet ? rangeText(rng) : t.text;
      } catch (e) { out += '#REF!'; }
      continue;
    }
    out += t.text;
  }
  return out;
}

/** Builds an .xlsx Blob from the workbook (values, formulas, formats, sizes, merges, frozen panes, names, lists). */
export async function writeXlsx(doc, engine, version) {
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], numFmts = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const keyIndex = new Map(), xfIndex = new Map([['{}', 0]]);
  const intern = (list, xml, offset = 0) => { const k = xml; if (!keyIndex.has(k)) { keyIndex.set(k, list.length + offset); list.push(xml); } return keyIndex.get(k); };
  const fontId = (st) => {
    if (!st.b && !st.i && !st.u && !st.x && !st.fc && !st.fs && !st.ff) return 0;
    const name = st.ff === 'serif' ? 'Times New Roman' : st.ff === 'mono' ? 'Consolas' : 'Calibri';
    return intern(fonts, `<font>${st.b ? '<b/>' : ''}${st.i ? '<i/>' : ''}${st.x ? '<strike/>' : ''}${st.u ? '<u/>' : ''}<sz val="${st.fs || 11}"/>${st.fc ? `<color rgb="FF${st.fc.slice(1).toUpperCase()}"/>` : ''}<name val="${name}"/></font>`);
  };
  const fillId = (st) => (st.bg ? intern(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${st.bg.slice(1).toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`) : 0);
  const borderId = (st) => {
    if (!st.bd) return 0;
    const side = (tag, ch) => { const on = st.bd.includes(ch) || st.bd.includes(ch.toUpperCase()); const thick = st.bd.includes(ch.toUpperCase()); return on ? `<${tag} style="${thick ? 'medium' : 'thin'}"><color rgb="FF${(st.bc || '#000000').slice(1).toUpperCase()}"/></${tag}>` : `<${tag}/>`; };
    return intern(borders, `<border>${side('left', 'l')}${side('right', 'r')}${side('top', 't')}${side('bottom', 'b')}<diagonal/></border>`);
  };
  const numFmtId = (st) => {
    const code = fmtCode(st);
    if (!code) return 0;
    const builtin = { '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '@': 49, '[h]:mm:ss': 46 }[code];
    if (builtin) return builtin;
    if (!keyIndex.has('fmt:' + code)) { keyIndex.set('fmt:' + code, 164 + numFmts.length); numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${xmlEsc(code)}"/>`); }
    return keyIndex.get('fmt:' + code);
  };
  const styleIndex = (st) => {
    if (!st) return 0;
    const k = JSON.stringify(st);
    if (xfIndex.has(k)) return xfIndex.get(k);
    const f = fontId(st), fi = fillId(st), b = borderId(st), n = numFmtId(st);
    const al = st.ha || st.va || st.w ? `<alignment${st.ha ? ` horizontal="${st.ha}"` : ''}${st.va ? ` vertical="${st.va === 'middle' ? 'center' : st.va}"` : ''}${st.w ? ' wrapText="1"' : ''}/>` : '';
    xfs.push(`<xf numFmtId="${n}" fontId="${f}" fillId="${fi}" borderId="${b}" xfId="0"${n ? ' applyNumberFormat="1"' : ''}${f ? ' applyFont="1"' : ''}${fi ? ' applyFill="1"' : ''}${b ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`);
    xfIndex.set(k, xfs.length - 1);
    return xfs.length - 1;
  };

  const sheetXml = doc.sheets.map((sheet) => {
    const ix = indexOf(sheet, version);
    let maxR = 0, maxC = 0;
    for (const k in sheet.cells) { const [rid, cid] = k.split(':'); maxR = Math.max(maxR, ix.rows.get(rid) ?? 0); maxC = Math.max(maxC, ix.cols.get(cid) ?? 0); }
    const rowsXml = [];
    for (let r = 0; r <= maxR; r++) {
      const rid = sheet.rows[r];
      const cellsXml = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = sheet.cells[rid + ':' + sheet.cols[c]];
        if (!cell) continue;
        const ref = colName(c) + (r + 1);
        const s = styleIndex(cell.st);
        const sAttr = s ? ` s="${s}"` : '';
        const raw = cell.v;
        if (raw === undefined) { cellsXml.push(`<c r="${ref}"${sAttr}/>`); continue; }
        if (raw[0] === '=' && raw.length > 1) {
          const val = engine.valueAt(sheet, r, c);
          const f = `<f>${xmlEsc(exportFormula(engine, raw.slice(1), sheet, r, c))}</f>`;
          if (typeof val === 'number') cellsXml.push(`<c r="${ref}"${sAttr}>${f}<v>${val}</v></c>`);
          else if (typeof val === 'boolean') cellsXml.push(`<c r="${ref}"${sAttr} t="b">${f}<v>${val ? 1 : 0}</v></c>`);
          else if (isErr(val)) cellsXml.push(`<c r="${ref}"${sAttr} t="e">${f}<v>${xmlEsc(val.code === '#CIRC!' || val.code === '#ERROR!' ? '#VALUE!' : val.code)}</v></c>`);
          else cellsXml.push(`<c r="${ref}"${sAttr} t="str">${f}<v>${xmlEsc(val ?? '')}</v></c>`);
          continue;
        }
        const val = engine.valueAt(sheet, r, c);
        if (typeof val === 'number') cellsXml.push(`<c r="${ref}"${sAttr}><v>${val}</v></c>`);
        else if (typeof val === 'boolean') cellsXml.push(`<c r="${ref}"${sAttr} t="b"><v>${val ? 1 : 0}</v></c>`);
        else if (isErr(val)) cellsXml.push(`<c r="${ref}"${sAttr} t="e"><v>${xmlEsc(val.code)}</v></c>`);
        else cellsXml.push(`<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(val ?? '')}</t></is></c>`);
      }
      const h = sheet.rowH[rid], hidden = sheet.hr[rid];
      if (cellsXml.length || h || hidden) rowsXml.push(`<row r="${r + 1}"${h ? ` ht="${(h * 0.75).toFixed(2)}" customHeight="1"` : ''}${hidden ? ' hidden="1"' : ''}>${cellsXml.join('')}</row>`);
    }
    const colsXml = sheet.cols.map((cid, c) => {
      const w = sheet.colW[cid], hidden = sheet.hc[cid];
      return w || hidden ? `<col min="${c + 1}" max="${c + 1}" width="${((w || 100) - 5) / 7}" customWidth="1"${hidden ? ' hidden="1"' : ''}/>` : '';
    }).join('');
    const merges = sheet.merges.map((m) => objRect(m, ix)).filter(Boolean).map((m) => `<mergeCell ref="${colName(m.c1)}${m.r1 + 1}:${colName(m.c2)}${m.r2 + 1}"/>`);
    const fr = sheet.freeze || { r: 0, c: 0 };
    const pane = fr.r || fr.c ? `<pane${fr.c ? ` xSplit="${fr.c}"` : ''}${fr.r ? ` ySplit="${fr.r}"` : ''} topLeftCell="${colName(fr.c)}${fr.r + 1}" activePane="${fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : '';
    const dvs = (sheet.dv || []).map((d) => {
      const rect = objRect(d, ix);
      if (!rect) return '';
      const sq = `${colName(rect.c1)}${rect.r1 + 1}:${colName(rect.c2)}${rect.r2 + 1}`;
      if (d.type === 'list') {
        const f = String(d.source || '')[0] === '=' ? xmlEsc(d.source.slice(1)) : '"' + xmlEsc(String(d.source || '').split(/[,;]/).map((x) => x.trim()).join(',')) + '"';
        return `<dataValidation type="list" allowBlank="1" showErrorMessage="${d.strict ? 1 : 0}" sqref="${sq}"><formula1>${f}</formula1></dataValidation>`;
      }
      return '';
    }).filter(Boolean);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:${colName(maxC)}${maxR + 1}"/><sheetViews><sheetView workbookViewId="0"${sheet.nogrid ? ' showGridLines="0"' : ''}>${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/>${colsXml ? `<cols>${colsXml}</cols>` : ''}<sheetData>${rowsXml.join('')}</sheetData>${scenariosXml(sheet)}${merges.length ? `<mergeCells count="${merges.length}">${merges.join('')}</mergeCells>` : ''}${dvs.length ? `<dataValidations count="${dvs.length}">${dvs.join('')}</dataValidations>` : ''}</worksheet>`;
  });

  const names = Object.entries(doc.names || {}).map(([n, ref]) => `<definedName name="${xmlEsc(n)}">${xmlEsc(exportFormula(engine, String(ref).replace(/^=/, ''), doc.sheets[0], 0, 0))}</definedName>`)
    .concat(doc.sheets.flatMap((sheet, si) => solverNamesXml(engine, sheet, si)));
  const files = [
    { name: '[Content_Types].xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${doc.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>` },
    { name: '_rels/.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${doc.sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${names.length ? `<definedNames>${names.join('')}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${doc.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${doc.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` }
  ];
  sheetXml.forEach((x, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, text: x }));
  files.push({ name: 'xl/styles.xml', text: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` });
  return zip(files);
}

/* ================================================================ CSV */

export function readCsv(text, locale, name = 'Sheet1') {
  const t = text.replace(/^﻿/, '');
  const firstLine = t.split(/\r?\n/)[0] || '';
  const delim = ['\t', ';', ','].map((d) => [d, firstLine.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (q) {
      if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"' && cur === '') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  const cells = [];
  let maxR = 0, maxC = 0;
  rows.forEach((r, ri) => r.forEach((v, ci) => {
    if (v === '') return;
    const p = parseInput(v, locale);
    cells.push({ r: ri, c: ci, v: p.raw, st: p.fmt || null });
    maxR = Math.max(maxR, ri); maxC = Math.max(maxC, ci);
  }));
  return { sheets: [{ name, cells, maxR, maxC }], names: {} };
}

/** CSV of a sheet. opts.delim: ',' | ';' | '\t'; opts.dec: '.' | ',' for numbers without a number format (numbers
 * with a format are written as they are shown). Fields that contain the separator, quotes or line breaks are quoted. */
export function writeCsv(engine, sheet, locale, opts = {}) {
  const delim = opts.delim || ',', dec = opts.dec || '.';
  let maxR = -1, maxC = -1;
  const ix = { rows: new Map(sheet.rows.map((id, i) => [id, i])), cols: new Map(sheet.cols.map((id, i) => [id, i])) };
  for (const k in sheet.cells) {
    const [rid, cid] = k.split(':');
    if (sheet.cells[k].v === undefined) continue;
    maxR = Math.max(maxR, ix.rows.get(rid) ?? -1); maxC = Math.max(maxC, ix.cols.get(cid) ?? -1);
  }
  const lines = [];
  for (let r = 0; r <= maxR; r++) {
    const line = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = cellAt(sheet, r, c);
      const v = engine.valueAt(sheet, r, c);
      const s = typeof v === 'number' && (!cell || !cell.st || !cell.st.nf || cell.st.nf === 'gen') ? numToText(v).replace('.', dec) : formatValue(v, cell && cell.st, locale);
      line.push(s.includes(delim) || /["\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
    }
    lines.push(line.join(delim));
  }
  return new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
}

export { parseRefText };
