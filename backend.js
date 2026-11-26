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
 * Spreadsheets for YouTrack – HTTP endpoints used by the issue widget. Every endpoint is issue-scoped: YouTrack only
 * calls it for users who can read the issue, and ctx.issue is that issue.
 *
 *   POST backend/sync            { rev, sel }  changes since rev (or the whole workbook) + who else is here
 *   POST backend/ops             { bid, baseRev, strict, ops }  applies one batch of changes
 *   POST backend/create          { sheets? }   adds a workbook to the issue (sheets: imported file)
 *   POST backend/delete          deletes the workbook and its history
 *   GET  backend/history         ?offset&limit  change history, newest first
 *   POST backend/history-delete  { id } or { all: true }
 *
 * Who may edit: users who can update the issue. With the project setting "Everyone who can see the issue
 * can edit" turned on, every user who can read the issue (except the guest account).
 *
 * Stored data is treated as untrusted: every batch is cleaned (sheet-core.js cleanOps) before it is applied, and the
 * cleaned batch is what goes into the log that other browsers replay. Workbooks saved by an older version are cleaned
 * when they are read and saved cleaned with the next change; log entries and history entries are cleaned when they
 * are sent.
 */
const entities = require('@jetbrains/youtrack-scripting-api/entities');
const core = require('./sheet-core');

const LOG_LIMIT = 300;            // batches kept for clients that are behind
const LOG_MAX_CHARS = 1500000;
const HISTORY_LIMIT = 500;        // history entries
const HISTORY_MAX_CHARS = 2000000;
const HISTORY_CELLS = 60;         // cells recorded per history entry
const HISTORY_EVENTS = 20;        // events (row, sheet and object changes) recorded per history entry
const HISTORY_EVENT_CHARS = 200;
const HISTORY_VALUE_CHARS = 2000; // longer values are recorded truncated (and can't be restored)
const HISTORY_MERGE_MS = 90000;   // edits by the same user within this window form one entry
const PRESENCE_TTL_MS = 30000;
const PRESENCE_WRITE_MS = 10000;
const NAME_CHARS = 200;           // display names and logins as stored and sent

class HttpError extends Error {
  constructor(code, message, extra) { super(message); this.code = code; this.extra = extra || null; }
}

function endpoint(method, path, fn) {
  return {
    scope: 'issue',
    method: method,
    path: path,
    handle: function (ctx) {
      try {
        const result = fn(ctx);
        ctx.response.json(result === undefined ? { ok: true } : result);
      } catch (e) {
        const code = e instanceof HttpError ? e.code : (e && e.code === 400 ? 400 : 500);
        ctx.response.code = code;
        ctx.response.json(Object.assign({ ok: false, status: code, error: String((e && e.message) || e) }, e && e.extra ? e.extra : {}));
      }
    }
  };
}

function body(ctx) {
  try { return ctx.request.json() || {}; } catch (e) { throw new HttpError(400, 'Invalid JSON body'); }
}

function parse(raw, fallback) {
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}

/** Text of a sent or stored value, at most max characters. Objects and arrays give '' (they are never converted: an
 * object can carry its own toString, and converting it may fail). */
function text(v, max) { return typeof v === 'string' ? v.slice(0, max) : typeof v === 'number' || typeof v === 'boolean' ? String(v).slice(0, max) : ''; }

function me(ctx) {
  const u = ctx.currentUser;
  return { login: text(u.login, NAME_CHARS), name: text(u.visibleName || u.fullName || u.login, NAME_CHARS) };
}

/** The guest account stands for every anonymous visitor (one shared user); it never edits. */
function isGuest(user) { return !user || user.login === 'guest'; }

/** True when the user is a member of one of the groups of an app setting (a set of UserGroup entities). */
function inGroups(user, groups) {
  if (!groups) return false;
  let hit = false;
  const visit = function (g) {
    if (hit || !g) return;
    try { if (user.isInGroup(g.name)) hit = true; } catch (e) { /* group not readable: ignore it */ }
  };
  try { if (typeof groups.forEach === 'function') groups.forEach(visit); } catch (e) { /* not a collection */ }
  return hit;
}

/** Whether the current user can see the issue (issue-level visibility restrictions count). */
function canView(ctx) {
  try { return ctx.issue.isVisibleTo(ctx.currentUser) !== false; } catch (e) { return true; }
}

/**
 * Whether the current user may create and edit the issue's spreadsheet. Checked on every request, so changes to
 * roles, groups and settings take effect without reloading. Role names are never looked at.
 *   1. The user must be able to see the issue (Issue.isVisibleTo: visibility restrictions on the issue count),
 *      and must not be the guest account.
 *   2. Members of the project's "Groups that can only view spreadsheets" never edit.
 *   3. Members of "Groups that can edit spreadsheets", or everyone when "Everyone who can see the issue can edit"
 *      is on, can edit.
 *   4. Otherwise YouTrack's own Update Issue permission in the issue's project decides, from whichever role grants it
 *      (global, project or custom roles; User.hasPermission, since YouTrack 2025.3). Users with only Read Issue
 *      can view the spreadsheet but not change it.
 * (Issue.canBeWrittenBy is a field-level check and does not reflect Update Issue, so it is not used.)
 */
function canEdit(ctx) {
  const s = ctx.settings || {};
  const user = ctx.currentUser, issue = ctx.issue;
  if (isGuest(user)) return false;
  if (!canView(ctx)) return false;
  if (inGroups(user, s.viewerGroups)) return false;
  if (s.readersCanEdit === true || inGroups(user, s.editorGroups)) return true;
  try { return user.hasPermission('UPDATE_ISSUE', issue.project) === true; } catch (e) {
    console.log('Spreadsheets for YouTrack: could not check the Update Issue permission: ' + e);
  }
  return false;
}

function requireView(ctx) {
  if (!canView(ctx)) throw new HttpError(403, 'You cannot see this issue.');
}

function requireEdit(ctx) {
  if (!canEdit(ctx)) throw new HttpError(403, 'You can view this spreadsheet but not edit it.');
}

// ---------------------------------------------------------------- storage (Issue extension properties)

/** The stored workbook as it is (null when there is none). */
function readDoc(issue) {
  const doc = parse(issue.extensionProperties.sheetDoc, null);
  return doc && typeof doc === 'object' && Array.isArray(doc.sheets) ? doc : null;
}
/** True when this version cleaned and saved the stored workbook. It only decides whether a cleaned copy is saved back:
 * what is stored is never trusted, so it is cleaned every time it is read. */
function isClean(doc) { return doc.sv === core.CLEAN_VERSION; }
/** The revision of a stored workbook as cleanDoc reads it (a number, 0 when it is missing or broken). */
function revOf(doc) { return typeof doc.rev === 'number' && isFinite(doc.rev) && doc.rev >= 0 ? doc.rev : 0; }
/** The stored workbook, cleaned. */
function getDoc(issue) {
  const raw = readDoc(issue);
  return raw ? core.cleanDoc(raw) : null;
}
function getLog(issue) { const l = parse(issue.extensionProperties.sheetLog, []); return Array.isArray(l) ? l : []; }
function getHistory(issue) { const h = parse(issue.extensionProperties.sheetHistory, []); return Array.isArray(h) ? h : []; }

function saveDoc(issue, doc) {
  const raw = JSON.stringify(doc);
  if (raw.length > core.LIMITS.docChars) {
    throw new HttpError(413, 'The spreadsheet is too large. Delete some rows or sheets and try again.');
  }
  issue.extensionProperties.sheetDoc = raw;
}

function saveLog(issue, log) {
  let list = log.slice(-LOG_LIMIT);
  let raw = JSON.stringify(list);
  while (raw.length > LOG_MAX_CHARS && list.length > 1) {
    list = list.slice(Math.ceil(list.length / 4));
    raw = JSON.stringify(list);
  }
  issue.extensionProperties.sheetLog = raw;
}

function saveHistory(issue, history) {
  let list = history.slice(0, HISTORY_LIMIT);
  let raw = JSON.stringify(list);
  while (raw.length > HISTORY_MAX_CHARS && list.length > 1) {
    list = list.slice(0, Math.floor(list.length * 3 / 4));
    raw = JSON.stringify(list);
  }
  issue.extensionProperties.sheetHistory = raw.length > HISTORY_MAX_CHARS ? '[]' : raw;
}

function meta(doc) {
  if (!doc) return null;
  return {
    sheets: doc.sheets.map(function (s) { return text(s && s.name, 31); }),
    cells: doc.sheets.reduce(function (n, s) { return n + (s && s.cells && typeof s.cells === 'object' ? Object.keys(s.cells).length : 0); }, 0),
    createdAt: text(doc.createdAt, 40), createdBy: text(doc.createdBy, NAME_CHARS),
    updatedAt: text(doc.updatedAt, 40), updatedBy: text(doc.updatedBy, NAME_CHARS), updatedByName: text(doc.updatedByName, NAME_CHARS)
  };
}

/** A logged batch as it is sent to browsers: its operations cleaned (entries written by older versions were not). */
function cleanBatch(b) {
  return { rev: b.rev, bid: text(b.bid, 40), by: text(b.by, NAME_CHARS), ops: core.cleanOps(b.ops, true) };
}

/** Batches after `rev` up to the workbook's revision `docRev`, or null when the log no longer reaches back that far
 * (send the whole workbook). */
function batchesSince(issue, docRev, rev) {
  if (typeof rev !== 'number' || rev < 0 || rev > docRev) return null;
  if (rev === docRev) return [];
  const log = getLog(issue).filter(function (b) { return b && typeof b === 'object' && typeof b.rev === 'number'; });
  if (!log.length || log[0].rev > rev + 1) return null;
  const list = log.filter(function (b) { return b.rev > rev; });
  // the batches must lead to the workbook's revision one by one; with a gap (a batch that was not logged) the browser
  // gets the whole workbook instead, so it can never stay behind without noticing
  if (list.length !== docRev - rev) return null;
  for (let i = 0; i < list.length; i++) if (list[i].rev !== rev + 1 + i) return null;
  return list.map(cleanBatch);
}

// ---------------------------------------------------------------- presence (User extension properties)

function presence(ctx, sel) {
  const issueId = ctx.issue.id;
  const now = Date.now();
  const u = ctx.currentUser;
  // the guest account is shared by every anonymous visitor: it is never shown and never written to
  if (sel !== undefined && !isGuest(u)) {
    const mine = sel && typeof sel === 'object' ? {
      issue: issueId,
      s: text(sel.s, 16),
      k: text(sel.k, 40),
      k2: text(sel.k2, 40),
      editing: !!sel.editing
    } : null;
    const old = parse(u.extensionProperties.sheetPresence, null);
    const changed = !mine || !old || old.issue !== mine.issue || old.s !== mine.s || old.k !== mine.k ||
      old.k2 !== mine.k2 || old.editing !== mine.editing;
    if (!mine) {
      if (u.extensionProperties.sheetPresenceIssue) {
        u.extensionProperties.sheetPresenceIssue = null;
        u.extensionProperties.sheetPresence = null;
      }
    } else if (changed || now - (old.at || 0) > PRESENCE_WRITE_MS) {
      mine.at = now;
      u.extensionProperties.sheetPresence = JSON.stringify(mine);
      if (u.extensionProperties.sheetPresenceIssue !== issueId) u.extensionProperties.sheetPresenceIssue = issueId;
    }
  }
  const out = [];
  // anonymous visitors do not see who is working on the issue
  if (isGuest(u)) return out;
  const set = entities.User.findByExtensionProperties({ sheetPresenceIssue: issueId });
  if (set) {
    set.forEach(function (other) {
      if (other.login === u.login || isGuest(other)) return;
      const p = parse(other.extensionProperties.sheetPresence, null);
      if (!p || typeof p !== 'object' || p.issue !== issueId || now - (p.at || 0) > PRESENCE_TTL_MS) return;
      out.push({
        login: text(other.login, NAME_CHARS), name: text(other.visibleName || other.login, NAME_CHARS),
        s: text(p.s, 16), k: text(p.k, 40), k2: text(p.k2, 40), editing: !!p.editing
      });
    });
  }
  return out;
}

// ---------------------------------------------------------------- history

function clip(cell) {
  if (!cell) return { c: null, trunc: false };
  const out = {};
  let trunc = false;
  if (cell.v !== undefined) {
    out.v = cell.v.length > HISTORY_VALUE_CHARS ? cell.v.slice(0, HISTORY_VALUE_CHARS) : cell.v;
    trunc = cell.v.length > HISTORY_VALUE_CHARS;
  }
  if (cell.st) out.st = cell.st;
  return { c: out, trunc: trunc };
}

function recordHistory(issue, doc, changes, user, rev) {
  if (!changes.length) return;
  const history = getHistory(issue);
  const now = Date.now();
  const cellChanges = changes.filter(function (c) { return c.kind === 'cell'; });
  // blank rows or columns added after the last one (the sheet growing as people move past its end) are not listed
  const other = changes.filter(function (c) { return c.kind !== 'cell' && !((c.kind === 'rows' || c.kind === 'cols') && c.added && c.end); });

  let events = other.map(function (c) {
    const sheet = core.findSheet(doc, c.s);
    const sn = sheet ? sheet.name : (c.name || '');
    if (c.kind === 'rows' || c.kind === 'cols') {
      const what = c.kind === 'rows' ? (c.added || c.deleted) + ((c.added || c.deleted) === 1 ? ' row' : ' rows')
        : (c.added || c.deleted) + ((c.added || c.deleted) === 1 ? ' column' : ' columns');
      const where = c.kind === 'rows' ? 'at row ' + (c.at + 1) : 'at column ' + core.colName(c.at);
      return (c.added ? 'Inserted ' : 'Deleted ') + what + ' ' + where + ' in ' + sn;
    }
    if (c.kind === 'addSheet') return 'Added sheet ' + c.name;
    if (c.kind === 'delSheet') return 'Deleted sheet ' + c.name;
    if (c.kind === 'moveSheet') return 'Moved sheet ' + c.name;
    if (c.kind === 'sheet' && c.renamed) return 'Renamed sheet ' + c.renamed.from + ' to ' + c.renamed.to;
    if (c.kind === 'obj') {
      const what = { tables: 'table', cf: 'conditional formatting rule', dv: 'data validation rule', charts: 'chart', solver: 'Solver model', scenarios: 'scenario' }[c.objKind] || 'item';
      const verb = c.action === 'add' ? 'Added ' : c.action === 'delete' ? 'Deleted ' : 'Changed ';
      return verb + what + (c.name ? ' ' + c.name : '') + ' in ' + sn;
    }
    if (c.kind === 'name') return (c.action === 'add' ? 'Defined name ' : c.action === 'delete' ? 'Deleted name ' : 'Changed name ') + c.name;
    return null;
  }).filter(Boolean).map(function (e) { return e.length > HISTORY_EVENT_CHARS ? e.slice(0, HISTORY_EVENT_CHARS) + '…' : e; });
  if (events.length > HISTORY_EVENTS) events = events.slice(0, HISTORY_EVENTS).concat(['… and ' + (events.length - HISTORY_EVENTS) + ' more changes']);

  // Rows or columns moved: the cell changes in the same batch are formula rewrites (or an undo restoring them);
  // the history lists the structural change only.
  const structural = other.some(function (c) { return c.kind === 'rows' || c.kind === 'cols' || c.kind === 'delSheet' || c.kind === 'addSheet'; });
  const changed = structural ? [] : cellChanges;
  // a batch can change many cells, but an entry keeps only a few: addresses are made for those only, with the row and
  // column positions of each sheet looked up once
  const sheets = new Map(), positions = new Map();
  doc.sheets.forEach(function (sh) { sheets.set(sh.id, sh); });
  const address = function (sheet, k) {
    let p = positions.get(sheet.id);
    if (!p) {
      p = { rows: new Map(), cols: new Map() };
      sheet.rows.forEach(function (id, i) { p.rows.set(id, i); });
      sheet.cols.forEach(function (id, i) { p.cols.set(id, i); });
      positions.set(sheet.id, p);
    }
    const parts = k.split(':');
    const r = p.rows.get(parts[0]), c = p.cols.get(parts[1]);
    return r === undefined || c === undefined ? '?' : core.colName(c) + (r + 1);
  };
  const entry = function (c) {
    const sheet = sheets.get(c.s);
    const from = clip(c.from), to = clip(c.to);
    return {
      s: c.s, k: c.k, sn: sheet ? sheet.name : '', a: sheet ? address(sheet, c.k) : '?',
      from: from.c, to: to.c, trunc: from.trunc || to.trunc
    };
  };

  // Merge into the user's previous entry when it is recent and only edited cells.
  const last = history[0];
  if (!events.length && !changed.length) return;
  if (!events.length && last && last.by === user.login && Array.isArray(last.events) && last.events.length === 0 && !last.structural &&
      Array.isArray(last.cells) && now - Date.parse(last.at) < HISTORY_MERGE_MS && last.total + changed.length <= 2000) {
    const listed = new Map();
    last.cells.forEach(function (x) { if (x && !listed.has(x.s + '|' + x.k)) listed.set(x.s + '|' + x.k, x); });
    changed.forEach(function (ch) {
      const prev = listed.get(ch.s + '|' + ch.k);
      if (prev) { const c = entry(ch); prev.to = c.to; prev.a = c.a; prev.trunc = prev.trunc || c.trunc; return; }
      last.total++;
      if (last.cells.length < HISTORY_CELLS) { const c = entry(ch); last.cells.push(c); listed.set(c.s + '|' + c.k, c); } else last.more = true;
    });
    last.at = new Date(now).toISOString();
    last.rev = rev;
    saveHistory(issue, history);
    return;
  }

  history.unshift({
    id: 'h' + now.toString(36) + Math.random().toString(36).slice(2, 6),
    rev: rev, at: new Date(now).toISOString(), by: user.login, byName: user.name,
    events: events,
    structural: structural,
    cells: changed.slice(0, HISTORY_CELLS).map(entry), total: changed.length, more: changed.length > HISTORY_CELLS
  });
  saveHistory(issue, history);
}

/** A history entry as it is sent to browsers (entries written by older versions were not limited or cleaned). */
function cleanEntry(e) {
  if (!e || typeof e !== 'object') return null;
  const cell = function (x) {
    if (!x || typeof x !== 'object') return null;
    const out = {};
    if (x.v !== undefined && x.v !== null) out.v = text(x.v, HISTORY_VALUE_CHARS);
    let st = null;
    try { st = core.cleanStyle(x.st); } catch (err) { st = null; }
    if (st) out.st = st;
    return out;
  };
  return {
    id: text(e.id, 40), rev: typeof e.rev === 'number' ? e.rev : 0, at: text(e.at, 40),
    by: text(e.by, NAME_CHARS), byName: text(e.byName || e.by, NAME_CHARS),
    events: (Array.isArray(e.events) ? e.events : []).slice(0, HISTORY_EVENTS + 1).map(function (x) { return text(x, HISTORY_EVENT_CHARS + 1); }),
    structural: !!e.structural,
    cells: (Array.isArray(e.cells) ? e.cells : []).slice(0, HISTORY_CELLS).map(function (c) {
      if (!c || typeof c !== 'object') return null;
      return { s: text(c.s, 16), k: text(c.k, 40), sn: text(c.sn, 31), a: text(c.a, 12), from: cell(c.from), to: cell(c.to), trunc: !!c.trunc };
    }).filter(Boolean),
    total: typeof e.total === 'number' ? e.total : 0, more: !!e.more
  };
}

// ---------------------------------------------------------------- endpoints

function sync(ctx) {
  requireView(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  const raw = readDoc(issue);
  const editable = canEdit(ctx);
  const result = {
    exists: !!raw, canEdit: editable, me: me(ctx),
    presence: presence(ctx, b.sel === undefined ? undefined : b.sel)
  };
  if (!raw) return result;
  const rev = revOf(raw);
  result.rev = rev;
  result.meta = meta(raw);
  if (b.metaOnly) return result;
  const batches = batchesSince(issue, rev, typeof b.rev === 'number' ? b.rev : -1);
  if (batches === null) {
    // the workbook is always sent cleaned; one that an older version saved is also stored cleaned when this user may
    // change it
    result.doc = core.cleanDoc(raw);
    if (!isClean(raw) && editable) { try { saveDoc(issue, result.doc); } catch (e) { /* stored as it was; cleaned again next time */ } }
  } else result.batches = batches;
  return result;
}

function ops(ctx) {
  requireEdit(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  const doc = getDoc(issue);
  if (!doc) throw new HttpError(404, 'This issue has no spreadsheet. It may have been deleted.', { gone: true });
  if (!Array.isArray(b.ops) || !b.ops.length) throw new HttpError(400, 'No changes');
  if (JSON.stringify(b.ops).length > core.LIMITS.batchChars) throw new HttpError(413, 'This change is too large. Change fewer cells at once.');
  const bid = text(b.bid, 40);
  const baseRev = typeof b.baseRev === 'number' ? b.baseRev : -1;
  // a browser that still shows a workbook which was deleted and created again
  if (baseRev >= 0 && typeof doc.base === 'number' && baseRev < doc.base) {
    throw new HttpError(404, 'This spreadsheet was deleted and created again.', { gone: true });
  }

  const log = getLog(issue);
  // Retried request that already went through: answer as if it was applied now.
  if (bid && log.some(function (x) { return x && x.bid === bid; })) {
    const since = batchesSince(issue, doc.rev, baseRev);
    return since === null ? { ok: true, rev: doc.rev, doc: doc } : { ok: true, rev: doc.rev, batches: since };
  }
  // only known operations with known fields are applied and logged
  const clean = core.cleanOps(b.ops);
  if (!clean.length) throw new HttpError(400, 'No changes');
  if ((b.strict || core.isStrict(clean)) && baseRev !== doc.rev) {
    const since = batchesSince(issue, doc.rev, baseRev);
    throw new HttpError(409, 'The sheet changed in the meantime.', since === null
      ? { conflict: true, rev: doc.rev, doc: doc } : { conflict: true, rev: doc.rev, batches: since });
  }

  const user = me(ctx);
  const changes = core.applyOps(doc, clean);
  // every accepted batch gets its own revision and log entry, also when it changed nothing (a value that someone else
  // had just set, a row that was deleted meanwhile): the browser that sent it takes "the next revision" as its own
  // batch, and replays are recognized by the batch id in the log
  doc.rev += 1;
  doc.updatedAt = new Date().toISOString();
  doc.updatedBy = user.login;
  doc.updatedByName = user.name;
  saveDoc(issue, doc);
  log.push({ rev: doc.rev, bid: bid, by: user.login, ops: clean });
  saveLog(issue, log);
  recordHistory(issue, doc, changes, user, doc.rev);
  return { ok: true, rev: doc.rev };
}

function create(ctx) {
  requireEdit(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  const existing = getDoc(issue);
  if (existing) return { ok: true, existed: true, rev: existing.rev, doc: existing };
  const user = me(ctx);
  let sheets = null;
  if (Array.isArray(b.sheets) && b.sheets.length) {
    if (b.sheets.length > core.LIMITS.sheets) throw new HttpError(400, 'Too many sheets');
    if (JSON.stringify(b.sheets).length > core.LIMITS.docChars) throw new HttpError(413, 'The file is too large.');
    sheets = b.sheets.map(core.cleanSheet);
  }
  // cleanDoc also gives duplicate sheet and table names a number
  const doc = core.cleanDoc(core.newWorkbook(user, sheets));
  // revisions never repeat across workbooks: a browser that still shows a deleted workbook gets the new one
  doc.rev = doc.base = Date.now();
  saveDoc(issue, doc);
  saveLog(issue, []);
  saveHistory(issue, [{
    id: 'h' + Date.now().toString(36), rev: doc.rev, at: doc.createdAt, by: user.login, byName: user.name,
    events: [sheets ? 'Imported a spreadsheet' : 'Created the spreadsheet'], structural: true, cells: [], total: 0
  }]);
  return { ok: true, rev: doc.rev, doc: doc };
}

function remove(ctx) {
  requireEdit(ctx);
  const issue = ctx.issue;
  issue.extensionProperties.sheetDoc = null;
  issue.extensionProperties.sheetLog = null;
  issue.extensionProperties.sheetHistory = null;
  return { ok: true };
}

function history(ctx) {
  requireView(ctx);
  const all = getHistory(ctx.issue);
  const offset = Math.max(0, parseInt(ctx.request.getParameter('offset') || '0', 10) || 0);
  const limit = Math.min(100, Math.max(1, parseInt(ctx.request.getParameter('limit') || '30', 10) || 30));
  return { total: all.length, entries: all.slice(offset, offset + limit).map(cleanEntry).filter(Boolean), canEdit: canEdit(ctx) };
}

function historyDelete(ctx) {
  requireEdit(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  if (b.all) { saveHistory(issue, []); return { ok: true }; }
  const list = getHistory(issue);
  const next = list.filter(function (h) { return !h || h.id !== b.id; });
  if (next.length === list.length) throw new HttpError(404, 'History entry not found');
  saveHistory(issue, next);
  return { ok: true };
}

exports.httpHandler = {
  endpoints: [
    endpoint('POST', 'sync', sync),
    endpoint('POST', 'ops', ops),
    endpoint('POST', 'create', create),
    endpoint('POST', 'delete', remove),
    endpoint('GET', 'history', history),
    endpoint('POST', 'history-delete', historyDelete)
  ]
};
