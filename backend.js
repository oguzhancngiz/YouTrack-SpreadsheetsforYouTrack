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
 * can edit" turned on, every user who can read the issue.
 */
const entities = require('@jetbrains/youtrack-scripting-api/entities');
const core = require('./sheet-core');

const LOG_LIMIT = 300;            // batches kept for clients that are behind
const LOG_MAX_CHARS = 1500000;
const HISTORY_LIMIT = 500;        // history entries
const HISTORY_CELLS = 60;         // cells recorded per history entry
const HISTORY_VALUE_CHARS = 2000; // longer values are recorded truncated (and can't be restored)
const HISTORY_MERGE_MS = 90000;   // edits by the same user within this window form one entry
const PRESENCE_TTL_MS = 30000;
const PRESENCE_WRITE_MS = 10000;

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

function me(ctx) {
  const u = ctx.currentUser;
  return { login: u.login, name: u.visibleName || u.fullName || u.login };
}

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
 *   1. The user must be able to see the issue (Issue.isVisibleTo: visibility restrictions on the issue count).
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

function getDoc(issue) { return core.upgrade(parse(issue.extensionProperties.sheetDoc, null)); }
function getLog(issue) { return parse(issue.extensionProperties.sheetLog, []); }
function getHistory(issue) { return parse(issue.extensionProperties.sheetHistory, []); }

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
  issue.extensionProperties.sheetHistory = JSON.stringify(history.slice(0, HISTORY_LIMIT));
}

function meta(doc) {
  if (!doc) return null;
  return {
    sheets: doc.sheets.map(function (s) { return s.name; }),
    cells: doc.sheets.reduce(function (n, s) { return n + Object.keys(s.cells).length; }, 0),
    createdAt: doc.createdAt, createdBy: doc.createdBy,
    updatedAt: doc.updatedAt, updatedBy: doc.updatedBy, updatedByName: doc.updatedByName
  };
}

/** Batches after `rev`, or null when the log no longer reaches back that far (send the whole workbook). */
function batchesSince(issue, doc, rev) {
  if (typeof rev !== 'number' || rev < 0 || rev > doc.rev) return null;
  if (rev === doc.rev) return [];
  const log = getLog(issue);
  if (!log.length || log[0].rev > rev + 1) return null;
  return log.filter(function (b) { return b.rev > rev; });
}

// ---------------------------------------------------------------- presence (User extension properties)

function presence(ctx, sel) {
  const issueId = ctx.issue.id;
  const now = Date.now();
  const u = ctx.currentUser;
  if (sel !== undefined) {
    const mine = sel && typeof sel === 'object' ? {
      issue: issueId,
      s: String(sel.s || '').slice(0, 16),
      k: String(sel.k || '').slice(0, 40),
      k2: String(sel.k2 || '').slice(0, 40),
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
  const set = entities.User.findByExtensionProperties({ sheetPresenceIssue: issueId });
  if (set) {
    set.forEach(function (other) {
      if (other.login === u.login) return;
      const p = parse(other.extensionProperties.sheetPresence, null);
      if (!p || p.issue !== issueId || now - (p.at || 0) > PRESENCE_TTL_MS) return;
      out.push({ login: other.login, name: other.visibleName || other.login, s: p.s, k: p.k, k2: p.k2, editing: !!p.editing });
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

  const events = other.map(function (c) {
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
  }).filter(Boolean);

  // Rows or columns moved: the cell changes in the same batch are formula rewrites (or an undo restoring them);
  // the history lists the structural change only.
  const structural = other.some(function (c) { return c.kind === 'rows' || c.kind === 'cols' || c.kind === 'delSheet' || c.kind === 'addSheet'; });
  const cells = (structural ? [] : cellChanges).map(function (c) {
    const sheet = core.findSheet(doc, c.s);
    const from = clip(c.from), to = clip(c.to);
    return {
      s: c.s, k: c.k, sn: sheet ? sheet.name : '', a: sheet ? core.cellLabel(sheet, c.k) : '?',
      from: from.c, to: to.c, trunc: from.trunc || to.trunc
    };
  });

  // Merge into the user's previous entry when it is recent and only edited cells.
  const last = history[0];
  if (!events.length && !cells.length) return;
  if (!events.length && last && last.by === user.login && last.events.length === 0 && !last.structural &&
      now - Date.parse(last.at) < HISTORY_MERGE_MS && last.total + cells.length <= 2000) {
    cells.forEach(function (c) {
      const prev = last.cells.filter(function (x) { return x.s === c.s && x.k === c.k; })[0];
      if (prev) { prev.to = c.to; prev.a = c.a; prev.trunc = prev.trunc || c.trunc; return; }
      last.total++;
      if (last.cells.length < HISTORY_CELLS) last.cells.push(c); else last.more = true;
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
    cells: cells.slice(0, HISTORY_CELLS), total: cells.length, more: cells.length > HISTORY_CELLS
  });
  saveHistory(issue, history);
}

// ---------------------------------------------------------------- endpoints

function sync(ctx) {
  requireView(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  const doc = getDoc(issue);
  const result = {
    exists: !!doc, canEdit: canEdit(ctx), me: me(ctx),
    presence: presence(ctx, b.sel === undefined ? undefined : b.sel)
  };
  if (!doc) return result;
  result.rev = doc.rev;
  result.meta = meta(doc);
  if (b.metaOnly) return result;
  const batches = batchesSince(issue, doc, typeof b.rev === 'number' ? b.rev : -1);
  if (batches === null) result.doc = doc; else result.batches = batches;
  return result;
}

function ops(ctx) {
  requireEdit(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  const doc = getDoc(issue);
  if (!doc) throw new HttpError(404, 'This issue has no spreadsheet. It may have been deleted.', { gone: true });
  if (!Array.isArray(b.ops) || !b.ops.length) throw new HttpError(400, 'No changes');
  const bid = String(b.bid || '').slice(0, 40);
  const baseRev = typeof b.baseRev === 'number' ? b.baseRev : -1;

  const log = getLog(issue);
  // Retried request that already went through: answer as if it was applied now.
  if (bid && log.some(function (x) { return x.bid === bid; })) {
    const since = batchesSince(issue, doc, baseRev);
    return since === null ? { ok: true, rev: doc.rev, doc: doc } : { ok: true, rev: doc.rev, batches: since };
  }
  if ((b.strict || core.isStrict(b.ops)) && baseRev !== doc.rev) {
    const since = batchesSince(issue, doc, baseRev);
    throw new HttpError(409, 'The sheet changed in the meantime.', since === null
      ? { conflict: true, rev: doc.rev, doc: doc } : { conflict: true, rev: doc.rev, batches: since });
  }

  const user = me(ctx);
  const changes = core.applyOps(doc, b.ops);
  doc.rev += 1;
  doc.updatedAt = new Date().toISOString();
  doc.updatedBy = user.login;
  doc.updatedByName = user.name;
  saveDoc(issue, doc);
  log.push({ rev: doc.rev, bid: bid, by: user.login, ops: b.ops });
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
    sheets = b.sheets.map(core.cleanSheet);
  }
  const doc = core.newWorkbook(user, sheets);
  doc.rev = 1;
  saveDoc(issue, doc);
  saveLog(issue, []);
  saveHistory(issue, [{
    id: 'h' + Date.now().toString(36), rev: 1, at: doc.createdAt, by: user.login, byName: user.name,
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
  return { total: all.length, entries: all.slice(offset, offset + limit), canEdit: canEdit(ctx) };
}

function historyDelete(ctx) {
  requireEdit(ctx);
  const b = body(ctx);
  const issue = ctx.issue;
  if (b.all) { saveHistory(issue, []); return { ok: true }; }
  const list = getHistory(issue);
  const next = list.filter(function (h) { return h.id !== b.id; });
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
