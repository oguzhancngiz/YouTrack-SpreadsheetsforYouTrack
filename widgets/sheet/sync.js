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
/* Spreadsheets for YouTrack – keeps the local workbook in step with the server.
 *
 * confirmed  workbook as the server has it (revision `rev`)
 * pending    batches made here that the server hasn't confirmed yet
 * view       confirmed + pending, which is what the user sees and edits
 *
 * Batches are sent one at a time. Cell edits merge with other people's edits (the last write to a cell wins).
 * Batches that move rows or columns are "strict": the server takes them only on the latest revision, so when
 * someone else changed the sheet first, the batch is rebuilt on the new state (regen) and sent again.
 *
 * Everything that comes from the server is cleaned before it is used (Core.cleanDoc for workbooks, Core.cleanOps for
 * batches): it was written by other people, possibly with an older version of the app.
 */
const Core = window.SheetCore;

const clone = (o) => (typeof structuredClone === 'function' ? structuredClone(o) : JSON.parse(JSON.stringify(o)));

function bidNew() { return 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

/** Inverse of one operation against the current state (call before applying it). Returns a list of ops. */
export function inverseOf(doc, op) {
  const sheet = op.s ? Core.findSheet(doc, op.s) : null;
  switch (op.t) {
    case 'set': {
      if (!sheet) return [];
      const c = sheet.cells[op.k] || {};
      const inv = { t: 'set', s: op.s, k: op.k };
      if (op.v !== undefined) inv.v = c.v === undefined ? null : c.v;
      if (op.st !== undefined) inv.st = c.st || null;
      if (op.n !== undefined) inv.n = c.n || null;
      return [inv];
    }
    case 'sheet': {
      if (!sheet) return [];
      const inv = { t: 'sheet', s: op.s };
      ['name', 'color', 'nogrid'].forEach((k) => { if (op[k] !== undefined) inv[k] = sheet[k] ?? null; });
      if (op.freeze) inv.freeze = { ...sheet.freeze };
      if (op.merges) inv.merges = clone(sheet.merges);
      ['colW', 'rowH', 'hr', 'hc'].forEach((k) => {
        if (!op[k]) return;
        inv[k] = {};
        Object.keys(op[k]).forEach((id) => { inv[k][id] = sheet[k][id] ?? null; });
      });
      return [inv];
    }
    case 'rows':
    case 'cols': {
      if (!sheet) return [];
      const list = op.t === 'rows' ? sheet.rows : sheet.cols;
      if (op.add) return [{ t: op.t, s: op.s, del: op.add.filter((id) => !list.includes(id)) }];
      const del = new Set(op.del.filter((id) => list.includes(id)));
      if (!del.size) return [];
      const out = [];
      // re-insert in runs, lowest index first, with the same ids
      let run = null;
      list.forEach((id, i) => {
        if (del.has(id)) {
          if (run && run.end === i - 1) { run.ids.push(id); run.end = i; }
          else { run = { at: i, end: i, ids: [id] }; out.push(run); }
        }
      });
      const ops = out.map((r) => ({ t: op.t, s: op.s, at: r.at, add: r.ids }));
      // restore the cells, sizes and visibility of the deleted rows/columns, and the ranges that shrank
      const part = op.t === 'rows' ? 0 : 1;
      for (const k in sheet.cells) {
        if (del.has(k.split(':')[part])) {
          const c = sheet.cells[k];
          ops.push({ t: 'set', s: op.s, k, v: c.v ?? null, st: c.st || null, n: c.n || null });
        }
      }
      const sizes = op.t === 'rows' ? 'rowH' : 'colW', hidden = op.t === 'rows' ? 'hr' : 'hc';
      const props = { t: 'sheet', s: op.s, [sizes]: {}, [hidden]: {}, merges: clone(sheet.merges) };
      del.forEach((id) => { if (sheet[sizes][id]) props[sizes][id] = sheet[sizes][id]; if (sheet[hidden][id]) props[hidden][id] = 1; });
      ops.push(props);
      Core.OBJ_KINDS.forEach((kind) => sheet[kind].forEach((o) => ops.push({ t: 'obj', s: op.s, kind, id: o.id, val: clone(o) })));
      return ops;
    }
    case 'obj': {
      if (!sheet) return [];
      const cur = (sheet[op.kind] || []).find((o) => o.id === op.id);
      return [{ t: 'obj', s: op.s, kind: op.kind, id: op.id, val: cur ? clone(cur) : null }];
    }
    case 'name': {
      const key = String(op.name).toUpperCase();
      return [{ t: 'name', name: key, ref: doc.names[key] ?? null }];
    }
    case 'addSheet': return [{ t: 'delSheet', s: op.sheet.id }];
    case 'delSheet': return sheet ? [{ t: 'addSheet', at: doc.sheets.indexOf(sheet), sheet: clone(sheet) }] : [];
    case 'moveSheet': return sheet ? [{ t: 'moveSheet', s: op.s, to: doc.sheets.indexOf(sheet) }] : [];
    default: return [];
  }
}

export class Store {
  /**
   * api(path, opts) calls the backend. Callbacks:
   *   onChange(reason)   the view changed ('local' | 'remote' | 'reset')
   *   onPresence(list), onStatus(state), onGone(), onReadonly(), onNotice(message)
   */
  constructor(api, callbacks) {
    this.api = api;
    this.cb = callbacks;
    this.confirmed = null;
    this.view = null;
    this.rev = -1;
    this.pending = [];
    this.inflight = null;
    this.canEdit = false;
    this.exists = false;
    this.me = null;
    this.sel = null;
    this.selDirty = false;
    this.pollTimer = null;
    this.pollMs = 2000;
    this.failures = 0;
    this.stopped = false;
  }

  get busy() { return !!this.inflight || this.pending.length > 0; }

  /** First load. Returns the sync response. */
  async load() {
    const res = await this.api('sync', { method: 'POST', body: { rev: -1, sel: this.sel } });
    this.absorb(res);
    return res;
  }

  absorb(res) {
    const could = this.canEdit;
    this.canEdit = !!res.canEdit;
    // permissions can change while the issue is open (roles, project or issue access): follow them right away
    if (this.loaded && could !== this.canEdit && this.cb.onPermission) this.cb.onPermission(this.canEdit);
    this.me = res.me || this.me;
    this.meta = res.meta || null;
    if (res.presence) this.cb.onPresence && this.cb.onPresence(res.presence);
    if (!res.exists) {
      const had = this.exists;
      this.exists = false;
      this.confirmed = this.view = null;
      this.rev = -1;
      this.pending = [];
      if (had) this.cb.onGone && this.cb.onGone();
      return;
    }
    this.exists = true;
    const doc = res.doc ? Core.cleanDoc(res.doc) : null;
    if (doc) {
      this.confirmed = doc;
      this.rev = doc.rev;
      this.rebuild();
      this.cb.onChange && this.cb.onChange('reset');
    } else if (res.batches && res.batches.length && this.confirmed) {
      this.applyRemote(res.batches);
    }
  }

  /**
   * Applies one batch from the server to the confirmed workbook. Returns false when it could not be applied: the
   * workbook here may then differ from the server's, so the whole workbook is loaded again.
   */
  applyBatch(b) {
    try {
      Core.applyOps(this.confirmed, Core.cleanOps(b.ops, true), { validate: false });
      return true;
    } catch (e) {
      console.warn('Spreadsheets for YouTrack: a change from the server could not be applied; loading the spreadsheet again', e);
      this.rev = -1;
      this.reload = true;
      return false;
    }
  }

  applyRemote(batches) {
    let changed = false;
    for (const b of batches) {
      if (!b || typeof b.rev !== 'number' || b.rev <= this.rev) continue;
      if (!this.applyBatch(b)) break;
      this.rev = b.rev;
      this.confirmed.rev = b.rev;
      const mine = this.pending.findIndex((p) => p.bid === b.bid);
      if (mine >= 0) this.pending.splice(mine, 1);
      if (this.inflight && this.inflight.bid === b.bid) this.inflight.confirmed = true;
      changed = true;
    }
    if (changed) {
      this.rebuild();
      this.cb.onChange && this.cb.onChange('remote');
    }
    if (this.reload) { this.reload = false; this.pull(); }
  }

  rebuild() {
    this.view = clone(this.confirmed);
    for (const p of this.pending) {
      // cleaned as the server will store them, so the view holds the same objects (with their ids) as the workbook
      // after the confirmation
      try { Core.applyOps(this.view, Core.cleanOps(p.ops, true), { validate: false }); } catch (e) { /* skip what no longer applies */ }
    }
  }

  /**
   * Applies a local change. Returns the inverse ops (for undo).
   *   opts.regen() → ops: rebuilds a strict batch on newer data after a conflict
   *   opts.undoEntry: object whose .ops are replaced when the batch is rebuilt
   */
  apply(ops, opts = {}) {
    if (!this.view || !ops.length) return [];
    const inverse = [];
    try {
      for (const op of ops) {
        inverse.unshift(...inverseOf(this.view, op));
        Core.applyOps(this.view, [op]);
      }
    } catch (e) {
      this.rebuild();
      throw e;
    }
    const batch = { bid: bidNew(), ops, strict: Core.isStrict(ops), regen: opts.regen || null, undoEntry: opts.undoEntry || null };
    this.pending.push(batch);
    // the bar shows "edited now by <me>" right away; the next poll brings the server's values
    this.meta = { ...(this.meta || {}), updatedAt: new Date().toISOString(), updatedByName: this.me ? this.me.name : null, updatedBy: this.me ? this.me.login : null };
    // the save is scheduled before the screen is redrawn, so an error while drawing never loses the change
    this.flushSoon();
    this.cb.onChange && this.cb.onChange('local');
    return inverse;
  }

  flushSoon() {
    clearTimeout(this.flushTimer);
    this.flushTimer = setTimeout(() => this.flush(), 30);
  }

  async flush() {
    if (this.inflight || !this.pending.length || this.stopped) return;
    // merge consecutive non-strict batches into one request
    const batch = this.pending[0];
    const sentRev = this.rev; // the revision this batch was made on (pulls may change this.rev while it is on its way)
    this.inflight = batch;
    this.cb.onStatus && this.cb.onStatus('saving');
    try {
      const res = await this.api('ops', { method: 'POST', body: { bid: batch.bid, baseRev: sentRev, strict: batch.strict, ops: batch.ops } });
      this.failures = 0;
      if (!batch.confirmed) {
        if (res.rev === this.rev + 1) {
          // only our batch happened in between (the server applied the cleaned batch, so clean it the same way)
          if (this.applyBatch(batch)) {
            this.rev = res.rev;
            this.confirmed.rev = res.rev;
          }
          const i = this.pending.indexOf(batch);
          if (i >= 0) this.pending.splice(i, 1);
          if (this.reload) { this.reload = false; this.rebuild(); this.pull(); }
        } else {
          batch.acked = true;
          await this.pull();
          const i = this.pending.indexOf(batch);
          if (i >= 0) { this.pending.splice(i, 1); this.rebuild(); this.cb.onChange && this.cb.onChange('remote'); }
        }
      }
      this.inflight = null;
      this.cb.onStatus && this.cb.onStatus(this.pending.length ? 'saving' : 'saved');
      if (this.pending.length) this.flush();
    } catch (e) {
      this.inflight = null;
      const data = e.data || {};
      if (e.status === 409 && data.conflict) {
        const doc = data.doc ? Core.cleanDoc(data.doc) : null;
        if (doc) { this.confirmed = doc; this.rev = doc.rev; }
        else if (Array.isArray(data.batches)) {
          for (const b of data.batches) {
            if (!b || typeof b.rev !== 'number' || b.rev <= this.rev) continue;
            if (!this.applyBatch(b)) break;
            this.rev = b.rev; this.confirmed.rev = b.rev;
            const mine = this.pending.findIndex((p) => p.bid === b.bid && p !== batch);
            if (mine >= 0) this.pending.splice(mine, 1);
          }
        }
        if (this.reload) { this.reload = false; this.pull(); }
        const i = this.pending.indexOf(batch);
        if (i >= 0) this.pending.splice(i, 1);
        this.rebuild();
        if (batch.regen && (batch.tries || 0) < 5) {
          let ops = null;
          try { ops = batch.regen(); } catch (x) { ops = null; }
          if (ops && ops.length) {
            const inverse = [];
            try {
              for (const op of ops) { inverse.unshift(...inverseOf(this.view, op)); Core.applyOps(this.view, [op]); }
              const again = { ...batch, bid: bidNew(), ops, strict: Core.isStrict(ops), tries: (batch.tries || 0) + 1 };
              this.pending.splice(Math.min(i < 0 ? 0 : i, this.pending.length), 0, again);
              if (batch.undoEntry) batch.undoEntry.ops = inverse;
            } catch (x) { this.rebuild(); }
          }
        } else {
          this.cb.onNotice && this.cb.onNotice('Someone else changed the sheet at the same time, so your last change was not applied. Please try again.');
          if (batch.undoEntry) batch.undoEntry.dead = true;
        }
        this.cb.onChange && this.cb.onChange('remote');
        this.flushSoon();
        return;
      }
      if (e.status === 404 && data.gone) {
        // the batch was made on a workbook that was deleted; if this browser already shows the one created after it
        // (a pull ran while the batch was on its way), only that batch is dropped
        if (this.exists && this.confirmed && typeof this.confirmed.base === 'number' && sentRev < this.confirmed.base) {
          const i = this.pending.indexOf(batch);
          if (i >= 0) this.pending.splice(i, 1);
          if (batch.undoEntry) batch.undoEntry.dead = true;
          this.rebuild();
          this.cb.onChange && this.cb.onChange('remote');
          this.cb.onStatus && this.cb.onStatus(this.pending.length ? 'saving' : 'saved');
          this.flushSoon();
          return;
        }
        this.pending = [];
        this.exists = false;
        this.confirmed = this.view = null;
        this.cb.onGone && this.cb.onGone();
        return;
      }
      if (e.status === 403) {
        this.canEdit = false;
        this.pending = [];
        this.rebuild();
        this.cb.onReadonly && this.cb.onReadonly(data.error || e.message);
        this.cb.onChange && this.cb.onChange('remote');
        return;
      }
      if (e.status === 400 || e.status === 413) {
        const i = this.pending.indexOf(batch);
        if (i >= 0) this.pending.splice(i, 1);
        this.rebuild();
        if (batch.undoEntry) batch.undoEntry.dead = true;
        this.cb.onNotice && this.cb.onNotice(data.error || e.message || 'The change was rejected.');
        this.cb.onChange && this.cb.onChange('remote');
        this.cb.onStatus && this.cb.onStatus(this.pending.length ? 'saving' : 'saved');
        this.flushSoon();
        return;
      }
      // network or server error: keep the batch and retry with the same id
      this.failures++;
      this.cb.onStatus && this.cb.onStatus('offline');
      setTimeout(() => this.flush(), Math.min(30000, 1000 * Math.pow(2, Math.min(this.failures, 5))));
    }
  }

  setSelection(sel) {
    const s = JSON.stringify(sel || null);
    if (s === JSON.stringify(this.sel || null)) return;
    this.sel = sel;
    this.selDirty = true;
    clearTimeout(this.selTimer);
    this.selTimer = setTimeout(() => { if (this.selDirty) this.pull(); }, 400);
  }

  async pull() {
    if (this.pulling) { this.pullAgain = true; return; }
    this.pulling = true;
    try {
      this.selDirty = false;
      const res = await this.api('sync', { method: 'POST', body: { rev: this.exists ? this.rev : -1, sel: this.sel } });
      this.failures = 0;
      // our own batches that the server already applied come back in res.batches
      this.absorb(res);
      if (this.cb.onStatus && !this.inflight) this.cb.onStatus(this.pending.length ? 'saving' : 'saved');
    } catch (e) {
      this.failures++;
      if (this.cb.onStatus) this.cb.onStatus('offline');
    } finally {
      this.pulling = false;
      if (this.pullAgain) { this.pullAgain = false; this.pull(); }
    }
  }

  /** Polls the server: fast while the sheet is open and visible, slowly otherwise. */
  startPolling(getInterval) {
    const tick = async () => {
      if (this.stopped) return;
      const ms = getInterval();
      if (ms > 0 && !this.inflight) await this.pull();
      this.pollTimer = setTimeout(tick, Math.max(1000, ms || 15000) + (this.failures ? Math.min(30000, 2000 * this.failures) : 0));
    };
    this.pollTimer = setTimeout(tick, getInterval() || 15000);
  }

  stop() { this.stopped = true; clearTimeout(this.pollTimer); }
}
