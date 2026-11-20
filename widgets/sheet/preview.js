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
/* Spreadsheets for YouTrack – development preview (not part of the release ZIP).
 *
 * Opened outside YouTrack, the widget talks to this simulated host. It runs the real backend.js and sheet-core.js
 * in the browser and keeps the "server" data in localStorage, so two browser tabs act as two users working on the
 * same issue:
 *   widgets/sheet/index.html?user=alice
 *   widgets/sheet/index.html?user=bob
 *   widgets/sheet/index.html?user=carol&readonly      (can view, not edit)
 *   widgets/sheet/window.html?user=alice              (options-menu window)
 *   …&locale=tr-TR  …&issue=2-99  …&reset             (clear the simulated data)
 */
const params = new URLSearchParams(location.search);
const LOGIN = params.get('user') || 'alice';
const ISSUE = params.get('issue') || '2-17';
const READONLY = params.has('readonly');
/* Simulated permissions of the signed-in preview user (changeable while the page is open, to test live updates):
   localStorage 'issue-sheets-preview:perm:<user>' or ?perm= : 'edit' (has Update Issue) | 'readonly' (only Read Issue) |
   'hidden' (the issue's visibility excludes the user).
   ?groups=Sheet editors,Sheet viewers puts the user in those groups; ?editorGroups=… / ?viewerGroups=… and
   ?readersCanEdit set the project's app settings. */
function permOf(login) {
  if (login !== LOGIN) return 'edit';
  return localStorage.getItem('issue-sheets-preview:perm:' + login) || params.get('perm') || (READONLY ? 'readonly' : 'edit');
}
const KEY = 'issue-sheets-preview:';
const USERS = { alice: 'Alice Smith', bob: 'Bob Yılmaz', carol: 'Carol Demir', dave: 'Dave Kaya' };

if (params.has('reset')) Object.keys(localStorage).filter((k) => k.startsWith(KEY)).forEach((k) => localStorage.removeItem(k));

function load(key, fallback) { try { return JSON.parse(localStorage.getItem(KEY + key)) ?? fallback; } catch (e) { return fallback; } }
function save(key, value) { localStorage.setItem(KEY + key, JSON.stringify(value)); }

async function loadModule(url, requireFn) {
  const src = await (await fetch(url, { cache: 'no-store' })).text();
  const module = { exports: {} };
  // eslint-disable-next-line no-new-func
  new Function('module', 'exports', 'require', src)(module, module.exports, requireFn);
  return module.exports;
}

/** ?demo: a realistic workbook with other people working on it (used for the Marketplace screenshots). */
function seedDemo(core) {
  Object.keys(localStorage).filter((k) => k.startsWith(KEY)).forEach((k) => localStorage.removeItem(k));
  const sh = core.newSheet('Budget', 100, 26);
  const set = (r, c, v, st) => { sh.cells[sh.rows[r] + ':' + sh.cols[c]] = { ...(v === null ? {} : { v: String(v) }), ...(st ? { st } : {}) }; };
  const usd = { nf: 'cur', cur: '$', dp: 2 };
  const rows = [
    ['Laptops for QA', 'Hardware', 'Ayşe', 4, 1450, 'Approved', 46356],
    ['4K monitors', 'Hardware', 'Bob', 6, 389.9, 'Ordered', 46349],
    ['CI build agents', 'Cloud', 'Deniz', 12, 96, 'Approved', 46363],
    ['Load test tool', 'Software', 'Carol', 1, 2400, 'In review', 46370],
    ['Design licenses', 'Software', 'Ayşe', 5, 59.9, 'Ordered', 46352],
    ['Staging database', 'Cloud', 'Deniz', 3, 310, 'Approved', 46377],
    ['Usability study', 'Services', 'Carol', 1, 3200, 'In review', 46384],
    ['Security audit', 'Services', 'Bob', 1, 5800, 'Draft', 46391]
  ];
  ['Item', 'Category', 'Owner', 'Qty', 'Unit price', 'Total', 'Status', 'Due'].forEach((t, c) => set(0, c, t));
  rows.forEach((row, i) => {
    const r = i + 1;
    set(r, 0, row[0]); set(r, 1, row[1]); set(r, 2, row[2]); set(r, 3, row[3]); set(r, 4, row[4], usd);
    set(r, 5, `=[@Qty]*[@[Unit price]]`, usd); set(r, 6, row[5]); set(r, 7, row[6], { nf: 'date' });
  });
  set(9, 0, 'Total'); set(9, 3, '=SUBTOTAL(109,Budget[Qty])'); set(9, 5, '=SUBTOTAL(109,Budget[Total])', usd);
  // summary with a chart
  set(0, 9, 'Category', { b: 1 }); set(0, 10, 'Total', { b: 1, ha: 'right' }); set(0, 11, 'Share', { b: 1, ha: 'right' });
  ['Hardware', 'Cloud', 'Software', 'Services'].forEach((cat, i) => {
    set(i + 1, 9, cat);
    set(i + 1, 10, `=SUMIF(Budget[Category],J${i + 2},Budget[Total])`, { nf: 'cur', cur: '$', dp: 0 });
    set(i + 1, 11, `=K${i + 2}/SUM($K$2:$K$5)`, { nf: 'pct', dp: 0 });
  });
  set(5, 9, 'Over 1,000', { i: 1, fc: '#6c707e' }); set(5, 10, '=COUNTIF(Budget[Total],">1000")', { ha: 'right' });
  const widths = [150, 92, 70, 48, 92, 104, 96, 92, 22, 90, 82, 58];
  widths.forEach((w, c) => { sh.colW[sh.cols[c]] = w; });
  sh.freeze = { r: 1, c: 0 };
  const id = (r, c) => ({ r1: sh.rows[r[0]], c1: sh.cols[c[0]], r2: sh.rows[r[1]], c2: sh.cols[c[1]] });
  sh.tables = [{ id: 'tbudget', name: 'Budget', ...id([0, 9], [0, 7]), header: true, totals: true, style: 'blue', banded: true, filter: true, filters: {} }];
  sh.dv = [{ id: 'dvstatus', ...id([1, 8], [6, 6]), type: 'list', source: 'Draft, In review, Approved, Ordered', strict: true }];
  sh.cf = [
    { id: 'cfbar', ...id([1, 8], [5, 5]), type: 'bar', color: '#3369d6' },
    { id: 'cfok', ...id([1, 8], [6, 6]), type: 'eq', a: 'Approved', style: { bg: '#d9ead3', fc: '#274e13' } },
    { id: 'cfrev', ...id([1, 8], [6, 6]), type: 'eq', a: 'In review', style: { bg: '#fff2cc', fc: '#7f6000' } },
    { id: 'cfdraft', ...id([1, 8], [6, 6]), type: 'eq', a: 'Draft', style: { fc: '#6c707e', i: 1 } }
  ];
  sh.charts = [{ id: 'chcat', type: 'doughnut', title: 'Budget by category', ...id([0, 4], [9, 10]), ar: sh.rows[7], ac: sh.cols[9], ox: 0, oy: 0, w: 236, h: 196 }];
  sh.cells[sh.rows[3] + ':' + sh.cols[6]].n = 'Approved in the planning call on 12 Nov.';

  const tm = core.newSheet('Test matrix', 60, 12);
  const tset = (r, c, v, st) => { tm.cells[tm.rows[r] + ':' + tm.cols[c]] = { v: String(v), ...(st ? { st } : {}) }; };
  ['Test case', 'Desktop', 'Laptop', 'Tablet', 'Phone', 'Passed'].forEach((t, c) => tset(0, c, t, { b: 1, ha: c ? 'center' : null }));
  [['Sign in with SSO', 1, 1, 1, 1], ['Upload attachment', 1, 1, 0, 1], ['Export to PDF', 1, 0, 0, 1], ['Dark theme', 1, 1, 1, 1]].forEach((row, i) => {
    tset(i + 1, 0, row[0]);
    for (let c = 1; c <= 4; c++) tset(i + 1, c, row[c] ? 'TRUE' : 'FALSE');
    tset(i + 1, 5, `=COUNTIF(B${i + 2}:E${i + 2},TRUE)/4`, { nf: 'pct', dp: 0 });
  });
  tm.dv = [{ id: 'dvcb', r1: tm.rows[1], c1: tm.cols[1], r2: tm.rows[4], c2: tm.cols[4], type: 'checkbox', strict: true }];
  tm.cf = [{ id: 'cfscale', r1: tm.rows[1], c1: tm.cols[5], r2: tm.rows[4], c2: tm.cols[5], type: 'scale3', colors: ['#e67c73', '#ffd666', '#57bb8a'] }];
  tm.colW[tm.cols[0]] = 170;

  // a capacity plan with a Solver model: the cheapest set of build agents for the required builds per hour
  const ag = core.newSheet('Build agents', 60, 12);
  const aset = (r, c, v, st) => { ag.cells[ag.rows[r] + ':' + ag.cols[c]] = { v: String(v), ...(st ? { st } : {}) }; };
  ['Agent type', 'Cost / month', 'Builds / hour', 'Agents', 'Max agents'].forEach((t, c) => aset(0, c, t, { b: 1, ha: c ? 'right' : null }));
  [['Small', 96, 4], ['Medium', 180, 9], ['Large', 320, 18]].forEach((row, i) => {
    aset(i + 1, 0, row[0]); aset(i + 1, 1, row[1], { nf: 'cur', cur: '$', dp: 0 }); aset(i + 1, 2, row[2]); aset(i + 1, 3, 2, { bg: '#fff2cc' }); aset(i + 1, 4, 6);
  });
  aset(5, 0, 'Monthly cost', { b: 1 }); aset(5, 1, '=SUMPRODUCT(B2:B4,D2:D4)', { b: 1, nf: 'cur', cur: '$', dp: 0 });
  aset(6, 0, 'Builds per hour'); aset(6, 2, '=SUMPRODUCT(C2:C4,D2:D4)');
  aset(7, 0, 'Required builds per hour'); aset(7, 2, 60);
  [150, 100, 100, 80, 90].forEach((w, c) => { ag.colW[ag.cols[c]] = w; });
  ag.solver = [{
    id: 'model', obj: '=$B$6', goal: 'min', value: '0', vars: ['=$D$2:$D$4'],
    cons: [{ l: '=$C$7', op: 'ge', r: '=$C$8' }, { l: '=$D$2:$D$4', op: 'le', r: '=$E$2:$E$4' }, { l: '=$D$2:$D$4', op: 'int', r: '' }],
    nonneg: true, method: 'lp', opts: { intTolerance: 0 }
  }];

  const doc = core.newWorkbook({ login: 'alice', name: USERS.alice }, [sh, tm, ag, core.newSheet('Notes')]);
  doc.rev = 1;
  doc.updatedBy = 'bob'; doc.updatedByName = USERS.bob;
  const at = (min) => new Date(Date.now() - min * 60000).toISOString();
  const k = (r, c) => sh.rows[r] + ':' + sh.cols[c];
  const history = [
    { id: 'h3', rev: 4, at: at(1), by: 'bob', byName: USERS.bob, events: [], cells: [{ s: sh.id, k: k(2, 6), sn: 'Budget', a: 'G3', from: { v: 'Approved' }, to: { v: 'Ordered' } }], total: 1 },
    { id: 'h2', rev: 3, at: at(14), by: 'carol', byName: USERS.carol, events: [], cells: [{ s: sh.id, k: k(4, 4), sn: 'Budget', a: 'E5', from: { v: '2150' }, to: { v: '2400' } }, { s: sh.id, k: k(7, 6), sn: 'Budget', a: 'G8', from: { v: 'Draft' }, to: { v: 'In review' } }], total: 2 },
    { id: 'h1', rev: 2, at: at(52), by: 'alice', byName: USERS.alice, events: ['Added chart Budget by category in Budget'], structural: false, cells: [], total: 0 },
    { id: 'h0', rev: 1, at: at(60), by: 'alice', byName: USERS.alice, events: ['Created the spreadsheet'], structural: true, cells: [], total: 0 }
  ];
  save('issue:' + ISSUE, { sheetDoc: JSON.stringify(doc), sheetLog: '[]', sheetHistory: JSON.stringify(history) });
  // a taller panel so that the 1280 x 800 screenshots are filled
  if (params.get('height')) localStorage.setItem(KEY + 'storage:' + LOGIN + ':height', params.get('height'));
  const now = Date.now();
  save('users', {
    bob: { sheetPresenceIssue: ISSUE, sheetPresence: JSON.stringify({ issue: ISSUE, s: sh.id, k: k(4, 6), k2: '', editing: false, at: now + 600000 }) },
    carol: { sheetPresenceIssue: ISSUE, sheetPresence: JSON.stringify({ issue: ISSUE, s: sh.id, k: k(7, 4), k2: '', editing: false, at: now + 600000 }) }
  });
}

/** ?select=F2:F9 and ?action=filter|history|cf|edit|solver|solve set the scene for screenshots. */
function stageScene() {
  const sel = params.get('select'), action = params.get('action');
  if (!sel && !action && !params.get('sheet')) return;
  const ready = () => document.querySelector('#grid td[data-r]');
  const wait = (fn) => (ready() ? setTimeout(fn, 300) : setTimeout(() => wait(fn), 100));
  wait(() => {
    const tab = params.get('sheet');
    if (tab) document.querySelectorAll('#tabs .tab')[+tab - 1].click();
    if (sel) {
      // click the first cell, shift-click the last one
      const pos = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 }; };
      const [a, b] = sel.split(':').map(pos);
      const click = (p, shift) => {
        const td = document.querySelector(`#grid td[data-r="${p.r}"][data-c="${p.c}"]`);
        if (!td) return;
        td.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, shiftKey: shift, button: 0 }));
        document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      };
      click(a, false);
      if (b) click(b, true);
    }
    if (action === 'filter') { const b = document.querySelectorAll('[data-flt]')[6]; if (b) b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); }
    if (action === 'history') { document.querySelectorAll('.mb-item')[1].click(); setTimeout(() => [...document.querySelectorAll('.menu-item')].find((b) => b.textContent.includes('Change history')).click(), 100); }
    if (action === 'cf') document.querySelector('[title="Conditional formatting"]').click();
    if (action === 'solver' || action === 'solve') {
      [...document.querySelectorAll('.mb-item')].find((b) => b.textContent === 'Data').click();
      setTimeout(() => {
        [...document.querySelectorAll('.menu-item')].find((b) => b.textContent.includes('Solver')).click();
        if (action === 'solve') {
          setTimeout(() => {
            [...document.querySelectorAll('.modal button')].find((b) => b.textContent === 'Solve').click();
            setTimeout(() => { const c = document.querySelector('[data-report="answer"]'); if (c) c.checked = true; const s = document.querySelector('[data-report="sensitivity"]'); if (s) s.checked = false; }, 600);
          }, 200);
        }
      }, 100);
    }
    if (action === 'edit') {
      const k = document.getElementById('kbd');
      k.value = '=SUMIF(Budget[Category],';
      k.dispatchEvent(new Event('input', { bubbles: true }));
    }
  });
}

export async function createPreviewHost() {
  const core = await loadModule('../../sheet-core.js', () => null);
  if (params.has('demo')) seedDemo(core);
  stageScene();
  const users = load('users', {});
  const entities = {
    User: {
      findByExtensionProperties(q) {
        const all = load('users', {});
        return Object.keys(all).filter((login) => Object.keys(q).every((k) => all[login][k] === q[k])).map((login) => userEntity(login));
      },
      findByLogin: (login) => userEntity(login)
    }
  };
  const backend = await loadModule('../../backend.js', (name) => (name === './sheet-core' ? core : name.endsWith('/entities') ? entities : null));
  if (!users[LOGIN]) { users[LOGIN] = {}; save('users', users); }

  function propsProxy(storeKey) {
    // reads and writes go straight to localStorage, like a database transaction per property
    return new Proxy({}, {
      get(_, prop) { const all = load(storeKey, {}); return all[prop] === undefined ? null : all[prop]; },
      set(_, prop, value) { const all = load(storeKey, {}); if (value === null || value === undefined) delete all[prop]; else all[prop] = value; save(storeKey, all); return true; }
    });
  }
  function userEntity(login) {
    return {
      login, visibleName: USERS[login] || login, fullName: USERS[login] || login,
      get extensionProperties() {
        return new Proxy({}, {
          get(_, prop) { const all = load('users', {}); return (all[login] || {})[prop] ?? null; },
          set(_, prop, value) { const all = load('users', {}); all[login] = all[login] || {}; if (value === null) delete all[login][prop]; else all[login][prop] = value; save('users', all); return true; }
        });
      },
      hasPermission: (key, project) => { if (!project) throw new Error('project expected'); const p = permOf(login); return key === 'UPDATE_ISSUE' ? p === 'edit' : p !== 'hidden'; },
      hasRole: () => false,
      isInGroup: (name) => login === LOGIN && (params.get('groups') || '').split(',').map((g) => g.trim()).includes(name)
    };
  }
  const issue = { id: ISSUE, idReadable: 'DEMO-17', project: { key: 'DEMO', name: 'Demo' }, isVisibleTo: (user) => permOf(user.login) !== 'hidden', extensionProperties: propsProxy('issue:' + ISSUE) };
  // the project's app settings; group settings are sets of UserGroup entities (only .name and forEach are used)
  const groupSet = (v) => { const list = (v || '').split(',').map((g) => g.trim()).filter(Boolean).map((name) => ({ name })); return { forEach: (fn) => list.forEach(fn), size: list.length }; };
  const settings = { readersCanEdit: params.has('readersCanEdit'), editorGroups: groupSet(params.get('editorGroups')), viewerGroups: groupSet(params.get('viewerGroups')) };

  async function fetchApp(path, opts = {}) {
    await new Promise((r) => setTimeout(r, 40 + Math.random() * 120));
    const [, name] = path.split('/');
    const method = (opts.method || 'GET').toUpperCase();
    const ep = backend.httpHandler.endpoints.find((e) => e.path === name && e.method === method);
    if (!ep) throw Object.assign(new Error('Not found'), { status: 404, data: { error: 'Not found' } });
    const body = opts.body === undefined ? '' : typeof opts.body === 'string' ? opts.body : JSON.stringify(opts.body);
    const query = opts.query || {};
    const res = { code: 200, body: null, json(o) { this.body = JSON.parse(JSON.stringify(o)); }, text(t) { this.body = t; } };
    const ctx = {
      currentUser: userEntity(LOGIN), issue, settings, globalStorage: { extensionProperties: propsProxy('global') },
      request: { body, json: () => (body ? JSON.parse(body) : {}), getParameter: (n) => (query[n] === undefined ? null : String(query[n])), getParameters: (n) => (query[n] === undefined ? [] : [String(query[n])]) },
      response: res
    };
    ep.handle(ctx);
    if (res.code >= 400) throw Object.assign(new Error((res.body && res.body.error) || 'Request failed'), { status: res.code, data: res.body });
    return res.body;
  }

  const storagePrefix = KEY + 'storage:' + LOGIN + ':';
  return {
    entity: { id: ISSUE, type: 'issue', idReadable: 'DEMO-17', summary: 'Preview issue' },
    locale: params.get('locale') || navigator.language || 'en-US',
    fetchApp,
    fetchYouTrack: async () => ({}),
    storage: {
      getItem: async (k) => localStorage.getItem(storagePrefix + k),
      setItem: async (k, v) => localStorage.setItem(storagePrefix + k, v),
      removeItem: async (k) => localStorage.removeItem(storagePrefix + k),
      clear: async () => {}, getKeys: async () => []
    },
    alert: (m) => console.log('[alert]', m)
  };
}

// Theme switch for the preview: ?theme=dark adds the class YouTrack sets in the frame
if (params.get('theme') === 'dark') document.body.classList.add('ring-ui-theme-dark');
