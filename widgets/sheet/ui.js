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
/* Spreadsheets for YouTrack – small UI kit in the Ring UI look: element builder, icons, menus, dropdowns, popovers, dialogs, toasts,
 * color picker. All text goes through textContent or esc(), never raw HTML from data. */

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/** h('div', {class: 'x', onclick}, child, …) */
export function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const k in attrs) {
      const v = attrs[k];
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'value') el.value = v;
      else if (k === 'checked') el.checked = !!v;
      else if (k === 'html') el.innerHTML = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
  }
  const add = (c) => {
    if (c === null || c === undefined || c === false) return;
    if (Array.isArray(c)) { c.forEach(add); return; }
    el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  };
  children.forEach(add);
  return el;
}

/* ---------------------------------------------------------------- icons (16 px, stroke = currentColor) */

const P = {
  undo: '<path d="M6 4 3 7l3 3"/><path d="M3 7h6.5a3.5 3.5 0 0 1 0 7H8"/>',
  redo: '<path d="m10 4 3 3-3 3"/><path d="M13 7H6.5a3.5 3.5 0 0 0 0 7H8"/>',
  bold: '<path d="M5 3h4a2.5 2.5 0 0 1 0 5H5zM5 8h4.5a2.5 2.5 0 0 1 0 5H5z" stroke-width="1.8"/>',
  italic: '<path d="M7 3h5M4 13h5M10 3 6 13"/>',
  underline: '<path d="M5 3v4.5a3 3 0 0 0 6 0V3M4 14h8"/>',
  strike: '<path d="M3 8h10M10.5 4.5C10 3.5 9 3 8 3c-1.7 0-3 .9-3 2.3 0 .8.4 1.4 1.2 1.8M5.5 11.3c.5 1 1.5 1.7 2.7 1.7 1.7 0 2.8-.9 2.8-2.2 0-.6-.2-1-.6-1.4"/>',
  textColor: '<path d="M4.5 11 8 2.5 11.5 11M6 7.5h4"/>',
  fill: '<path d="m3 8 5-5 5 5-5 5z"/><path d="M13 10.5s1.2 1.4 1.2 2.2a1.2 1.2 0 0 1-2.4 0c0-.8 1.2-2.2 1.2-2.2z"/>',
  borders: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M8 2.5v11M2.5 8h11" stroke-dasharray="1.5 1.5"/>',
  left: '<path d="M2.5 4h11M2.5 7h7M2.5 10h11M2.5 13h7"/>',
  center: '<path d="M2.5 4h11M4.5 7h7M2.5 10h11M4.5 13h7"/>',
  right: '<path d="M2.5 4h11M6.5 7h7M2.5 10h11M6.5 13h7"/>',
  vtop: '<path d="M2.5 2.5h11M8 5v8M5.5 7.5 8 5l2.5 2.5"/>',
  vmid: '<path d="M2.5 8h11M8 2v3.5M8 10.5V14M6 4l2 1.5L10 4M6 12l2-1.5 2 1.5"/>',
  vbottom: '<path d="M2.5 13.5h11M8 3v8M5.5 8.5 8 11l2.5-2.5"/>',
  wrap: '<path d="M2.5 4h11M2.5 8h9a2 2 0 0 1 0 4H8M9.5 10.5 8 12l1.5 1.5M2.5 12h3"/>',
  merge: '<rect x="2.5" y="3.5" width="11" height="9" rx="1"/><path d="M5 8h6M6.5 6.5 5 8l1.5 1.5M9.5 6.5 11 8l-1.5 1.5"/>',
  pct: '<path d="M4 12 12 4"/><circle cx="4.8" cy="4.8" r="1.6"/><circle cx="11.2" cy="11.2" r="1.6"/>',
  currency: '<path d="M10.8 5.2C10.3 4.3 9.3 3.8 8 3.8c-1.7 0-2.8.8-2.8 2s1 1.7 2.8 2.1 2.8 1 2.8 2.2-1.2 2.1-2.8 2.1c-1.4 0-2.4-.6-2.9-1.5M8 2v12"/>',
  // ".0" with an arrow to the left, ".00" with an arrow to the right (drawn, not text, so they stay sharp)
  decDec: '<circle cx="3" cy="9.3" r=".9" fill="currentColor" stroke="none"/><rect x="5.5" y="2.5" width="4" height="7" rx="2"/><path d="M14 13H6.5M8.5 11l-2 2 2 2"/>',
  decInc: '<circle cx="1.9" cy="9.3" r=".9" fill="currentColor" stroke="none"/><rect x="3.8" y="2.5" width="4" height="7" rx="2"/><rect x="9.6" y="2.5" width="4" height="7" rx="2"/><path d="M3.5 13H11M9 11l2 2-2 2"/>',
  sum: '<path d="M12 3.5H4l4.5 4.5L4 12.5h8"/>',
  filter: '<path d="M2.5 3.5h11l-4.3 5v4l-2.4 1v-5z"/>',
  sortAsc: '<path d="M4.5 3v10M2.5 5l2-2 2 2M9 4h4.5M9 8h3.5M9 12h2"/>',
  sortDesc: '<path d="M4.5 3v10M2.5 11l2 2 2-2M9 4h2M9 8h3.5M9 12h4.5"/>',
  chart: '<path d="M2.5 13.5h11M4.5 11V8M7.5 11V4.5M10.5 11V6.5"/>',
  table: '<rect x="2.5" y="3" width="11" height="10" rx="1"/><path d="M2.5 6.5h11M2.5 9.8h11M6.5 6.5V13"/>',
  note: '<path d="M3 2.5h10v8l-3 3H3z"/><path d="M10 13.5v-3h3M5.5 6h5M5.5 8.5h3"/>',
  fx: '<path d="M9 3.4C8.6 2.8 8 2.5 7.3 2.5c-1 0-1.6.7-1.8 1.7L4.2 13.5M3.8 6.5h4"/><path d="m9.2 7.5 4 6m0-6-4 6"/>',
  cf: '<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M8 2.5v11M2.5 8h11"/><rect x="8" y="8" width="5.5" height="5.5" fill="currentColor"/>',
  more: '<circle cx="3.5" cy="8" r="1.1" fill="currentColor"/><circle cx="8" cy="8" r="1.1" fill="currentColor"/><circle cx="12.5" cy="8" r="1.1" fill="currentColor"/>',
  chevron: '<path d="m6 4 4 4-4 4"/>',
  down: '<path d="m4 6 4 4 4-4"/>',
  up: '<path d="M8 13V3M4 7l4-4 4 4"/>',
  plus: '<path d="M8 3v10M3 8h10"/>',
  close: '<path d="m4 4 8 8M12 4l-8 8"/>',
  error: '<circle cx="8" cy="8" r="6"/><path d="M8 4.8v3.6"/><circle cx="8" cy="11" r=".9" fill="currentColor" stroke="none"/>',
  history: '<path d="M2.8 8a5.2 5.2 0 1 0 1.5-3.7"/><path d="M2.5 2.8v2.7h2.7M8 5.2V8l2 1.5"/>',
  sheets: '<path d="M2.5 4h11M2.5 8h11M2.5 12h11"/>',
  clear: '<path d="m9.5 3 4 4-6 6H4l-1.5-1.5z"/><path d="M6 6.5l4 4M8 13h5.5"/>',
  link: '<path d="M7 9a2.5 2.5 0 0 0 3.5 0l2-2A2.5 2.5 0 0 0 9 3.5l-.7.7M9 7a2.5 2.5 0 0 0-3.5 0l-2 2A2.5 2.5 0 0 0 7 12.5l.7-.7"/>',
  search: '<circle cx="7" cy="7" r="4"/><path d="m10 10 3.5 3.5"/>',
  freeze: '<rect x="2.5" y="2.5" width="11" height="11" rx="1"/><path d="M2.5 6h11M6 2.5v11" stroke-width="2"/>',
  paint: '<rect x="3" y="2.5" width="9" height="4" rx="1"/><path d="M12 4.5h1.5V8H8v2.5M7 10.5h2V14H7z"/>',
  check: '<path d="m3.5 8.5 3 3 6-7"/>',
  window: '<rect x="2.5" y="3" width="11" height="10" rx="1.5"/><path d="M2.5 6h11"/>',
  download: '<path d="M8 2.5v8M5 7.5l3 3 3-3M3 13h10"/>',
  upload: '<path d="M8 10.5v-8M5 5.5l3-3 3 3M3 13h10"/>',
  trash: '<path d="M3 4.5h10M6.5 4.5V3h3v1.5M4.5 4.5l.7 9h5.6l.7-9"/>',
  eye: '<path d="M1.5 8S4 3.5 8 3.5 14.5 8 14.5 8 12 12.5 8 12.5 1.5 8 1.5 8z"/><circle cx="8" cy="8" r="2"/>',
  dropdown: '<path d="m4.5 6.5 3.5 3.5 3.5-3.5"/>',
  checkbox: '<rect x="2.5" y="2.5" width="11" height="11" rx="2"/><path d="m5 8 2 2 4-4.5"/>',
  calendar: '<rect x="2.5" y="3.5" width="11" height="10" rx="1.5"/><path d="M2.5 6.5h11M5.5 2v3M10.5 2v3"/>',
  pick: '<rect x="2.5" y="2.5" width="8" height="8" rx="1" stroke-dasharray="2 1.6"/><path d="M8.5 8.5 13.5 10.5 11.3 11.3 10.5 13.5z"/>',
  solver: '<path d="M2.5 13.5h11M2.5 13.5V2.5"/><path d="M4.5 11 7 7.5l2.5 2L13 4"/><circle cx="13" cy="4" r="1.2" fill="currentColor"/>',
  whatif: '<path d="M6 6a2 2 0 1 1 2.8 1.8c-.5.3-.8.7-.8 1.2v.5"/><path d="M8 12h.01"/><circle cx="8" cy="8" r="5.5"/>'
};
export function icon(name, size = 16) {
  return `<svg class="ic" width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[name] || ''}</svg>`;
}
export function iconEl(name, size) { const s = h('span', { class: 'ic-wrap', html: icon(name, size) }); return s; }

/* ---------------------------------------------------------------- toast (Ring UI alert look) */

let toastTimer;
/* In the issue the frame is only as high as the page content (YouTrack sizes it from the content). A short page,
   such as the one-row state of an issue without a spreadsheet, has no room for a message: while one shows, the page
   reserves room for it under the content, so the frame grows, and it shrinks again when the message goes. */
let roomObserver = null;
function fitToast() {
  const body = document.body;
  const el = document.getElementById('toast');
  const app = document.getElementById('app');
  if (!el || !app || !body.classList.contains('mode-inline')) return;
  body.style.minHeight = '';
  el.style.bottom = '';
  if (el.hidden) return;
  const bottom = parseFloat(getComputedStyle(el).bottom) || 0;
  if (app.offsetHeight >= el.offsetHeight + bottom + 8) return; // tall enough: the message floats over the content
  el.style.bottom = '8px';
  body.style.minHeight = Math.ceil(app.offsetHeight + 8 + el.offsetHeight + 8) + 'px';
}
function toastRoom() {
  if (!document.body.classList.contains('mode-inline')) return;
  fitToast();
  // the content can change while the message shows (a deleted spreadsheet leaves only the one-row state)
  if (!roomObserver && typeof ResizeObserver !== 'undefined' && document.getElementById('app')) {
    roomObserver = new ResizeObserver(() => fitToast());
    roomObserver.observe(document.getElementById('app'));
  }
}
function hideToast(el) { el.hidden = true; fitToast(); }
/** Shows a short message: 3.2 s, errors 8 s. */
export function toast(msg, kind) {
  const error = kind === 'error';
  const ms = error ? 8000 : 3200;
  clearTimeout(toastTimer);
  const body = document.body;
  let el = document.getElementById('toast');
  if (!el) { el = h('div', { id: 'toast' }); body.appendChild(el); }
  el.className = 'toast' + (kind ? ' ' + kind : '');
  el.setAttribute('role', error ? 'alert' : 'status');
  el.setAttribute('aria-live', error ? 'assertive' : 'polite');
  el.innerHTML = '';
  if (error) el.appendChild(h('span', { class: 'toast-ic', html: icon('error') }));
  el.appendChild(h('div', { class: 'toast-text' }, msg));
  const x = h('button', { class: 'icon-btn toast-x', type: 'button', title: 'Close', 'aria-label': 'Close', html: icon('close') });
  x.addEventListener('mousedown', (e) => e.preventDefault());
  x.addEventListener('click', () => { clearTimeout(toastTimer); hideToast(el); });
  el.appendChild(x);
  el.hidden = false;
  // replay the entrance animation for a new message
  el.style.animation = 'none';
  void el.offsetWidth;
  el.style.animation = '';
  toastRoom();
  toastTimer = setTimeout(() => { hideToast(el); }, ms);
}

/* ---------------------------------------------------------------- popups */

let openPopups = [];
export function closePopups(except) {
  openPopups.filter((p) => p !== except && !(except && except.contains && p.contains(except))).forEach((p) => { p.remove(); if (p._onClose) p._onClose(); });
  openPopups = openPopups.filter((p) => p.isConnected);
}
document.addEventListener('mousedown', (e) => {
  if (!openPopups.length) return;
  if (e.target.closest && e.target.closest('[data-popup-anchor]')) return; // the anchor's own click toggles its popup
  // a click inside a popup closes only the popups opened above it (such as a dropdown list inside the filter popover)
  let i = openPopups.length - 1;
  while (i >= 0 && !openPopups[i].contains(e.target)) i--;
  if (i >= 0) {
    const above = openPopups.slice(i + 1);
    openPopups = openPopups.slice(0, i + 1);
    above.forEach((p) => { p.remove(); if (p._onClose) p._onClose(); });
    return;
  }
  closePopups();
}, true);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && openPopups.length) {
    e.stopPropagation(); e.preventDefault();
    const p = openPopups.pop();
    p.remove();
    if (p._onClose) p._onClose();
    if (p._parentItem && p._parentItem.isConnected) p._parentItem.focus({ preventScroll: true }); // back to the item that opened the submenu
  }
}, true);

/** Places an absolutely positioned popup next to a rect, inside the viewport. */
export function place(el, rect, prefer = 'below') {
  el.style.visibility = 'hidden';
  document.body.appendChild(el);
  const w = el.offsetWidth, hh = el.offsetHeight;
  const vw = document.documentElement.clientWidth, vh = Math.max(document.documentElement.clientHeight, window.innerHeight);
  let x = prefer === 'right' ? rect.right : rect.left;
  let y = prefer === 'right' ? rect.top : rect.bottom + 2;
  if (x + w > vw - 4) x = prefer === 'right' ? Math.max(4, rect.left - w) : Math.max(4, vw - w - 4);
  if (y + hh > vh - 4) {
    const below = vh - 4 - (rect.bottom + 2), above = rect.top - 2 - 4;
    if (prefer === 'right') y = Math.max(4, vh - hh - 4);
    else if (hh <= above) y = rect.top - hh - 2;
    else if (Math.max(above, below) >= 80) {
      // too tall for either side: it scrolls on the larger side and never covers its anchor (a second click on the
      // anchor then still closes it)
      const room = Math.max(above, below);
      el.style.maxHeight = room + 'px';
      el.style.overflowY = 'auto';
      y = below >= above ? rect.bottom + 2 : rect.top - 2 - room;
    } else y = Math.max(4, rect.top - hh - 2);
  }
  if (y < 4) y = 4;
  el.style.left = x + 'px';
  el.style.top = y + 'px';
  el.style.visibility = '';
}

export function popover(content, anchorRect, opts = {}) {
  if (!opts.keep) closePopups();
  const el = h('div', { class: 'popover ' + (opts.class || ''), role: opts.role || 'dialog' }, content);
  place(el, anchorRect, opts.prefer);
  openPopups.push(el);
  el._onClose = opts.onClose || null;
  return el;
}

const isItem = (it) => it && it !== '-' && !it.header;
/** Without the disabled items (for people who can only view): no greyed-out entries, no empty submenus, no doubled separators. */
function pruneMenu(items) {
  const out = [];
  for (const it of items.filter(Boolean)) {
    if (it === '-') { if (out.length && out[out.length - 1] !== '-') out.push('-'); continue; }
    if (it.disabled) continue;
    if (Array.isArray(it.sub)) {
      const sub = pruneMenu(it.sub);
      if (!sub.some(isItem)) continue;
      out.push({ ...it, sub });
      continue;
    }
    out.push(it);
  }
  while (out.length && out[out.length - 1] === '-') out.pop();
  return out;
}

/**
 * Menu. items: [{ label, kbd, icon, action, disabled, checked, sub: [...] } | '-' ]
 * opts: { prefer, onDone, onClose, focusIndex, minWidth, hideDisabled, focus, keepFor }
 * keepFor: an element in a popover (a dropdown in the filter popover): opening the menu and picking an item keep the
 * popups that contain it open.
 * Returns the menu element, or null when there is nothing to show.
 */
export function menu(items, anchorRect, opts = {}) {
  if (!opts.sub) closePopups(opts.keepFor);
  let list = items.filter(Boolean);
  if (opts.hideDisabled) list = pruneMenu(list);
  if (!list.some(isItem)) return null;
  const el = h('div', { class: 'menu', role: 'menu' });
  // the check and icon columns only take room when an item uses them
  if (list.some((it) => isItem(it) && it.checked !== undefined)) el.classList.add('has-check');
  if (list.some((it) => isItem(it) && (it.icon || it.swatch))) el.classList.add('has-icon');
  if (opts.minWidth) el.style.minWidth = Math.ceil(opts.minWidth) + 'px';
  let subOpen = null;
  const buttons = [];
  const closeSub = () => { if (subOpen) { subOpen.remove(); openPopups = openPopups.filter((p) => p !== subOpen); subOpen = null; } };
  list.forEach((it) => {
    if (it === '-') { el.appendChild(h('div', { class: 'menu-sep', role: 'separator' })); return; }
    if (it.header) { el.appendChild(h('div', { class: 'menu-header' }, it.header)); return; }
    const b = h('button', { class: 'menu-item' + (it.danger ? ' danger' : ''), role: it.checked !== undefined ? 'menuitemcheckbox' : 'menuitem', 'aria-checked': it.checked !== undefined ? String(!!it.checked) : null, type: 'button', disabled: it.disabled ? true : null },
      h('span', { class: 'mi-check', html: it.checked ? icon('check', 14) : '' }),
      h('span', { class: 'mi-icon', html: it.icon ? icon(it.icon, 16) : (it.swatch ? `<span class="swatch" style="background:${esc(it.swatch)}"></span>` : '') }),
      h('span', { class: 'mi-label' }, it.label),
      h('span', { class: 'mi-kbd' }, it.kbd || ''),
      it.sub ? h('span', { class: 'mi-sub', html: icon('chevron', 12) }) : null);
    b._label = String(it.label ?? '');
    buttons.push(b);
    const openSub = () => {
      closeSub();
      if (!it.sub || it.disabled) return;
      subOpen = menu(typeof it.sub === 'function' ? it.sub() : it.sub, b.getBoundingClientRect(), { sub: true, prefer: 'right', onDone: opts.onDone, hideDisabled: opts.hideDisabled, parentItem: b, keepFor: opts.keepFor });
    };
    // from the keyboard a submenu takes the focus
    b._openSubFocus = () => { openSub(); const f = subOpen && subOpen.querySelector('.menu-item:not(:disabled)'); if (f) f.focus({ preventScroll: true }); };
    b.addEventListener('mouseenter', () => { if (it.sub) openSub(); else closeSub(); });
    b.addEventListener('click', (e) => {
      e.preventDefault();
      if (it.sub) { if (e.detail === 0) b._openSubFocus(); else openSub(); return; }
      closePopups(opts.keepFor);
      setTimeout(() => { if (it.action) it.action(); if (opts.onDone) opts.onDone(); }, 0);
    });
    el.appendChild(b);
  });
  place(el, anchorRect, opts.prefer);
  openPopups.push(el);
  el._onClose = opts.onClose || null;
  el._parentItem = opts.parentItem || null;
  let typed = '', typedAt = 0;
  el.addEventListener('keydown', (e) => {
    const i = buttons.indexOf(document.activeElement);
    const n = buttons.length;
    const go = (from, dir) => { for (let k = 1; k <= n; k++) { const b = buttons[(from + dir * k + n * 2) % n]; if (!b.disabled) { b.focus(); return; } } };
    if (e.key === 'ArrowDown') { e.preventDefault(); go(i < 0 ? -1 : i, 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); go(i < 0 ? 0 : i, -1); }
    else if (e.key === 'Home') { e.preventDefault(); go(-1, 1); }
    else if (e.key === 'End') { e.preventDefault(); go(n, -1); }
    else if (e.key === 'ArrowRight' && i >= 0) { e.preventDefault(); buttons[i]._openSubFocus(); }
    else if (e.key === 'ArrowLeft' && opts.parentItem) {
      // closes this submenu and goes back to its item
      e.preventDefault();
      el.remove();
      openPopups = openPopups.filter((p) => p !== el);
      opts.parentItem.focus({ preventScroll: true });
    }
    else if (e.key.length === 1 && e.key !== ' ' && !e.ctrlKey && !e.metaKey && !e.altKey) {
      // type-ahead: jump to the next item that starts with the typed letters
      const now = Date.now();
      typed = now - typedAt < 800 ? typed + e.key.toLowerCase() : e.key.toLowerCase();
      typedAt = now;
      const start = typed.length === 1 ? i + 1 : Math.max(i, 0);
      for (let k = 0; k < n; k++) {
        const b = buttons[(start + k) % n];
        if (!b.disabled && b._label.toLowerCase().startsWith(typed)) { b.focus(); break; }
      }
      e.preventDefault();
    }
  });
  if (opts.focus !== false && !opts.sub) {
    setTimeout(() => {
      const want = opts.focusIndex !== undefined ? buttons[opts.focusIndex] : null;
      const f = want && !want.disabled ? want : buttons.find((b) => !b.disabled);
      if (f) { f.focus({ preventScroll: true }); if (want) f.scrollIntoView({ block: 'nearest' }); }
    }, 0);
  }
  return el;
}

/* ---------------------------------------------------------------- dialogs */

/**
 * Modal dialog. Returns { el, body, close }. buttons: [{ label, primary, danger, action(close) → false keeps it open }]
 * As in Ring UI: the primary button comes first, then the others in the given order, on the left; buttons that only
 * delete something (danger, not primary) and opts.footLeft sit on the right.
 */
export function dialog(title, body, buttons = [], opts = {}) {
  closePopups();
  const prevFocus = document.activeElement;
  const back = h('div', { class: 'modal-back' });
  const box = h('div', { class: 'modal' + (opts.wide ? ' wide' : '') + (opts.class ? ' ' + opts.class : ''), role: 'dialog', 'aria-modal': 'true', 'aria-label': title });
  const head = h('div', { class: 'modal-head' }, h('h2', null, title));
  const content = h('div', { class: 'modal-body' }, body);
  const foot = h('div', { class: 'modal-foot' });
  const end = h('div', { class: 'modal-foot-end' });
  const list = buttons.filter(Boolean);
  list.filter((b) => b.primary).concat(list.filter((b) => !b.primary)).forEach((b) => {
    const btn = h('button', { class: 'btn' + (b.primary ? ' primary' : '') + (b.danger ? ' danger' : ''), type: 'button' }, b.label);
    btn.addEventListener('click', async () => {
      if (b.action) {
        btn.disabled = true;
        let keep;
        try { keep = await b.action(close); } finally { btn.disabled = false; }
        if (keep === false) return;
      }
      close();
    });
    (b.danger && !b.primary ? end : foot).appendChild(btn);
  });
  if (opts.footLeft) end.appendChild(opts.footLeft);
  if (end.children.length) foot.appendChild(end);
  box.append(head, content, foot.children.length ? foot : '');
  back.appendChild(box);
  document.body.appendChild(back);
  back.addEventListener('mousedown', (e) => { if (e.target === back && !opts.sticky) close(); });
  // Enter submits, but never with a button that deletes something
  const submit = () => foot.querySelector('.btn.primary:not(.danger):not(:disabled)');
  box.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.stopPropagation(); close(); }
    if (e.key === 'Enter' && !e.shiftKey && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') {
      const primary = submit();
      if (primary) { e.preventDefault(); primary.click(); }
    }
  });
  // the focus always goes into the dialog (Esc and Enter are handled on the box), never to a button that deletes
  box.tabIndex = -1;
  setTimeout(() => {
    const f = box.querySelector('[autofocus]') || box.querySelector('input:not([type=checkbox]):not([type=radio]),select,textarea,.dd') || submit()
      || (foot.querySelector('.btn.primary.danger') ? foot.querySelector('.btn:not(.danger)') : null)
      || content.querySelector('.btn.primary:not(.danger):not(:disabled)') || foot.querySelector('.btn:not(.danger):not(:disabled)') || box;
    if (f) { f.focus(); if (f.select && f.tagName === 'INPUT') f.select(); }
  }, 0);
  let closed = false;
  // the topmost visible dialog (not one that is hidden while cells are picked on the sheet)
  const isTop = () => {
    if (closed || back.style.display === 'none') return false;
    const all = document.querySelectorAll('.modal-back');
    return all[all.length - 1] === back;
  };
  const inPopup = (t) => openPopups.some((p) => p.contains(t));
  const focusables = () => [...box.querySelectorAll('button, input, select, textarea, a[href], [tabindex]:not([tabindex="-1"])')]
    .filter((x) => !x.disabled && x.offsetParent !== null);
  // Tab and Shift+Tab cycle inside the dialog; when the focus left it (the focused button was redrawn away), Esc still
  // closes the topmost dialog and Tab goes back into it
  const onDocKey = (e) => {
    if (e.defaultPrevented || !isTop() || inPopup(e.target)) return;
    if (e.key === 'Tab') {
      const f = focusables();
      if (!f.length) { e.preventDefault(); box.focus({ preventScroll: true }); return; }
      const i = f.indexOf(document.activeElement);
      if (i < 0 || (e.shiftKey && i === 0) || (!e.shiftKey && i === f.length - 1)) {
        e.preventDefault();
        f[e.shiftKey ? (i < 0 ? f.length - 1 : f.length - 1) : 0].focus({ preventScroll: true });
      }
      return;
    }
    if (e.key !== 'Escape' || box.contains(e.target)) return;
    e.stopPropagation(); e.preventDefault();
    close();
  };
  // a click or script that moves the focus behind the dialog brings it back into the dialog
  const onFocusIn = (e) => { if (isTop() && !box.contains(e.target) && !inPopup(e.target)) box.focus({ preventScroll: true }); };
  document.addEventListener('keydown', onDocKey, true);
  document.addEventListener('focusin', onFocusIn);
  function close() {
    if (closed) return;
    closed = true;
    document.removeEventListener('keydown', onDocKey, true);
    document.removeEventListener('focusin', onFocusIn);
    closePopups();
    back.remove();
    if (opts.onClose) opts.onClose();
    // back to where the focus was; when that element is gone, to the dialog underneath
    const under = [...document.querySelectorAll('.modal')].pop();
    const to = prevFocus && prevFocus.isConnected && prevFocus.focus ? prevFocus : under;
    if (to) to.focus({ preventScroll: true });
  }
  return { el: box, body: content, close };
}

export function confirmDialog(title, text, okLabel = 'OK', danger = false) {
  return new Promise((resolve) => {
    let ok = false;
    dialog(title, h('p', { class: 'dlg-text' }, text), [
      { label: 'Cancel' },
      { label: okLabel, primary: true, danger, action: () => { ok = true; } }
    ], { onClose: () => resolve(ok) });
  });
}

export function promptDialog(title, label, value = '', opts = {}) {
  return new Promise((resolve) => {
    let result = null;
    const input = h('input', { type: opts.type || 'text', class: 'input', value, autofocus: true, min: opts.min, max: opts.max, step: opts.step });
    dialog(title, h('label', { class: 'fld' }, h('span', null, label), input, opts.help ? h('span', { class: 'help' }, opts.help) : null), [
      { label: 'Cancel' },
      { label: opts.ok || 'OK', primary: true, action: () => { result = input.value; } }
    ], { onClose: () => resolve(result) });
  });
}

/** Field helpers for dialogs. */
export function field(label, control, help) {
  return h('label', { class: 'fld' }, h('span', null, label), control, help ? h('span', { class: 'help' }, help) : null);
}
/**
 * Dropdown in the Ring UI look, used instead of a native <select> (whose list follows the OS, not YouTrack).
 * It behaves like a select for the callers: .value reads and sets the value, picking an item fires a bubbling
 * "change" event, .disabled works. Opens with a click, Enter, Space or the arrow keys.
 */
export function select(options, value, attrs = {}) {
  const opts = (Array.isArray(options) ? options.map((o) => (Array.isArray(o) ? o : [o, o])) : Object.entries(options)).map(([v, l]) => [String(v), String(l)]);
  const { class: cls, ...rest } = attrs || {};
  const extra = String(cls || '').split(/\s+/).filter((c) => c && c !== 'input');
  const label = h('span', { class: 'dd-label' });
  const b = h('button', { ...rest, type: 'button', class: ['btn', 'dd', ...extra].join(' '), 'aria-haspopup': 'listbox', 'aria-expanded': 'false', 'data-popup-anchor': '' },
    label, h('span', { class: 'dd-chev', html: icon('dropdown') }));
  let cur = opts.length ? opts[0][0] : '';
  const paint = () => { const o = opts.find(([v]) => v === cur); label.textContent = o ? o[1] : ''; };
  Object.defineProperty(b, 'value', {
    configurable: true,
    get: () => cur,
    set: (v) => { cur = v === null || v === undefined ? '' : String(v); paint(); }
  });
  if (value !== undefined && value !== null) b.value = value; else paint();
  let open = null;
  const show = () => {
    if (b.disabled) return;
    const r = b.getBoundingClientRect();
    const items = opts.map(([v, l]) => ({
      label: l, checked: v === cur,
      action: () => { if (v === cur) return; b.value = v; b.dispatchEvent(new Event('change', { bubbles: true })); }
    }));
    // keepFor: a dropdown in a popover (the table filter) keeps that popover open
    open = menu(items, r, {
      focusIndex: Math.max(0, opts.findIndex(([v]) => v === cur)), minWidth: r.width, keepFor: b,
      onClose: () => { open = null; b.setAttribute('aria-expanded', 'false'); if (b.isConnected) b.focus({ preventScroll: true }); }
    });
    if (open) b.setAttribute('aria-expanded', 'true');
  };
  b.addEventListener('click', (e) => { e.preventDefault(); if (open && open.isConnected) closePopups(b); else show(); });
  b.addEventListener('keydown', (e) => {
    if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && !e.altKey) { e.preventDefault(); if (!open || !open.isConnected) show(); }
  });
  return b;
}
export function checkbox(label, checked, attrs = {}) {
  const input = h('input', { type: 'checkbox', checked, ...attrs });
  return { el: h('label', { class: 'check' }, input, h('span', null, label)), input };
}

/* ---------------------------------------------------------------- colors */

export const COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
  '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
  '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
  '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
  '#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
  '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130'
];

/** Color picker popover. onPick(hex | null). */
export function colorPicker(anchorRect, current, onPick, resetLabel = 'Automatic') {
  const grid = h('div', { class: 'swatches', role: 'listbox' });
  COLORS.forEach((c) => grid.appendChild(h('button', {
    class: 'sw' + (current === c ? ' on' : ''), type: 'button', title: c, 'aria-label': c, style: { background: c },
    onclick: () => { closePopups(); onPick(c); }
  })));
  const custom = h('input', { type: 'color', value: current || '#3369d6', 'aria-label': 'Custom color' });
  custom.addEventListener('change', () => { closePopups(); onPick(custom.value.toLowerCase()); });
  const content = h('div', { class: 'color-pop' },
    h('button', { class: 'btn small', type: 'button', onclick: () => { closePopups(); onPick(null); } }, resetLabel),
    grid,
    h('label', { class: 'custom-color' }, h('span', null, 'Custom'), custom));
  return popover(content, anchorRect);
}

const HEX = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i;
/** Relative luminance (WCAG) of a #rrggbb color, or null for anything else. */
export function luminance(hex) {
  const m = HEX.exec(hex || '');
  if (!m) return null;
  const [r, g, b] = [m[1], m[2], m[3]].map((x) => parseInt(x, 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Text color that stays readable on a background color. */
export function contrastText(bg) {
  const L = luminance(bg);
  if (L === null) return null;
  return L > 0.2 ? '#1f2023' : '#f5f6f8'; // 0.2: dark and light text give the same contrast here
}

/* Colors that users pick are stored for the light theme. In the dark theme light fills become dark ones of the same
   hue and dark text becomes light, so that colored cells stay readable. */
const hslCache = new Map();
function toHsl(hex) {
  const m = HEX.exec(hex);
  const [r, g, b] = [m[1], m[2], m[3]].map((x) => parseInt(x, 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
  let hue = 0, s = 0;
  if (d) {
    s = d / (1 - Math.abs(2 * l - 1));
    hue = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    hue *= 60;
    if (hue < 0) hue += 360;
  }
  return [hue, s, l];
}
function hslHex(hue, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((hue / 60) % 2) - 1)), m = l - c / 2;
  const [r, g, b] = hue < 60 ? [c, x, 0] : hue < 120 ? [x, c, 0] : hue < 180 ? [0, c, x] : hue < 240 ? [0, x, c] : hue < 300 ? [x, 0, c] : [c, 0, x];
  return '#' + [r, g, b].map((v) => Math.round(Math.min(1, Math.max(0, v + m)) * 255).toString(16).padStart(2, '0')).join('');
}
function cachedHsl(kind, hex, fn) {
  const key = kind + hex.toLowerCase();
  let v = hslCache.get(key);
  if (v === undefined) {
    if (hslCache.size > 2000) hslCache.clear();
    v = HEX.test(hex) ? fn(...toHsl(hex)) : hex;
    hslCache.set(key, v);
  }
  return v;
}
/** A dark fill of the same hue for the dark theme. */
export function darkFill(hex) { return cachedHsl('f', hex, (hue, s, l) => hslHex(hue, s * 0.7, Math.min(0.28, Math.max(0.15, 1.15 - l)))); }
/** A light text color of the same hue for the dark theme. */
export function liftText(hex) { return cachedHsl('t', hex, (hue, s, l) => hslHex(hue, s, 1 - 0.6 * l)); }

/**
 * Classes and inline declarations for a fill and a text color that people picked (#rrggbb, stored for the light theme):
 * class f with --cbg (and --cbg-d, a darker fill for the dark theme), class c with --cfc (and --cfc-d, a lighter text).
 * app.css draws them for cells and for the rule previews of conditional formatting. Returns { cls, parts }.
 */
export function themedColors(bg, fc) {
  const parts = [];
  let cls = '';
  const fill = bg && HEX.test(bg) ? bg : null;
  if (fill) {
    cls += ' f';
    parts.push('--cbg:' + fill);
    if (luminance(fill) > 0.4) parts.push('--cbg-d:' + darkFill(fill));
  }
  if (fc && HEX.test(fc)) {
    cls += ' c';
    parts.push('--cfc:' + fc);
    if (luminance(fc) < 0.25) parts.push('--cfc-d:' + liftText(fc));
  } else if (!fc && fill && luminance(fill) <= 0.4) parts.push('color:' + contrastText(fill)); // mid and dark fills keep their color in both themes
  return { cls, parts };
}

/** The current color as the bar under a button's label (a .cbar child). Dark text colors show lighter in the dark theme,
 * as they do in the cells. */
export function paintBar(btn, color, isText) {
  btn.style.setProperty('--cur', color || (isText ? 'var(--text)' : 'transparent'));
  const L = isText && color ? luminance(color) : null;
  if (L !== null && L < 0.25) btn.style.setProperty('--cur-d', liftText(color));
  else btn.style.removeProperty('--cur-d');
}

export function initials(name) {
  const parts = String(name || '?').trim().split(/\s+/);
  return ((parts[0] || '?')[0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}
export const PEER_COLORS = ['#e5793a', '#1f9e6e', '#c94f7c', '#8a63d2', '#d4a72c', '#2a9fbd', '#b24a3b', '#5c8f2b'];
export function peerColor(login) {
  let hsh = 0;
  for (const ch of String(login)) hsh = (hsh * 31 + ch.charCodeAt(0)) >>> 0;
  return PEER_COLORS[hsh % PEER_COLORS.length];
}

export function relTime(iso, locale) {
  if (!iso) return '';
  const s = Math.round((Date.now() - Date.parse(iso)) / 1000);
  try {
    const rtf = new Intl.RelativeTimeFormat(locale || 'en', { numeric: 'auto' });
    if (s < 45) return rtf.format(0, 'second');
    if (s < 3600) return rtf.format(-Math.round(s / 60), 'minute');
    if (s < 86400) return rtf.format(-Math.round(s / 3600), 'hour');
    if (s < 86400 * 7) return rtf.format(-Math.round(s / 86400), 'day');
  } catch (e) { /* ignore */ }
  try { return new Intl.DateTimeFormat(locale || 'en', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(iso)); } catch (e) { return iso.slice(0, 16).replace('T', ' '); }
}
