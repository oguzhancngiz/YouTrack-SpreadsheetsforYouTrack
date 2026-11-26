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
/* Spreadsheets for YouTrack – charts drawn as inline SVG (column, bar, stacked, line, area, pie, doughnut, scatter).
 * Colors for text and grid lines come from the Ring UI variables, so charts follow the YouTrack theme. */

export const CHART_TYPES = {
  column: 'Column', stacked: 'Stacked column', bar: 'Bar', line: 'Line', area: 'Area', pie: 'Pie', doughnut: 'Doughnut', scatter: 'Scatter'
};
export const PALETTE = ['#3369d6', '#e5793a', '#1f9e6e', '#c94f7c', '#8a63d2', '#d4a72c', '#2a9fbd', '#7a7f8c', '#b24a3b', '#5c8f2b'];
// Upper bounds for one chart, so a huge range cannot freeze the page: series, points per series, points in all.
export const CHART_LIMITS = { series: 50, points: 2000, total: 20000 };

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const HEX_RE = /^#[0-9a-fA-F]{6}$/;
// A color reaches the SVG only as #rrggbb; anything else gets the first palette color.
const paint = (c) => (typeof c === 'string' && HEX_RE.test(c) ? c : PALETTE[0]);
// Points per series when a chart has n series.
const pointLimit = (n) => Math.min(CHART_LIMITS.points, Math.floor(CHART_LIMITS.total / Math.max(1, n)));

/**
 * values: 2-D array of computed values; texts: same shape, as shown in the cells.
 * Returns { categories, series: [{ name, values }] }.
 */
export function chartData(values, texts) {
  if (!values.length || !values[0].length) return { categories: [], series: [] };
  const isNum = (v) => typeof v === 'number';
  const firstRowHeader = values[0].some((v, j) => v !== null && !isNum(v) && (j > 0 || values[0].length === 1));
  const firstColCats = values.slice(firstRowHeader ? 1 : 0).some((r) => r[0] !== null && !isNum(r[0])) || (values[0].length > 1 && firstRowHeader && values[0][0] === null);
  const top = firstRowHeader ? 1 : 0, left = firstColCats ? 1 : 0;
  // only the part of a huge range that can be drawn (CHART_LIMITS)
  const colEnd = Math.min(values[0].length, left + CHART_LIMITS.series);
  const rowEnd = Math.min(values.length, top + pointLimit(colEnd - left));
  const categories = [];
  for (let i = top; i < rowEnd; i++) categories.push(left ? texts[i][0] : String(i - top + 1));
  const series = [];
  for (let j = left; j < colEnd; j++) {
    series.push({
      name: top ? texts[0][j] || 'Series ' + (j - left + 1) : 'Series ' + (j - left + 1),
      values: values.slice(top, rowEnd).map((r) => (isNum(r[j]) ? r[j] : null))
    });
  }
  return { categories, series };
}

function niceTicks(min, max, count = 5) {
  // always ends: bad input gets a 0..1 axis, and values too close for doubles to step between are widened like equal ones
  if (!Number.isFinite(min) || !Number.isFinite(max)) return [0, 1];
  if (min > max) [min, max] = [max, min];
  const big = Math.max(Math.abs(min), Math.abs(max));
  if (max - min <= big * 1e-9) { if (big === 0) return [0, 1]; const d = big * 0.1; min -= d; max += d; }
  const span = max - min;
  const step0 = span / count;
  const mag = Math.pow(10, Math.floor(Math.log10(step0)));
  const err = step0 / mag;
  const step = (err >= 7.5 ? 10 : err >= 3.5 ? 5 : err >= 1.5 ? 2 : 1) * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  // steps that overflow (near the largest double) or vanish (tiny values): keep just the two ends
  if (!(step > 0) || !Number.isFinite(lo) || !Number.isFinite(hi + step / 2)) return Number.isFinite(span) ? [min, max] : [0, 1];
  const out = [];
  for (let v = lo; v <= hi + step / 2 && out.length < 50; v += step) out.push(+v.toPrecision(12));
  return out;
}
// One number format per locale and digit count: making a new one for every point is slow on big charts.
const formats = new Map();
function numFormat(locale, digits) {
  const key = String(locale) + '|' + digits;
  let f = formats.get(key);
  if (!f) {
    if (formats.size > 20) formats.clear();
    f = new Intl.NumberFormat(locale, { maximumFractionDigits: digits });
    formats.set(key, f);
  }
  return f;
}
function short(n, locale) {
  const a = Math.abs(n);
  try {
    if (a >= 1e9) return numFormat(locale, 1).format(n / 1e9) + 'B';
    if (a >= 1e6) return numFormat(locale, 1).format(n / 1e6) + 'M';
    if (a >= 1e4) return numFormat(locale, 1).format(n / 1e3) + 'K';
    return numFormat(locale, 2).format(n);
  } catch (e) { return String(n); }
}
const trunc = (s, n) => (String(s).length > n ? String(s).slice(0, n - 1) + '…' : String(s));
// Smallest and largest of a list, without spreading a long list into call arguments.
function extent(list, min, max) {
  for (const v of list) { if (v < min) min = v; if (v > max) max = v; }
  return [min, max];
}
// The data as drawn: arrays only, finite numbers or null, and no more than CHART_LIMITS.
function cleanData(data) {
  const src = data && Array.isArray(data.series) ? data.series.slice(0, CHART_LIMITS.series) : [];
  const categories = data && Array.isArray(data.categories) ? data.categories.slice(0, pointLimit(src.length)) : [];
  const series = src.map((s) => {
    const vals = s && Array.isArray(s.values) ? s.values : [];
    return { name: (s && s.name) ?? '', values: categories.map((_, i) => (typeof vals[i] === 'number' && Number.isFinite(vals[i]) ? vals[i] : null)) };
  });
  return { categories, series };
}

export function renderChart(spec, data, w, h, locale) {
  // stored chart objects may be malformed (older data, other clients): bad values count as absent
  if (!spec || typeof spec !== 'object') spec = {};
  w = Number.isFinite(w) ? Math.max(0, Math.min(20000, w)) : 480;
  h = Number.isFinite(h) ? Math.max(0, Math.min(20000, h)) : 280;
  const type = Object.keys(CHART_TYPES).includes(spec.type) ? spec.type : 'column';
  const title = typeof spec.title === 'string' ? spec.title : '';
  const { categories, series } = cleanData(data);
  // stored colors are used only when they are #rrggbb strings, otherwise the palette
  const own = Array.isArray(spec.colors) ? spec.colors.slice(0, 64).filter((c) => typeof c === 'string' && HEX_RE.test(c)) : [];
  const colors = own.length ? own : PALETTE;
  const colorAt = (i) => paint(colors[i % colors.length]);
  const parts = [];
  const titleH = title ? 26 : 6;
  if (title) parts.push(`<text x="${w / 2}" y="18" text-anchor="middle" class="ch-title">${esc(trunc(title, 80))}</text>`);
  if (!series.length || !categories.length) {
    parts.push(`<text x="${w / 2}" y="${h / 2}" text-anchor="middle" class="ch-muted">Select a range with numbers to chart</text>`);
    return svg(w, h, parts);
  }
  const showLegend = spec.legend !== false && (series.length > 1 || type === 'pie' || type === 'doughnut');

  if (type === 'pie' || type === 'doughnut') {
    const vals = series[0].values.map((v) => Math.max(0, v || 0));
    const total = vals.reduce((s, v) => s + v, 0) || 1;
    const legendW = showLegend ? Math.min(180, w * 0.38) : 0;
    const cx = (w - legendW) / 2, cy = titleH + (h - titleH) / 2, r = Math.max(10, Math.min(w - legendW, h - titleH) / 2 - 12);
    const inner = type === 'doughnut' ? r * 0.55 : 0;
    let a0 = -Math.PI / 2;
    const single = vals.filter(Boolean).length === 1;
    vals.forEach((v, i) => {
      if (!v) return;
      const a1 = a0 + (v / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const p = (rad, a) => `${cx + rad * Math.cos(a)},${cy + rad * Math.sin(a)}`;
      const d = single
        ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` + (inner ? `M${cx - inner},${cy}a${inner},${inner} 0 1,1 ${2 * inner},0a${inner},${inner} 0 1,1 ${-2 * inner},0` : '')
        : inner
          ? `M${p(r, a0)}A${r},${r} 0 ${large} 1 ${p(r, a1)}L${p(inner, a1)}A${inner},${inner} 0 ${large} 0 ${p(inner, a0)}Z`
          : `M${cx},${cy}L${p(r, a0)}A${r},${r} 0 ${large} 1 ${p(r, a1)}Z`;
      const pct = Math.round((v / total) * 1000) / 10;
      parts.push(`<path d="${d}" fill="${esc(colorAt(i))}" class="ch-slice" fill-rule="evenodd"><title>${esc(trunc(categories[i], 100))}: ${esc(short(v, locale))} (${pct}%)</title></path>`);
      if (pct >= 6 && r > 50) {
        const mid = (a0 + a1) / 2, lr = inner ? (r + inner) / 2 : r * 0.65;
        parts.push(`<text x="${cx + lr * Math.cos(mid)}" y="${cy + lr * Math.sin(mid) + 4}" text-anchor="middle" class="ch-onbar">${pct}%</text>`);
      }
      a0 = a1;
    });
    if (showLegend) legend(parts, categories.map((c, i) => ({ name: c, color: colorAt(i) })), w - legendW + 8, titleH + 8, legendW - 12, h - titleH);
    return svg(w, h, parts);
  }

  const legendH = showLegend ? 22 : 0;
  const all = [];
  series.forEach((s) => s.values.forEach((v) => { if (v !== null) all.push(v); }));
  let [min, max] = extent(all, 0, 0);
  if (type === 'stacked') {
    max = 0; min = 0;
    categories.forEach((_, i) => {
      let pos = 0, neg = 0;
      series.forEach((s) => { const v = s.values[i] || 0; if (v >= 0) pos += v; else neg += v; });
      max = Math.max(max, pos); min = Math.min(min, neg);
    });
  }
  if (type === 'scatter') [min, max] = extent(all, Infinity, -Infinity);
  const ticks = niceTicks(min, max);
  const tmin = ticks[0], tmax = ticks[ticks.length - 1];
  const horizontal = type === 'bar';
  const labelW = horizontal ? Math.min(120, 8 + categories.reduce((m, c) => Math.max(m, String(c).length), 0) * 6.5) : Math.max(...ticks.map((t) => short(t, locale).length)) * 7 + 10;
  const x0 = 8 + labelW, x1 = w - 12, y0 = titleH + 4, y1 = h - 22 - legendH;
  const plotW = Math.max(10, x1 - x0), plotH = Math.max(10, y1 - y0);
  const val = (v) => (horizontal ? x0 + ((v - tmin) / (tmax - tmin)) * plotW : y1 - ((v - tmin) / (tmax - tmin)) * plotH);

  // grid and value axis
  ticks.forEach((t) => {
    if (horizontal) {
      const x = val(t);
      parts.push(`<line x1="${x}" y1="${y0}" x2="${x}" y2="${y1}" class="ch-grid"/>`);
      parts.push(`<text x="${x}" y="${y1 + 15}" text-anchor="middle" class="ch-axis">${esc(short(t, locale))}</text>`);
    } else {
      const y = val(t);
      parts.push(`<line x1="${x0}" y1="${y}" x2="${x1}" y2="${y}" class="ch-grid"/>`);
      parts.push(`<text x="${x0 - 6}" y="${y + 4}" text-anchor="end" class="ch-axis">${esc(short(t, locale))}</text>`);
    }
  });
  const n = categories.length;
  const band = (horizontal ? plotH : plotW) / n;
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? plotH : plotW) / (horizontal ? 18 : 60)))));
  categories.forEach((c, i) => {
    if (i % every) return;
    if (horizontal) parts.push(`<text x="${x0 - 6}" y="${y0 + band * (i + 0.5) + 4}" text-anchor="end" class="ch-axis">${esc(trunc(c, 18))}</text>`);
    else if (type !== 'scatter') parts.push(`<text x="${x0 + band * (i + 0.5)}" y="${y1 + 15}" text-anchor="middle" class="ch-axis">${esc(trunc(c, Math.max(4, Math.floor(band * every / 7))))}</text>`);
  });
  const zero = val(Math.max(tmin, Math.min(tmax, 0)));

  if (type === 'column' || type === 'bar') {
    const groupW = band * 0.75, bw = groupW / series.length;
    series.forEach((s, si) => s.values.forEach((v, i) => {
      if (v === null) return;
      const pos = (horizontal ? y0 : x0) + band * i + (band - groupW) / 2 + bw * si;
      const a = val(v);
      const tip = `<title>${esc(trunc(s.name, 100))} · ${esc(trunc(categories[i], 100))}: ${esc(short(v, locale))}</title>`;
      if (horizontal) parts.push(`<rect x="${Math.min(a, zero)}" y="${pos}" width="${Math.max(1, Math.abs(a - zero))}" height="${Math.max(1, bw - 1)}" fill="${esc(colorAt(si))}" rx="1">${tip}</rect>`);
      else parts.push(`<rect x="${pos}" y="${Math.min(a, zero)}" width="${Math.max(1, bw - 1)}" height="${Math.max(1, Math.abs(a - zero))}" fill="${esc(colorAt(si))}" rx="1">${tip}</rect>`);
    }));
  } else if (type === 'stacked') {
    const bw = band * 0.6;
    categories.forEach((c, i) => {
      let pos = 0, neg = 0;
      series.forEach((s, si) => {
        const v = s.values[i];
        if (!v) return;
        const from = v >= 0 ? pos : neg;
        const to = from + v;
        if (v >= 0) pos = to; else neg = to;
        const ya = val(from), yb = val(to);
        parts.push(`<rect x="${x0 + band * i + (band - bw) / 2}" y="${Math.min(ya, yb)}" width="${bw}" height="${Math.max(1, Math.abs(ya - yb))}" fill="${esc(colorAt(si))}"><title>${esc(trunc(s.name, 100))} · ${esc(trunc(c, 100))}: ${esc(short(v, locale))}</title></rect>`);
      });
    });
  } else if (type === 'scatter') {
    // first series = x values, the others = y values
    const xs = series.length > 1 ? series[0].values : categories.map((_, i) => i + 1);
    const ys = series.length > 1 ? series.slice(1) : series;
    const xv = xs.filter((v) => v !== null);
    const xt = niceTicks(...extent(xv, Infinity, -Infinity));
    const xmin = xt[0], xmax = xt[xt.length - 1];
    const px = (v) => x0 + ((v - xmin) / (xmax - xmin || 1)) * plotW;
    xt.forEach((t) => parts.push(`<text x="${px(t)}" y="${y1 + 15}" text-anchor="middle" class="ch-axis">${esc(short(t, locale))}</text>`));
    ys.forEach((s, si) => s.values.forEach((v, i) => {
      if (v === null || xs[i] === null) return;
      parts.push(`<circle cx="${px(xs[i])}" cy="${val(v)}" r="4" fill="${esc(colorAt(si))}"><title>${esc(trunc(s.name, 100))}: (${esc(short(xs[i], locale))}, ${esc(short(v, locale))})</title></circle>`);
    }));
  } else {
    // line / area
    series.forEach((s, si) => {
      const pts = s.values.map((v, i) => (v === null ? null : [x0 + band * (i + 0.5), val(v)]));
      const segs = [];
      let cur = [];
      pts.forEach((p) => { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } });
      if (cur.length) segs.push(cur);
      const color = colorAt(si);
      segs.forEach((seg) => {
        const d = seg.map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1)).join('');
        if (type === 'area') parts.push(`<path d="${d}L${seg[seg.length - 1][0]},${zero}L${seg[0][0]},${zero}Z" fill="${esc(color)}" fill-opacity="0.18"/>`);
        parts.push(`<path d="${d}" fill="none" stroke="${esc(color)}" stroke-width="2" stroke-linejoin="round"/>`);
      });
      if (n <= 40) pts.forEach((p, i) => { if (p) parts.push(`<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${esc(color)}"><title>${esc(trunc(s.name, 100))} · ${esc(trunc(categories[i], 100))}: ${esc(short(s.values[i], locale))}</title></circle>`); });
    });
  }
  parts.push(horizontal ? `<line x1="${zero}" y1="${y0}" x2="${zero}" y2="${y1}" class="ch-zero"/>` : `<line x1="${x0}" y1="${zero}" x2="${x1}" y2="${zero}" class="ch-zero"/>`);
  if (showLegend) legend(parts, (type === 'scatter' && series.length > 1 ? series.slice(1) : series).map((s, i) => ({ name: s.name, color: colorAt(i) })), x0, h - legendH + 2, plotW, legendH, true);
  return svg(w, h, parts);
}

function legend(parts, items, x, y, w, h, row) {
  if (row) {
    let cx = x;
    for (const it of items) {
      const label = trunc(it.name, 24);
      const iw = 18 + label.length * 6.5 + 12;
      if (cx + iw > x + w) break;
      parts.push(`<rect x="${cx}" y="${y + 4}" width="10" height="10" rx="2" fill="${esc(paint(it.color))}"/><text x="${cx + 14}" y="${y + 13}" class="ch-legend">${esc(label)}</text>`);
      cx += iw;
    }
    return;
  }
  items.slice(0, Math.floor(h / 18)).forEach((it, i) => {
    parts.push(`<rect x="${x}" y="${y + i * 18}" width="10" height="10" rx="2" fill="${esc(paint(it.color))}"/><text x="${x + 14}" y="${y + i * 18 + 9}" class="ch-legend">${esc(trunc(it.name, Math.max(6, Math.floor(w / 6.5) - 3)))}</text>`);
  });
}

function svg(w, h, parts) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" class="chart-svg" role="img">${parts.join('')}</svg>`;
}
