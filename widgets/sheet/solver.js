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
/* Spreadsheets for YouTrack – Solver engines (optimization), independent of the spreadsheet.
 *
 * A problem is given as numbers and a function:
 *   n            number of variables
 *   x0           starting values
 *   lb, ub       bounds (±Infinity when absent)
 *   isInt        integer variables (binary = integer with bounds 0..1)
 *   difGroups    groups of variables that must take all different integer values 1..N (N = group size)
 *   sense        1 to minimize, -1 to maximize
 *   ctype        per general constraint: 'le' (g ≤ 0), 'ge' (g ≥ 0), 'eq' (g = 0), with g = left − right
 *   evaluate(x)  → { f, g: [...] } (NaN or non-finite values mean an error value in a cell)
 *
 * Engines: simplexLP (with branch and bound for integers), grg (augmented Lagrangian with projected quasi-Newton
 * steps, optional multistart), evolutionary (genetic algorithm with local search). All are async so the page stays
 * responsive; hooks.shouldStop() stops them and hooks.onIteration(info) can pause them ("Show Iteration Results").
 */

/** Solver options. As in Excel, population size, random seed, convergence and "require bounds" are shared by the
 * GRG Multistart and Evolutionary settings. 0 means "no limit" for the limits. */
export const DEFAULT_OPTIONS = {
  precision: 1e-6,        // constraint precision
  scaling: true,          // use automatic scaling
  showIter: false,        // show iteration results
  ignoreInt: false,       // ignore integer constraints
  intTolerance: 1,        // integer optimality (%)
  maxTime: 0,             // seconds
  iterations: 0,
  maxSubproblems: 0,
  maxFeasible: 0,
  convergence: 1e-4,
  derivatives: 'forward', // 'forward' | 'central' (GRG)
  multistart: false,      // GRG
  population: 100,        // GRG Multistart and Evolutionary
  seed: 0,                // 0 = a different random seed each time
  requireBounds: true,    // GRG Multistart and Evolutionary
  mutation: 0.075,        // Evolutionary
  maxNoImprove: 30        // Evolutionary: seconds without improvement
};

export const MESSAGES = {
  found: 'Solver found a solution. All Constraints and optimality conditions are satisfied.',
  converged: 'Solver has converged to the current solution. All Constraints are satisfied.',
  noImprove: 'Solver cannot improve the current solution. All Constraints are satisfied.',
  intTol: 'Solver found an integer solution within tolerance. All Constraints are satisfied.',
  global: 'Solver converged in probability to a global solution. All Constraints are satisfied.',
  infeasible: 'Solver could not find a feasible solution.',
  unbounded: 'The Objective Cell values do not converge.',
  linearity: 'The linearity conditions required by this LP Solver are not satisfied.',
  stopped: 'Solver stopped at user\'s request.',
  time: 'Solver has reached the maximum time limit.',
  iterations: 'Solver has reached the maximum iteration limit.',
  subproblems: 'Stop chosen when the maximum number of subproblems was reached.',
  feasibleLimit: 'Stop chosen when the maximum number of feasible (integer) solutions was reached.',
  error: 'Solver encountered an error value in the Objective Cell or a Constraint cell.',
  bounds: 'All Variable Cells must have both upper and lower bounds when the Require Bounds option is set.',
  tooLarge: 'The problem is too large for Solver to handle.'
};

const BIG = 1e30;
const finite = (v) => typeof v === 'number' && isFinite(v);

/** Seeded random numbers (mulberry32); seed 0 picks a random seed. */
export function rng(seed) {
  let a = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Shared run state: time limit, iteration count, stop requests, yielding to the page. */
class Run {
  constructor(opts, hooks) {
    this.opts = opts;
    this.hooks = hooks || {};
    this.start = Date.now();
    this.paused = 0;            // ms spent in pause dialogs (not counted as solving time)
    this.iterations = 0;
    this.subproblems = 0;
    this.lastYield = Date.now();
    this.stopReason = null;
    this.timeLimit = opts.maxTime > 0 ? opts.maxTime : 0;
    this.iterLimit = opts.iterations > 0 ? opts.iterations : 0;
  }
  /** Clock without pauses, in ms. */
  now() { return Date.now() - this.paused; }
  get elapsed() { return (this.now() - this.start) / 1000; }
  hardLimit() {
    if (this.timeLimit && this.elapsed > this.timeLimit) return 'time';
    if (this.iterLimit && this.iterations >= this.iterLimit) return 'iterations';
    return null;
  }
  limitReached() {
    if (this.stopReason) return this.stopReason;
    if (this.hooks.shouldStop && this.hooks.shouldStop()) this.stopReason = 'stopped';
    else if (!this.hooks.onLimit) this.stopReason = this.hardLimit();
    return this.stopReason;
  }
  async pause(fn) {
    const t = Date.now();
    try { return await fn(); } finally { this.paused += Date.now() - t; this.lastYield = Date.now(); }
  }
  /**
   * Called once per iteration. Lets the page paint every ~40 ms, pauses when "Show Iteration Results" is on and,
   * when a time or iteration limit is reached, asks whether to continue (hooks.onLimit), as Excel does.
   */
  async iteration(info) {
    this.iterations++;
    const lim = !this.stopReason && this.hardLimit();
    if (lim && this.hooks.onLimit) {
      const go = await this.pause(() => this.hooks.onLimit(lim, { ...info, iteration: this.iterations }));
      if (!go) this.stopReason = lim;
      else if (lim === 'time') this.timeLimit = this.elapsed + this.opts.maxTime;
      else this.iterLimit = this.iterations + this.opts.iterations;
      return;
    }
    if (this.opts.showIter && this.hooks.onIteration && !this.stopReason) {
      const go = await this.pause(() => this.hooks.onIteration({ ...info, iteration: this.iterations }));
      if (!go) this.stopReason = 'stopped';
      return;
    }
    if (Date.now() - this.lastYield > 40) {
      if (this.hooks.onProgress) this.hooks.onProgress({ ...info, iteration: this.iterations, elapsed: this.elapsed });
      await new Promise((r) => setTimeout(r, 0));
      this.lastYield = Date.now();
    }
  }
}

/* ================================================================ evaluation helpers */

/**
 * Largest constraint violation of an evaluation, relative to the size of each right side (max(1, |right|)), so that
 * "Constraint Precision" means the same for a budget of 1,000,000 as for a share of 0.5.
 */
function violation(P, ev) {
  const g = ev.g, lhs = ev.lhs, rhs = ev.rhs;
  let v = 0;
  P.ctype.forEach((t, k) => {
    const x = g[k];
    if (!finite(x)) { v = Infinity; return; }
    const right = rhs && finite(rhs[k]) ? rhs[k] : lhs && finite(lhs[k]) ? lhs[k] - x : 0;
    const d = t === 'le' ? x : t === 'ge' ? -x : Math.abs(x);
    v = Math.max(v, d / Math.max(1, Math.abs(right)));
  });
  return v;
}
/**
 * Change of constraint k between two evaluations. When both sides are known, each side is differenced on its own:
 * a constant right side then drops out exactly, instead of rounding the difference to the last digit of a large
 * right side (1e8 and more).
 */
function gDelta(e, base, k) {
  return e.lhs && e.rhs && base.lhs && base.rhs ? (e.lhs[k] - base.lhs[k]) - (e.rhs[k] - base.rhs[k]) : e.g[k] - base.g[k];
}
const bad = (e) => !finite(e.f) || e.g.some((v) => !finite(v));
/** Bounds that no value can meet (lb > ub, or an integer variable with no whole number between its bounds). */
const conflictingBounds = (lb, ub, isInt) => lb.some((v, j) => v > ub[j] + 1e-12 || (isInt[j] && Math.ceil(v - 1e-9) > Math.floor(ub[j] + 1e-9)));
function boundViolation(x, lb, ub) {
  let v = 0;
  for (let j = 0; j < x.length; j++) v = Math.max(v, lb[j] - x[j], x[j] - ub[j]);
  return Math.max(0, v);
}
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const project = (x, lb, ub) => x.map((v, j) => clamp(v, lb[j], ub[j]));

/* ================================================================ linear model extraction */

/**
 * Reads the linear form of the problem: f(x) = c·x + c0, g_k(x) = a_k·x + b_k. Also tests linearity at extra points.
 * Returns { c, c0, A, b, linear, nonlinearF, nonlinearG: [bool] }.
 */
export function linearize(P, x0) {
  const n = P.n;
  const base = P.evaluate(x0);
  if (bad(base)) return { error: 'error' };
  const sides = !!(base.lhs && base.rhs);
  const c = new Array(n), A = P.ctype.map(() => new Array(n));
  for (let j = 0; j < n; j++) {
    const x = x0.slice();
    x[j] += 1;
    const e = P.evaluate(x);
    if (bad(e)) return { error: 'error' };
    c[j] = e.f - base.f;
    e.g.forEach((v, k) => { A[k][j] = gDelta(e, base, k); });
  }
  const c0 = base.f - dot(c, x0);
  const b = base.g.map((v, k) => (sides ? (base.lhs[k] - dot(A[k], x0)) - base.rhs[k] : v - dot(A[k], x0)));
  // linearity test at two other points
  const nonlinearG = P.ctype.map(() => false);
  let nonlinearF = false;
  const R = rng(12345);
  for (let t = 0; t < 2; t++) {
    const x = x0.map((v) => v + (R() * 4 - 2) * Math.max(1, Math.abs(v)));
    const e = P.evaluate(x);
    const tol = (pred, got, size = 0) => Math.abs(pred - got) > 1e-6 * Math.max(1, Math.abs(pred), Math.abs(got), size);
    if (!finite(e.f) || tol(dot(c, x) + c0, e.f)) nonlinearF = true;
    e.g.forEach((v, k) => {
      const size = sides && e.lhs ? Math.max(Math.abs(e.lhs[k]), Math.abs(e.rhs[k])) : 0;
      if (!finite(v) || tol(dot(A[k], x) + b[k], v, size)) nonlinearG[k] = true;
    });
  }
  return { c, c0, A, b, linear: !nonlinearF && !nonlinearG.some(Boolean), nonlinearF, nonlinearG };
}
function dot(a, b) { let s = 0; for (let i = 0; i < a.length; i++) s += a[i] * b[i]; return s; }

/* ================================================================ dense two-phase simplex */

/**
 * minimize cost·y  subject to  rows (a·y  le|ge|eq  b),  y ≥ 0.
 * Returns { status: 'optimal'|'infeasible'|'unbounded'|'stopped', y, obj, duals (per row, ∂obj/∂b), reduced (per y),
 *           ranging: { rhs: [[dec, inc]], cost: [[dec, inc]] } }.
 */
async function simplexCore(cost, rows, run, twins = new Map(), tol = 1e-9) {
  const ny = cost.length, m = rows.length;
  // flip rows so that b ≥ 0
  const R = rows.map((r) => (r.b < 0 ? { a: r.a.map((v) => -v), b: -r.b, rel: r.rel === 'le' ? 'ge' : r.rel === 'ge' ? 'le' : 'eq', flip: -1 } : { a: r.a.slice(), b: r.b, rel: r.rel, flip: 1 }));
  // columns: y (ny) | slack/surplus (one per le/ge row) | artificial (per ge/eq row)
  const slackCol = new Array(m).fill(-1), artCol = new Array(m).fill(-1);
  let ncol = ny;
  R.forEach((r, i) => { if (r.rel !== 'eq') slackCol[i] = ncol++; });
  R.forEach((r, i) => { if (r.rel !== 'le') artCol[i] = ncol++; });
  const T = R.map((r, i) => {
    const row = new Float64Array(ncol + 1);
    for (let j = 0; j < ny; j++) row[j] = r.a[j];
    if (slackCol[i] >= 0) row[slackCol[i]] = r.rel === 'le' ? 1 : -1;
    if (artCol[i] >= 0) row[artCol[i]] = 1;
    row[ncol] = r.b;
    return row;
  });
  const basis = R.map((r, i) => (r.rel === 'le' ? slackCol[i] : artCol[i]));
  const initCol = basis.slice(); // column i of B^-1 lives in column initCol[i]
  const isArt = new Uint8Array(ncol);
  const artRow = new Map();
  artCol.forEach((c, i) => { if (c >= 0) { isArt[c] = 1; artRow.set(c, i); } });
  const redundant = new Uint8Array(m); // rows that are combinations of others (their artificial stays basic at 0)

  async function optimize(costVec, blockArt) {
    // reduced costs d_j = c_j − c_B B^-1 A_j, kept as a row
    const d = new Float64Array(ncol + 1);
    for (let j = 0; j < ncol; j++) d[j] = costVec[j] || 0;
    for (let i = 0; i < m; i++) {
      const cb = costVec[basis[i]] || 0;
      if (cb) for (let j = 0; j <= ncol; j++) d[j] -= cb * T[i][j];
    }
    let degenerate = 0;
    for (;;) {
      if (run.limitReached()) return 'stopped';
      // entering column: most negative reduced cost (Bland's rule after degenerate steps)
      let enter = -1, best = -1e-9;
      for (let j = 0; j < ncol; j++) {
        if (blockArt && isArt[j]) continue;
        if (d[j] < best) {
          if (degenerate > 50) { enter = j; break; }
          best = d[j]; enter = j;
        }
      }
      if (enter < 0) return 'optimal';
      let leave = -1, ratio = Infinity;
      for (let i = 0; i < m; i++) {
        if (redundant[i]) continue;
        const a = T[i][enter];
        if (a > 1e-9) {
          const r = T[i][ncol] / a;
          if (r < ratio - 1e-12 || (Math.abs(r - ratio) <= 1e-12 && basis[i] < basis[leave])) { ratio = r; leave = i; }
        }
      }
      if (leave < 0) return 'unbounded';
      degenerate = ratio < 1e-12 ? degenerate + 1 : 0;
      pivot(leave, enter, d);
      await run.iteration({ phase: blockArt ? 2 : 1 });
    }
  }
  function pivot(r, c, d) {
    const pr = T[r], pv = pr[c];
    for (let j = 0; j <= ncol; j++) pr[j] /= pv;
    for (let i = 0; i < m; i++) {
      if (i === r) continue;
      const f = T[i][c];
      if (f) { const row = T[i]; for (let j = 0; j <= ncol; j++) row[j] -= f * pr[j]; }
    }
    const f = d[c];
    if (f) for (let j = 0; j <= ncol; j++) d[j] -= f * pr[j];
    basis[r] = c;
  }

  // phase 1
  if (artCol.some((c) => c >= 0)) {
    const c1 = new Float64Array(ncol);
    artCol.forEach((c) => { if (c >= 0) c1[c] = 1; });
    const st = await optimize(c1, false);
    if (st === 'stopped') return { status: 'stopped' };
    // every artificial still in the basis must be zero (within the constraint precision of its own row)
    for (let i = 0; i < m; i++) {
      if (!isArt[basis[i]]) continue;
      const scale = Math.max(1, Math.abs(R[artRow.get(basis[i])].b));
      if (T[i][ncol] > Math.max(tol, 1e-9) * scale) return { status: 'infeasible' };
    }
    // drive the (zero) artificials out of the basis: a degenerate pivot on the largest entry keeps every value;
    // a row whose entries are all round-off compared with its original coefficients repeats other rows
    for (let i = 0; i < m; i++) {
      if (!isArt[basis[i]]) continue;
      T[i][ncol] = 0;
      const origin = R[artRow.get(basis[i])];
      const rowScale = Math.max(1e-300, ...origin.a.map(Math.abs));
      let c = -1, best = 0;
      for (let j = 0; j < ncol; j++) {
        if (isArt[j]) continue;
        const a = Math.abs(T[i][j]);
        if (a > best) { best = a; c = j; }
      }
      if (c >= 0 && best > 1e-10 * Math.max(1, rowScale)) pivot(i, c, new Float64Array(ncol + 1));
      else redundant[i] = 1;
    }
  }
  // phase 2
  const c2 = new Float64Array(ncol);
  for (let j = 0; j < ny; j++) c2[j] = cost[j];
  const st = await optimize(c2, true);
  if (st !== 'optimal') return { status: st };

  const y = new Array(ny).fill(0);
  for (let i = 0; i < m; i++) if (basis[i] < ny) y[basis[i]] = T[i][ncol];
  const obj = dot(cost, y);
  // duals: π = c_B B^-1  (B^-1 column i = tableau column initCol[i])
  const cB = basis.map((b) => (b < ny ? cost[b] : 0));
  const pi = new Array(m).fill(0);
  for (let i = 0; i < m; i++) { let s = 0; for (let r = 0; r < m; r++) s += cB[r] * T[r][initCol[i]]; pi[i] = s; }
  // for a ≥ row the artificial column is +e_i, so the same formula holds; flipped rows change sign
  const duals = pi.map((p, i) => p * R[i].flip);
  const reduced = new Array(ny);
  for (let j = 0; j < ny; j++) { let s = cost[j]; for (let i = 0; i < m; i++) s -= pi[i] * R[i].a[j]; reduced[j] = s; }
  // RHS ranging: x_B(Δ) = x_B + Δ·B^-1 e_i ≥ 0
  const rhs = R.map((r, i) => {
    let inc = BIG, dec = BIG;
    for (let k = 0; k < m; k++) {
      const col = T[k][initCol[i]], xb = T[k][ncol];
      // an artificial left in the basis (a repeated row) must stay exactly 0: any change that moves it is infeasible
      if (isArt[basis[k]] && Math.abs(col) > 1e-11) { dec = 0; inc = 0; break; }
      if (col > 1e-11) dec = Math.min(dec, xb / col);
      else if (col < -1e-11) inc = Math.min(inc, -xb / col);
    }
    return r.flip > 0 ? [dec, inc] : [inc, dec];
  });
  // cost ranging (minimization): reduced costs of nonbasic columns must stay ≥ 0
  const rowOf = new Map(basis.map((b, i) => [b, i]));
  const dvec = new Float64Array(ncol);
  for (let j = 0; j < ncol; j++) {
    if (isArt[j]) { dvec[j] = Infinity; continue; }
    let s = j < ny ? cost[j] : 0;
    for (let i = 0; i < m; i++) s -= cB[i] * T[i][j];
    dvec[j] = s;
  }
  const costRange = cost.map((cj, j) => {
    // nonbasic at its lower bound: the cost may rise freely and fall by the reduced cost before the variable enters
    if (!rowOf.has(j)) return [Math.max(0, reduced[j]), BIG];
    const r = rowOf.get(j);
    let inc = BIG, dec = BIG;
    for (let k = 0; k < ncol; k++) {
      if (rowOf.has(k) || isArt[k] || twins.get(j) === k) continue;
      const alpha = T[r][k];
      if (alpha > 1e-11) inc = Math.min(inc, dvec[k] / alpha);
      else if (alpha < -1e-11) dec = Math.min(dec, -dvec[k] / alpha);
    }
    return [dec, inc];
  });
  return { status: 'optimal', y, obj, duals, reduced, basic: new Set(basis), ranging: { rhs, cost: costRange } };
}

/**
 * Bounded LP in the user's variables: minimize c·x subject to A x (rel) −b … i.e. a_k·x + b_k (le|ge|eq) 0, lb ≤ x ≤ ub.
 * Maps x to y ≥ 0 (shift, mirror or split) and back.
 */
async function solveLP(lin, ctype, lb, ub, run, tol) {
  const n = lin.c.length;
  const map = []; // per x: { kind: 'shift'|'mirror'|'split', col, col2, base }
  let ny = 0;
  for (let j = 0; j < n; j++) {
    if (isFinite(lb[j])) map.push({ kind: 'shift', col: ny++, base: lb[j] });
    else if (isFinite(ub[j])) map.push({ kind: 'mirror', col: ny++, base: ub[j] });
    else { map.push({ kind: 'split', col: ny, col2: ny + 1, base: 0 }); ny += 2; }
  }
  const toY = (coef) => {
    // coefficient vector over x → over y, plus constant from the shifts
    const a = new Array(ny).fill(0); let k0 = 0;
    coef.forEach((v, j) => {
      const mp = map[j];
      if (mp.kind === 'shift') { a[mp.col] += v; k0 += v * mp.base; }
      else if (mp.kind === 'mirror') { a[mp.col] -= v; k0 += v * mp.base; }
      else { a[mp.col] += v; a[mp.col2] -= v; }
    });
    return { a, k0 };
  };
  const rows = [];
  lin.A.forEach((row, k) => {
    const { a, k0 } = toY(row);
    rows.push({ a, rel: ctype[k], b: -(lin.b[k] + k0), src: { kind: 'g', k } });
  });
  for (let j = 0; j < n; j++) {
    if (isFinite(lb[j]) && isFinite(ub[j])) {
      const a = new Array(ny).fill(0); a[map[j].col] = 1;
      rows.push({ a, rel: 'le', b: ub[j] - lb[j], src: { kind: 'ub', j } });
    }
  }
  const { a: cy, k0: cconst } = toY(lin.c);
  const twins = new Map();
  map.forEach((mp) => { if (mp.kind === 'split') { twins.set(mp.col, mp.col2); twins.set(mp.col2, mp.col); } });
  const res = await simplexCore(cy, rows, run, twins, tol);
  if (res.status !== 'optimal') return { status: res.status };
  const x = map.map((mp) => (mp.kind === 'shift' ? mp.base + res.y[mp.col] : mp.kind === 'mirror' ? mp.base - res.y[mp.col] : res.y[mp.col] - res.y[mp.col2]));
  const F = res.obj + cconst + lin.c0;
  // sensitivity in x terms (minimization of F)
  const gDuals = new Array(ctype.length).fill(0), gRange = new Array(ctype.length).fill([BIG, BIG]);
  const ubDual = new Array(n).fill(0);
  rows.forEach((r, i) => {
    if (r.src.kind === 'g') { gDuals[r.src.k] = res.duals[i]; gRange[r.src.k] = res.ranging.rhs[i]; }
    else ubDual[r.src.j] = res.duals[i];
  });
  // bounds are not listed as constraints: their shadow price shows up in the variable's reduced cost
  const reduced = map.map((mp, j) => {
    const d = res.reduced[mp.col] + ubDual[j];
    return mp.kind === 'mirror' ? -d : d;
  });
  const isBasic = (col) => res.basic.has(col);
  const costRange = map.map((mp) => {
    if (mp.kind === 'shift') return res.ranging.cost[mp.col];
    if (mp.kind === 'mirror') return res.ranging.cost[mp.col].slice().reverse();
    // free variable x = y1 − y2: the basic one of the pair carries the ranging (y2's cost is −c, so it is reversed)
    if (isBasic(mp.col)) return res.ranging.cost[mp.col];
    if (isBasic(mp.col2)) return res.ranging.cost[mp.col2].slice().reverse();
    return [Math.max(0, res.reduced[mp.col]), Math.max(0, res.reduced[mp.col2])];
  });
  return { status: 'optimal', x, F, gDuals, gRange, reduced, costRange };
}

/* ================================================================ branch and bound (integers, binary, alldifferent) */

function intViolation(x, isInt, difGroups, tol = 1e-6) {
  let best = -1, frac = 0;
  for (let j = 0; j < x.length; j++) {
    if (!isInt[j]) continue;
    const f = Math.abs(x[j] - Math.round(x[j]));
    if (f > tol && f > frac) { frac = f; best = j; }
  }
  if (best >= 0) return { kind: 'int', j: best, v: x[best] };
  for (const grp of difGroups) {
    const seen = new Map();
    for (const j of grp) {
      const v = Math.round(x[j]);
      if (seen.has(v)) return { kind: 'dif', i: seen.get(v), j, v };
      seen.set(v, j);
    }
  }
  return null;
}

const STOP_STATUSES = ['stopped', 'time', 'iterations'];
const LIMIT_STATUSES = STOP_STATUSES.concat(['subproblems', 'feasibleLimit']);

/** relax(lb, ub) → { status, x, F } minimizing F. */
async function branchAndBound(relax, lb0, ub0, isInt, difGroups, opts, run) {
  const stack = [{ lb: lb0.slice(), ub: ub0.slice() }];
  let best = null, feasibleCount = 0, rootStatus = null, limit = null;
  // Integer Optimality is relative to the incumbent, whatever the size of the objective; plus round-off
  const gap = (F) => (opts.intTolerance / 100) * Math.abs(F) + 1e-9 * Math.max(1, Math.abs(F));
  while (stack.length) {
    if ((limit = run.limitReached())) break;
    if (opts.maxSubproblems > 0 && run.subproblems >= opts.maxSubproblems) { limit = 'subproblems'; break; }
    if (opts.maxFeasible > 0 && feasibleCount >= opts.maxFeasible) { limit = 'feasibleLimit'; break; }
    const node = stack.pop();
    run.subproblems++;
    const r = await relax(node.lb, node.ub);
    if (rootStatus === null) rootStatus = r.status;
    // a stop or limit inside a subproblem stops the search; this node is not finished
    if (STOP_STATUSES.includes(r.status)) { limit = run.stopReason || r.status; break; }
    if (r.status === 'unbounded' && !best) return { status: 'unbounded' };
    if (r.status !== 'optimal') continue;
    if (best && r.F >= best.F - gap(best.F)) continue;
    const v = intViolation(r.x, isInt, difGroups);
    if (!v) { best = r; feasibleCount++; continue; }
    const child = (changes) => {
      const c = { lb: node.lb.slice(), ub: node.ub.slice() };
      for (const [j, lo, hi] of changes) { if (lo !== null) c.lb[j] = Math.max(c.lb[j], lo); if (hi !== null) c.ub[j] = Math.min(c.ub[j], hi); }
      return changes.every(([j]) => c.lb[j] <= c.ub[j]) ? c : null;
    };
    let kids;
    if (v.kind === 'int') {
      const down = child([[v.j, null, Math.floor(v.v)]]), up = child([[v.j, Math.ceil(v.v), null]]);
      kids = v.v - Math.floor(v.v) >= 0.5 ? [down, up] : [up, down]; // the last one is explored first
    } else {
      // x_i = x_j = v is not allowed: x_j < v, x_j > v, or x_j = v with x_i < v or x_i > v
      kids = [child([[v.j, null, v.v - 1]]), child([[v.j, v.v + 1, null]]),
        child([[v.j, v.v, v.v], [v.i, null, v.v - 1]]), child([[v.j, v.v, v.v], [v.i, v.v + 1, null]])];
    }
    kids.forEach((k) => { if (k) stack.push(k); });
  }
  if (!limit && run.stopReason) limit = run.stopReason;
  if (!best) return { status: limit || 'infeasible' };
  return { ...best, status: limit ? 'limit' : 'optimal', limit, withinTolerance: opts.intTolerance > 0 };
}

/* ================================================================ Simplex LP engine */

export async function simplexLP(P, opts, hooks) {
  const run = new Run(opts, hooks);
  if (conflictingBounds(P.lb, P.ub, opts.ignoreInt ? [] : P.isInt)) return result('infeasible', run);
  const lin = linearize(P, P.x0.map((v, j) => clamp(finite(v) ? v : 0, P.lb[j], P.ub[j])));
  if (lin.error) return result('error', run);
  if (!lin.linear) return { ...result('linearity', run), linearity: lin };
  const isInt = opts.ignoreInt ? P.isInt.map(() => false) : P.isInt;
  const difs = opts.ignoreInt ? [] : P.difGroups;
  const relax = async (lb, ub) => {
    const cc = lin.c.map((v) => v * P.sense);
    const r = await solveLP({ c: cc, c0: lin.c0 * P.sense, A: lin.A, b: lin.b }, P.ctype, lb, ub, run, opts.precision);
    return r.status === 'optimal' ? { ...r, status: 'optimal' } : r;
  };
  const hasInt = isInt.some(Boolean) || difs.length > 0;
  const r = hasInt ? await branchAndBound(relax, P.lb, P.ub, isInt, difs, opts, run) : await relax(P.lb, P.ub);
  if (r.status === 'unbounded') return result('unbounded', run);
  if (!r.x) return result(r.status === 'stopped' ? run.stopReason || 'stopped' : r.status, run);
  const status = r.limit ? r.limit : hasInt ? (opts.intTolerance > 0 ? 'intTol' : 'found') : 'found';
  return { ...result(status, run), x: r.x, feasible: true, hasInt, linear: lin,
    sensitivity: hasInt ? null : { kind: 'lp', reduced: r.reduced, gDuals: r.gDuals, gRange: r.gRange, costRange: r.costRange } };
}

/* ================================================================ GRG Nonlinear engine */

/**
 * Local solve: augmented Lagrangian; inner problems on the box use projected quasi-Newton (BFGS) steps with finite
 * difference gradients.
 */
async function grgLocal(P, lb, ub, x0, opts, run, iterLimit) {
  const n = P.n, m = P.ctype.length;
  let x = project(x0.map((v) => (finite(v) ? v : 0)), lb, ub);
  const e0 = P.evaluate(x);
  if (bad(e0)) return { status: 'error' };
  const sense = P.sense;
  const hstep = (v) => (opts.derivatives === 'central' ? 1e-5 : 1e-7) * Math.max(1, Math.abs(v));
  // Automatic scaling: the objective and each constraint are divided by how much they change over the size of the
  // variables, so that daily returns of 0.0003 and budgets of 1e8 are solved alike.
  let fScale = 1;
  const gScale = new Array(m).fill(1);
  if (opts.scaling) {
    const xs = Math.max(1, ...x.map(Math.abs));
    let fg = 0;
    const gg = new Array(m).fill(0);
    for (let j = 0; j < n; j++) {
      const h = hstep(x[j]), xp = x.slice(), up = x[j] + h <= ub[j] ? h : -h;
      xp[j] += up;
      const ep = P.evaluate(xp);
      if (finite(ep.f)) fg = Math.max(fg, Math.abs((ep.f - e0.f) / up));
      for (let k = 0; k < m; k++) { const d = gDelta(ep, e0, k); if (finite(d)) gg[k] = Math.max(gg[k], Math.abs(d / up)); }
    }
    const scaleOf = (grad, val) => {
      const size = grad * xs > 0 ? grad * xs : Math.abs(val);
      return size > 0 ? clamp(1 / size, 1e-8, 1e8) : 1;
    };
    fScale = scaleOf(fg, e0.f);
    for (let k = 0; k < m; k++) gScale[k] = scaleOf(gg[k], e0.g[k]);
  }
  let lam = new Array(m).fill(0), mu = 10;
  // canonical inequality form: h(x) ≤ 0 for le/ge
  const hOf = (g, k) => (P.ctype[k] === 'ge' ? -g[k] : g[k]) * gScale[k];
  function AL(ev) {
    if (bad(ev)) return Infinity;
    let s = sense * ev.f * fScale;
    for (let k = 0; k < m; k++) {
      const h = hOf(ev.g, k);
      if (P.ctype[k] === 'eq') s += lam[k] * h + 0.5 * mu * h * h;
      else { const t = Math.max(0, lam[k] + mu * h); s += (t * t - lam[k] * lam[k]) / (2 * mu); }
    }
    return s;
  }
  function gradient(xc, ev) {
    const g = new Array(n), base = AL(ev);
    const parts = []; // ∂f, ∂g per variable (for sensitivity)
    for (let j = 0; j < n; j++) {
      const h = hstep(xc[j]);
      const xp = xc.slice();
      const up = xc[j] + h <= ub[j] ? h : -h;
      xp[j] = xc[j] + up;
      const ep = P.evaluate(xp);
      if (opts.derivatives === 'central' && xc[j] - up >= lb[j] && xc[j] - up <= ub[j]) {
        const xm = xc.slice(); xm[j] = xc[j] - up;
        const em = P.evaluate(xm);
        g[j] = (AL(ep) - AL(em)) / (2 * up);
        parts.push({ df: (ep.f - em.f) / (2 * up), dg: ep.g.map((v, k) => gDelta(ep, em, k) / (2 * up)) });
      } else {
        g[j] = (AL(ep) - base) / up;
        parts.push({ df: (ep.f - ev.f) / up, dg: ep.g.map((v, k) => gDelta(ep, ev, k) / up) });
      }
      if (!finite(g[j])) g[j] = 0;
    }
    return { g, parts };
  }
  const tolG = 1e-7;
  const isFeasible = (ev, xx) => violation(P, ev) <= opts.precision && boundViolation(xx, lb, ub) <= 1e-9;
  let ev = e0, history = [], status = 'converged', lastParts = null, kkt = false, done = false;
  // the best feasible point seen, returned when the method ends somewhere infeasible
  let bestFeasible = isFeasible(ev, x) ? { x: x.slice(), ev } : null;
  const keepFeasible = () => {
    if (isFeasible(ev, x) && (!bestFeasible || sense * ev.f < sense * bestFeasible.ev.f)) bestFeasible = { x: x.slice(), ev };
  };
  for (let outer = 0; outer < 60 && !done; outer++) {
    // inner: projected BFGS on the augmented Lagrangian
    let H = null, F = AL(ev), { g, parts } = gradient(x, ev), steepest = false;
    lastParts = parts;
    kkt = false;
    for (let it = 0; it < 200; it++) {
      if (run.limitReached()) return { status: run.stopReason, x, ev, lam, parts: lastParts };
      if (iterLimit && run.iterations >= iterLimit) break;
      const free = x.map((v, j) => !((v <= lb[j] + 1e-12 && g[j] > 0) || (v >= ub[j] - 1e-12 && g[j] < 0)));
      const pg = g.map((v, j) => (free[j] ? v : 0));
      const pgn = Math.max(0, ...pg.map(Math.abs));
      if (pgn < tolG) { kkt = true; break; }
      if (!H) H = Array.from({ length: n }, (_, i) => Array.from({ length: n }, (__, j) => (i === j ? 1 : 0)));
      let d = new Array(n).fill(0);
      for (let i = 0; i < n; i++) { if (!free[i]) continue; let s = 0; for (let j = 0; j < n; j++) if (free[j]) s -= H[i][j] * g[j]; d[i] = s; }
      let slope = dot(d, pg);
      if (!(slope < 0)) { H = null; d = pg.map((v) => -v); slope = dot(d, pg); }
      // backtracking line search on the projected path
      let t = 1, xn, en, Fn;
      for (let ls = 0; ls < 40; ls++) {
        xn = project(x.map((v, j) => v + t * d[j]), lb, ub);
        en = P.evaluate(xn);
        Fn = AL(en);
        if (Fn <= F + 1e-4 * dot(g, xn.map((v, j) => v - x[j]))) break;
        t *= 0.5;
      }
      // no progress: retry once along the steepest descent, then accept the point
      if (!(Fn < F) || !finite(Fn)) { if (!steepest) { H = null; steepest = true; continue; } kkt = pgn < 1e-5; break; }
      steepest = false;
      const s = xn.map((v, j) => v - x[j]);
      const gr = gradient(xn, en);
      const yv = gr.g.map((v, j) => v - g[j]);
      const sy = dot(s, yv);
      if (H && sy > 1e-10 * Math.sqrt(dot(s, s) * dot(yv, yv))) {
        // BFGS update of the inverse Hessian
        const Hy = H.map((row) => dot(row, yv));
        const yHy = dot(yv, Hy);
        for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) {
          H[i][j] += ((sy + yHy) * s[i] * s[j]) / (sy * sy) - (Hy[i] * s[j] + s[i] * Hy[j]) / sy;
        }
      }
      const rel = Math.abs(F - Fn) / Math.max(1, Math.abs(F));
      x = xn; ev = en; F = Fn; g = gr.g; lastParts = gr.parts;
      await run.iteration({ objective: ev.f, x });
      if (rel < 1e-12 && Math.max(...s.map(Math.abs)) < 1e-10) break;
    }
    // the objective runs away: the problem is unbounded
    if (Math.abs(ev.f) > BIG || x.some((v) => Math.abs(v) > 1e15)) return { status: 'unbounded', x, ev };
    // multipliers, and how far they still moved: a positive multiplier on a slack row (no complementarity) or a
    // changing multiplier means the point is not yet optimal for the real problem
    let comp = 0;
    for (let k = 0; k < m; k++) {
      const h = hOf(ev.g, k), old = lam[k];
      lam[k] = P.ctype[k] === 'eq' ? old + mu * h : Math.max(0, old + mu * h);
      comp = Math.max(comp, Math.abs(lam[k] - old) / mu);
    }
    const realViol = violation(P, ev);
    keepFeasible();
    history.push(ev.f);
    const recent = history.slice(-5);
    const conv = recent.length >= 3 && (Math.max(...recent) - Math.min(...recent)) / Math.max(1e-10, Math.abs(ev.f), 1) <= opts.convergence;
    if (realViol <= opts.precision && comp <= opts.precision && (kkt || conv)) { status = kkt ? 'found' : 'converged'; done = true; break; }
    if (realViol > opts.precision && outer > 2) mu = Math.min(mu * 10, 1e12);
    if (run.limitReached()) return { status: run.stopReason, x, ev, lam, parts: lastParts };
    if (outer === 59) status = realViol <= opts.precision ? 'converged' : 'infeasible';
  }
  // the same tolerance as the status above and as the reports: never call a point feasible that a report marks violated
  const feasible = isFeasible(ev, x);
  if (!feasible && bestFeasible) {
    return { status: 'noImprove', x: bestFeasible.x, ev: bestFeasible.ev, F: sense * bestFeasible.ev.f, lam: null, parts: null, feasible: true };
  }
  return { status: feasible ? status : 'infeasible', x, ev, F: sense * ev.f, lam: lam.map((l, k) => l * gScale[k] / fScale), parts: lastParts, feasible };
}

export async function grg(P, opts, hooks) {
  const run = new Run(opts, hooks);
  const isInt = opts.ignoreInt ? P.isInt.map(() => false) : P.isInt;
  const difs = opts.ignoreInt ? [] : P.difGroups;
  const hasInt = isInt.some(Boolean) || difs.length > 0;
  if (conflictingBounds(P.lb, P.ub, isInt)) return result('infeasible', run);
  if (opts.multistart && opts.requireBounds && P.lb.some((v, j) => !isFinite(v) || !isFinite(P.ub[j]))) return result('bounds', run);
  const R = rng(opts.seed);
  const localFrom = async (lb, ub, start) => {
    const r = await grgLocal(P, lb, ub, start, opts, run);
    const ok = r.status === 'found' || r.status === 'converged' || r.status === 'noImprove';
    return { ...r, status: ok ? 'optimal' : r.status, F: r.ev ? P.sense * r.ev.f : Infinity, grgStatus: r.status };
  };
  const relax = async (lb, ub) => {
    if (!opts.multistart) return localFrom(lb, ub, P.x0);
    // multistart: several starting points spread over the bounds, keep the best feasible local solution
    let best = null;
    const starts = Math.max(1, Math.min(opts.population, 200));
    for (let s = 0; s < starts; s++) {
      if (run.limitReached()) break;
      const start = s === 0 ? P.x0 : P.x0.map((_, j) => {
        const lo = isFinite(lb[j]) ? lb[j] : -1000, hi = isFinite(ub[j]) ? ub[j] : 1000;
        return lo + R() * (hi - lo);
      });
      const r = await localFrom(lb, ub, start);
      if (r.status === 'optimal' && (!best || r.F < best.F - 1e-9)) best = r;
    }
    return best ? { ...best, multistart: true } : { status: 'infeasible' };
  };
  const r = hasInt ? await branchAndBound(relax, P.lb, P.ub, isInt, difs, opts, run) : await relax(P.lb, P.ub);
  if (r.status === 'error') return result('error', run);
  if (r.status === 'unbounded') return result('unbounded', run);
  if (!r.x) return result(LIMIT_STATUSES.includes(r.status) ? r.status : run.stopReason || 'infeasible', run);
  if (r.status !== 'optimal' && r.status !== 'limit' && !r.feasible) {
    if (run.stopReason) return { ...result(run.stopReason, run), x: r.x, feasible: false };
    return result(r.status === 'infeasible' ? 'infeasible' : r.status, run);
  }
  let status = r.limit || (hasInt ? (opts.intTolerance > 0 ? 'intTol' : 'found') : r.multistart ? 'global'
    : r.grgStatus === 'found' ? 'found' : r.grgStatus === 'noImprove' ? 'noImprove' : 'converged');
  if (run.stopReason && !r.limit) status = run.stopReason;
  // sensitivity: reduced gradient and Lagrange multipliers (∂objective/∂right side)
  let sensitivity = null;
  if (!hasInt && r.parts && r.lam) {
    const reducedGradient = r.parts.map((p) => {
      let s = p.df;
      P.ctype.forEach((t, k) => { const sign = t === 'ge' ? -1 : 1; s += P.sense * r.lam[k] * sign * p.dg[k]; });
      return s;
    });
    // ∂objective/∂(right side): relaxing a ≤ row lowers h, tightening a ≥ row raises it
    const multipliers = P.ctype.map((t, k) => (t === 'ge' ? P.sense : -P.sense) * r.lam[k]);
    sensitivity = { kind: 'nlp', reducedGradient, multipliers };
  }
  return { ...result(status, run), x: r.x, feasible: true, hasInt, sensitivity };
}

/* ================================================================ Evolutionary engine */

export async function evolutionary(P, opts, hooks) {
  const run = new Run(opts, hooks);
  const n = P.n;
  const isInt = opts.ignoreInt ? P.isInt.map(() => false) : P.isInt;
  const difs = opts.ignoreInt ? [] : P.difGroups;
  const inDif = new Array(n).fill(-1);
  difs.forEach((grp, gi) => grp.forEach((j) => { inDif[j] = gi; }));
  const lb = P.lb.slice(), ub = P.ub.slice();
  // AllDifferent members take the values 1..N, within any bounds of their own
  difs.forEach((grp) => grp.forEach((j) => { lb[j] = Math.max(lb[j], 1); ub[j] = Math.min(ub[j], grp.length); }));
  if (conflictingBounds(lb, ub, isInt)) return result('infeasible', run);
  if (opts.requireBounds && lb.some((v, j) => inDif[j] < 0 && (!isFinite(v) || !isFinite(ub[j])))) return result('bounds', run);
  const lo = lb.map((v, j) => (isFinite(v) ? v : (isFinite(ub[j]) ? ub[j] - 2000 : -1000)));
  const hi = ub.map((v, j) => (isFinite(v) ? v : lo[j] + 2000));
  const R = rng(opts.seed);
  const gauss = () => { let u = 0, v = 0; while (!u) u = R(); while (!v) v = R(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); };
  const shuffle = (a) => { for (let i = a.length - 1; i > 0; i--) { const k = Math.floor(R() * (i + 1)); [a[i], a[k]] = [a[k], a[i]]; } return a; };
  const repair = (x) => {
    for (let j = 0; j < n; j++) {
      if (inDif[j] >= 0) continue;
      x[j] = clamp(x[j], lo[j], hi[j]);
      if (isInt[j]) x[j] = clamp(Math.round(x[j]), Math.ceil(lo[j]), Math.floor(hi[j]));
    }
    return x;
  };
  const randomPerm = (grp) => shuffle(grp.map((_, i) => i + 1));
  const make = (x) => {
    const ev = P.evaluate(x);
    const broken = bad(ev);
    // permutations of AllDifferent groups can break a member's own bounds: that counts as a violation too
    return { x, f: broken ? NaN : ev.f, v: broken ? Infinity : Math.max(violation(P, ev), boundViolation(x, lb, ub)), ev };
  };
  // feasible beats infeasible; then objective (or smaller violation)
  const better = (a, b) => {
    const fa = a.v <= opts.precision, fb = b.v <= opts.precision;
    if (fa !== fb) return fa;
    if (!fa) return a.v < b.v;
    return P.sense * a.f < P.sense * b.f;
  };
  const size = Math.max(10, Math.min(opts.population, 1000));
  let pop = [];
  const start = repair(P.x0.map((v, j) => (finite(v) ? v : lo[j])));
  difs.forEach((grp) => {
    // keep the start values only when they are exactly 1..N
    const vals = grp.map((j) => start[j]).sort((a, b) => a - b);
    if (!vals.every((v, i) => v === i + 1)) randomPerm(grp).forEach((v, i) => { start[grp[i]] = v; });
  });
  pop.push(make(start));
  while (pop.length < size) {
    const x = new Array(n);
    for (let j = 0; j < n; j++) x[j] = lo[j] + R() * (hi[j] - lo[j]);
    difs.forEach((grp) => randomPerm(grp).forEach((v, i) => { x[grp[i]] = v; }));
    pop.push(make(repair(x)));
  }
  const sortPop = () => pop.sort((a, b) => (better(a, b) ? -1 : better(b, a) ? 1 : 0));
  sortPop();
  let best = pop[0], lastImprove = run.now(), status = 'converged';
  const tournament = () => { const a = pop[Math.floor(R() * pop.length)], b = pop[Math.floor(R() * pop.length)]; return better(a, b) ? a : b; };
  for (let gen = 0; ; gen++) {
    const lim = run.limitReached();
    if (lim) { status = lim; break; }
    if (opts.maxNoImprove > 0 && (run.now() - lastImprove) / 1000 > opts.maxNoImprove) { status = 'noImprove'; break; }
    if (opts.maxSubproblems > 0 && run.subproblems >= opts.maxSubproblems) { status = 'subproblems'; break; }
    const children = [];
    for (let c = 0; c < size; c++) {
      const a = tournament(), b = tournament();
      const x = new Array(n);
      for (let j = 0; j < n; j++) {
        if (inDif[j] >= 0) continue;
        const w = R() * 1.5 - 0.25; // blend crossover
        x[j] = a.x[j] + w * (b.x[j] - a.x[j]);
        if (R() < opts.mutation) x[j] += gauss() * 0.1 * (hi[j] - lo[j]);
        if (R() < opts.mutation * 0.2) x[j] = lo[j] + R() * (hi[j] - lo[j]);
      }
      difs.forEach((grp) => {
        // order crossover on the permutation, then a swap mutation
        const pa = grp.map((j) => a.x[j]), pb = grp.map((j) => b.x[j]);
        const i1 = Math.floor(R() * grp.length), i2 = i1 + Math.floor(R() * (grp.length - i1));
        const child = new Array(grp.length).fill(0), used = new Set();
        for (let i = i1; i <= i2; i++) { child[i] = pa[i]; used.add(pa[i]); }
        let p = 0;
        for (let i = 0; i < grp.length; i++) { if (i >= i1 && i <= i2) continue; while (used.has(pb[p])) p++; child[i] = pb[p]; used.add(pb[p]); }
        if (R() < Math.max(opts.mutation, 0.1)) { const s1 = Math.floor(R() * grp.length), s2 = Math.floor(R() * grp.length); [child[s1], child[s2]] = [child[s2], child[s1]]; }
        child.forEach((v, i) => { x[grp[i]] = v; });
      });
      children.push(make(repair(x)));
      run.subproblems++;
    }
    pop = pop.concat(children);
    sortPop();
    pop = pop.slice(0, size);
    // local search around the best member (pattern steps on continuous variables)
    if (gen % 5 === 0) {
      let cur = pop[0];
      for (let j = 0; j < n; j++) {
        if (inDif[j] >= 0) continue;
        for (const dir of [1, -1]) {
          const step = isInt[j] ? 1 : 0.01 * Math.max(1e-6, hi[j] - lo[j]);
          const x = cur.x.slice(); x[j] += dir * step;
          const cand = make(repair(x));
          if (better(cand, cur)) cur = cand;
        }
      }
      if (better(cur, pop[0])) { pop[0] = cur; sortPop(); }
    }
    if (better(pop[0], best)) { best = pop[0]; lastImprove = run.now(); }
    await run.iteration({ objective: best.f, x: best.x, generation: gen });
    // convergence: the best part of the population agrees on the objective
    const feas = pop.filter((p) => p.v <= opts.precision);
    if (feas.length >= pop.length * 0.99) {
      const fs = feas.map((p) => p.f);
      const spread = (Math.max(...fs) - Math.min(...fs)) / Math.max(1, Math.abs(best.f));
      if (spread <= opts.convergence) { status = 'converged'; break; }
    }
  }
  if (!(best.v <= opts.precision)) return { ...result(run.stopReason || 'infeasible', run), x: best.x, feasible: false };
  const stats = Array.from({ length: n }, (_, j) => {
    const vals = pop.map((p) => p.x[j]);
    const meanV = vals.reduce((s, v) => s + v, 0) / vals.length;
    const sd = Math.sqrt(vals.reduce((s, v) => s + (v - meanV) ** 2, 0) / Math.max(1, vals.length - 1));
    return { best: best.x[j], mean: meanV, sd, max: Math.max(...vals), min: Math.min(...vals) };
  });
  const gStats = P.ctype.map((_, k) => {
    const lhs = (p) => (p.ev.lhs ? p.ev.lhs[k] : p.ev.g[k]);
    const vals = pop.map(lhs).filter(finite);
    const meanV = vals.reduce((s, v) => s + v, 0) / Math.max(1, vals.length);
    const sd = Math.sqrt(vals.reduce((s, v) => s + (v - meanV) ** 2, 0) / Math.max(1, vals.length - 1));
    return { best: lhs(best), mean: meanV, sd, max: Math.max(...vals), min: Math.min(...vals) };
  });
  return { ...result(status, run), x: best.x, feasible: true, population: { vars: stats, constraints: gStats } };
}

/* ================================================================ helpers for reports */

function result(status, run) {
  return { status, message: MESSAGES[status] || status, time: run.elapsed, iterations: run.iterations, subproblems: run.subproblems };
}

/**
 * Feasibility analysis: a smallest set of constraints that cannot be met together (deletion filter).
 * isFeasible(activeConstraintIndexes) → Promise<boolean>.
 */
export async function findIIS(count, isFeasible, run) {
  let active = Array.from({ length: count }, (_, i) => i);
  if (await isFeasible(active)) return [];
  for (let i = 0; i < count; i++) {
    if (run && run.limitReached && run.limitReached()) break;
    const without = active.filter((k) => k !== i);
    if (!(await isFeasible(without))) active = without;
  }
  return active;
}

/** Limits report: how far each variable can move (others fixed) while all constraints hold. */
export function variableLimits(P, x, lin) {
  return x.map((xj, j) => {
    let lo = P.lb[j], hi = P.ub[j];
    if (lin) {
      // linear: each constraint a·x + b (rel) 0 bounds x_j
      lin.A.forEach((row, k) => {
        const a = row[j];
        if (Math.abs(a) < 1e-12) return;
        const rest = dot(row, x) - a * xj + lin.b[k];
        const lim = -rest / a;
        const t = P.ctype[k];
        if (t === 'eq') { lo = Math.max(lo, lim); hi = Math.min(hi, lim); }
        else if ((t === 'le') === (a > 0)) hi = Math.min(hi, lim);
        else lo = Math.max(lo, lim);
      });
    } else {
      // nonlinear: walk outwards from the solution and bisect the feasibility border
      const feasibleAt = (v) => { const z = x.slice(); z[j] = v; const e = P.evaluate(z); return finite(e.f) && violation(P, e) <= 1e-6; };
      const edge = (dir) => {
        let good = xj, step = Math.max(1, Math.abs(xj));
        let bound = dir > 0 ? hi : lo;
        for (let k = 0; k < 40; k++) {
          const t = good + dir * step;
          if ((dir > 0 && t > bound) || (dir < 0 && t < bound)) { if (isFinite(bound) && feasibleAt(bound)) return bound; break; }
          if (feasibleAt(t)) { good = t; step *= 2; } else { bound = t; break; }
        }
        let a = good, b = bound;
        if (!isFinite(b)) return dir > 0 ? BIG : -BIG;
        for (let k = 0; k < 50; k++) { const mid = (a + b) / 2; if (feasibleAt(mid)) a = mid; else b = mid; }
        return a;
      };
      lo = edge(-1); hi = edge(1);
    }
    const objAt = (v) => {
      if (!isFinite(v) || Math.abs(v) >= BIG) return null;
      const z = x.slice(); z[j] = v;
      const e = P.evaluate(z);
      return e.objCell !== undefined ? e.objCell : e.f;
    };
    return { lower: isFinite(lo) ? lo : -BIG, upper: isFinite(hi) ? hi : BIG, objLower: objAt(lo), objUpper: objAt(hi) };
  });
}

/**
 * Linearity report: which functions are linear (objective, each constraint) and which variables occur linearly
 * (their partial derivatives are the same everywhere).
 */
export function linearityReport(P, x0) {
  const lin = linearize(P, x0);
  if (lin.error) return null;
  const R = rng(777);
  const far = x0.map((v) => v + (R() * 6 - 3) * Math.max(1, Math.abs(v)));
  const e0 = P.evaluate(x0), e1 = P.evaluate(far);
  const vars = x0.map((_, j) => {
    const a = x0.slice(); a[j] += 1;
    const b = far.slice(); b[j] += 1;
    const ea = P.evaluate(a), eb = P.evaluate(b);
    const same = (p0, p1, q0, q1) => {
      const d0 = p1 - p0, d1 = q1 - q0;
      return finite(d0) && finite(d1) && Math.abs(d0 - d1) <= 1e-6 * Math.max(1, Math.abs(d0), Math.abs(d1));
    };
    if (!same(e0.f, ea.f, e1.f, eb.f)) return false;
    return e0.g.every((v, k) => same(v, ea.g[k], e1.g[k], eb.g[k]));
  });
  return { objective: !lin.nonlinearF, constraints: lin.nonlinearG.map((v) => !v), vars };
}
