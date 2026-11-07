/**
 * Office One spreadsheet function library.
 *
 * Every function is registered here with its arity and implementation. The evaluator
 * (`formula.ts`) owns parsing, references, caching and dependency tracking, and hands each
 * function a `Ctx` that resolves arguments lazily, so `IF(TRUE,1,1/0)` never evaluates the
 * error branch.
 *
 * Errors are raised by throwing a `FormulaError`; the evaluator turns that into the cell value.
 */
import { FormulaError, toStr, type Scalar } from './formulaCore'

export type Val = Scalar | Scalar[][]
const err = (c: string) => new FormulaError(c)
export const isErr = (v: unknown): v is FormulaError => v instanceof FormulaError

export interface Ctx {
  /** Number of arguments supplied. */
  readonly n: number
  /** True when argument `i` was supplied and is not an empty placeholder (`SUM(1,,2)`). */
  has(i: number): boolean
  /** Raw value of argument `i` (may be a FormulaError; use for IFERROR-style handling). */
  scalar(i: number): Scalar
  /** Argument `i` flattened to a list of values. */
  flat(i: number): Scalar[]
  /** Argument `i` as a 2-D grid. */
  grid(i: number): Scalar[][]
  /** True when argument `i` is a cell reference or range rather than a literal/expression. */
  isRef(i: number): boolean
  num(i: number, dflt?: number): number
  int(i: number, dflt?: number): number
  str(i: number, dflt?: string): string
  bool(i: number, dflt?: boolean): boolean
  /** Numbers across arguments `from..`, using aggregate rules: ranges skip text and blanks. */
  nums(from?: number, to?: number): number[]
  /** Every value across arguments `from..`. */
  all(from?: number): Scalar[]
}

export interface FnDef { min: number; max: number; fn: (c: Ctx) => Val; volatile?: boolean }

/* ------------------------------------------------------------------ helpers */

const EPS = 1e-10
const round = (n: number, d: number) => { const f = Math.pow(10, d); return Math.sign(n) * Math.round(Math.abs(n) * f + EPS) / f }
const need = (ok: boolean, code = '#VALUE!') => { if (!ok) throw err(code) }

/** Excel serial dates: day 1 is 1900-01-01, with Excel's 1900 leap-year quirk preserved above day 60. */
const EPOCH = Date.UTC(1899, 11, 30)
export const toSerial = (d: Date) => (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - EPOCH) / 86400000
export const fromSerial = (n: number) => new Date(EPOCH + Math.floor(n) * 86400000)
const todaySerial = () => { const d = new Date(); return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / 86400000 }
const nowSerial = () => { const d = new Date(); return todaySerial() + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 }
const timePart = (n: number) => { const f = n - Math.floor(n); return Math.round(f * 86400) }

export function wildcardRe(pat: string): RegExp {
  return new RegExp('^' + pat.replace(/~\*/g, '\u0000').replace(/~\?/g, '\u0001').replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0000/g, '\\*').replace(/\u0001/g, '\\?') + '$', 'i')
}
export const looseEq = (a: Scalar, b: Scalar): boolean => {
  if (typeof a === 'number' && typeof b === 'number') return a === b
  if (a === null || b === null) return (a === null || a === '') && (b === null || b === '')
  return toStr(a).toLowerCase() === toStr(b).toLowerCase()
}
export const compare = (a: Scalar, b: Scalar): number => {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  const x = toStr(a).toLowerCase(), y = toStr(b).toLowerCase()
  return x < y ? -1 : x > y ? 1 : 0
}
const asNumOrNull = (v: Scalar): number | null =>
  typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v)) ? Number(v) : null

/** Build a predicate from a COUNTIF/SUMIF-style criterion such as 5, ">=3" or "app*". */
export function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c === 'number' || typeof c === 'boolean') return v => looseEq(v, c)
  const s = toStr(c)
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s)!
  const op = m[1] ?? '=', rhs = m[2]
  const rn = rhs.trim() !== '' && !Number.isNaN(Number(rhs)) ? Number(rhs) : null
  return v => {
    if (isErr(v)) return false
    if (rn !== null) {
      const vn = asNumOrNull(v)
      if (vn === null) return op === '<>'
      return op === '=' ? vn === rn : op === '<>' ? vn !== rn : op === '<' ? vn < rn : op === '>' ? vn > rn : op === '<=' ? vn <= rn : vn >= rn
    }
    if (op === '=' || op === '<>') { const hit = wildcardRe(rhs).test(toStr(v)); return op === '=' ? hit : !hit }
    const a = toStr(v).toLowerCase(), b = rhs.toLowerCase()
    return op === '<' ? a < b : op === '>' ? a > b : op === '<=' ? a <= b : a >= b
  }
}

/** Rows matching every (range, criterion) pair of a *IFS function, as index lists. */
function ifsMatches(c: Ctx, firstPair: number): number[] {
  const pairs: { vals: Scalar[]; test: (v: Scalar) => boolean }[] = []
  for (let i = firstPair; i + 1 < c.n; i += 2) pairs.push({ vals: c.flat(i), test: criterion(c.scalar(i + 1)) })
  need(pairs.length > 0, '#VALUE!')
  const len = pairs[0].vals.length
  for (const p of pairs) need(p.vals.length === len, '#VALUE!')
  const out: number[] = []
  for (let i = 0; i < len; i++) if (pairs.every(p => p.test(p.vals[i]))) out.push(i)
  return out
}

const stats = (xs: number[]) => {
  const n = xs.length, mean = n ? xs.reduce((a, b) => a + b, 0) / n : 0
  const ss = xs.reduce((a, b) => a + (b - mean) ** 2, 0)
  return { n, mean, ss }
}
/** Linear interpolation percentile, matching PERCENTILE.INC. */
function percentile(xs: number[], p: number): number {
  need(xs.length > 0, '#NUM!'); need(p >= 0 && p <= 1, '#NUM!')
  const s = [...xs].sort((a, b) => a - b)
  const idx = p * (s.length - 1), lo = Math.floor(idx), hi = Math.ceil(idx)
  return lo === hi ? s[lo] : s[lo] + (s[hi] - s[lo]) * (idx - lo)
}

/* --------------------------------------------------------------- financial */
/** Present/future value relation shared by PMT, PV, FV, NPER. type 0 = end of period, 1 = start. */
function pmtOf(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper
  const p = Math.pow(1 + rate, nper)
  return -(pv * p + fv) * rate / ((1 + rate * type) * (p - 1))
}
function fvOf(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + pmt * nper)
  const p = Math.pow(1 + rate, nper)
  return -(pv * p + pmt * (1 + rate * type) * (p - 1) / rate)
}

/* ------------------------------------------------------------------ registry */

export const FUNCTIONS: Record<string, FnDef> = {}
const def = (names: string, min: number, max: number, fn: (c: Ctx) => Val, volatile = false) => {
  for (const n of names.split(' ')) FUNCTIONS[n] = { min, max, fn, volatile }
}
const INF = 255

/* ---- math ---- */
def('SUM', 1, INF, c => c.nums().reduce((a, b) => a + b, 0))
def('PRODUCT', 1, INF, c => { const n = c.nums(); return n.length ? n.reduce((a, b) => a * b, 1) : 0 })
def('SUMSQ', 1, INF, c => c.nums().reduce((a, b) => a + b * b, 0))
def('ABS', 1, 1, c => Math.abs(c.num(0)))
def('SIGN', 1, 1, c => Math.sign(c.num(0)))
def('SQRT', 1, 1, c => { const n = c.num(0); need(n >= 0, '#NUM!'); return Math.sqrt(n) })
def('INT', 1, 1, c => Math.floor(c.num(0)))
def('TRUNC', 1, 2, c => { const n = c.num(0), d = c.has(1) ? c.int(1) : 0, f = Math.pow(10, d); return Math.trunc(n * f) / f })
def('ROUND', 1, 2, c => round(c.num(0), c.has(1) ? c.int(1) : 0))
def('ROUNDUP', 1, 2, c => { const n = c.num(0), f = Math.pow(10, c.has(1) ? c.int(1) : 0); return Math.sign(n) * Math.ceil(Math.abs(n) * f - EPS) / f })
def('ROUNDDOWN', 1, 2, c => { const n = c.num(0), f = Math.pow(10, c.has(1) ? c.int(1) : 0); return Math.sign(n) * Math.floor(Math.abs(n) * f + EPS) / f })
def('MROUND', 2, 2, c => { const n = c.num(0), m = c.num(1); if (m === 0) return 0; need(Math.sign(n) === Math.sign(m) || n === 0, '#NUM!'); return Math.round(n / m) * m })
def('CEILING CEILING.MATH', 1, 2, c => { const n = c.num(0), s = c.has(1) ? c.num(1) : 1; if (s === 0) return 0; return Math.ceil(n / s) * s })
def('FLOOR FLOOR.MATH', 1, 2, c => { const n = c.num(0), s = c.has(1) ? c.num(1) : 1; if (s === 0) throw err('#DIV/0!'); return Math.floor(n / s) * s })
def('EVEN', 1, 1, c => { const n = c.num(0); return Math.sign(n) * Math.ceil(Math.abs(n) / 2) * 2 || 0 })
def('ODD', 1, 1, c => { const n = c.num(0); if (n === 0) return 1; const a = Math.ceil(Math.abs(n)); return Math.sign(n) * (a % 2 === 0 ? a + 1 : a) })
def('MOD', 2, 2, c => { const a = c.num(0), b = c.num(1); if (b === 0) throw err('#DIV/0!'); return a - b * Math.floor(a / b) })
def('QUOTIENT', 2, 2, c => { const b = c.num(1); if (b === 0) throw err('#DIV/0!'); return Math.trunc(c.num(0) / b) })
def('POWER', 2, 2, c => { const r = Math.pow(c.num(0), c.num(1)); need(Number.isFinite(r) || Math.abs(r) === Infinity, '#NUM!'); need(!Number.isNaN(r), '#NUM!'); return r })
def('EXP', 1, 1, c => Math.exp(c.num(0)))
def('LN', 1, 1, c => { const n = c.num(0); need(n > 0, '#NUM!'); return Math.log(n) })
def('LOG10', 1, 1, c => { const n = c.num(0); need(n > 0, '#NUM!'); return Math.log10(n) })
def('LOG', 1, 2, c => { const n = c.num(0), b = c.has(1) ? c.num(1) : 10; need(n > 0 && b > 0 && b !== 1, '#NUM!'); return Math.log(n) / Math.log(b) })
def('PI', 0, 0, () => Math.PI)
def('DEGREES', 1, 1, c => (c.num(0) * 180) / Math.PI)
def('RADIANS', 1, 1, c => (c.num(0) * Math.PI) / 180)
for (const t of ['SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'SINH', 'COSH', 'TANH'] as const) {
  def(t, 1, 1, c => { const r = (Math[t.toLowerCase() as 'sin'] as (x: number) => number)(c.num(0)); need(!Number.isNaN(r), '#NUM!'); return r })
}
def('ATAN2', 2, 2, c => Math.atan2(c.num(1), c.num(0)))
def('FACT', 1, 1, c => { const n = Math.floor(c.num(0)); need(n >= 0 && n <= 170, '#NUM!'); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r })
def('COMBIN', 2, 2, c => { const n = Math.floor(c.num(0)), k = Math.floor(c.num(1)); need(n >= 0 && k >= 0 && k <= n, '#NUM!'); let r = 1; for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i; return Math.round(r) })
def('PERMUT', 2, 2, c => { const n = Math.floor(c.num(0)), k = Math.floor(c.num(1)); need(n >= 0 && k >= 0 && k <= n, '#NUM!'); let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r })
def('GCD', 1, INF, c => { const g = (a: number, b: number): number => (b ? g(b, a % b) : a); return c.nums().map(Math.abs).map(Math.floor).reduce((a, b) => g(a, b), 0) })
def('LCM', 1, INF, c => { const g = (a: number, b: number): number => (b ? g(b, a % b) : a); return c.nums().map(Math.abs).map(Math.floor).reduce((a, b) => (a && b ? (a * b) / g(a, b) : 0), 1) })
def('RAND', 0, 0, () => Math.random(), true)
def('RANDBETWEEN', 2, 2, c => { const lo = Math.ceil(c.num(0)), hi = Math.floor(c.num(1)); need(hi >= lo, '#NUM!'); return lo + Math.floor(Math.random() * (hi - lo + 1)) }, true)
def('SUMPRODUCT', 1, INF, c => {
  const grids = Array.from({ length: c.n }, (_, i) => c.grid(i).flat())
  const len = grids[0].length
  for (const g of grids) need(g.length === len, '#VALUE!')
  let total = 0
  for (let i = 0; i < len; i++) { let p = 1; for (const g of grids) p *= typeof g[i] === 'number' ? (g[i] as number) : 0; total += p }
  return total
})
def('SUMIF', 2, 3, c => {
  const range = c.flat(0), test = criterion(c.scalar(1)), sr = c.has(2) ? c.flat(2) : range
  let s = 0; range.forEach((v, i) => { if (test(v)) { const x = sr[i]; if (typeof x === 'number') s += x } }); return s
})
def('SUMIFS', 3, INF, c => { const sr = c.flat(0); return ifsMatches(c, 1).reduce((a, i) => a + (typeof sr[i] === 'number' ? (sr[i] as number) : 0), 0) })

/* ---- statistics ---- */
def('AVERAGE', 1, INF, c => { const n = c.nums(); if (!n.length) throw err('#DIV/0!'); return n.reduce((a, b) => a + b, 0) / n.length })
def('AVERAGEA', 1, INF, c => { const v = c.all(); const n = v.filter(x => x !== null && x !== ''); if (!n.length) throw err('#DIV/0!'); return n.reduce((a: number, b) => a + (typeof b === 'number' ? b : typeof b === 'boolean' ? (b ? 1 : 0) : 0), 0) / n.length })
def('MIN', 1, INF, c => { const n = c.nums(); return n.length ? Math.min(...n) : 0 })
def('MAX', 1, INF, c => { const n = c.nums(); return n.length ? Math.max(...n) : 0 })
def('COUNT', 1, INF, c => c.all().filter(v => typeof v === 'number').length)
def('COUNTA', 1, INF, c => c.all().filter(v => v !== null && v !== '').length)
def('COUNTBLANK', 1, 1, c => c.flat(0).filter(v => v === null || v === '').length)
def('COUNTIF', 2, 2, c => { const t = criterion(c.scalar(1)); return c.flat(0).filter(t).length })
def('COUNTIFS', 2, INF, c => ifsMatches(c, 0).length)
def('AVERAGEIF', 2, 3, c => {
  const range = c.flat(0), test = criterion(c.scalar(1)), ar = c.has(2) ? c.flat(2) : range
  let s = 0, n = 0; range.forEach((v, i) => { if (test(v)) { const x = ar[i]; if (typeof x === 'number') { s += x; n++ } } })
  if (!n) throw err('#DIV/0!'); return s / n
})
def('AVERAGEIFS', 3, INF, c => {
  const ar = c.flat(0), hits = ifsMatches(c, 1).filter(i => typeof ar[i] === 'number')
  if (!hits.length) throw err('#DIV/0!')
  return hits.reduce((a, i) => a + (ar[i] as number), 0) / hits.length
})
def('MINIFS', 3, INF, c => { const ar = c.flat(0); const v = ifsMatches(c, 1).map(i => ar[i]).filter(x => typeof x === 'number') as number[]; return v.length ? Math.min(...v) : 0 })
def('MAXIFS', 3, INF, c => { const ar = c.flat(0); const v = ifsMatches(c, 1).map(i => ar[i]).filter(x => typeof x === 'number') as number[]; return v.length ? Math.max(...v) : 0 })
def('MEDIAN', 1, INF, c => { const s = c.nums().sort((a, b) => a - b); need(s.length > 0, '#NUM!'); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2 })
def('MODE MODE.SNGL', 1, INF, c => {
  const n = c.nums(); const count = new Map<number, number>()
  for (const x of n) count.set(x, (count.get(x) ?? 0) + 1)
  let best: number | null = null, bc = 1
  for (const x of n) { const k = count.get(x)!; if (k > bc) { bc = k; best = x } }
  if (best === null) throw err('#N/A'); return best
})
def('STDEV STDEV.S', 1, INF, c => { const { n, ss } = stats(c.nums()); need(n > 1, '#DIV/0!'); return Math.sqrt(ss / (n - 1)) })
def('STDEVP STDEV.P', 1, INF, c => { const { n, ss } = stats(c.nums()); need(n > 0, '#DIV/0!'); return Math.sqrt(ss / n) })
def('VAR VAR.S', 1, INF, c => { const { n, ss } = stats(c.nums()); need(n > 1, '#DIV/0!'); return ss / (n - 1) })
def('VARP VAR.P', 1, INF, c => { const { n, ss } = stats(c.nums()); need(n > 0, '#DIV/0!'); return ss / n })
def('LARGE', 2, 2, c => { const s = c.nums(0, 0).sort((a, b) => b - a), k = c.int(1); need(k >= 1 && k <= s.length, '#NUM!'); return s[k - 1] })
def('SMALL', 2, 2, c => { const s = c.nums(0, 0).sort((a, b) => a - b), k = c.int(1); need(k >= 1 && k <= s.length, '#NUM!'); return s[k - 1] })
def('RANK RANK.EQ', 2, 3, c => {
  const x = c.num(0), xs = c.nums(1, 1), asc = c.has(2) && c.bool(2)
  const s = [...xs].sort((a, b) => (asc ? a - b : b - a))
  const i = s.findIndex(v => v === x); need(i >= 0, '#N/A'); return i + 1
})
def('PERCENTILE PERCENTILE.INC', 2, 2, c => percentile(c.nums(0, 0), c.num(1)))
def('QUARTILE QUARTILE.INC', 2, 2, c => { const q = c.int(1); need(q >= 0 && q <= 4, '#NUM!'); return percentile(c.nums(0, 0), q / 4) })
def('GEOMEAN', 1, INF, c => { const n = c.nums(); need(n.length > 0 && n.every(x => x > 0), '#NUM!'); return Math.exp(n.reduce((a, b) => a + Math.log(b), 0) / n.length) })
def('HARMEAN', 1, INF, c => { const n = c.nums(); need(n.length > 0 && n.every(x => x > 0), '#NUM!'); return n.length / n.reduce((a, b) => a + 1 / b, 0) })
const pairXY = (c: Ctx) => {
  const ys = c.flat(0), xs = c.flat(1); need(ys.length === xs.length, '#N/A')
  const px: number[] = [], py: number[] = []
  for (let i = 0; i < ys.length; i++) if (typeof ys[i] === 'number' && typeof xs[i] === 'number') { py.push(ys[i] as number); px.push(xs[i] as number) }
  need(px.length > 1, '#DIV/0!')
  const mx = px.reduce((a, b) => a + b, 0) / px.length, my = py.reduce((a, b) => a + b, 0) / py.length
  let sxy = 0, sxx = 0, syy = 0
  for (let i = 0; i < px.length; i++) { sxy += (px[i] - mx) * (py[i] - my); sxx += (px[i] - mx) ** 2; syy += (py[i] - my) ** 2 }
  return { mx, my, sxy, sxx, syy }
}
def('CORREL PEARSON', 2, 2, c => { const { sxy, sxx, syy } = pairXY(c); need(sxx > 0 && syy > 0, '#DIV/0!'); return sxy / Math.sqrt(sxx * syy) })
def('SLOPE', 2, 2, c => { const { sxy, sxx } = pairXY(c); need(sxx !== 0, '#DIV/0!'); return sxy / sxx })
def('INTERCEPT', 2, 2, c => { const { mx, my, sxy, sxx } = pairXY(c); need(sxx !== 0, '#DIV/0!'); return my - (sxy / sxx) * mx })

/* ---- logical ---- */
/** Picks `g`'s value at (r,c), stretching single rows/columns the way Excel broadcasts. */
const pickAt = (g: Scalar[][], r: number, c: number): Scalar => {
  const row = g.length === 1 ? g[0] : g[r]
  if (!row) return null
  return (row.length === 1 ? row[0] : row[c]) ?? null
}
const asBool = (v: Scalar): boolean => {
  if (isErr(v)) throw v
  if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; const n = Number(v); if (v !== '' && !Number.isNaN(n)) return n !== 0; throw err('#VALUE!') }
  return !!v
}
def('IF', 2, 3, c => {
  const cond = c.grid(0)
  // A range condition makes IF an array formula, as in modern Excel: =SUM(IF(A1:A5>3,1,0))
  if (cond.length > 1 || (cond[0]?.length ?? 0) > 1) {
    const yes = c.grid(1), no = c.has(2) ? c.grid(2) : [[false]]
    return cond.map((row, r) => row.map((v, i) => {
      try { return asBool(v) ? pickAt(yes, r, i) : pickAt(no, r, i) } catch (e) { return isErr(e) ? e : err('#VALUE!') }
    }))
  }
  return c.bool(0) ? c.scalar(1) : c.has(2) ? c.scalar(2) : false
})
def('IFS', 2, INF, c => { for (let i = 0; i + 1 < c.n; i += 2) if (c.bool(i)) return c.scalar(i + 1); throw err('#N/A') })
def('SWITCH', 2, INF, c => {
  const v = c.scalar(0)
  let i = 1
  for (; i + 1 < c.n; i += 2) if (looseEq(v, c.scalar(i))) return c.scalar(i + 1)
  return i < c.n ? c.scalar(i) : err('#N/A')
})
def('AND', 1, INF, c => { const b = boolArgs(c); need(b.length > 0); return b.every(Boolean) })
def('OR', 1, INF, c => { const b = boolArgs(c); need(b.length > 0); return b.some(Boolean) })
def('XOR', 1, INF, c => { const b = boolArgs(c); need(b.length > 0); return b.filter(Boolean).length % 2 === 1 })
def('NOT', 1, 1, c => !c.bool(0))
def('TRUE', 0, 0, () => true)
def('FALSE', 0, 0, () => false)
def('IFERROR', 2, 2, c => { const v = c.scalar(0); return isErr(v) ? c.scalar(1) : v })
def('IFNA', 2, 2, c => { const v = c.scalar(0); return isErr(v) && v.code === '#N/A' ? c.scalar(1) : v })
function boolArgs(c: Ctx): boolean[] {
  const out: boolean[] = []
  for (let i = 0; i < c.n; i++) {
    const ref = c.isRef(i)
    for (const v of c.flat(i)) {
      if (isErr(v)) throw v
      if (ref && (v === null || typeof v === 'string')) continue // ranges ignore text and blanks
      if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE' || u === 'FALSE') { out.push(u === 'TRUE'); continue } const n = Number(v); if (v !== '' && !Number.isNaN(n)) { out.push(n !== 0); continue } throw err('#VALUE!') }
      out.push(typeof v === 'number' ? v !== 0 : !!v)
    }
  }
  return out
}

/* ---- information ---- */
def('ISBLANK', 1, 1, c => { const v = c.scalar(0); return v === null || v === '' })
def('ISNUMBER', 1, 1, c => typeof c.scalar(0) === 'number')
def('ISTEXT', 1, 1, c => typeof c.scalar(0) === 'string')
def('ISNONTEXT', 1, 1, c => typeof c.scalar(0) !== 'string')
def('ISLOGICAL', 1, 1, c => typeof c.scalar(0) === 'boolean')
def('ISERROR', 1, 1, c => isErr(c.scalar(0)))
def('ISERR', 1, 1, c => { const v = c.scalar(0); return isErr(v) && v.code !== '#N/A' })
def('ISNA', 1, 1, c => { const v = c.scalar(0); return isErr(v) && v.code === '#N/A' })
def('ISEVEN', 1, 1, c => Math.floor(Math.abs(c.num(0))) % 2 === 0)
def('ISODD', 1, 1, c => Math.floor(Math.abs(c.num(0))) % 2 === 1)
def('NA', 0, 0, () => err('#N/A'))
def('N', 1, 1, c => { const v = c.scalar(0); if (isErr(v)) throw v; return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0 })
def('T', 1, 1, c => { const v = c.scalar(0); if (isErr(v)) throw v; return typeof v === 'string' ? v : '' })
def('TYPE', 1, 1, c => { const v = c.scalar(0); return isErr(v) ? 16 : typeof v === 'number' ? 1 : typeof v === 'boolean' ? 4 : 2 })
def('ERROR.TYPE', 1, 1, c => {
  const v = c.scalar(0); if (!isErr(v)) throw err('#N/A')
  const order = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']
  const i = order.indexOf(v.code); return i < 0 ? 8 : i + 1
})

/* ---- text ---- */
def('CONCAT CONCATENATE', 1, INF, c => c.all().map(toStr).join(''))
def('TEXTJOIN', 3, INF, c => {
  const sep = c.str(0), skip = c.bool(1)
  const parts: string[] = []
  for (let i = 2; i < c.n; i++) for (const v of c.flat(i)) { if (isErr(v)) throw v; if (skip && (v === null || v === '')) continue; parts.push(toStr(v)) }
  return parts.join(sep)
})
def('LEN', 1, 1, c => c.str(0).length)
def('LEFT', 1, 2, c => { const n = c.has(1) ? c.int(1) : 1; need(n >= 0); return c.str(0).slice(0, n) })
def('RIGHT', 1, 2, c => { const n = c.has(1) ? c.int(1) : 1; need(n >= 0); return n === 0 ? '' : c.str(0).slice(-n) })
def('MID', 3, 3, c => { const s = c.str(0), start = c.int(1), n = c.int(2); need(start >= 1 && n >= 0, '#VALUE!'); return s.slice(start - 1, start - 1 + n) })
def('UPPER', 1, 1, c => c.str(0).toUpperCase())
def('LOWER', 1, 1, c => c.str(0).toLowerCase())
def('PROPER', 1, 1, c => c.str(0).replace(/[A-Za-zÀ-ɏ]+/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()))
def('TRIM', 1, 1, c => c.str(0).replace(/\s+/g, ' ').trim())
def('CLEAN', 1, 1, c => c.str(0).replace(/[\x00-\x1F\x7F]/g, ''))
def('REPT', 2, 2, c => { const n = c.int(1); need(n >= 0, '#VALUE!'); const s = c.str(0); need(s.length * n <= 32767, '#VALUE!'); return s.repeat(n) })
def('EXACT', 2, 2, c => c.str(0) === c.str(1))
def('FIND', 2, 3, c => { const start = c.has(2) ? c.int(2) : 1; need(start >= 1, '#VALUE!'); const i = c.str(1).indexOf(c.str(0), start - 1); need(i >= 0, '#VALUE!'); return i + 1 })
def('SEARCH', 2, 3, c => {
  const start = c.has(2) ? c.int(2) : 1; need(start >= 1, '#VALUE!')
  const hay = c.str(1), re = wildcardRe('*' + c.str(0) + '*')
  for (let i = start - 1; i < hay.length; i++) { for (let j = hay.length; j >= i; j--) if (re.test(hay.slice(0, i) + hay.slice(i, j) + hay.slice(j))) break }
  const i = hay.toLowerCase().indexOf(c.str(0).toLowerCase(), start - 1)
  need(i >= 0, '#VALUE!'); return i + 1
})
def('SUBSTITUTE', 3, 4, c => {
  const s = c.str(0), find = c.str(1), rep = c.str(2)
  if (find === '') return s
  if (!c.has(3)) return s.split(find).join(rep)
  const nth = c.int(3); need(nth >= 1, '#VALUE!')
  let i = -1, count = 0
  while (count < nth) { i = s.indexOf(find, i + 1); if (i < 0) return s; count++ }
  return s.slice(0, i) + rep + s.slice(i + find.length)
})
def('REPLACE', 4, 4, c => { const s = c.str(0), start = c.int(1), n = c.int(2); need(start >= 1 && n >= 0, '#VALUE!'); return s.slice(0, start - 1) + c.str(3) + s.slice(start - 1 + n) })
def('CHAR', 1, 1, c => { const n = c.int(0); need(n >= 1 && n <= 255, '#VALUE!'); return String.fromCharCode(n) })
def('UNICHAR', 1, 1, c => { const n = c.int(0); need(n >= 1 && n <= 0x10ffff, '#VALUE!'); return String.fromCodePoint(n) })
def('CODE UNICODE', 1, 1, c => { const s = c.str(0); need(s.length > 0, '#VALUE!'); return s.codePointAt(0)! })
def('VALUE', 1, 1, c => { const s = c.str(0).trim().replace(/[,\s]/g, ''); const n = Number(s.replace(/^\$/, '').replace(/%$/, '')); need(!Number.isNaN(n) && s !== '', '#VALUE!'); return s.endsWith('%') ? n / 100 : n })
def('NUMBERVALUE', 1, 3, c => { const n = Number(c.str(0).trim()); need(!Number.isNaN(n), '#VALUE!'); return n })
def('TEXTBEFORE', 2, 2, c => { const s = c.str(0), d = c.str(1), i = s.indexOf(d); return i < 0 ? err('#N/A') : s.slice(0, i) })
def('TEXTAFTER', 2, 2, c => { const s = c.str(0), d = c.str(1), i = s.indexOf(d); return i < 0 ? err('#N/A') : s.slice(i + d.length) })
def('FIXED', 1, 3, c => {
  const n = c.num(0), d = c.has(1) ? c.int(1) : 2, noCommas = c.has(2) && c.bool(2)
  const s = round(n, d).toFixed(Math.max(0, d))
  return noCommas ? s : Number(s).toLocaleString('en-US', { minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d) })
})
def('DOLLAR', 1, 2, c => {
  const d = c.has(1) ? c.int(1) : 2
  return round(c.num(0), d).toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d) })
})
/** TEXT with the common Excel number/date patterns. */
def('TEXT', 2, 2, c => {
  const v = c.scalar(0), f = c.str(1)
  if (isErr(v)) throw v
  const n = typeof v === 'number' ? v : Number(toStr(v))
  if (/[ymdhs]/i.test(f) && !Number.isNaN(n)) return formatSerial(n, f)
  if (Number.isNaN(n)) return toStr(v)
  if (f === 'General' || f === '') return toStr(n)
  if (/%/.test(f)) { const d = (f.split('.')[1] ?? '').replace(/[^0#]/g, '').length; return (n * 100).toFixed(d) + '%' }
  const dec = (f.split('.')[1] ?? '').replace(/[^0#]/g, '').length
  const grouped = f.includes(',')
  const body = grouped ? Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec }) : Math.abs(n).toFixed(dec)
  const sign = n < 0 ? '-' : ''
  return (f.includes('$') ? sign + '$' + body : sign + body)
})
function pad(n: number, w = 2) { return String(Math.floor(Math.abs(n))).padStart(w, '0') }
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
function formatSerial(serial: number, f: string): string {
  const d = fromSerial(serial), secs = timePart(serial)
  const h24 = Math.floor(secs / 3600), mi = Math.floor((secs % 3600) / 60), se = secs % 60
  const ampm = /am\/pm|a\/p/i.test(f)
  const h = ampm ? (h24 % 12 || 12) : h24
  return f.replace(/\\?(am\/pm|a\/p|yyyy|yyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|.)/gi, tok => {
    const t = tok.toLowerCase()
    switch (t) {
      case 'yyyy': case 'yyy': return String(d.getUTCFullYear())
      case 'yy': return pad(d.getUTCFullYear() % 100)
      case 'mmmm': return MONTHS[d.getUTCMonth()]
      case 'mmm': return MONTHS[d.getUTCMonth()].slice(0, 3)
      case 'dddd': return DAYS[d.getUTCDay()]
      case 'ddd': return DAYS[d.getUTCDay()].slice(0, 3)
      case 'dd': return pad(d.getUTCDate())
      case 'd': return String(d.getUTCDate())
      case 'hh': return pad(h)
      case 'h': return String(h)
      case 'ss': return pad(se)
      case 's': return String(se)
      case 'am/pm': case 'a/p': return h24 < 12 ? 'AM' : 'PM'
      case 'mm': return /[hs]/i.test(f) && /[h:]\s*$|:\s*$/.test(f.slice(0, f.indexOf(tok))) ? pad(mi) : pad(d.getUTCMonth() + 1)
      case 'm': return pad(d.getUTCMonth() + 1).replace(/^0/, '')
      default: return tok
    }
  })
}

/* ---- date and time ---- */
def('TODAY', 0, 0, () => todaySerial(), true)
def('NOW', 0, 0, () => nowSerial(), true)
def('DATE', 3, 3, c => { const y = c.int(0), m = c.int(1), d = c.int(2); return (Date.UTC(y < 1900 ? y + 1900 : y, m - 1, d) - EPOCH) / 86400000 })
def('TIME', 3, 3, c => ((c.int(0) * 3600 + c.int(1) * 60 + c.int(2)) % 86400) / 86400)
def('YEAR', 1, 1, c => fromSerial(c.num(0)).getUTCFullYear())
def('MONTH', 1, 1, c => fromSerial(c.num(0)).getUTCMonth() + 1)
def('DAY', 1, 1, c => fromSerial(c.num(0)).getUTCDate())
def('HOUR', 1, 1, c => Math.floor(timePart(c.num(0)) / 3600))
def('MINUTE', 1, 1, c => Math.floor((timePart(c.num(0)) % 3600) / 60))
def('SECOND', 1, 1, c => timePart(c.num(0)) % 60)
def('WEEKDAY', 1, 2, c => { const d = fromSerial(c.num(0)).getUTCDay(), t = c.has(1) ? c.int(1) : 1; return t === 2 ? (d === 0 ? 7 : d) : t === 3 ? (d + 6) % 7 : d + 1 })
def('WEEKNUM', 1, 2, c => { const d = fromSerial(c.num(0)); const start = Date.UTC(d.getUTCFullYear(), 0, 1); return Math.floor((Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - start) / 86400000 / 7) + 1 })
def('DAYS', 2, 2, c => Math.floor(c.num(0)) - Math.floor(c.num(1)))
def('DAYS360', 2, 3, c => {
  const a = fromSerial(c.num(0)), b = fromSerial(c.num(1))
  const d1 = Math.min(a.getUTCDate(), 30), d2 = b.getUTCDate() > 30 && d1 === 30 ? 30 : b.getUTCDate()
  return (b.getUTCFullYear() - a.getUTCFullYear()) * 360 + (b.getUTCMonth() - a.getUTCMonth()) * 30 + (d2 - d1)
})
def('EDATE', 2, 2, c => { const d = fromSerial(c.num(0)), m = c.int(1); const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m, 1)); const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate(); return (Date.UTC(t.getUTCFullYear(), t.getUTCMonth(), Math.min(d.getUTCDate(), last)) - EPOCH) / 86400000 })
def('EOMONTH', 2, 2, c => { const d = fromSerial(c.num(0)), m = c.int(1); return (Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m + 1, 0) - EPOCH) / 86400000 })
def('DATEVALUE', 1, 1, c => { const t = Date.parse(c.str(0) + 'T00:00:00Z'); const u = Number.isNaN(t) ? Date.parse(c.str(0)) : t; need(!Number.isNaN(u), '#VALUE!'); const d = new Date(u); return (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - EPOCH) / 86400000 })
def('TIMEVALUE', 1, 1, c => { const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(c.str(0).trim()); need(!!m, '#VALUE!'); let h = Number(m![1]); if (/pm/i.test(m![4] ?? '') && h < 12) h += 12; if (/am/i.test(m![4] ?? '') && h === 12) h = 0; return (h * 3600 + Number(m![2]) * 60 + Number(m![3] ?? 0)) / 86400 })
def('DATEDIF', 3, 3, c => {
  const a = fromSerial(c.num(0)), b = fromSerial(c.num(1)), u = c.str(2).toUpperCase()
  need(c.num(1) >= c.num(0), '#NUM!')
  let years = b.getUTCFullYear() - a.getUTCFullYear(), months = years * 12 + (b.getUTCMonth() - a.getUTCMonth())
  if (b.getUTCDate() < a.getUTCDate()) months--
  years = Math.floor(months / 12)
  switch (u) {
    case 'Y': return years
    case 'M': return months
    case 'D': return Math.floor(c.num(1)) - Math.floor(c.num(0))
    case 'MD': { const prev = new Date(Date.UTC(b.getUTCFullYear(), b.getUTCMonth(), 0)).getUTCDate(); return b.getUTCDate() >= a.getUTCDate() ? b.getUTCDate() - a.getUTCDate() : b.getUTCDate() + prev - a.getUTCDate() }
    case 'YM': return months - years * 12
    case 'YD': { const anniv = Date.UTC(b.getUTCFullYear() - (b.getUTCMonth() < a.getUTCMonth() || (b.getUTCMonth() === a.getUTCMonth() && b.getUTCDate() < a.getUTCDate()) ? 1 : 0), a.getUTCMonth(), a.getUTCDate()); return Math.round((Date.UTC(b.getUTCFullYear(), b.getUTCMonth(), b.getUTCDate()) - anniv) / 86400000) }
    default: throw err('#NUM!')
  }
})
const isWeekend = (serial: number) => { const d = fromSerial(serial).getUTCDay(); return d === 0 || d === 6 }
def('NETWORKDAYS', 2, 3, c => {
  const a = Math.floor(c.num(0)), b = Math.floor(c.num(1))
  const holidays = new Set(c.has(2) ? c.flat(2).filter(v => typeof v === 'number').map(v => Math.floor(v as number)) : [])
  const [lo, hi] = a <= b ? [a, b] : [b, a]
  let n = 0
  for (let s = lo; s <= hi; s++) if (!isWeekend(s) && !holidays.has(s)) n++
  return a <= b ? n : -n
})
def('WORKDAY', 2, 3, c => {
  let s = Math.floor(c.num(0)); let left = c.int(1)
  const holidays = new Set(c.has(2) ? c.flat(2).filter(v => typeof v === 'number').map(v => Math.floor(v as number)) : [])
  const step = left >= 0 ? 1 : -1
  left = Math.abs(left)
  while (left > 0) { s += step; if (!isWeekend(s) && !holidays.has(s)) left-- }
  return s
})

/* ---- lookup and reference ---- */
def('CHOOSE', 2, INF, c => { const i = c.int(0); need(i >= 1 && i < c.n, '#VALUE!'); return c.scalar(i) })
def('VLOOKUP', 3, 4, c => {
  const key = c.scalar(0); if (isErr(key)) throw key
  const table = c.grid(1), col = c.int(2)
  const approx = !c.has(3) || c.bool(3)
  need(col >= 1, '#VALUE!'); need(col <= (table[0]?.length ?? 0), '#REF!')
  const ci = col - 1
  if (!approx) {
    const row = table.find(r => looseEq(r[0], key) || (typeof key === 'string' && /[*?]/.test(key) && wildcardRe(key).test(toStr(r[0]))))
    return row ? row[ci] : err('#N/A')
  }
  let best = -1
  for (let i = 0; i < table.length; i++) { const v = table[i][0]; if (v === null || isErr(v)) continue; if (typeof v !== typeof key) continue; if (compare(v, key) <= 0) best = i; else break }
  return best >= 0 ? table[best][ci] : err('#N/A')
})
def('HLOOKUP', 3, 4, c => {
  const key = c.scalar(0); if (isErr(key)) throw key
  const table = c.grid(1), row = c.int(2)
  const approx = !c.has(3) || c.bool(3)
  need(row >= 1, '#VALUE!'); need(row <= table.length, '#REF!')
  const header = table[0] ?? []
  let hit = -1
  for (let i = 0; i < header.length; i++) {
    const v = header[i]
    if (approx) { if (v !== null && typeof v === typeof key && compare(v, key) <= 0) hit = i; else if (v !== null && typeof v === typeof key) break }
    else if (looseEq(v, key)) { hit = i; break }
  }
  return hit >= 0 ? table[row - 1][hit] : err('#N/A')
})
def('LOOKUP', 2, 3, c => {
  const key = c.scalar(0), vec = c.flat(1), res = c.has(2) ? c.flat(2) : vec
  let best = -1
  for (let i = 0; i < vec.length; i++) { const v = vec[i]; if (v === null || isErr(v)) continue; if (compare(v, key) <= 0) best = i }
  return best >= 0 ? res[best] ?? err('#N/A') : err('#N/A')
})
def('MATCH', 2, 3, c => {
  const key = c.scalar(0), vec = c.flat(1), type = c.has(2) ? c.int(2) : 1
  if (type === 0) {
    const i = vec.findIndex(v => looseEq(v, key) || (typeof key === 'string' && /[*?]/.test(key) && wildcardRe(key).test(toStr(v))))
    need(i >= 0, '#N/A'); return i + 1
  }
  let best = -1
  for (let i = 0; i < vec.length; i++) { const v = vec[i]; if (v === null || isErr(v)) continue; if (type === 1 ? compare(v, key) <= 0 : compare(v, key) >= 0) best = i; else break }
  need(best >= 0, '#N/A'); return best + 1
})
def('XMATCH', 2, 4, c => {
  const key = c.scalar(0), vec = c.flat(1), mode = c.has(2) ? c.int(2) : 0, dir = c.has(3) ? c.int(3) : 1
  const order = Array.from(vec.keys()); if (dir < 0) order.reverse()
  for (const i of order) if (mode === 2 ? wildcardRe(toStr(key)).test(toStr(vec[i])) : looseEq(vec[i], key)) return i + 1
  if (mode === 1 || mode === -1) {
    let hit = -1
    for (const i of order) { const v = vec[i]; if (v === null || isErr(v)) continue; const d = compare(v, key); if ((mode === 1 && d > 0) || (mode === -1 && d < 0)) { if (hit < 0 || (mode === 1 ? compare(v, vec[hit]) < 0 : compare(v, vec[hit]) > 0)) hit = i } }
    if (hit >= 0) return hit + 1
  }
  return err('#N/A')
})
def('XLOOKUP', 3, 6, c => {
  const key = c.scalar(0); if (isErr(key)) throw key
  const look = c.flat(1), ret = c.grid(2)
  const retFlat = ret.length === 1 ? ret[0] : ret.map(r => r[0])
  need(look.length === retFlat.length, '#VALUE!')
  const mode = c.has(4) ? c.int(4) : 0, dir = c.has(5) ? c.int(5) : 1
  const order = Array.from(look.keys()); if (dir < 0) order.reverse()
  let hit = -1
  for (const i of order) if (mode === 2 ? wildcardRe(toStr(key)).test(toStr(look[i])) : looseEq(look[i], key)) { hit = i; break }
  if (hit < 0 && (mode === 1 || mode === -1)) {
    for (const i of order) { const v = look[i]; if (v === null || isErr(v)) continue; const d = compare(v, key); if ((mode === 1 && d > 0) || (mode === -1 && d < 0)) { if (hit < 0 || (mode === 1 ? compare(v, look[hit]) < 0 : compare(v, look[hit]) > 0)) hit = i } }
  }
  if (hit >= 0) return retFlat[hit]
  return c.has(3) ? c.scalar(3) : err('#N/A')
})
def('INDEX', 2, 3, c => {
  const g = c.grid(0), r = c.int(1), col = c.has(2) ? c.int(2) : 0
  const oneRow = g.length === 1
  if (r === 0 && col === 0) return g
  if (col === 0 && oneRow) { need(r >= 1 && r <= g[0].length, '#REF!'); return g[0][r - 1] }
  if (col === 0) { need(r >= 1 && r <= g.length, '#REF!'); return g[r - 1].length === 1 ? g[r - 1][0] : [g[r - 1]] }
  need(r >= 1 && r <= g.length, '#REF!'); need(col >= 1 && col <= g[0].length, '#REF!')
  return g[r - 1][col - 1]
})
def('ROWS', 1, 1, c => c.grid(0).length)
def('COLUMNS', 1, 1, c => c.grid(0)[0]?.length ?? 0)
def('TRANSPOSE', 1, 1, c => { const g = c.grid(0); return g[0].map((_, i) => g.map(r => r[i])) })
def('UNIQUE', 1, 3, c => {
  const g = c.grid(0)
  const seen = new Set<string>(), out: Scalar[][] = []
  for (const row of g) { const k = row.map(toStr).join('\u0000'); if (!seen.has(k)) { seen.add(k); out.push(row) } }
  return out.length ? out : [[err('#CALC!')]]
})
def('SORT', 1, 4, c => {
  const g = c.grid(0).map(r => [...r])
  const idx = (c.has(1) ? c.int(1) : 1) - 1, asc = !c.has(2) || c.num(2) >= 0
  need(idx >= 0 && idx < (g[0]?.length ?? 0), '#VALUE!')
  g.sort((a, b) => (asc ? 1 : -1) * compare(a[idx], b[idx]))
  return g
})
def('FILTER', 2, 3, c => {
  const g = c.grid(0), cond = c.flat(1)
  need(cond.length === g.length || cond.length === (g[0]?.length ?? 0), '#VALUE!')
  const out = g.filter((_, i) => { const v = cond[i]; return typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false })
  if (out.length) return out
  return c.has(2) ? c.scalar(2) : err('#CALC!')
})
def('SEQUENCE', 1, 4, c => {
  const rows = c.int(0), cols = c.has(1) ? c.int(1) : 1, start = c.has(2) ? c.num(2) : 1, step = c.has(3) ? c.num(3) : 1
  need(rows >= 1 && cols >= 1, '#VALUE!'); need(rows * cols <= 100000, '#NUM!')
  const out: Scalar[][] = []
  let v = start
  for (let r = 0; r < rows; r++) { const row: Scalar[] = []; for (let col = 0; col < cols; col++) { row.push(v); v += step } out.push(row) }
  return out
})

/* ---- financial ---- */
def('PMT', 3, 5, c => pmtOf(c.num(0), c.num(1), c.num(2), c.has(3) ? c.num(3) : 0, c.has(4) ? c.int(4) : 0))
def('FV', 3, 5, c => fvOf(c.num(0), c.num(1), c.num(2), c.has(3) ? c.num(3) : 0, c.has(4) ? c.int(4) : 0))
def('PV', 3, 5, c => {
  const rate = c.num(0), nper = c.num(1), pmt = c.num(2), fv = c.has(3) ? c.num(3) : 0, type = c.has(4) ? c.int(4) : 0
  if (rate === 0) return -(fv + pmt * nper)
  const p = Math.pow(1 + rate, nper)
  return -(fv + pmt * (1 + rate * type) * (p - 1) / rate) / p
})
def('NPER', 3, 5, c => {
  const rate = c.num(0), pmt = c.num(1), pv = c.num(2), fv = c.has(3) ? c.num(3) : 0, type = c.has(4) ? c.int(4) : 0
  if (rate === 0) { need(pmt !== 0, '#NUM!'); return -(pv + fv) / pmt }
  const a = pmt * (1 + rate * type)
  const v = (a - fv * rate) / (a + pv * rate)
  need(v > 0, '#NUM!')
  return Math.log(v) / Math.log(1 + rate)
})
def('RATE', 3, 6, c => {
  const nper = c.num(0), pmt = c.num(1), pv = c.num(2), fv = c.has(3) ? c.num(3) : 0, type = c.has(4) ? c.int(4) : 0
  let r = c.has(5) ? c.num(5) : 0.1
  for (let i = 0; i < 100; i++) {
    const f = fvOf(r, nper, pmt, pv, type) - fv
    const d = (fvOf(r + 1e-6, nper, pmt, pv, type) - fv - f) / 1e-6
    if (Math.abs(d) < 1e-12) break
    const next = r - f / d
    if (!Number.isFinite(next)) break
    if (Math.abs(next - r) < 1e-10) return next
    r = next
  }
  need(Number.isFinite(r), '#NUM!'); return r
})
def('NPV', 2, INF, c => { const rate = c.num(0); let s = 0; let i = 1; for (let a = 1; a < c.n; a++) for (const v of c.flat(a)) { if (typeof v === 'number') s += v / Math.pow(1 + rate, i); i++ } return s })
def('IRR', 1, 2, c => {
  const flows = c.flat(0).filter(v => typeof v === 'number') as number[]
  need(flows.length > 1, '#NUM!')
  let r = c.has(1) ? c.num(1) : 0.1
  for (let i = 0; i < 200; i++) {
    let f = 0, d = 0
    flows.forEach((v, t) => { f += v / Math.pow(1 + r, t); d -= (t * v) / Math.pow(1 + r, t + 1) })
    if (Math.abs(d) < 1e-12) break
    const next = r - f / d
    if (!Number.isFinite(next) || next <= -1) break
    if (Math.abs(next - r) < 1e-10) return next
    r = next
  }
  throw err('#NUM!')
})
def('SLN', 3, 3, c => { const life = c.num(2); need(life !== 0, '#DIV/0!'); return (c.num(0) - c.num(1)) / life })
def('SYD', 4, 4, c => { const cost = c.num(0), sal = c.num(1), life = c.num(2), per = c.num(3); need(life > 0, '#NUM!'); return ((cost - sal) * (life - per + 1) * 2) / (life * (life + 1)) })
def('IPMT', 4, 6, c => {
  const rate = c.num(0), per = c.num(1), nper = c.num(2), pv = c.num(3), fv = c.has(4) ? c.num(4) : 0, type = c.has(5) ? c.int(5) : 0
  need(per >= 1 && per <= nper, '#NUM!')
  const pmt = pmtOf(rate, nper, pv, fv, type)
  const bal = fvOf(rate, per - 1, pmt, pv, type)
  return per === 1 && type === 1 ? 0 : bal * rate
})
def('PPMT', 4, 6, c => {
  const rate = c.num(0), per = c.num(1), nper = c.num(2), pv = c.num(3), fv = c.has(4) ? c.num(4) : 0, type = c.has(5) ? c.int(5) : 0
  const pmt = pmtOf(rate, nper, pv, fv, type)
  const bal = fvOf(rate, per - 1, pmt, pv, type)
  return pmt - (per === 1 && type === 1 ? 0 : bal * rate)
})

/* ---- engineering ---- */
const baseConv = (s: string, from: number) => { const n = parseInt(s.trim(), from); need(!Number.isNaN(n), '#NUM!'); return n }
def('DEC2BIN', 1, 2, c => { const n = c.int(0); need(n >= -512 && n < 512, '#NUM!'); return (n < 0 ? (n >>> 0).toString(2).slice(-10) : n.toString(2)).padStart(c.has(1) ? c.int(1) : 0, '0') })
def('DEC2HEX', 1, 2, c => { const n = c.int(0); return (n < 0 ? (n >>> 0).toString(16) : n.toString(16)).toUpperCase().padStart(c.has(1) ? c.int(1) : 0, '0') })
def('DEC2OCT', 1, 2, c => c.int(0).toString(8).padStart(c.has(1) ? c.int(1) : 0, '0'))
def('BIN2DEC', 1, 1, c => { const s = c.str(0); const n = baseConv(s, 2); return s.length === 10 && s[0] === '1' ? n - 1024 : n })
def('HEX2DEC', 1, 1, c => baseConv(c.str(0), 16))
def('OCT2DEC', 1, 1, c => baseConv(c.str(0), 8))
def('BASE', 2, 3, c => { const r = c.int(1); need(r >= 2 && r <= 36, '#NUM!'); return c.int(0).toString(r).toUpperCase().padStart(c.has(2) ? c.int(2) : 0, '0') })
def('DECIMAL', 2, 2, c => { const r = c.int(1); need(r >= 2 && r <= 36, '#NUM!'); return baseConv(c.str(0), r) })
def('BITAND', 2, 2, c => c.int(0) & c.int(1))
def('BITOR', 2, 2, c => c.int(0) | c.int(1))
def('BITXOR', 2, 2, c => c.int(0) ^ c.int(1))

/** Function names in the order shown in the Formulas ribbon. */
export const FUNCTION_GROUPS: Record<string, string[]> = {
  'Math & Trig': ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'PRODUCT', 'ABS', 'SIGN', 'SQRT', 'POWER', 'EXP', 'LN', 'LOG', 'MOD', 'QUOTIENT', 'INT', 'TRUNC', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'MROUND', 'CEILING', 'FLOOR', 'EVEN', 'ODD', 'FACT', 'COMBIN', 'GCD', 'LCM', 'PI', 'DEGREES', 'RADIANS', 'SIN', 'COS', 'TAN', 'RAND', 'RANDBETWEEN'],
  Statistical: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MIN', 'MAX', 'MINIFS', 'MAXIFS', 'MEDIAN', 'MODE', 'STDEV', 'STDEVP', 'VAR', 'VARP', 'LARGE', 'SMALL', 'RANK', 'PERCENTILE', 'QUARTILE', 'GEOMEAN', 'HARMEAN', 'CORREL', 'SLOPE', 'INTERCEPT'],
  Logical: ['IF', 'IFS', 'SWITCH', 'AND', 'OR', 'XOR', 'NOT', 'TRUE', 'FALSE', 'IFERROR', 'IFNA'],
  Text: ['CONCAT', 'CONCATENATE', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'CLEAN', 'REPT', 'EXACT', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'TEXT', 'VALUE', 'FIXED', 'DOLLAR', 'CHAR', 'CODE', 'TEXTBEFORE', 'TEXTAFTER'],
  'Date & Time': ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'DAYS', 'DAYS360', 'EDATE', 'EOMONTH', 'DATEDIF', 'DATEVALUE', 'TIMEVALUE', 'NETWORKDAYS', 'WORKDAY'],
  'Lookup & Reference': ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'XMATCH', 'CHOOSE', 'ROWS', 'COLUMNS', 'TRANSPOSE', 'UNIQUE', 'SORT', 'FILTER', 'SEQUENCE'],
  Financial: ['PMT', 'PV', 'FV', 'NPER', 'RATE', 'NPV', 'IRR', 'IPMT', 'PPMT', 'SLN', 'SYD'],
  Information: ['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'ISERROR', 'ISERR', 'ISNA', 'ISEVEN', 'ISODD', 'NA', 'N', 'T', 'TYPE', 'ERROR.TYPE'],
  Engineering: ['DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'BIN2DEC', 'HEX2DEC', 'OCT2DEC', 'BASE', 'DECIMAL', 'BITAND', 'BITOR', 'BITXOR'],
}

/** Short signature hints shown by the formula autocomplete. */
export const FUNCTION_HINTS: Record<string, string> = {
  SUM: 'number1, [number2], …', AVERAGE: 'number1, [number2], …', COUNT: 'value1, [value2], …', COUNTA: 'value1, [value2], …',
  MIN: 'number1, [number2], …', MAX: 'number1, [number2], …', IF: 'logical_test, value_if_true, [value_if_false]',
  IFS: 'test1, value1, [test2, value2], …', SWITCH: 'expression, value1, result1, …, [default]',
  AND: 'logical1, [logical2], …', OR: 'logical1, [logical2], …', NOT: 'logical', XOR: 'logical1, [logical2], …',
  IFERROR: 'value, value_if_error', IFNA: 'value, value_if_na',
  ROUND: 'number, [num_digits]', ROUNDUP: 'number, [num_digits]', ROUNDDOWN: 'number, [num_digits]',
  ABS: 'number', PRODUCT: 'number1, [number2], …', MOD: 'number, divisor', POWER: 'number, power', SQRT: 'number',
  COUNTIF: 'range, criteria', COUNTIFS: 'range1, criteria1, …', SUMIF: 'range, criteria, [sum_range]',
  SUMIFS: 'sum_range, range1, criteria1, …', AVERAGEIF: 'range, criteria, [average_range]', AVERAGEIFS: 'average_range, range1, criteria1, …',
  VLOOKUP: 'lookup_value, table_array, col_index, [range_lookup]', HLOOKUP: 'lookup_value, table_array, row_index, [range_lookup]',
  XLOOKUP: 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]',
  INDEX: 'array, row_num, [column_num]', MATCH: 'lookup_value, lookup_array, [match_type]', CHOOSE: 'index_num, value1, …',
  CONCAT: 'text1, [text2], …', CONCATENATE: 'text1, [text2], …', TEXTJOIN: 'delimiter, ignore_empty, text1, …',
  LEFT: 'text, [num_chars]', RIGHT: 'text, [num_chars]', MID: 'text, start_num, num_chars', LEN: 'text',
  UPPER: 'text', LOWER: 'text', PROPER: 'text', TRIM: 'text', TEXT: 'value, format_text',
  SUBSTITUTE: 'text, old_text, new_text, [instance]', REPLACE: 'old_text, start_num, num_chars, new_text',
  FIND: 'find_text, within_text, [start_num]', SEARCH: 'find_text, within_text, [start_num]',
  TODAY: '( )', NOW: '( )', DATE: 'year, month, day', TIME: 'hour, minute, second', YEAR: 'serial', MONTH: 'serial', DAY: 'serial',
  EOMONTH: 'start_date, months', EDATE: 'start_date, months', DATEDIF: 'start_date, end_date, unit',
  NETWORKDAYS: 'start_date, end_date, [holidays]', WORKDAY: 'start_date, days, [holidays]',
  PMT: 'rate, nper, pv, [fv], [type]', FV: 'rate, nper, pmt, [pv], [type]', PV: 'rate, nper, pmt, [fv], [type]',
  NPER: 'rate, pmt, pv, [fv], [type]', RATE: 'nper, pmt, pv, [fv], [type], [guess]', NPV: 'rate, value1, …', IRR: 'values, [guess]',
  UNIQUE: 'array', SORT: 'array, [sort_index], [sort_order]', FILTER: 'array, include, [if_empty]', SEQUENCE: 'rows, [columns], [start], [step]',
  TRANSPOSE: 'array', SUMPRODUCT: 'array1, [array2], …', MEDIAN: 'number1, …', LARGE: 'array, k', SMALL: 'array, k',
}

export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort()
