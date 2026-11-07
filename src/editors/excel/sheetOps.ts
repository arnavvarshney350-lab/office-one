/** Pure sheet operations: sorting, range parsing, conditional formats, find and replace, chart data. */
import { Evaluator, toStr, type Scalar } from './formula'
import { colName, parseAddr, shiftFormula, toAddr, type Cell, type Chart, type CondRule, type Sheet, type Workbook } from './model'
import { FormulaError } from './formulaCore'

export interface Rect { r1: number; c1: number; r2: number; c2: number }

export function parseRange(a: string): Rect | null {
  const [x, y] = a.toUpperCase().split(':')
  const p = parseAddr(x), q = parseAddr(y ?? x)
  if (!p || !q) return null
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) }
}
export const rangeToA1 = (r: Rect) => (r.r1 === r.r2 && r.c1 === r.c2 ? toAddr(r.c1, r.r1) : `${toAddr(r.c1, r.r1)}:${toAddr(r.c2, r.r2)}`)

/* ------------------------------------------------------------------ sort */
const kindOf = (v: Scalar) => (v === null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'boolean' ? 2 : v instanceof FormulaError ? 4 : 1)

/**
 * Sort the rows of `rect` by one column, with Excel's ordering (numbers, then text, then
 * logicals; blanks always last). Formulas move with their rows and have their relative
 * references shifted so they still point at the same row-mates.
 */
export function sortRange(wb: Workbook, si: number, rect: Rect, byCol: number, ascending: boolean, hasHeader: boolean): Workbook {
  const ev = new Evaluator(wb)
  const start = hasHeader ? rect.r1 + 1 : rect.r1
  if (start >= rect.r2) return wb
  const rows: { r: number; key: Scalar; cells: (Cell | undefined)[] }[] = []
  const sheet = wb.sheets[si]
  for (let r = start; r <= rect.r2; r++) {
    const cells: (Cell | undefined)[] = []
    for (let c = rect.c1; c <= rect.c2; c++) cells.push(sheet.cells[toAddr(c, r)])
    rows.push({ r, key: ev.get(si, toAddr(byCol, r), false), cells })
  }
  const dir = ascending ? 1 : -1
  const sorted = [...rows].sort((a, b) => {
    const ka = kindOf(a.key), kb = kindOf(b.key)
    if (ka === 3 || kb === 3) return ka === kb ? 0 : ka === 3 ? 1 : -1 // blanks last whichever the direction
    if (ka !== kb) return dir * (ka - kb)
    if (typeof a.key === 'number' && typeof b.key === 'number') return dir * (a.key - b.key)
    return dir * toStr(a.key).localeCompare(toStr(b.key), undefined, { sensitivity: 'base', numeric: true })
  })
  const cells = { ...sheet.cells }
  for (let i = 0; i < sorted.length; i++) {
    const target = start + i
    const from = sorted[i].r
    sorted[i].cells.forEach((cell, j) => {
      const addr = toAddr(rect.c1 + j, target)
      if (!cell) { delete cells[addr]; return }
      cells[addr] = cell.v.startsWith('=') ? { ...cell, v: shiftRows(cell.v, target - from) } : cell
    })
  }
  return { ...wb, sheets: wb.sheets.map((s, i) => (i === si ? { ...s, cells } : s)) }
}
function shiftRows(formula: string, dr: number): string {
  if (dr === 0) return formula
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g, (m, ca, col, ra, row) => (ra ? m : `${ca}${col}${ra}${Math.max(1, Number(row) + dr)}`))
}

/* ----------------------------------------------------------- conditional formats */
export interface CfResult { bg?: string; color?: string; bar?: number }

/** Evaluate the rules for one cell. Colour scales use the min/max supplied by `scaleOf`. */
export function cfFor(rules: CondRule[] | undefined, addr: string, value: Scalar, ctx: {
  scaleOf: (rule: CondRule) => { min: number; max: number } | null
  rankOf: (rule: CondRule) => number[]
  dupCount: (rule: CondRule, v: Scalar) => number
}): CfResult | null {
  if (!rules?.length) return null
  const p = parseAddr(addr)
  if (!p) return null
  for (const rule of rules) {
    const rg = parseRange(rule.range)
    if (!rg || p.r < rg.r1 || p.r > rg.r2 || p.c < rg.c1 || p.c > rg.c2) continue
    const n = typeof value === 'number' ? value : null
    const a = Number(rule.a), b = Number(rule.b)
    let hit = false
    switch (rule.kind) {
      case 'gt': hit = n !== null && n > a; break
      case 'lt': hit = n !== null && n < a; break
      case 'eq': hit = n !== null && !Number.isNaN(a) ? n === a : toStr(value).toLowerCase() === (rule.a ?? '').toLowerCase(); break
      case 'between': hit = n !== null && n >= Math.min(a, b) && n <= Math.max(a, b); break
      case 'text': hit = value !== null && toStr(value).toLowerCase().includes((rule.a ?? '').toLowerCase()) && (rule.a ?? '') !== ''; break
      case 'dup': hit = value !== null && value !== '' && ctx.dupCount(rule, value) > 1; break
      case 'top': case 'bottom': {
        const ranks = ctx.rankOf(rule); const k = Math.max(1, Number(rule.a) || 1)
        hit = n !== null && (rule.kind === 'top' ? n >= (ranks[k - 1] ?? Infinity) : n <= (ranks[k - 1] ?? -Infinity))
        break
      }
      case 'scale': {
        const sc = ctx.scaleOf(rule)
        if (n !== null && sc && sc.max > sc.min) {
          const t = (n - sc.min) / (sc.max - sc.min)
          // white -> rule colour
          return { bg: mix('#ffffff', rule.bg ?? '#63be7b', t) }
        }
        break
      }
    }
    if (hit) return { bg: rule.bg ?? '#fde68a', color: rule.color }
  }
  return null
}
function mix(a: string, b: string, t: number): string {
  const f = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16)
  const c = [0, 1, 2].map(i => Math.round(f(a, i) + (f(b, i) - f(a, i)) * Math.max(0, Math.min(1, t))))
  return `#${c.map(x => x.toString(16).padStart(2, '0')).join('')}`
}

/** Per-render helper that memoises the range statistics the rules need. */
export function cfContext(ev: Evaluator, si: number) {
  const cache = new Map<string, unknown>()
  const values = (rule: CondRule): Scalar[] => {
    const k = 'v' + rule.id
    if (cache.has(k)) return cache.get(k) as Scalar[]
    const rg = parseRange(rule.range); const out: Scalar[] = []
    if (rg) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push(ev.displayValue(si, toAddr(c, r)))
    cache.set(k, out); return out
  }
  return {
    scaleOf: (rule: CondRule) => { const n = values(rule).filter(v => typeof v === 'number') as number[]; return n.length ? { min: Math.min(...n), max: Math.max(...n) } : null },
    rankOf: (rule: CondRule) => { const k = 'r' + rule.id; if (!cache.has(k)) cache.set(k, (values(rule).filter(v => typeof v === 'number') as number[]).sort((a, b) => (rule.kind === 'top' ? b - a : a - b))); return cache.get(k) as number[] },
    dupCount: (rule: CondRule, v: Scalar) => { const k = 'd' + rule.id; if (!cache.has(k)) { const m = new Map<string, number>(); for (const x of values(rule)) { const s = toStr(x).toLowerCase(); m.set(s, (m.get(s) ?? 0) + 1) } cache.set(k, m) } return (cache.get(k) as Map<string, number>).get(toStr(v).toLowerCase()) ?? 0 },
  }
}

/* ------------------------------------------------------------ find & replace */
export interface FindHit { addr: string }
export function findCells(wb: Workbook, si: number, q: string, opts: { matchCase: boolean; whole: boolean; inFormulas: boolean }): FindHit[] {
  if (!q) return []
  const ev = new Evaluator(wb)
  const needle = opts.matchCase ? q : q.toLowerCase()
  const out: FindHit[] = []
  const cells = wb.sheets[si].cells
  const addrs = Object.keys(cells).sort((a, b) => { const p = parseAddr(a)!, r = parseAddr(b)!; return p.r - r.r || p.c - r.c })
  for (const addr of addrs) {
    const cell = cells[addr]
    const text = opts.inFormulas || !cell.v.startsWith('=') ? cell.v : toStr(ev.get(si, addr, false))
    const hay = opts.matchCase ? text : text.toLowerCase()
    if (opts.whole ? hay === needle : hay.includes(needle)) out.push({ addr })
  }
  return out
}
export function replaceInCell(text: string, q: string, to: string, opts: { matchCase: boolean; whole: boolean }): string {
  if (opts.whole) return (opts.matchCase ? text === q : text.toLowerCase() === q.toLowerCase()) ? to : text
  const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), opts.matchCase ? 'g' : 'gi')
  return text.replace(re, () => to)
}

/* ---------------------------------------------------------------- charts */
export interface ChartData { categories: string[]; series: { name: string; values: number[] }[] }
/** Read a chart's data range: first row = series names, first column = categories. */
export function chartData(ev: Evaluator, si: number, chart: Chart): ChartData {
  const rg = parseRange(chart.range)
  if (!rg) return { categories: [], series: [] }
  const cell = (r: number, c: number) => ev.displayValue(si, toAddr(c, r))
  const hasHeader = rg.r2 > rg.r1 && typeof cell(rg.r1, rg.c1 + 1) !== 'number' && rg.c2 > rg.c1
  const hasCats = rg.c2 > rg.c1 && (rg.r2 === rg.r1 || typeof cell(rg.r1 + (hasHeader ? 1 : 0), rg.c1) !== 'number')
  const firstDataRow = rg.r1 + (hasHeader ? 1 : 0)
  const firstDataCol = rg.c1 + (hasCats ? 1 : 0)
  const categories: string[] = []
  for (let r = firstDataRow; r <= rg.r2; r++) categories.push(hasCats ? toStr(cell(r, rg.c1)) : String(r - firstDataRow + 1))
  const series: ChartData['series'] = []
  for (let c = firstDataCol; c <= rg.c2; c++) {
    const values: number[] = []
    for (let r = firstDataRow; r <= rg.r2; r++) { const v = cell(r, c); values.push(typeof v === 'number' ? v : Number(v) || 0) }
    series.push({ name: hasHeader ? toStr(cell(rg.r1, c)) || colName(c) : `Series ${series.length + 1}`, values })
  }
  return { categories, series }
}
export type { Sheet }

/**
 * Fill handle: extend the source block `S` to the cell `to`, like dragging Excel's fill handle.
 * Two or more numbers continue as a linear series, a lone "Item 1" counts up, formulas shift their
 * references, anything else repeats. Returns null when `to` is inside the block.
 */
export function fillRange(cells: Record<string, Cell>, S: Rect, to: { r: number; c: number }): { cells: Record<string, Cell>; rect: Rect } | null {
  const dv = to.r > S.r2 ? to.r - S.r2 : to.r < S.r1 ? to.r - S.r1 : 0
  const dh = to.c > S.c2 ? to.c - S.c2 : to.c < S.c1 ? to.c - S.c1 : 0
  if (!dv && !dh) return null
  const vertical = Math.abs(dv) >= Math.abs(dh)
  const count = Math.abs(vertical ? dv : dh), sign = (vertical ? dv : dh) > 0 ? 1 : -1
  const lines = vertical ? S.c2 - S.c1 + 1 : S.r2 - S.r1 + 1, n = vertical ? S.r2 - S.r1 + 1 : S.c2 - S.c1 + 1
  const out = { ...cells }
  const isNum = (c?: Cell) => !!c && !c.v.startsWith('=') && c.v.trim() !== '' && Number.isFinite(Number(c.v))
  for (let line = 0; line < lines; line++) {
    const at = (i: number) => (vertical ? toAddr(S.c1 + line, S.r1 + i) : toAddr(S.c1 + i, S.r1 + line))
    const src = Array.from({ length: n }, (_, i) => cells[at(i)])
    const series = n >= 2 && src.every(isNum)
    const first = series ? Number(src[0]!.v) : 0, last = series ? Number(src[n - 1]!.v) : 0, step = series ? (last - first) / (n - 1) : 0
    for (let k = 1; k <= count; k++) {
      const i = sign > 0 ? n - 1 + k : -k
      const from = ((i % n) + n) % n, s = src[from]
      if (!s) { delete out[at(i)]; continue }
      let v = s.v
      const text = /^(.*?)(\d+)$/.exec(v)
      if (series) v = String(Number((sign > 0 ? last + step * k : first - step * k).toFixed(10)))
      else if (v.startsWith('=')) v = vertical ? shiftFormula(v, 0, i - from) : shiftFormula(v, i - from, 0)
      else if (n === 1 && text && !isNum(s) && Number(text[2]) + sign * k >= 0) v = text[1] + String(Number(text[2]) + sign * k).padStart(text[2].length, '0')
      out[at(i)] = { ...s, v }
    }
  }
  const rect = vertical ? (sign > 0 ? { ...S, r2: S.r2 + count } : { ...S, r1: S.r1 - count }) : (sign > 0 ? { ...S, c2: S.c2 + count } : { ...S, c1: S.c1 - count })
  return { cells: out, rect }
}
