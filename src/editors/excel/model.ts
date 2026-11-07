export interface CellStyle {
  b?: boolean; i?: boolean; u?: boolean; strike?: boolean; color?: string; bg?: string
  font?: string; size?: number; wrap?: boolean
  /** Digits after the decimal point for number, currency and percent formats. */
  dec?: number
  align?: 'left' | 'center' | 'right'
  fmt?: 'general' | 'number' | 'currency' | 'percent' | 'date' | 'time' | 'scientific'
  border?: boolean
}
export interface Cell { v: string; s?: CellStyle }
export type CondKind = 'gt' | 'lt' | 'eq' | 'between' | 'text' | 'dup' | 'top' | 'bottom' | 'scale'
/** A conditional-format rule applied to an A1-style range. */
export interface CondRule { id: string; range: string; kind: CondKind; a?: string; b?: string; bg?: string; color?: string }
export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'area'
/** A chart built from a data range: first row = series names, first column = categories. */
export interface Chart { id: string; type: ChartType; range: string; title: string; x: number; y: number; w: number; h: number }
export interface Sheet {
  id: string; name: string
  cells: Record<string, Cell>
  colW: Record<number, number>
  rows: number; cols: number
  /** Rows and columns kept visible while scrolling. */
  freeze?: { rows: number; cols: number }
  cf?: CondRule[]
  charts?: Chart[]
}
export interface Workbook { sheets: Sheet[]; active: number }

export const DEFAULT_ROWS = 100
export const DEFAULT_COLS = 26

export function colName(i: number): string {
  let s = ''
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}
export function colIndex(name: string): number {
  let n = 0
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
export const toAddr = (c: number, r: number) => `${colName(c)}${r + 1}`
export function parseAddr(a: string): { c: number; r: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a.trim())
  return m ? { c: colIndex(m[1]), r: Number(m[2]) - 1 } : null
}

let seq = 0
export const newSheet = (name: string): Sheet => ({
  id: `s${Date.now().toString(36)}${seq++}`, name, cells: {}, colW: {}, rows: DEFAULT_ROWS, cols: DEFAULT_COLS,
})
export const emptyWorkbook = (): Workbook => ({ sheets: [newSheet('Sheet1')], active: 0 })

const ONE = '(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)'
const REF_RE = new RegExp(`((?:'[^']+'|[A-Za-z_][\\w.]*)!)?${ONE}(?::${ONE})?(?![\\w(])`, 'g')
interface Pt { c: number; r: number }

/**
 * Rewrite every cell reference in a formula (text inside "quotes" is left alone).
 * `fn` maps single references; `rangeFn`, when given, maps whole A1:B2 ranges at once.
 */
export function mapRefs(
  formula: string,
  fn: (c: number, r: number, absC: boolean, absR: boolean, sheetPrefix: string | undefined) => Pt | null,
  rangeFn?: (a: Pt, b: Pt, sheetPrefix: string | undefined) => [Pt, Pt] | null,
): string {
  return formula.split(/("(?:[^"]|"")*")/).map((part, i) => i % 2 === 1 ? part : part.replace(REF_RE, (_m, pre: string | undefined, ca: string, col: string, ra: string, row: string, cb?: string, col2?: string, rb?: string, row2?: string) => {
    const p = pre ? pre.slice(0, -1) : undefined
    const fmt = (pt: Pt, abs1: string, abs2: string) => `${abs1}${colName(pt.c)}${abs2}${pt.r + 1}`
    if (col2 === undefined) {
      const out = fn(colIndex(col), Number(row) - 1, !!ca, !!ra, p)
      return out ? `${pre ?? ''}${fmt(out, ca, ra)}` : '#REF!'
    }
    const A = { c: colIndex(col), r: Number(row) - 1 }, B = { c: colIndex(col2), r: Number(row2) - 1 }
    if (rangeFn) {
      const out = rangeFn(A, B, p)
      return out ? `${pre ?? ''}${fmt(out[0], ca, ra)}:${fmt(out[1], cb ?? '', rb ?? '')}` : '#REF!'
    }
    const oa = fn(A.c, A.r, !!ca, !!ra, p), ob = fn(B.c, B.r, !!cb, !!rb, p)
    return oa && ob ? `${pre ?? ''}${fmt(oa, ca, ra)}:${fmt(ob, cb ?? '', rb ?? '')}` : '#REF!'
  })).join('')
}

/** Shift relative references in a formula (used when pasting/filling). */
export function shiftFormula(f: string, dc: number, dr: number): string {
  return mapRefs(f, (c, r, absC, absR) => {
    const nc = absC ? c : c + dc, nr = absR ? r : r + dr
    return nc < 0 || nr < 0 ? null : { c: nc, r: nr }
  })
}

type Axis = 'row' | 'col'
/** Insert (n>0) or delete (n<0) rows/columns on one sheet and rewrite formulas that point into it. */
function restructure(wb: Workbook, si: number, axis: Axis, at: number, n: number): Workbook {
  const target = wb.sheets[si]
  const norm = (x: string | undefined) => x?.replace(/^'|'$/g, '').toLowerCase()
  const shiftIdx = (i: number): number | null => {
    if (n > 0) return i >= at ? i + n : i
    const cnt = -n
    if (i >= at && i < at + cnt) return null
    return i >= at + cnt ? i - cnt : i
  }
  const sheets = wb.sheets.map((sh, idx) => {
    const cells: Record<string, Cell> = {}
    for (const [addr, cell] of Object.entries(sh.cells)) {
      let v = cell.v
      if (v.startsWith('=')) {
        const points = (pre: string | undefined) => (pre ? norm(pre) === target.name.toLowerCase() : idx === si)
        const at1 = (pt: Pt, x: number): Pt => (axis === 'row' ? { c: pt.c, r: x } : { c: x, r: pt.r })
        v = mapRefs(v, (c, r, _ac, _ar, pre) => {
          if (!points(pre)) return { c, r }
          const x = shiftIdx(axis === 'row' ? r : c); if (x === null) return null
          return at1({ c, r }, x)
        }, (a, b, pre) => {
          if (!points(pre)) return [a, b]
          const ia = axis === 'row' ? a.r : a.c, ib = axis === 'row' ? b.r : b.c
          const lo = Math.min(ia, ib), hi = Math.max(ia, ib)
          if (n > 0) return [at1(a, ia >= at ? ia + n : ia), at1(b, ib >= at ? ib + n : ib)]
          const cnt = -n
          // deleting rows/cols inside a range shrinks it; deleting all of it is a #REF!
          const nlo = lo >= at + cnt ? lo - cnt : lo >= at ? at : lo
          const nhi = hi >= at + cnt ? hi - cnt : hi >= at ? at - 1 : hi
          if (nhi < nlo) return null
          return ia <= ib ? [at1(a, nlo), at1(b, nhi)] : [at1(a, nhi), at1(b, nlo)]
        })
      }
      let key = addr
      if (idx === si) {
        const p = parseAddr(addr)!
        const x = shiftIdx(axis === 'row' ? p.r : p.c)
        if (x === null) continue // the cell was deleted
        key = axis === 'row' ? toAddr(p.c, x) : toAddr(x, p.r)
      }
      cells[key] = v === cell.v ? cell : { ...cell, v }
    }
    if (idx !== si) return { ...sh, cells }
    const colW: Record<number, number> = {}
    for (const [k, w] of Object.entries(sh.colW)) { const x = axis === 'col' ? shiftIdx(Number(k)) : Number(k); if (x !== null) colW[x] = w }
    return { ...sh, cells, colW, rows: axis === 'row' ? Math.max(1, sh.rows + n) : sh.rows, cols: axis === 'col' ? Math.max(1, sh.cols + n) : sh.cols }
  })
  return { ...wb, sheets }
}
export const insertRows = (wb: Workbook, si: number, at: number, n = 1) => restructure(wb, si, 'row', at, n)
export const deleteRows = (wb: Workbook, si: number, at: number, n = 1) => restructure(wb, si, 'row', at, -n)
export const insertCols = (wb: Workbook, si: number, at: number, n = 1) => restructure(wb, si, 'col', at, n)
export const deleteCols = (wb: Workbook, si: number, at: number, n = 1) => restructure(wb, si, 'col', at, -n)
