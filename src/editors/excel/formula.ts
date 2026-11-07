/**
 * Formula evaluator: tokenizer, parser and a lazy, cached evaluator with dependency tracking
 * and dynamic-array spilling. The function library itself lives in `functions.ts`.
 */
import { colName, parseAddr, type Cell, type CellStyle, type Workbook } from './model'
import { FormulaError, toStr, type Scalar } from './formulaCore'
import { FUNCTIONS, isErr, type Ctx, type Val } from './functions'

export { FormulaError, toStr } from './formulaCore'
export type { Scalar } from './formulaCore'
export { FUNCTION_GROUPS, FUNCTION_HINTS, FUNCTION_NAMES } from './functions'

const err = (c: string) => new FormulaError(c)

/* ============================== tokenizer ============================== */
type TokType = 'num' | 'str' | 'id' | 'op' | 'cell' | 'col' | 'lp' | 'rp' | 'comma' | 'semi' | 'lbrace' | 'rbrace' | 'err'
interface Tok { t: TokType; v: string; sheet?: string }

const PREFIX = "(?:('[^']+'|[A-Za-z_][\\w.]*)!)?"
const RE_CELL = new RegExp(`^${PREFIX}(\\$?[A-Za-z]{1,3}\\$?\\d{1,7})(?![\\w(])`)
const RE_COL = new RegExp(`^${PREFIX}(\\$?[A-Za-z]{1,3})(?![\\w(!])`)
const RE_NUM = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i
const RE_ID = /^[A-Za-z_][\w.]*/
const RE_OP = /^(<=|>=|<>|[+\-*/^&=<>%:])/
const BRACES: Record<string, TokType> = { '(': 'lp', ')': 'rp', '{': 'lbrace', '}': 'rbrace' }

function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (/\s/.test(ch)) { i++; continue }
    if (ch === '"') {
      let s = ''; i++
      while (i < src.length) {
        if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue } break }
        s += src[i++]
      }
      if (src[i] !== '"') throw err('#ERROR!')
      i++; out.push({ t: 'str', v: s }); continue
    }
    const rest = src.slice(i)
    if (rest.startsWith('#REF!')) { out.push({ t: 'err', v: '#REF!' }); i += 5; continue }
    const num = RE_NUM.exec(rest)
    if (num) { out.push({ t: 'num', v: num[0] }); i += num[0].length; continue }
    const cell = RE_CELL.exec(rest)
    if (cell) { out.push({ t: 'cell', v: cell[2], sheet: cell[1] }); i += cell[0].length; continue }
    const col = RE_COL.exec(rest)
    if (col && (out.at(-1)?.v === ':' || /^\s*:/.test(rest.slice(col[0].length)))) { out.push({ t: 'col', v: col[2], sheet: col[1] }); i += col[0].length; continue }
    const id = RE_ID.exec(rest)
    if (id) { out.push({ t: 'id', v: id[0] }); i += id[0].length; continue }
    const op = RE_OP.exec(rest)
    if (op) { out.push({ t: 'op', v: op[0] }); i += op[0].length; continue }
    if (BRACES[ch]) { out.push({ t: BRACES[ch], v: ch }); i++; continue }
    if (ch === ',') { out.push({ t: 'comma', v: ',' }); i++; continue }
    if (ch === ';') { out.push({ t: 'semi', v: ';' }); i++; continue }
    throw err('#ERROR!')
  }
  return out
}

/* =============================== parser ================================ */
/** A point in a reference. `r === null` means a whole-column reference such as `A:A`. */
interface Pt { c: number; r: number | null }
type Node =
  | { k: 'num'; v: number } | { k: 'str'; v: string } | { k: 'bool'; v: boolean } | { k: 'empty' } | { k: 'err'; v: string }
  | { k: 'ref'; sheet?: string; p: Pt }
  | { k: 'range'; sheet?: string; a: Pt; b: Pt }
  | { k: 'un'; op: string; a: Node } | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'fn'; name: string; args: Node[] }
  | { k: 'array'; rows: Node[][] }

const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 }
const colIdx = (s: string) => { let n = 0; for (const ch of s.replace(/\$/g, '').toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1 }
function ptOf(t: Tok): Pt {
  if (t.t === 'col') return { c: colIdx(t.v), r: null }
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(t.v)
  if (!m) throw err('#REF!')
  return { c: colIdx(m[1]), r: Number(m[2]) - 1 }
}

function parse(src: string): Node {
  const toks = tokenize(src)
  let p = 0
  const peek = () => toks[p]
  const next = () => toks[p++]

  function primary(): Node {
    const t = next()
    if (!t) throw err('#ERROR!')
    if (t.t === 'num') return { k: 'num', v: Number(t.v) }
    if (t.t === 'str') return { k: 'str', v: t.v }
    if (t.t === 'err') return { k: 'err', v: t.v }
    if (t.t === 'cell' || t.t === 'col') {
      // a `:` here (with any surrounding spaces) makes this a range
      if (peek()?.t === 'op' && peek()!.v === ':') {
        next()
        const b = next()
        if (!b || (b.t !== 'cell' && b.t !== 'col')) throw err('#ERROR!')
        return { k: 'range', sheet: t.sheet ?? b.sheet, a: ptOf(t), b: ptOf(b) }
      }
      if (t.t === 'col') throw err('#NAME?')
      return { k: 'ref', sheet: t.sheet, p: ptOf(t) }
    }
    if (t.t === 'lp') { const e = expr(1); if (next()?.t !== 'rp') throw err('#ERROR!'); return e }
    if (t.t === 'lbrace') {
      // array literal: {1,2;3,4} — commas separate columns, semicolons separate rows
      const rows: Node[][] = [[]]
      if (peek()?.t !== 'rbrace') {
        for (;;) {
          rows[rows.length - 1].push(expr(1))
          const sep = peek()
          if (sep?.t === 'comma') { next(); continue }
          if (sep?.t === 'semi') { next(); rows.push([]); continue }
          break
        }
      }
      if (next()?.t !== 'rbrace') throw err('#ERROR!')
      const w = rows[0].length
      if (rows.some(r => r.length !== w)) throw err('#VALUE!')
      return { k: 'array', rows }
    }
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) return { k: 'un', op: t.v, a: postfix() }
    if (t.t === 'id') {
      const up = t.v.toUpperCase()
      if (peek()?.t === 'lp') {
        next()
        const args: Node[] = []
        if (peek()?.t !== 'rp') {
          for (;;) {
            const sep = () => peek()?.t === 'comma' || peek()?.t === 'semi'
            if (sep() || peek()?.t === 'rp' || !peek()) args.push({ k: 'empty' }); else args.push(expr(1))
            if (sep()) { next(); continue }
            break
          }
        }
        if (next()?.t !== 'rp') throw err('#ERROR!')
        return { k: 'fn', name: up, args }
      }
      if (up === 'TRUE') return { k: 'bool', v: true }
      if (up === 'FALSE') return { k: 'bool', v: false }
      throw err('#NAME?')
    }
    throw err('#ERROR!')
  }
  function postfix(): Node {
    let n = primary()
    while (peek()?.t === 'op' && peek()!.v === '%') { next(); n = { k: 'bin', op: '/', a: n, b: { k: 'num', v: 100 } } }
    return n
  }
  function expr(minPrec: number): Node {
    let left = postfix()
    for (;;) {
      const t = peek()
      if (!t || t.t !== 'op' || PREC[t.v] === undefined || PREC[t.v] < minPrec) break
      const prec = PREC[t.v]
      next()
      left = { k: 'bin', op: t.v, a: left, b: expr(t.v === '^' ? prec : prec + 1) }
    }
    return left
  }
  const n = expr(1)
  if (p < toks.length) throw err('#ERROR!')
  return n
}

/* ============================== coercion =============================== */
function toNum(v: Scalar): number {
  if (isErr(v)) throw v
  if (v === null || v === '') return 0
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  const n = Number(v.trim())
  if (Number.isNaN(n)) throw err('#VALUE!')
  return n
}
function truthy(v: Scalar): boolean {
  if (isErr(v)) throw v
  if (typeof v === 'string') {
    const u = v.toUpperCase()
    if (u === 'TRUE') return true
    if (u === 'FALSE') return false
    const n = Number(v)
    if (v !== '' && !Number.isNaN(n)) return n !== 0
    throw err('#VALUE!')
  }
  return !!v
}
const looseEq = (a: Scalar, b: Scalar) => {
  if (typeof a === 'number' && typeof b === 'number') return a === b
  if (a === null || b === null) return (a === null || a === '') && (b === null || b === '')
  return toStr(a).toLowerCase() === toStr(b).toLowerCase()
}

/** One binary operation on two plain values. Errors propagate from either side. */
function binScalar(a: Scalar, b: Scalar, op: string): Scalar {
  if (isErr(a)) return a
  if (isErr(b)) return b
  if (op === '&') return toStr(a) + toStr(b)
  if (op === '=') return looseEq(a, b)
  if (op === '<>') return !looseEq(a, b)
  if (op === '<' || op === '>' || op === '<=' || op === '>=') {
    const text = (typeof a === 'string' || typeof b === 'string') && !(typeof a === 'number' || typeof b === 'number')
    const d = text
      ? (toStr(a).toLowerCase() < toStr(b).toLowerCase() ? -1 : toStr(a).toLowerCase() > toStr(b).toLowerCase() ? 1 : 0)
      : toNum(a) - toNum(b)
    return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0
  }
  const x = toNum(a), y = toNum(b)
  switch (op) {
    case '+': return x + y
    case '-': return x - y
    case '*': return x * y
    case '/': if (y === 0) throw err('#DIV/0!'); return x / y
    case '^': { const r = Math.pow(x, y); if (Number.isNaN(r)) throw err('#NUM!'); return r }
  }
  throw err('#ERROR!')
}

/** Element-wise application of a binary operator, with scalars and single rows/columns stretched. */
function broadcast(a: Val, b: Val, op: string): Scalar[][] {
  const ga = Array.isArray(a) ? a : [[a]]
  const gb = Array.isArray(b) ? b : [[b]]
  const rows = Math.max(ga.length, gb.length)
  const cols = Math.max(ga[0]?.length ?? 0, gb[0]?.length ?? 0)
  const pick = (g: Scalar[][], r: number, c: number): Scalar => {
    const row = g.length === 1 ? g[0] : g[r]
    if (!row) return err('#N/A')
    const v = row.length === 1 ? row[0] : row[c]
    return v === undefined ? err('#N/A') : v
  }
  const out: Scalar[][] = []
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = []
    for (let c = 0; c < cols; c++) {
      try { row.push(binScalar(pick(ga, r, c), pick(gb, r, c), op)) } catch (e) { if (!isErr(e)) throw e; row.push(e) }
    }
    out.push(row)
  }
  return out
}

/* =============================== evaluator ============================= */
interface RangeDep { sheet: number; r1: number; c1: number; r2: number; c2: number; dependent: string }
const key = (sheet: number, addr: string) => `${sheet}!${addr}`
const MAX_SPILL = 20000

/**
 * Lazy, cached evaluator. Every read performed while evaluating a formula is recorded, so
 * `invalidate()` after an edit clears only that cell and whatever (transitively) depends on it.
 */
export class Evaluator {
  private wb: Workbook
  private cache = new Map<string, Scalar>()
  private stack: string[] = []
  private visiting = new Set<string>()
  private astCache = new Map<string, Node | FormulaError>()
  private dependents = new Map<string, Set<string>>()
  private rangeDeps: RangeDep[] = []
  /** Values produced by a dynamic array spilling out of its origin cell. */
  private spillCells = new Map<string, Scalar>()
  private spillOwner = new Map<string, string>()
  private spilled = new Map<string, string[]>()
  private volatileCells = new Set<string>()
  private scanned = new Set<number>()

  constructor(wb: Workbook) { this.wb = wb }

  setWorkbook(wb: Workbook) { this.wb = wb }

  /** Cells currently cached (used by tests to prove recalculation is incremental). */
  get cachedCount() { return this.cache.size }

  invalidate(sheet: number, addr: string) {
    const todo = [key(sheet, addr)]
    const seen = new Set<string>()
    while (todo.length) {
      const k = todo.pop()!
      if (seen.has(k)) continue
      seen.add(k)
      this.cache.delete(k)
      this.dependents.get(k)?.forEach(d => todo.push(d))
      const owner = this.spillOwner.get(k)
      if (owner && owner !== k) todo.push(owner)
      for (const sk of this.spilled.get(k) ?? []) { this.spillCells.delete(sk); this.spillOwner.delete(sk); todo.push(sk) }
      this.spilled.delete(k)
      const bang = k.indexOf('!')
      const s = Number(k.slice(0, bang)), p = parseAddr(k.slice(bang + 1))
      if (p) for (const rd of this.rangeDeps) if (rd.sheet === s && p.r >= rd.r1 && p.r <= rd.r2 && p.c >= rd.c1 && p.c <= rd.c2) todo.push(rd.dependent)
    }
    if (this.rangeDeps.length) this.rangeDeps = this.rangeDeps.filter(d => !seen.has(d.dependent))
    this.scanned.clear()
  }

  /** Drop cached results of volatile formulas (RAND, TODAY, NOW) so they recompute. */
  invalidateVolatile() { for (const k of this.volatileCells) { const i = k.indexOf('!'); this.invalidate(Number(k.slice(0, i)), k.slice(i + 1)) } }

  private link(sheet: number, addr: string) {
    const cur = this.stack[this.stack.length - 1]
    if (!cur) return
    const k = key(sheet, addr)
    let set = this.dependents.get(k)
    if (!set) this.dependents.set(k, (set = new Set()))
    set.add(cur)
  }

  rawCell(sheet: number, addr: string): Cell | undefined { return this.wb.sheets[sheet]?.cells[addr] }

  /** Evaluate every formula on a sheet once, so dynamic arrays register their spill ranges. */
  ensureFormulas(sheet: number) {
    if (this.scanned.has(sheet)) return
    this.scanned.add(sheet)
    const cells = this.wb.sheets[sheet]?.cells
    if (!cells) return
    for (const addr of Object.keys(cells)) if (cells[addr].v.startsWith('=')) this.get(sheet, addr, false)
  }

  /** Value shown in a cell: its own result, or a value spilled into it by a neighbouring array. */
  displayValue(sheet: number, addr: string): Scalar {
    const cell = this.rawCell(sheet, addr)
    if (cell && cell.v !== '') return this.get(sheet, addr, false)
    const k = key(sheet, addr)
    if (this.spillCells.has(k)) return this.spillCells.get(k)!
    return null
  }
  /** True when the value at this address came from another cell's dynamic array. */
  isSpilled(sheet: number, addr: string): boolean {
    const k = key(sheet, addr)
    return this.spillCells.has(k) && this.spillOwner.get(k) !== k
  }

  get(sheet: number, addr: string, link = true): Scalar {
    const k = key(sheet, addr)
    if (link) this.link(sheet, addr)
    const hit = this.cache.get(k)
    if (hit !== undefined || this.cache.has(k)) return hit as Scalar
    const cell = this.rawCell(sheet, addr)
    if (!cell || cell.v === '') {
      if (this.spillCells.has(k)) return this.spillCells.get(k)!
      return null
    }
    if (cell.v.startsWith('=')) {
      if (this.stack.includes(k)) return err('#CIRC!')
      this.precompute(sheet, addr)
      if (this.cache.has(k)) return this.cache.get(k) as Scalar
    }
    return this.compute(sheet, k, cell)
  }

  private ast(formula: string): Node | FormulaError {
    let a = this.astCache.get(formula)
    if (!a) {
      try { a = parse(formula.slice(1)) } catch (e) { a = e instanceof FormulaError ? e : err('#ERROR!') }
      this.astCache.set(formula, a)
    }
    return a
  }

  private compute(sheet: number, k: string, cell: Cell): Scalar {
    let val: Scalar
    if (cell.v.startsWith('=')) {
      this.stack.push(k)
      try {
        const a = this.ast(cell.v)
        if (isErr(a)) val = a
        else {
          const r = this.ev(a, sheet)
          val = Array.isArray(r) ? this.spill(sheet, k, r) : r
        }
      } catch (e) { val = e instanceof FormulaError ? e : err('#ERROR!') }
      this.stack.pop()
    } else {
      const t = cell.v.trim()
      if (t !== '' && !Number.isNaN(Number(t)) && /^[-+]?(\d|\.\d)/.test(t)) val = Number(t)
      else if (/^(true|false)$/i.test(t)) val = t.toUpperCase() === 'TRUE'
      else val = cell.v
    }
    if (!(isErr(val) && val.code === '#CIRC!' && this.stack.length > 0)) this.cache.set(k, val)
    return val
  }

  /** Write a dynamic array into the cells below/right of its origin, Excel-style. */
  private spill(sheet: number, originKey: string, rows: Scalar[][]): Scalar {
    const addr = originKey.slice(originKey.indexOf('!') + 1)
    const p = parseAddr(addr)
    const h = rows.length, w = rows[0]?.length ?? 0
    if (!p || h === 0 || w === 0) return rows[0]?.[0] ?? null
    if (h === 1 && w === 1) return rows[0][0]
    if (h * w > MAX_SPILL) return err('#SPILL!')
    const sh = this.wb.sheets[sheet]
    if (p.r + h > sh.rows || p.c + w > sh.cols) return err('#SPILL!')
    const keys: string[] = []
    for (let r = 0; r < h; r++) {
      for (let c = 0; c < w; c++) {
        if (r === 0 && c === 0) continue
        const a = `${colName(p.c + c)}${p.r + r + 1}`
        const occupant = sh.cells[a]
        if (occupant && occupant.v !== '') return err('#SPILL!')
        const kk = key(sheet, a)
        const other = this.spillOwner.get(kk)
        if (other && other !== originKey) return err('#SPILL!')
        keys.push(kk)
      }
    }
    for (const kk of keys) { this.spillCells.delete(kk); this.spillOwner.delete(kk) }
    let i = 0
    for (let r = 0; r < h; r++) {
      for (let c = 0; c < w; c++) {
        if (r === 0 && c === 0) continue
        const kk = keys[i++]
        this.spillCells.set(kk, rows[r][c] ?? null)
        this.spillOwner.set(kk, originKey)
      }
    }
    this.spilled.set(originKey, keys)
    this.spillOwner.set(originKey, originKey)
    return rows[0][0]
  }

  private sheetIndex(prefix: string | undefined, cur: number): number {
    if (!prefix) return cur
    const name = prefix.replace(/^'|'$/g, '').toLowerCase()
    const i = this.wb.sheets.findIndex(s => s.name.toLowerCase() === name)
    if (i < 0) throw err('#REF!')
    return i
  }

  private readRange(n: Extract<Node, { k: 'range' }>, sheet: number): Scalar[][] {
    const sh = this.sheetIndex(n.sheet, sheet)
    const sheetObj = this.wb.sheets[sh]
    const whole = n.a.r === null || n.b.r === null
    const r1 = whole ? 0 : Math.min(n.a.r!, n.b.r!)
    const r2 = whole ? Math.max(0, sheetObj.rows - 1) : Math.max(n.a.r!, n.b.r!)
    const c1 = Math.min(n.a.c, n.b.c), c2 = Math.max(n.a.c, n.b.c)
    if (r1 < 0 || c1 < 0) throw err('#REF!')
    const cur = this.stack[this.stack.length - 1]
    if (cur) this.rangeDeps.push({ sheet: sh, r1, c1, r2, c2, dependent: cur })
    const rows: Scalar[][] = []
    for (let r = r1; r <= r2; r++) {
      const row: Scalar[] = []
      // reads inside a range are covered by the single range dependency registered above
      for (let c = c1; c <= c2; c++) row.push(this.get(sh, `${colName(c)}${r + 1}`, false))
      rows.push(row)
    }
    return rows
  }

  /** Formula cells this formula reads directly, used to resolve precedents without recursion. */
  private precedents(sheet: number, formula: string): [number, string][] {
    const a = this.ast(formula)
    if (isErr(a)) return []
    const out: [number, string][] = []
    const walk = (n: Node) => {
      switch (n.k) {
        case 'ref':
          try { if (n.p.r !== null) out.push([this.sheetIndex(n.sheet, sheet), `${colName(n.p.c)}${n.p.r + 1}`]) } catch { /* bad sheet: reported when evaluated */ }
          break
        case 'range':
          try {
            const sh = this.sheetIndex(n.sheet, sheet)
            const cells = this.wb.sheets[sh].cells
            const whole = n.a.r === null || n.b.r === null
            const r1 = whole ? 0 : Math.min(n.a.r!, n.b.r!), r2 = whole ? this.wb.sheets[sh].rows - 1 : Math.max(n.a.r!, n.b.r!)
            const c1 = Math.min(n.a.c, n.b.c), c2 = Math.max(n.a.c, n.b.c)
            const area = (r2 - r1 + 1) * (c2 - c1 + 1)
            if (area > Object.keys(cells).length) {
              for (const ad of Object.keys(cells)) { const p = parseAddr(ad); if (p && p.r >= r1 && p.r <= r2 && p.c >= c1 && p.c <= c2) out.push([sh, ad]) }
            } else {
              for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push([sh, `${colName(c)}${r + 1}`])
            }
          } catch { /* bad sheet: reported when evaluated */ }
          break
        case 'un': walk(n.a); break
        case 'bin': walk(n.a); walk(n.b); break
        case 'fn': n.args.forEach(walk); break
        case 'array': n.rows.forEach(r => r.forEach(walk)); break
        default: break
      }
    }
    walk(a)
    return out
  }

  /** Depth-first evaluation driven by an explicit work list, so deep chains cannot overflow. */
  private precompute(sheet: number, addr: string) {
    const work: { sheet: number; addr: string; expanded: boolean }[] = [{ sheet, addr, expanded: false }]
    while (work.length) {
      const top = work[work.length - 1]
      const k = key(top.sheet, top.addr)
      const cell = this.rawCell(top.sheet, top.addr)
      if (this.cache.has(k) || !cell?.v.startsWith('=')) { work.pop(); this.visiting.delete(k); continue }
      if (!top.expanded) {
        if (this.visiting.has(k)) { work.pop(); continue } // cycle: reported when the cell is evaluated
        top.expanded = true
        this.visiting.add(k)
        for (const [ps, pa] of this.precedents(top.sheet, cell.v)) {
          const pk = key(ps, pa)
          if (this.cache.has(pk) || this.visiting.has(pk)) continue
          if (this.rawCell(ps, pa)?.v.startsWith('=')) work.push({ sheet: ps, addr: pa, expanded: false })
        }
        continue
      }
      work.pop()
      this.compute(top.sheet, k, cell)
      this.visiting.delete(k)
    }
  }

  private ev(n: Node, sheet: number): Val {
    switch (n.k) {
      case 'num': case 'str': case 'bool': return n.v
      case 'empty': return null
      case 'err': return err(n.v)
      case 'ref': {
        if (n.p.r === null) throw err('#REF!')
        return this.get(this.sheetIndex(n.sheet, sheet), `${colName(n.p.c)}${n.p.r + 1}`)
      }
      case 'range': return this.readRange(n, sheet)
      case 'array': return n.rows.map(r => r.map(x => this.scalar(x, sheet)))
      case 'un': { const x = toNum(this.scalar(n.a, sheet)); return n.op === '-' ? -x : x }
      case 'bin': {
        const a = this.ev(n.a, sheet), b = this.ev(n.b, sheet)
        // Operating on a range broadcasts element-wise, so `A1:A5>3` yields an array of booleans
        // the way it does in Excel (this is what makes FILTER and SUMPRODUCT conditions work).
        if (Array.isArray(a) || Array.isArray(b)) return broadcast(a, b, n.op)
        return binScalar(a, b, n.op)
      }
      case 'fn': return this.call(n, sheet)
    }
  }

  private scalar(n: Node, sheet: number): Scalar {
    const v = this.ev(n, sheet)
    return Array.isArray(v) ? (v[0]?.[0] ?? null) : v
  }

  private call(n: Extract<Node, { k: 'fn' }>, sheet: number): Val {
    const def = FUNCTIONS[n.name]
    if (!def) throw err('#NAME?')
    const args = n.args
    const supplied = args.length === 1 && args[0].k === 'empty' ? 0 : args.length
    if (supplied < def.min || supplied > def.max) throw err('#ERROR!')
    if (def.volatile) { const cur = this.stack[this.stack.length - 1]; if (cur) this.volatileCells.add(cur) }

    const cache = new Map<number, Val>()
    const val = (i: number): Val => {
      if (cache.has(i)) return cache.get(i)!
      // An error raised while evaluating an argument becomes a value here, so IFERROR can see it.
      // The typed accessors below re-throw it, which keeps normal propagation intact.
      let v: Val
      try { v = args[i] === undefined ? null : this.ev(args[i], sheet) } catch (e) { if (!isErr(e)) throw e; v = e }
      cache.set(i, v)
      return v
    }
    const one = (i: number): Scalar => { const v = val(i); return Array.isArray(v) ? (v[0]?.[0] ?? null) : v }
    const grid = (i: number): Scalar[][] => { const v = val(i); return Array.isArray(v) ? v : [[v]] }
    const flat = (i: number): Scalar[] => grid(i).flat()
    const isRef = (i: number) => args[i]?.k === 'range' || args[i]?.k === 'ref'
    const guard = (v: Scalar): Scalar => { if (isErr(v)) throw v; return v }

    const ctx: Ctx = {
      n: supplied,
      has: i => i < supplied && args[i] !== undefined && args[i].k !== 'empty',
      scalar: one,
      flat, grid, isRef,
      num: (i, d) => (ctx.has(i) ? toNum(guard(one(i))) : d !== undefined ? d : toNum(guard(one(i)))),
      int: (i, d) => Math.trunc(ctx.num(i, d)),
      str: (i, d) => (ctx.has(i) ? toStr(guard(one(i))) : d !== undefined ? d : toStr(guard(one(i)))),
      bool: (i, d) => (ctx.has(i) ? truthy(guard(one(i))) : d !== undefined ? d : truthy(guard(one(i)))),
      all: (from = 0) => {
        const out: Scalar[] = []
        for (let i = from; i < supplied; i++) for (const v of flat(i)) { if (isErr(v)) throw v; out.push(v) }
        return out
      },
      nums: (from = 0, to = supplied - 1) => {
        const out: number[] = []
        for (let i = from; i <= to && i < supplied; i++) {
          const direct = !isRef(i)
          for (const v of flat(i)) {
            if (isErr(v)) throw v
            if (typeof v === 'number') out.push(v)
            else if (typeof v === 'boolean') { if (direct) out.push(v ? 1 : 0) }
            else if (direct && v !== null && v !== '') out.push(toNum(v))
          }
        }
        return out
      },
    }
    return def.fn(ctx)
  }
}

/* ================================ display ============================== */
/**
 * Excel closes unbalanced parentheses when you commit a formula. Quoted text is ignored so
 * `="a("` is left alone.
 */
/**
 * Entry-time check, like Excel's "There's a problem with this formula": returns a message for a formula
 * that cannot be parsed or calls a known function with the wrong number of arguments, else null.
 * Unknown function names are allowed (they evaluate to #NAME?).
 */
export function validateFormula(text: string): string | null {
  if (!text.startsWith('=') || text.length < 2) return null
  let tree: Node
  try { tree = parse(text.slice(1)) } catch { return "There's a problem with this formula. Check the brackets, quotes and commas." }
  const walk = (n: Node): string | null => {
    if (n.k === 'fn') {
      const def = FUNCTIONS[n.name]
      const supplied = n.args.length === 1 && n.args[0].k === 'empty' ? 0 : n.args.length
      if (def && (supplied < def.min || supplied > def.max)) {
        const need = def.max >= 255 ? `at least ${def.min}` : def.min === def.max ? `${def.min}` : `${def.min} to ${def.max}`
        return `${n.name} needs ${need} argument${def.max === 1 ? '' : 's'}, but ${supplied} ${supplied === 1 ? 'was' : 'were'} given.`
      }
      for (const a of n.args) { const m = walk(a); if (m) return m }
    } else if (n.k === 'un') return walk(n.a)
    else if (n.k === 'bin') return walk(n.a) ?? walk(n.b)
    else if (n.k === 'array') { for (const r of n.rows) for (const a of r) { const m = walk(a); if (m) return m } }
    return null
  }
  return walk(tree)
}

export function balanceFormula(input: string): string {
  if (!input.startsWith('=')) return input
  let depth = 0, inStr = false
  for (let i = 1; i < input.length; i++) {
    const ch = input[i]
    if (ch === '"') { if (inStr && input[i + 1] === '"') { i++; continue } inStr = !inStr; continue }
    if (inStr) continue
    if (ch === '(') depth++
    else if (ch === ')') depth = Math.max(0, depth - 1)
  }
  return depth > 0 ? input + ')'.repeat(depth) : input
}

export function formatValue(v: Scalar, fmt: CellStyle['fmt'] = 'general', dec?: number): string {
  if (v === null) return ''
  if (isErr(v)) return v.code
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'string') return v
  const d = dec !== undefined ? Math.max(0, Math.min(15, dec)) : undefined
  switch (fmt) {
    case 'number': return v.toLocaleString(undefined, { minimumFractionDigits: d ?? 2, maximumFractionDigits: d ?? 2 })
    case 'currency': return v.toLocaleString(undefined, { style: 'currency', currency: 'USD', minimumFractionDigits: d ?? 2, maximumFractionDigits: d ?? 2 })
    case 'percent': return (v * 100).toLocaleString(undefined, { minimumFractionDigits: d ?? 0, maximumFractionDigits: d ?? 2 }) + '%'
    case 'date': { const dt = new Date(Math.round((v - 25569) * 86400000)); return Number.isNaN(dt.getTime()) ? String(v) : dt.toLocaleDateString() }
    case 'time': { const s = Math.round((v - Math.floor(v)) * 86400); return `${String(Math.floor(s / 3600)).padStart(2, '0')}:${String(Math.floor((s % 3600) / 60)).padStart(2, '0')}` }
    case 'scientific': return v.toExponential(d ?? 2).toUpperCase()
    default: return d !== undefined ? v.toFixed(d) : Number.isInteger(v) ? String(v) : String(Number(v.toPrecision(12)))
  }
}
