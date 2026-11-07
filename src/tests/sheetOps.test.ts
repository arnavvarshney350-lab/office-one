import { describe, expect, it } from 'vitest'
import { Evaluator, validateFormula } from '../editors/excel/formula'
import { emptyWorkbook, type CondRule, type Workbook } from '../editors/excel/model'
import { cfContext, cfFor, chartData, fillRange, findCells, parseRange, replaceInCell, sortRange } from '../editors/excel/sheetOps'

const wb = (cells: Record<string, string>): Workbook => { const w = emptyWorkbook(); for (const [k, v] of Object.entries(cells)) w.sheets[0].cells[k] = { v }; return w }
const col = (w: Workbook, c: string, n: number) => Array.from({ length: n }, (_, i) => w.sheets[0].cells[`${c}${i + 1}`]?.v || null) // blank and missing cells both read as null

describe('sortRange', () => {
  it('sorts numbers ascending and descending, keeping rows together', () => {
    const w = wb({ A1: '3', B1: 'c', A2: '1', B2: 'a', A3: '2', B3: 'b' })
    const asc = sortRange(w, 0, { r1: 0, c1: 0, r2: 2, c2: 1 }, 0, true, false)
    expect([col(asc, 'A', 3), col(asc, 'B', 3)]).toEqual([['1', '2', '3'], ['a', 'b', 'c']])
    const desc = sortRange(w, 0, { r1: 0, c1: 0, r2: 2, c2: 1 }, 0, false, false)
    expect(col(desc, 'A', 3)).toEqual(['3', '2', '1'])
  })
  it('leaves a header row in place and puts blanks last in either direction', () => {
    const w = wb({ A1: 'Name', A2: 'pear', A3: '', A4: 'apple' })
    expect(col(sortRange(w, 0, { r1: 0, c1: 0, r2: 3, c2: 0 }, 0, true, true), 'A', 4)).toEqual(['Name', 'apple', 'pear', null])
    expect(col(sortRange(w, 0, { r1: 0, c1: 0, r2: 3, c2: 0 }, 0, false, true), 'A', 4)).toEqual(['Name', 'pear', 'apple', null])
  })
  it('orders numbers before text', () => { expect(col(sortRange(wb({ A1: 'b', A2: '10', A3: '2' }), 0, { r1: 0, c1: 0, r2: 2, c2: 0 }, 0, true, false), 'A', 3)).toEqual(['2', '10', 'b']) })
  it('formulas move with their row and keep pointing at their row-mates', () => {
    const w = wb({ A1: '5', B1: '=A1*2', A2: '1', B2: '=A2*2' })
    const s = sortRange(w, 0, { r1: 0, c1: 0, r2: 1, c2: 1 }, 0, true, false)
    expect(col(s, 'B', 2)).toEqual(['=A1*2', '=A2*2'])
    const ev = new Evaluator(s)
    expect([ev.get(0, 'B1'), ev.get(0, 'B2')]).toEqual([2, 10])
  })
})

describe('conditional formatting', () => {
  const run = (rule: Partial<CondRule>, cells: Record<string, string>, addr: string) => {
    const w = wb(cells); const ev = new Evaluator(w)
    const r = { id: 'x', range: 'A1:A5', kind: 'gt', bg: '#ff0000', ...rule } as CondRule
    return cfFor([r], addr, ev.displayValue(0, addr), cfContext(ev, 0))
  }
  const D = { A1: '1', A2: '5', A3: '9', A4: 'apple', A5: '5' }
  it('greater than / less than / equal / between', () => {
    expect(run({ kind: 'gt', a: '4' }, D, 'A2')?.bg).toBe('#ff0000'); expect(run({ kind: 'gt', a: '4' }, D, 'A1')).toBeNull()
    expect(run({ kind: 'lt', a: '4' }, D, 'A1')?.bg).toBe('#ff0000')
    expect(run({ kind: 'eq', a: '9' }, D, 'A3')).not.toBeNull(); expect(run({ kind: 'eq', a: 'APPLE' }, D, 'A4')).not.toBeNull()
    expect(run({ kind: 'between', a: '2', b: '6' }, D, 'A2')).not.toBeNull(); expect(run({ kind: 'between', a: '2', b: '6' }, D, 'A3')).toBeNull()
  })
  it('text contains, duplicates, top and bottom', () => {
    expect(run({ kind: 'text', a: 'ppl' }, D, 'A4')).not.toBeNull()
    expect(run({ kind: 'dup' }, D, 'A2')).not.toBeNull(); expect(run({ kind: 'dup' }, D, 'A3')).toBeNull()
    expect(run({ kind: 'top', a: '1' }, D, 'A3')).not.toBeNull(); expect(run({ kind: 'top', a: '1' }, D, 'A2')).toBeNull()
    expect(run({ kind: 'bottom', a: '1' }, D, 'A1')).not.toBeNull()
  })
  it('colour scale shades by position in the range and ignores cells outside the range', () => {
    const lo = run({ kind: 'scale', bg: '#00ff00' }, D, 'A1')!.bg, hi = run({ kind: 'scale', bg: '#00ff00' }, D, 'A3')!.bg
    expect(lo).toBe('#ffffff'); expect(hi).toBe('#00ff00')
    expect(run({ kind: 'gt', a: '0', range: 'B1:B2' }, D, 'A2')).toBeNull()
  })
  it('conditions follow recalculated values', () => { expect(run({ kind: 'gt', a: '10' }, { A1: '=5*3' }, 'A1')).not.toBeNull() })
})

describe('find and replace', () => {
  const w = wb({ A1: 'Hello', A2: 'hello world', A3: '=1+1', B1: 'HELLO' })
  it('finds with case and whole-cell options', () => {
    expect(findCells(w, 0, 'hello', { matchCase: false, whole: false, inFormulas: false }).map(h => h.addr)).toEqual(['A1', 'B1', 'A2'])
    expect(findCells(w, 0, 'hello', { matchCase: true, whole: false, inFormulas: false }).map(h => h.addr)).toEqual(['A2'])
    expect(findCells(w, 0, 'hello', { matchCase: false, whole: true, inFormulas: false }).map(h => h.addr)).toEqual(['A1', 'B1'])
  })
  it('searches formula text or the computed value', () => {
    expect(findCells(w, 0, '1+1', { matchCase: false, whole: false, inFormulas: true }).length).toBe(1)
    expect(findCells(w, 0, '2', { matchCase: false, whole: true, inFormulas: false }).map(h => h.addr)).toEqual(['A3'])
  })
  it('replaces text, treating the search as literal (no regex surprises)', () => {
    expect(replaceInCell('a.b.c', '.', '-', { matchCase: false, whole: false })).toBe('a-b-c')
    expect(replaceInCell('Foo foo', 'foo', 'X', { matchCase: false, whole: false })).toBe('X X')
    expect(replaceInCell('Foo foo', 'foo', 'X', { matchCase: true, whole: false })).toBe('Foo X')
    expect(replaceInCell('abc', 'abc', '$&', { matchCase: false, whole: true })).toBe('$&')
  })
})

describe('chart data and ranges', () => {
  it('reads headers, categories and series', () => {
    const w = wb({ A1: '', B1: 'Sales', C1: 'Costs', A2: 'Q1', B2: '10', C2: '4', A3: 'Q2', B3: '20', C3: '6' })
    const d = chartData(new Evaluator(w), 0, { id: 'c', type: 'column', range: 'A1:C3', title: '', x: 0, y: 0, w: 1, h: 1 })
    expect(d.categories).toEqual(['Q1', 'Q2']); expect(d.series).toEqual([{ name: 'Sales', values: [10, 20] }, { name: 'Costs', values: [4, 6] }])
  })
  it('works without a header row or category column', () => {
    const d = chartData(new Evaluator(wb({ A1: '3', A2: '4', A3: '5' })), 0, { id: 'c', type: 'line', range: 'A1:A3', title: '', x: 0, y: 0, w: 1, h: 1 })
    expect(d.series[0].values).toEqual([3, 4, 5])
  })
  it('parses A1 ranges in any order', () => { expect(parseRange('C3:A1')).toEqual({ r1: 0, c1: 0, r2: 2, c2: 2 }); expect(parseRange('nope')).toBeNull() })
})

describe('fill handle (fillRange)', () => {
  const grid = (o: Record<string, string>) => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, { v }]))
  const vals = (c: Record<string, { v: string }>, ...a: string[]) => a.map(x => c[x]?.v)
  it('continues a numeric series from two or more cells', () => {
    const r = fillRange(grid({ D1: '1', D2: '2' }), { r1: 0, r2: 1, c1: 3, c2: 3 }, { r: 5, c: 3 })!
    expect(vals(r.cells, 'D3', 'D4', 'D5', 'D6')).toEqual(['3', '4', '5', '6']); expect(r.rect).toEqual({ r1: 0, r2: 5, c1: 3, c2: 3 })
  })
  it('copies a lone number and counts up a lone "Item 1"', () => {
    expect(vals(fillRange(grid({ A1: '7' }), { r1: 0, r2: 0, c1: 0, c2: 0 }, { r: 2, c: 0 })!.cells, 'A2', 'A3')).toEqual(['7', '7'])
    expect(vals(fillRange(grid({ A1: 'Item 1' }), { r1: 0, r2: 0, c1: 0, c2: 0 }, { r: 2, c: 0 })!.cells, 'A2', 'A3')).toEqual(['Item 2', 'Item 3'])
  })
  it('shifts formula references, right and up too', () => {
    expect(vals(fillRange(grid({ B1: '=A1*2' }), { r1: 0, r2: 0, c1: 1, c2: 1 }, { r: 2, c: 1 })!.cells, 'B2', 'B3')).toEqual(['=A2*2', '=A3*2'])
    expect(vals(fillRange(grid({ A1: '1', B1: '3' }), { r1: 0, r2: 0, c1: 0, c2: 1 }, { r: 0, c: 3 })!.cells, 'C1', 'D1')).toEqual(['5', '7'])
    const up = fillRange(grid({ A5: '10', A6: '20' }), { r1: 4, r2: 5, c1: 0, c2: 0 }, { r: 2, c: 0 })!
    expect(vals(up.cells, 'A4', 'A3')).toEqual(['0', '-10']); expect(up.rect.r1).toBe(2)
  })
  it('does nothing when released inside the block', () => expect(fillRange(grid({ A1: '1' }), { r1: 0, r2: 0, c1: 0, c2: 0 }, { r: 0, c: 0 })).toBeNull())
})

describe('validateFormula (entry-time check)', () => {
  it('rejects wrong argument counts and syntax errors with a message', () => {
    expect(validateFormula('=SUM()')).toMatch(/SUM needs at least 1 argument/)
    expect(validateFormula('=ABS(1,2)')).toMatch(/ABS needs 1 argument/)
    expect(validateFormula('=1+')).toMatch(/problem with this formula/)
    expect(validateFormula('=IF(ABS(),1,2)')).toMatch(/ABS/)
  })
  it('accepts valid formulas, unknown names (#NAME?) and plain text', () => {
    for (const f of ['=SUM(A1:A4)', '=1+2', '=FOO(1)', '=IF(A1>1,"a","b")', 'hello', '=']) expect(validateFormula(f)).toBeNull()
  })
})
