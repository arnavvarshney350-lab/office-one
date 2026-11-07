import { describe, expect, it } from 'vitest'
import { Evaluator, FormulaError, formatValue } from '../editors/excel/formula'
import { FUNCTION_GROUPS, FUNCTIONS } from '../editors/excel/functions'
import { deleteCols, deleteRows, emptyWorkbook, insertCols, insertRows, newSheet, shiftFormula, type Workbook } from '../editors/excel/model'

function wbOf(cells: Record<string, string>, extra?: Record<string, Record<string, string>>): Workbook {
  const wb = emptyWorkbook()
  for (const [k, v] of Object.entries(cells)) wb.sheets[0].cells[k] = { v }
  for (const [name, c] of Object.entries(extra ?? {})) { const s = newSheet(name); for (const [k, v] of Object.entries(c)) s.cells[k] = { v }; wb.sheets.push(s) }
  return wb
}
const val = (cells: Record<string, string>, addr: string, extra?: Record<string, Record<string, string>>) => new Evaluator(wbOf(cells, extra)).get(0, addr)
const code = (v: unknown) => (v instanceof FormulaError ? v.code : v)
/** evaluate one formula in Z1 against optional data */
const f = (formula: string, data: Record<string, string> = {}) => code(val({ ...data, Z1: formula }, 'Z1'))

describe('required examples', () => {
  it('Example 1: A1=10, A2=20 -> SUM over the range is 30; single-cell SUM(A1) is 10', () => {
    expect(val({ A1: '10', A2: '20', A3: '=SUM(A1:A2)' }, 'A3')).toBe(30)
    expect(val({ A1: '10', A2: '20', A3: '=SUM(A1)' }, 'A3')).toBe(10)
  })
  it('Example 2: AVERAGE(B1:B2) = 37.5; AVERAGE(B1) = 50', () => {
    expect(val({ B1: '50', B2: '25', B3: '=AVERAGE(B1:B2)' }, 'B3')).toBe(37.5)
    expect(val({ B1: '50', B2: '25', B3: '=AVERAGE(B1)' }, 'B3')).toBe(50)
  })
  it('Example 3: IF(C1>50,"PASS","FAIL") = PASS', () => { expect(val({ C1: '100', C2: '=IF(C1>50,"PASS","FAIL")' }, 'C2')).toBe('PASS') })
  it('Example 4: D1+D2 = 30', () => { expect(val({ D1: '10', D2: '20', D3: '=D1+D2' }, 'D3')).toBe(30) })
  it('Example 5: PRODUCT(E1:E2) = 25; PRODUCT(E1) = 5', () => {
    expect(val({ E1: '5', E2: '5', E3: '=PRODUCT(E1:E2)' }, 'E3')).toBe(25)
    expect(val({ E1: '5', E2: '5', E3: '=PRODUCT(E1)' }, 'E3')).toBe(5)
  })
  it('Example 6: MAX(F1:F2) = 20; MAX(F1) = 15', () => {
    expect(val({ F1: '15', F2: '20', F3: '=MAX(F1:F2)' }, 'F3')).toBe(20)
    expect(val({ F1: '15', F2: '20', F3: '=MAX(F1)' }, 'F3')).toBe(15)
  })
})

describe('basic functions', () => {
  const d = { A1: '1', A2: '2', A3: 'x', A4: '', A5: '9' }
  it('SUM', () => { expect(f('=SUM(A1:A5)', d)).toBe(12); expect(f('=SUM(A1,A2,10)', d)).toBe(13); expect(f('=SUM(1,"2")')).toBe(3) })
  it('AVERAGE', () => { expect(f('=AVERAGE(A1:A5)', d)).toBe(4); expect(f('=AVERAGE(A3:A4)', d)).toBe('#DIV/0!') })
  it('COUNT / COUNTA', () => { expect(f('=COUNT(A1:A5)', d)).toBe(3); expect(f('=COUNTA(A1:A5)', d)).toBe(4) })
  it('MIN / MAX', () => { expect(f('=MIN(A1:A5)', d)).toBe(1); expect(f('=MAX(A1:A5)', d)).toBe(9); expect(f('=MAX(A3:A4)', d)).toBe(0) })
})

describe('logical functions', () => {
  it('IF incl. nested and missing else', () => {
    expect(f('=IF(1>2,"a","b")')).toBe('b'); expect(f('=IF(1>2,"a")')).toBe(false)
    expect(f('=IF(A1>90,"A",IF(A1>50,"B","C"))', { A1: '75' })).toBe('B')
  })
  it('IF only evaluates the chosen branch', () => { expect(f('=IF(TRUE,1,1/0)')).toBe(1) })
  it('AND / OR / NOT', () => {
    expect(f('=AND(1<2,2<3)')).toBe(true); expect(f('=AND(1<2,2>3)')).toBe(false)
    expect(f('=OR(1>2,2<3)')).toBe(true); expect(f('=OR(1>2,2>3)')).toBe(false)
    expect(f('=NOT(1>2)')).toBe(true); expect(f('=NOT(TRUE)')).toBe(false)
    expect(f('=AND(A1:A3)', { A1: '1', A2: '1', A3: '0' })).toBe(false)
    expect(f('=IF(AND(A1>1,OR(A2="x",NOT(A3))),"yes","no")', { A1: '5', A2: 'y', A3: 'FALSE' })).toBe('yes')
  })
  it('IFERROR', () => { expect(f('=IFERROR(1/0,"oops")')).toBe('oops'); expect(f('=IFERROR(4,"oops")')).toBe(4) })
})

describe('mathematical functions', () => {
  it('ROUND', () => { expect(f('=ROUND(2.345,2)')).toBe(2.35); expect(f('=ROUND(-2.5)')).toBe(-3); expect(f('=ROUND(1234,-2)')).toBe(1200) })
  it('ABS', () => { expect(f('=ABS(-7.5)')).toBe(7.5) })
  it('PRODUCT', () => { expect(f('=PRODUCT(2,3,4)')).toBe(24); expect(f('=PRODUCT(A1:A3)', { A1: '2', A2: '3', A3: '5' })).toBe(30) })
  it('MOD', () => { expect(f('=MOD(10,3)')).toBe(1); expect(f('=MOD(-3,2)')).toBe(1); expect(f('=MOD(5,0)')).toBe('#DIV/0!') })
  it('POWER', () => { expect(f('=POWER(2,10)')).toBe(1024); expect(f('=POWER(-1,0.5)')).toBe('#NUM!') })
  it('operators and precedence', () => { expect(f('=2+3*4^2')).toBe(50); expect(f('=(2+3)*4')).toBe(20); expect(f('=-A1+1', { A1: '5' })).toBe(-4); expect(f('=10%')).toBe(0.1) })
})

describe('conditional functions', () => {
  const d = { A1: '1', A2: '5', A3: '10', A4: 'apple', B1: '2', B2: '3', B3: '4', B4: '8' }
  it('COUNTIF', () => { expect(f('=COUNTIF(A1:A4,">3")', d)).toBe(2); expect(f('=COUNTIF(A1:A4,"app*")', d)).toBe(1); expect(f('=COUNTIF(A1:A4,5)', d)).toBe(1) })
  it('SUMIF', () => { expect(f('=SUMIF(A1:A3,">=5")', d)).toBe(15); expect(f('=SUMIF(A1:A3,">1",B1:B3)', d)).toBe(7) })
  it('AVERAGEIF', () => { expect(f('=AVERAGEIF(A1:A3,">1")', d)).toBe(7.5); expect(f('=AVERAGEIF(A1:A3,">1",B1:B3)', d)).toBe(3.5); expect(f('=AVERAGEIF(A1:A3,">100")', d)).toBe('#DIV/0!') })
})

describe('lookup functions', () => {
  const t = { A1: 'apple', B1: '1.5', C1: 'red', A2: 'banana', B2: '0.5', C2: 'yellow', A3: 'cherry', B3: '4', C3: 'dark red', D1: '10', D2: '20', D3: '30', E1: 'ten', E2: 'twenty', E3: 'thirty' }
  it('VLOOKUP exact', () => { expect(f('=VLOOKUP("banana",A1:C3,2,FALSE)', t)).toBe(0.5); expect(f('=VLOOKUP("CHERRY",A1:C3,3,FALSE)', t)).toBe('dark red'); expect(f('=VLOOKUP("kiwi",A1:C3,2,FALSE)', t)).toBe('#N/A') })
  it('VLOOKUP approximate', () => { expect(f('=VLOOKUP(25,D1:E3,2)', t)).toBe('twenty'); expect(f('=VLOOKUP(5,D1:E3,2)', t)).toBe('#N/A'); expect(f('=VLOOKUP(99,D1:E3,2,TRUE)', t)).toBe('thirty') })
  it('VLOOKUP errors', () => { expect(f('=VLOOKUP("apple",A1:C3,9,FALSE)', t)).toBe('#REF!'); expect(f('=VLOOKUP("apple",A1:C3,0,FALSE)', t)).toBe('#VALUE!') })
  it('XLOOKUP', () => {
    expect(f('=XLOOKUP("banana",A1:A3,B1:B3)', t)).toBe(0.5)
    expect(f('=XLOOKUP("kiwi",A1:A3,B1:B3,"none")', t)).toBe('none')
    expect(f('=XLOOKUP("kiwi",A1:A3,B1:B3)', t)).toBe('#N/A')
    expect(f('=XLOOKUP(25,D1:D3,E1:E3,"-",-1)', t)).toBe('twenty')
    expect(f('=XLOOKUP(25,D1:D3,E1:E3,"-",1)', t)).toBe('thirty')
    expect(f('=XLOOKUP("ch*",A1:A3,B1:B3,"-",2)', t)).toBe(4)
    expect(f('=XLOOKUP(10,D1:D3,E1:E2)', t)).toBe('#VALUE!')
  })
  it('lookups compose with other functions', () => { expect(f('=VLOOKUP("apple",A1:C3,2,FALSE)*2+XLOOKUP("banana",A1:A3,B1:B3)', t)).toBe(3.5) })
})

describe('text functions', () => {
  it('CONCAT / CONCATENATE', () => {
    expect(f('=CONCAT("a",1,TRUE)')).toBe('a1TRUE'); expect(f('=CONCATENATE("x"," ","y")')).toBe('x y')
    expect(f('=CONCAT(A1:A3)', { A1: 'a', A2: 'b', A3: 'c' })).toBe('abc'); expect(f('="a"&"b"&1')).toBe('ab1')
  })
  it('LEFT / RIGHT', () => { expect(f('=LEFT("Office",3)')).toBe('Off'); expect(f('=RIGHT("Office",3)')).toBe('ice'); expect(f('=LEFT("abc")')).toBe('a'); expect(f('=RIGHT("abc",0)')).toBe(''); expect(f('=LEFT("abc",-1)')).toBe('#VALUE!') })
  it('LEN / UPPER / LOWER', () => { expect(f('=LEN("hello")')).toBe(5); expect(f('=UPPER("aBc")')).toBe('ABC'); expect(f('=LOWER("aBc")')).toBe('abc'); expect(f('=LEN(A1)', { A1: '12345' })).toBe(5) })
  it('nests', () => { expect(f('=UPPER(LEFT(A1,2))&LOWER(RIGHT(A1,2))', { A1: 'Spreadsheet' })).toBe('SPet') })
})

describe('references', () => {
  it('relative references shift, absolute ones stay', () => {
    expect(shiftFormula('=A1+$B$2+C$3+$D4', 1, 1)).toBe('=B2+$B$2+D$3+$D5')
    expect(shiftFormula('=SUM(A1:B2)', 0, 2)).toBe('=SUM(A3:B4)')
    expect(shiftFormula('=A1&"A1"', 1, 0)).toBe('=B1&"A1"')
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!')
  })
  it('absolute references evaluate', () => { expect(val({ A1: '4', B1: '=$A$1*2' }, 'B1')).toBe(8) })
  it('cross-sheet references', () => {
    expect(val({ A1: "=Data!B2+1" }, 'A1', { Data: { B2: '41' } })).toBe(42)
    expect(val({ A1: "=SUM('My Data'!A1:A2)" }, 'A1', { 'My Data': { A1: '1', A2: '2' } })).toBe(3)
    expect(code(val({ A1: '=Nope!A1' }, 'A1'))).toBe('#REF!')
  })
  it('whole-column ranges', () => { expect(f('=SUM(A:A)', { A1: '1', A2: '2', A50: '3' })).toBe(6) })
  it('case-insensitive function names and refs', () => { expect(val({ A1: '3', B1: '=sum(a1,2)' }, 'B1')).toBe(5) })
})

describe('errors', () => {
  it('reports the right error code and propagates', () => {
    expect(f('=1/0')).toBe('#DIV/0!'); expect(f('=FOO(1)')).toBe('#NAME?'); expect(f('=(1+')).toBe('#ERROR!'); expect(f('=1+"a"')).toBe('#VALUE!')
    expect(code(val({ A1: '=A1' }, 'A1'))).toBe('#CIRC!')
    expect(code(val({ A1: '=B1', B1: '=A1' }, 'A1'))).toBe('#CIRC!')
    expect(f('=A1+1', { A1: '=1/0' })).toBe('#DIV/0!'); expect(f('=SUM(A1:A2)', { A1: '1', A2: '=1/0' })).toBe('#DIV/0!')
  })
  it('never throws on garbage input', () => { for (const g of ['=', '=)', '=SUM(', '=1+*2', '="abc', '=A1:', '=@@', '=SUM(,,)', '=IF()']) expect(() => val({ Z1: g }, 'Z1')).not.toThrow() })
  it('every function reports #ERROR! on wrong argument counts instead of crashing', () => { for (const fn of ['ABS', 'IF', 'MOD', 'POWER', 'VLOOKUP', 'XLOOKUP', 'LEFT', 'LEN', 'NOT']) expect(f(`=${fn}()`)).toBe('#ERROR!') })
})

describe('dependency tracking and recalculation', () => {
  it('editing a referenced cell recalculates dependents', () => {
    const wb = wbOf({ A1: '10', A2: '20', A3: '=SUM(A1:A2)', B1: '=A3*2', C1: '=B1+A1' })
    const ev = new Evaluator(wb)
    expect([ev.get(0, 'A3'), ev.get(0, 'B1'), ev.get(0, 'C1')]).toEqual([30, 60, 70])
    const wb2: Workbook = { ...wb, sheets: [{ ...wb.sheets[0], cells: { ...wb.sheets[0].cells, A1: { v: '15' } } }] }
    ev.setWorkbook(wb2); ev.invalidate(0, 'A1')
    expect([ev.get(0, 'A3'), ev.get(0, 'B1'), ev.get(0, 'C1')]).toEqual([35, 70, 85])
  })
  it('invalidation is incremental: unrelated cells stay cached', () => {
    const wb = wbOf({ A1: '1', B1: '=A1+1', X1: '5', Y1: '=X1*2' })
    const ev = new Evaluator(wb); ev.get(0, 'B1'); ev.get(0, 'Y1')
    const before = ev.cachedCount
    ev.invalidate(0, 'A1')
    expect(ev.cachedCount).toBe(before - 2) // A1 and B1 dropped, X1/Y1 untouched
  })
  it('range dependencies invalidate when any cell inside the range changes (even an empty one)', () => {
    const wb = wbOf({ A1: '1', S1: '=SUM(A1:A3)' })
    const ev = new Evaluator(wb); expect(ev.get(0, 'S1')).toBe(1)
    ev.setWorkbook({ ...wb, sheets: [{ ...wb.sheets[0], cells: { ...wb.sheets[0].cells, A3: { v: '9' } } }] }); ev.invalidate(0, 'A3')
    expect(ev.get(0, 'S1')).toBe(10)
  })
  it('handles a long dependency chain without recursion problems', () => {
    const cells: Record<string, string> = { A1: '1' }
    for (let i = 2; i <= 800; i++) cells[`A${i}`] = `=A${i - 1}+1`
    expect(val(cells, 'A800')).toBe(800)
  })
  it('evaluates a large sheet quickly', () => {
    const cells: Record<string, string> = {}
    for (let r = 1; r <= 5000; r++) { cells[`A${r}`] = String(r); cells[`B${r}`] = `=A${r}*2` }
    cells.C1 = '=SUM(B1:B5000)'
    const t = performance.now(); expect(val(cells, 'C1')).toBe(25005000); expect(performance.now() - t).toBeLessThan(1500)
  })
})

describe('row/column insert and delete rewrite references', () => {
  const base = () => wbOf({ A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: '=A2*2' })
  it('insert row shifts cells and formulas', () => {
    const w = insertRows(base(), 0, 1) // new row at index 1 (row 2)
    expect(w.sheets[0].cells.A3.v).toBe('2'); expect(w.sheets[0].cells.A4.v).toBe('=SUM(A1:A3)'); expect(w.sheets[0].cells.B1.v).toBe('=A3*2')
    expect(new Evaluator(w).get(0, 'A4')).toBe(3)
  })
  it('delete row turns direct references into #REF! and shrinks ranges', () => {
    const w = deleteRows(base(), 0, 1)
    expect(w.sheets[0].cells.B1.v).toBe('=#REF!*2'); expect(new Evaluator(w).get(0, 'B1')).toBeInstanceOf(FormulaError); expect(code(new Evaluator(w).get(0, 'B1'))).toBe('#REF!'); expect(w.sheets[0].cells.A2.v).toBe('=SUM(A1:A1)')
  })
  it('insert and delete columns', () => {
    const w = insertCols(wbOf({ A1: '1', B1: '=A1+1', C1: '=B1' }), 0, 1)
    expect(w.sheets[0].cells.C1.v).toBe('=A1+1'); expect(w.sheets[0].cells.D1.v).toBe('=C1')
    expect(deleteCols(w, 0, 1).sheets[0].cells.B1.v).toBe('=A1+1')
  })
  it('formulas on other sheets that point at the changed sheet are rewritten too', () => {
    const w = insertRows(wbOf({ A1: '7' }, { Two: { A1: '=Sheet1!A1+1' } }), 0, 0)
    expect(w.sheets[1].cells.A1.v).toBe('=Sheet1!A2+1')
  })
})

describe('display and coverage', () => {
  it('formats values', () => { expect(formatValue(0.1 + 0.2)).toBe('0.3'); expect(formatValue(1234.5, 'number')).toMatch(/1,234\.50/); expect(formatValue(0.256, 'percent')).toBe('25.6%'); expect(formatValue(new FormulaError('#N/A'))).toBe('#N/A') })
  it('every function advertised in the ribbon exists in the registry', () => {
    for (const [group, names] of Object.entries(FUNCTION_GROUPS)) for (const n of names) expect(FUNCTIONS[n], `${group}/${n}`).toBeDefined()
  })
})
