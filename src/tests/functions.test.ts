import { describe, expect, it } from 'vitest'
import { balanceFormula, Evaluator, FormulaError } from '../editors/excel/formula'
import { FUNCTIONS } from '../editors/excel/functions'
import { emptyWorkbook, newSheet, type Workbook } from '../editors/excel/model'

/** Data available to every case below. */
const DATA: Record<string, string> = {
  A1: '1', A2: '2', A3: '3', A4: '4', A5: '5',
  B1: '10', B2: '20', B3: '30', B4: '40', B5: '50',
  C1: 'apple', C2: 'Banana', C3: 'cherry', C4: '', C5: 'TRUE',
  D1: '-2.5', D2: '0', D3: '7.456', D4: '100', D5: '2',
  E1: '45000', E2: '45031', E3: '0.5',
  F1: 'a', F2: 'b', F3: 'a',
}
function wbOf(extra: Record<string, string> = {}): Workbook {
  const wb = emptyWorkbook()
  for (const [k, v] of Object.entries({ ...DATA, ...extra })) wb.sheets[0].cells[k] = { v }
  return wb
}
const code = (v: unknown) => (v instanceof FormulaError ? v.code : v)
/** Evaluate one formula and return its value (errors come back as their code string). */
function f(formula: string, extra: Record<string, string> = {}): unknown {
  const wb = wbOf(extra)
  wb.sheets[0].cells.Z100 = { v: formula }
  return code(new Evaluator(wb).get(0, 'Z100'))
}
const near = (v: unknown, want: number, eps = 1e-6) => { expect(typeof v).toBe('number'); expect(Math.abs((v as number) - want)).toBeLessThan(eps) }

/**
 * One passing sample per registered function. The final test asserts this table covers
 * the whole registry, so a new function cannot be added without a test.
 */
const CASES: [string, string, unknown][] = [
  // math
  ['SUM', '=SUM(A1:A5)', 15], ['PRODUCT', '=PRODUCT(A1:A3)', 6], ['SUMSQ', '=SUMSQ(3,4)', 25],
  ['ABS', '=ABS(D1)', 2.5], ['SIGN', '=SIGN(D1)', -1], ['SQRT', '=SQRT(16)', 4],
  ['INT', '=INT(D3)', 7], ['TRUNC', '=TRUNC(D3,1)', 7.4], ['ROUND', '=ROUND(D3,2)', 7.46],
  ['ROUNDUP', '=ROUNDUP(7.41,1)', 7.5], ['ROUNDDOWN', '=ROUNDDOWN(7.49,1)', 7.4], ['MROUND', '=MROUND(10,3)', 9],
  ['CEILING', '=CEILING(7.1,1)', 8], ['CEILING.MATH', '=CEILING.MATH(7.1)', 8], ['FLOOR', '=FLOOR(7.9,1)', 7], ['FLOOR.MATH', '=FLOOR.MATH(7.9)', 7],
  ['EVEN', '=EVEN(3)', 4], ['ODD', '=ODD(4)', 5], ['MOD', '=MOD(10,3)', 1], ['QUOTIENT', '=QUOTIENT(10,3)', 3],
  ['POWER', '=POWER(2,10)', 1024], ['EXP', '=ROUND(EXP(1),4)', 2.7183], ['LN', '=LN(1)', 0],
  ['LOG10', '=LOG10(1000)', 3], ['LOG', '=LOG(8,2)', 3], ['PI', '=ROUND(PI(),4)', 3.1416],
  ['DEGREES', '=DEGREES(PI())', 180], ['RADIANS', '=ROUND(RADIANS(180),4)', 3.1416],
  ['SIN', '=ROUND(SIN(0),4)', 0], ['COS', '=COS(0)', 1], ['TAN', '=ROUND(TAN(0),4)', 0],
  ['ASIN', '=ROUND(ASIN(1),4)', 1.5708], ['ACOS', '=ACOS(1)', 0], ['ATAN', '=ROUND(ATAN(1),4)', 0.7854],
  ['SINH', '=ROUND(SINH(0),4)', 0], ['COSH', '=COSH(0)', 1], ['TANH', '=ROUND(TANH(0),4)', 0],
  ['ATAN2', '=ROUND(ATAN2(1,1),4)', 0.7854], ['FACT', '=FACT(5)', 120], ['COMBIN', '=COMBIN(5,2)', 10], ['PERMUT', '=PERMUT(5,2)', 20],
  ['GCD', '=GCD(12,18)', 6], ['LCM', '=LCM(4,6)', 12],
  ['SUMPRODUCT', '=SUMPRODUCT(A1:A3,B1:B3)', 140],
  ['SUMIF', '=SUMIF(A1:A5,">3")', 9], ['SUMIFS', '=SUMIFS(B1:B5,A1:A5,">3")', 90],
  // statistics
  ['AVERAGE', '=AVERAGE(A1:A5)', 3], ['AVERAGEA', '=AVERAGEA(A1:A5)', 3],
  ['MIN', '=MIN(A1:A5)', 1], ['MAX', '=MAX(A1:A5)', 5],
  ['COUNT', '=COUNT(A1:C5)', 10], ['COUNTA', '=COUNTA(C1:C5)', 4], ['COUNTBLANK', '=COUNTBLANK(C1:C5)', 1],
  ['COUNTIF', '=COUNTIF(A1:A5,">=3")', 3], ['COUNTIFS', '=COUNTIFS(A1:A5,">=2",B1:B5,"<50")', 3],
  ['AVERAGEIF', '=AVERAGEIF(A1:A5,">3")', 4.5], ['AVERAGEIFS', '=AVERAGEIFS(B1:B5,A1:A5,">3")', 45],
  ['MINIFS', '=MINIFS(B1:B5,A1:A5,">3")', 40], ['MAXIFS', '=MAXIFS(B1:B5,A1:A5,"<3")', 20],
  ['MEDIAN', '=MEDIAN(A1:A5)', 3], ['MODE', '=MODE(1,2,2,3)', 2], ['MODE.SNGL', '=MODE.SNGL(1,2,2,3)', 2],
  ['STDEV', '=ROUND(STDEV(A1:A5),4)', 1.5811], ['STDEV.S', '=ROUND(STDEV.S(A1:A5),4)', 1.5811],
  ['STDEVP', '=ROUND(STDEVP(A1:A5),4)', 1.4142], ['STDEV.P', '=ROUND(STDEV.P(A1:A5),4)', 1.4142],
  ['VAR', '=VAR(A1:A5)', 2.5], ['VAR.S', '=VAR.S(A1:A5)', 2.5], ['VARP', '=VARP(A1:A5)', 2], ['VAR.P', '=VAR.P(A1:A5)', 2],
  ['LARGE', '=LARGE(A1:A5,2)', 4], ['SMALL', '=SMALL(A1:A5,2)', 2],
  ['RANK', '=RANK(4,A1:A5)', 2], ['RANK.EQ', '=RANK.EQ(4,A1:A5)', 2],
  ['PERCENTILE', '=PERCENTILE(A1:A5,0.5)', 3], ['PERCENTILE.INC', '=PERCENTILE.INC(A1:A5,0.5)', 3],
  ['QUARTILE', '=QUARTILE(A1:A5,2)', 3], ['QUARTILE.INC', '=QUARTILE.INC(A1:A5,2)', 3],
  ['GEOMEAN', '=ROUND(GEOMEAN(1,4),4)', 2], ['HARMEAN', '=HARMEAN(1,1)', 1],
  ['CORREL', '=ROUND(CORREL(A1:A5,B1:B5),4)', 1], ['PEARSON', '=ROUND(PEARSON(A1:A5,B1:B5),4)', 1],
  ['SLOPE', '=SLOPE(B1:B5,A1:A5)', 10], ['INTERCEPT', '=INTERCEPT(B1:B5,A1:A5)', 0],
  // logical
  ['IF', '=IF(A1<A2,"lt","ge")', 'lt'], ['IFS', '=IFS(A1>5,"big",A1>0,"small")', 'small'],
  ['SWITCH', '=SWITCH(2,1,"one",2,"two","other")', 'two'],
  ['AND', '=AND(TRUE,1=1)', true], ['OR', '=OR(FALSE,1=1)', true], ['XOR', '=XOR(TRUE,FALSE)', true], ['NOT', '=NOT(FALSE)', true],
  ['TRUE', '=TRUE()', true], ['FALSE', '=FALSE()', false],
  ['IFERROR', '=IFERROR(1/0,"safe")', 'safe'], ['IFNA', '=IFNA(NA(),"gone")', 'gone'],
  // information
  ['ISBLANK', '=ISBLANK(C4)', true], ['ISNUMBER', '=ISNUMBER(A1)', true], ['ISTEXT', '=ISTEXT(C1)', true],
  ['ISNONTEXT', '=ISNONTEXT(A1)', true], ['ISLOGICAL', '=ISLOGICAL(TRUE)', true],
  ['ISERROR', '=ISERROR(1/0)', true], ['ISERR', '=ISERR(1/0)', true], ['ISNA', '=ISNA(NA())', true],
  ['ISEVEN', '=ISEVEN(4)', true], ['ISODD', '=ISODD(3)', true], ['NA', '=NA()', '#N/A'],
  ['N', '=N(A1)', 1], ['T', '=T(C1)', 'apple'], ['TYPE', '=TYPE(A1)', 1], ['ERROR.TYPE', '=ERROR.TYPE(1/0)', 2],
  // text
  ['CONCAT', '=CONCAT("a","b",1)', 'ab1'], ['CONCATENATE', '=CONCATENATE("x","y")', 'xy'],
  ['TEXTJOIN', '=TEXTJOIN("-",TRUE,C1:C3)', 'apple-Banana-cherry'],
  ['LEN', '=LEN("hello")', 5], ['LEFT', '=LEFT("Office",3)', 'Off'], ['RIGHT', '=RIGHT("Office",3)', 'ice'],
  ['MID', '=MID("Office",2,3)', 'ffi'], ['UPPER', '=UPPER("aB")', 'AB'], ['LOWER', '=LOWER("Ab")', 'ab'],
  ['PROPER', '=PROPER("hello world")', 'Hello World'], ['TRIM', '=TRIM("  a  b  ")', 'a b'],
  ['CLEAN', '=CLEAN("a b")', 'a b'], ['REPT', '=REPT("ab",3)', 'ababab'], ['EXACT', '=EXACT("a","a")', true],
  ['FIND', '=FIND("f","Office")', 2], ['SEARCH', '=SEARCH("F","Office")', 2],
  ['SUBSTITUTE', '=SUBSTITUTE("a-b-c","-","+")', 'a+b+c'], ['REPLACE', '=REPLACE("abcdef",2,3,"X")', 'aXef'],
  ['CHAR', '=CHAR(65)', 'A'], ['UNICHAR', '=UNICHAR(8364)', '€'], ['CODE', '=CODE("A")', 65], ['UNICODE', '=UNICODE("A")', 65],
  ['VALUE', '=VALUE("1,234")', 1234], ['NUMBERVALUE', '=NUMBERVALUE("12.5")', 12.5],
  ['TEXTBEFORE', '=TEXTBEFORE("a-b","-")', 'a'], ['TEXTAFTER', '=TEXTAFTER("a-b","-")', 'b'],
  ['FIXED', '=FIXED(1234.567,2)', '1,234.57'], ['DOLLAR', '=DOLLAR(5,2)', '$5.00'],
  ['TEXT', '=TEXT(0.25,"0.00%")', '25.00%'],
  // date and time
  ['TODAY', '=TODAY()>44000', true], ['NOW', '=NOW()>44000', true],
  ['DATE', '=DATE(2024,3,15)', 45366], ['TIME', '=ROUND(TIME(12,0,0),4)', 0.5],
  ['YEAR', '=YEAR(45366)', 2024], ['MONTH', '=MONTH(45366)', 3], ['DAY', '=DAY(45366)', 15],
  ['HOUR', '=HOUR(0.5)', 12], ['MINUTE', '=MINUTE(0.5)', 0], ['SECOND', '=SECOND(0.5)', 0],
  ['WEEKDAY', '=WEEKDAY(45366)', 6], ['WEEKNUM', '=WEEKNUM(45366)', 11],
  ['DAYS', '=DAYS(45366,45365)', 1], ['DAYS360', '=DAYS360(45366,45396)', 29],
  ['EDATE', '=EDATE(45366,1)', 45397], ['EOMONTH', '=EOMONTH(45366,0)', 45382],
  ['DATEDIF', '=DATEDIF(45000,45366,"M")', 12], ['DATEVALUE', '=DATEVALUE("2024-03-15")', 45366],
  ['TIMEVALUE', '=ROUND(TIMEVALUE("12:00"),4)', 0.5],
  ['NETWORKDAYS', '=NETWORKDAYS(45366,45372)', 5], ['WORKDAY', '=WORKDAY(45366,1)', 45369],
  // lookup
  ['CHOOSE', '=CHOOSE(2,"a","b","c")', 'b'],
  ['VLOOKUP', '=VLOOKUP(3,A1:B5,2,FALSE)', 30], ['HLOOKUP', '=HLOOKUP(1,A1:B2,2,FALSE)', 2],
  ['LOOKUP', '=LOOKUP(3,A1:A5,B1:B5)', 30],
  ['MATCH', '=MATCH(3,A1:A5,0)', 3], ['XMATCH', '=XMATCH(3,A1:A5)', 3],
  ['XLOOKUP', '=XLOOKUP(3,A1:A5,B1:B5)', 30],
  ['INDEX', '=INDEX(A1:B5,2,2)', 20], ['ROWS', '=ROWS(A1:A5)', 5], ['COLUMNS', '=COLUMNS(A1:B5)', 2],
  ['TRANSPOSE', '=INDEX(TRANSPOSE(A1:A5),1,2)', 2],
  ['UNIQUE', '=INDEX(UNIQUE(F1:F3),1,1)', 'a'], ['SORT', '=INDEX(SORT(A1:A5,1,-1),1,1)', 5],
  ['FILTER', '=INDEX(FILTER(A1:A5,A1:A5>3),1,1)', 4], ['SEQUENCE', '=INDEX(SEQUENCE(3,1,10,5),3,1)', 20],
  // financial
  ['PMT', '=ROUND(PMT(0.05/12,360,200000),2)', -1073.64],
  ['FV', '=ROUND(FV(0.05/12,120,-100),2)', 15528.23],
  ['PV', '=ROUND(PV(0.05/12,120,-100),2)', 9428.14],
  ['NPER', '=ROUND(NPER(0.05/12,-1073.64,200000),0)', 360],
  ['RATE', '=ROUND(RATE(360,-1073.64,200000)*12,4)', 0.05],
  ['NPV', '=ROUND(NPV(0.1,100,100),2)', 173.55],
  ['IRR', '=ROUND(IRR(G1:G3),4)', 0.1],
  ['IPMT', '=ROUND(IPMT(0.05/12,1,360,200000),2)', -833.33],
  ['PPMT', '=ROUND(PPMT(0.05/12,1,360,200000),2)', -240.31],
  ['SLN', '=SLN(1000,100,5)', 180], ['SYD', '=SYD(1000,100,5,1)', 300],
  // engineering
  ['DEC2BIN', '=DEC2BIN(10)', '1010'], ['DEC2HEX', '=DEC2HEX(255)', 'FF'], ['DEC2OCT', '=DEC2OCT(8)', '10'],
  ['BIN2DEC', '=BIN2DEC("1010")', 10], ['HEX2DEC', '=HEX2DEC("FF")', 255], ['OCT2DEC', '=OCT2DEC("10")', 8],
  ['BASE', '=BASE(255,16)', 'FF'], ['DECIMAL', '=DECIMAL("FF",16)', 255],
  ['BITAND', '=BITAND(12,10)', 8], ['BITOR', '=BITOR(12,10)', 14], ['BITXOR', '=BITXOR(12,10)', 6],
]

describe('function library', () => {
  const extra = { G1: '-1000', G2: '500', G3: '660' }
  for (const [name, formula, want] of CASES) {
    it(`${name}  ${formula}`, () => {
      const got = f(formula, extra)
      if (typeof want === 'number' && !Number.isInteger(want)) near(got, want, 0.011)
      else expect(got).toEqual(want)
    })
  }

  it('covers every function in the registry', () => {
    const tested = new Set(CASES.map(c => c[0]))
    const elsewhere = ['RAND', 'RANDBETWEEN'] // non-deterministic: asserted by range in their own test
    const missing = Object.keys(FUNCTIONS).filter(n => !tested.has(n) && !elsewhere.includes(n))
    expect(missing, 'functions without a test case').toEqual([])
  })

  it('no case returned an error value', () => {
    const extra2 = { G1: '-1000', G2: '500', G3: '660' }
    const failures = CASES.filter(([, formula, want]) => {
      if (typeof want === 'string' && want.startsWith('#')) return false
      const v = f(formula, extra2)
      return typeof v === 'string' && v.startsWith('#')
    }).map(c => `${c[0]} -> ${f(c[1], extra2)}`)
    expect(failures).toEqual([])
  })
})

describe('volatile functions and RAND', () => {
  it('RAND and RANDBETWEEN stay in range', () => {
    for (let i = 0; i < 20; i++) {
      const r = f('=RAND()') as number
      expect(r).toBeGreaterThanOrEqual(0); expect(r).toBeLessThan(1)
      const b = f('=RANDBETWEEN(5,7)') as number
      expect([5, 6, 7]).toContain(b)
    }
  })
})

describe('dynamic arrays spill into neighbouring cells', () => {
  function spill(formula: string, cells: Record<string, string> = {}) {
    const wb = emptyWorkbook()
    for (const [k, v] of Object.entries(cells)) wb.sheets[0].cells[k] = { v }
    wb.sheets[0].cells.A1 = { v: formula }
    const ev = new Evaluator(wb)
    ev.ensureFormulas(0)
    return (addr: string) => code(ev.displayValue(0, addr))
  }
  it('SEQUENCE fills a block', () => {
    const at = spill('=SEQUENCE(3,2)')
    expect([at('A1'), at('B1'), at('A2'), at('B2'), at('A3'), at('B3')]).toEqual([1, 2, 3, 4, 5, 6])
  })
  it('SORT and UNIQUE spill down a column', () => {
    const at = spill('=SORT(D1:D3)', { D1: '3', D2: '1', D3: '2' })
    expect([at('A1'), at('A2'), at('A3')]).toEqual([1, 2, 3])
    const u = spill('=UNIQUE(D1:D4)', { D1: 'a', D2: 'b', D3: 'a', D4: 'c' })
    expect([u('A1'), u('A2'), u('A3'), u('A4')]).toEqual(['a', 'b', 'c', null])
  })
  it('a blocked spill reports #SPILL! and leaves the obstruction intact', () => {
    const at = spill('=SEQUENCE(3)', { A2: 'in the way' })
    expect(at('A1')).toBe('#SPILL!')
    expect(at('A2')).toBe('in the way')
  })
  it('spilled values can be read by other formulas and update when the source changes', () => {
    const wb = emptyWorkbook()
    Object.assign(wb.sheets[0].cells, { D1: { v: '3' }, D2: { v: '1' }, A1: { v: '=SORT(D1:D2)' }, C1: { v: '=A2+100' } })
    const ev = new Evaluator(wb)
    ev.ensureFormulas(0)
    expect(ev.get(0, 'C1')).toBe(103)
    const wb2: Workbook = { ...wb, sheets: [{ ...wb.sheets[0], cells: { ...wb.sheets[0].cells, D1: { v: '9' } } }] }
    ev.setWorkbook(wb2); ev.invalidate(0, 'D1'); ev.ensureFormulas(0)
    expect(ev.get(0, 'C1')).toBe(109)
  })
})

describe('tokenizer robustness', () => {
  it('accepts spaces inside and around references', () => {
    expect(f('=SUM( A1 : A3 )')).toBe(6)
    expect(f('= SUM(A1:A3)')).toBe(6)
    expect(f('=SUM(A1 , A2)')).toBe(3)
    expect(f('= A1 + A2 ')).toBe(3)
  })
  it('accepts lower-case and mixed-case names and references', () => {
    expect(f('=sum(a1:a3)')).toBe(6)
    expect(f('=Sum(A1:a3)')).toBe(6)
    expect(f('=vlookup(3,a1:b5,2,false)')).toBe(30)
  })
  it('accepts semicolons as argument separators', () => { expect(f('=SUM(A1;A2)')).toBe(3) })
  it('handles whole-column ranges with spaces', () => { expect(f('=COUNT(A : A)')).toBe(5) })
  it('reports #NAME? for unknown functions rather than crashing', () => {
    expect(f('=NOTAFUNCTION(1)')).toBe('#NAME?')
    expect(f('=XYZZY()')).toBe('#NAME?')
  })
  it('never throws on malformed input', () => {
    for (const g of ['=', '=)', '=SUM(', '=1+*2', '="abc', '=A1:', '=@@', '=SUM(,,)', '=IF()', '=((((', '=1 2 3', '=:::', '=A1:::A2']) {
      expect(() => f(g), g).not.toThrow()
    }
  })
  it('enforces argument counts', () => {
    expect(f('=ABS()')).toBe('#ERROR!')
    expect(f('=ABS(1,2)')).toBe('#ERROR!')
    expect(f('=MID("abc",1)')).toBe('#ERROR!')
  })
})

describe('cross-sheet and structural behaviour', () => {
  it('reads other sheets by name, including names needing quotes', () => {
    const wb = emptyWorkbook()
    const s2 = newSheet('Data'); s2.cells.A1 = { v: '7' }
    const s3 = newSheet('My Data'); s3.cells.A1 = { v: '5' }
    wb.sheets.push(s2, s3)
    wb.sheets[0].cells.A1 = { v: "=Data!A1 + 'My Data'!A1" }
    expect(new Evaluator(wb).get(0, 'A1')).toBe(12)
  })
  it('reports #REF! for a missing sheet', () => { expect(f('=Missing!A1')).toBe('#REF!') })
})

describe('array literals and parenthesis balancing', () => {
  it('evaluates array constants', () => {
    expect(f('=SUM({1,2,3})')).toBe(6)
    expect(f('=SUM({1,2;3,4})')).toBe(10)
    expect(f('=SUMPRODUCT({1,2},{3,4})')).toBe(11)
    expect(f('=INDEX({10,20;30,40},2,1)')).toBe(30)
    expect(f('=COUNT({1,"a",3})')).toBe(2)
  })
  it('rejects ragged array literals', () => { expect(f('=SUM({1,2;3})')).toBe('#VALUE!') })
  it('closes unbalanced parentheses on commit, as Excel does', () => {
    expect(balanceFormula('=SUM(A1:A3')).toBe('=SUM(A1:A3)')
    expect(balanceFormula('=IF(A1>1,SUM(A1:A2')).toBe('=IF(A1>1,SUM(A1:A2))')
    expect(balanceFormula('=SUM(A1:A3)')).toBe('=SUM(A1:A3)')
    expect(balanceFormula('="a("')).toBe('="a("')
    expect(balanceFormula('plain text (')).toBe('plain text (')
  })
  it('array comparisons broadcast', () => {
    expect(f('=SUM(IF(A1:A5>3,1,0))', {})).toBe(2)
    expect(f('=SUMPRODUCT((A1:A5>3)*1)')).toBe(2)
    expect(f('=SUM(IF(A1:A5>3,B1:B5,0))')).toBe(90)
    expect(f('=COUNT(FILTER(A1:A5,A1:A5>=4))')).toBe(2)
  })
})
