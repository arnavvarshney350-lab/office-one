import type { FileKind } from '../types'
import { newSheet, toAddr, type CellStyle, type CondRule, type Sheet, type Workbook } from '../editors/excel/model'

export interface StarterTemplate { id: string; kind: FileKind; name: string; description: string; build: () => unknown }

/* ---------------------------------------------------------------- Excel helpers */
/** Build a sheet from a grid of strings; `styles` maps A1 addresses to formatting. */
function sheetOf(name: string, rows: string[][], styles: Record<string, CellStyle> = {}, opts: { colW?: Record<number, number>; cf?: CondRule[]; freeze?: Sheet['freeze'] } = {}): Sheet {
  const s = newSheet(name)
  rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') s.cells[toAddr(c, r)] = { v } }))
  for (const [addr, st] of Object.entries(styles)) s.cells[addr] = { v: s.cells[addr]?.v ?? '', s: { ...s.cells[addr]?.s, ...st } }
  if (opts.colW) s.colW = opts.colW
  if (opts.cf) s.cf = opts.cf
  if (opts.freeze) s.freeze = opts.freeze
  return s
}
const head: CellStyle = { b: true, bg: '#e7efff', color: '#1e3a8a' }
const money: CellStyle = { fmt: 'currency' }
const wb = (s: Sheet): Workbook => ({ sheets: [s], active: 0 })
const range = (col: string, from: number, to: number, f: (r: number) => CellStyle): Record<string, CellStyle> => Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [`${col}${from + i}`, f(from + i)]))

function budget(): Workbook {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
  const rows: string[][] = [
    ['Monthly budget', ...months, 'Total'],
    ['Income'],
    ['Salary', '4200', '4200', '4200', '4200', '4200', '4200', '=SUM(B3:G3)'],
    ['Other income', '150', '0', '300', '120', '0', '250', '=SUM(B4:G4)'],
    ['Total income', ...months.map((_, i) => `=SUM(${'BCDEFG'[i]}3:${'BCDEFG'[i]}4)`), '=SUM(B5:G5)'],
    ['Expenses'],
    ['Rent', '1400', '1400', '1400', '1400', '1400', '1400', '=SUM(B7:G7)'],
    ['Groceries', '520', '480', '510', '530', '495', '560', '=SUM(B8:G8)'],
    ['Transport', '160', '150', '170', '155', '165', '180', '=SUM(B9:G9)'],
    ['Utilities', '210', '195', '180', '170', '165', '190', '=SUM(B10:G10)'],
    ['Total expenses', ...months.map((_, i) => `=SUM(${'BCDEFG'[i]}7:${'BCDEFG'[i]}10)`), '=SUM(B11:G11)'],
    [''],
    ['Savings', ...months.map((_, i) => `=${'BCDEFG'[i]}5-${'BCDEFG'[i]}11`), '=SUM(B13:G13)'],
    ['Savings rate', ...months.map((_, i) => `=IF(${'BCDEFG'[i]}5=0,0,${'BCDEFG'[i]}13/${'BCDEFG'[i]}5)`), '=IF(H5=0,0,H13/H5)'],
  ]
  const styles: Record<string, CellStyle> = {}
  for (const c of 'ABCDEFGH') styles[`${c}1`] = head
  for (const r of [2, 6]) styles[`A${r}`] = { b: true, color: '#475467' }
  for (const r of [5, 11, 13]) for (const c of 'ABCDEFGH') styles[`${c}${r}`] = { b: true, ...(c === 'A' ? {} : money) }
  for (const r of [3, 4, 7, 8, 9, 10]) for (const c of 'BCDEFGH') styles[`${c}${r}`] = money
  for (const c of 'BCDEFGH') styles[`${c}14`] = { fmt: 'percent', dec: 1 }
  const cf: CondRule[] = [{ id: 'neg', range: 'B13:H13', kind: 'lt', a: '0', bg: '#fecaca' }, { id: 'pos', range: 'B13:H13', kind: 'gt', a: '0', bg: '#bbf7d0' }]
  return wb(sheetOf('Budget', rows, styles, { colW: { 0: 150 }, cf, freeze: { rows: 1, cols: 1 } }))
}

function invoice(): Workbook {
  const rows: string[][] = [
    ['INVOICE'], ['Bill to:', 'Customer name'], ['Invoice #', 'INV-0001'], ['Date', '=TODAY()'], [''],
    ['Description', 'Qty', 'Unit price', 'Amount'],
    ['Design work', '10', '85', '=B7*C7'], ['Development', '24', '95', '=B8*C8'], ['Hosting (annual)', '1', '240', '=B9*C9'], ['', '', '', ''], ['', '', '', ''],
    ['', '', 'Subtotal', '=SUM(D7:D11)'], ['', '', 'Tax rate', '0.08'], ['', '', 'Tax', '=D12*D13'], ['', '', 'Total due', '=D12+D14'],
  ]
  const styles: Record<string, CellStyle> = { A1: { b: true, size: 22, color: '#1e3a8a' }, D4: { fmt: 'date', align: 'left' }, D13: { fmt: 'percent' }, D15: { b: true, fmt: 'currency', bg: '#e7efff', size: 15 }, C15: { b: true, align: 'right' }, A2: { b: true }, A3: { b: true }, A4: { b: true } }
  for (const c of 'ABCD') styles[`${c}6`] = head
  Object.assign(styles, range('C', 7, 11, () => money), range('D', 7, 12, () => money), { D14: money })
  return wb(sheetOf('Invoice', rows, styles, { colW: { 0: 220, 1: 90, 2: 110, 3: 120 } }))
}

function gradebook(): Workbook {
  const names = ['Ava Chen', 'Ben Okafor', 'Chloe Martin', 'Dev Patel', 'Ella Rossi', 'Finn Larsen']
  const scores = [[88, 92, 79], [65, 70, 58], [95, 98, 91], [72, 64, 80], [55, 61, 49], [83, 77, 90]]
  const rows: string[][] = [['Student', 'Quiz 1', 'Quiz 2', 'Exam', 'Average', 'Result'],
    ...names.map((n, i) => [n, ...scores[i].map(String), `=ROUND(AVERAGE(B${i + 2}:D${i + 2}),1)`, `=IF(E${i + 2}>=60,"Pass","Needs support")`]),
    [''], ['Class average', '=ROUND(AVERAGE(B2:B7),1)', '=ROUND(AVERAGE(C2:C7),1)', '=ROUND(AVERAGE(D2:D7),1)', '=ROUND(AVERAGE(E2:E7),1)', '=COUNTIF(F2:F7,"Pass")&" of "&COUNTA(A2:A7)&" passing"']]
  const styles: Record<string, CellStyle> = { A9: { b: true } }
  for (const c of 'ABCDEF') styles[`${c}1`] = head
  for (const c of 'BCDE') styles[`${c}9`] = { b: true }
  const cf: CondRule[] = [{ id: 'low', range: 'B2:E7', kind: 'lt', a: '60', bg: '#fecaca' }, { id: 'top', range: 'E2:E7', kind: 'top', a: '1', bg: '#bbf7d0' }]
  return wb(sheetOf('Grades', rows, styles, { colW: { 0: 150, 5: 140 }, cf, freeze: { rows: 1, cols: 0 } }))
}

/* ---------------------------------------------------------------- Word helpers */
const check = (items: string[]) => `<ul data-type="taskList">${items.map(t => `<li data-type="taskItem" data-checked="false"><label><input type="checkbox"><span></span></label><div><p>${t}</p></div></li>`).join('')}</ul>`
const doc = (html: string) => ({ html })

export const STARTER_TEMPLATES: StarterTemplate[] = [
  { id: 'meeting', kind: 'docx', name: 'Meeting notes', description: 'Agenda, decisions and action items',
    build: () => doc(`<h1>Meeting notes</h1><p><strong>Date:</strong> &nbsp; <strong>Attendees:</strong> &nbsp; <strong>Facilitator:</strong></p><h2>Agenda</h2><ol><li><p>Welcome and updates</p></li><li><p>Main discussion</p></li><li><p>Next steps</p></li></ol><h2>Notes</h2><p></p><h2>Decisions</h2><ul><li><p></p></li></ul><h2>Action items</h2>${check(['Owner — task — due date'])}`) },
  { id: 'letter', kind: 'docx', name: 'Business letter', description: 'A clean formal letter layout',
    build: () => doc(`<p>Your Name</p><p>Street address, City, Postcode</p><p></p><p>Date</p><p></p><p>Recipient name</p><p>Company</p><p>Address</p><p></p><p>Dear Recipient,</p><p>Open with the purpose of your letter in a sentence or two. Then give the details the reader needs, keeping paragraphs short and specific.</p><p>Close by saying what you would like to happen next, and thank them for their time.</p><p></p><p>Sincerely,</p><p></p><p>Your Name</p>`) },
  { id: 'proposal', kind: 'docx', name: 'Project proposal', description: 'Summary, scope, timeline and budget',
    build: () => doc(`<h1>Project proposal</h1><p style="text-align:center"><em>Prepared by · Date</em></p><h2>Summary</h2><p>One paragraph describing the problem, the proposed solution and the outcome.</p><h2>Goals</h2><ul><li><p>Goal one</p></li><li><p>Goal two</p></li></ul><h2>Scope</h2><p>What is included, and what is not.</p><h2>Timeline</h2><table><tbody><tr><th><p>Phase</p></th><th><p>Dates</p></th><th><p>Deliverable</p></th></tr><tr><td><p>Discovery</p></td><td><p></p></td><td><p></p></td></tr><tr><td><p>Build</p></td><td><p></p></td><td><p></p></td></tr><tr><td><p>Launch</p></td><td><p></p></td><td><p></p></td></tr></tbody></table><h2>Budget</h2><p>Estimated cost and assumptions.</p>`) },
  { id: 'resume', kind: 'docx', name: 'Resume', description: 'One-page resume with clear sections',
    build: () => doc(`<h1 style="text-align:center">Your Name</h1><p style="text-align:center">City · email@example.com · +00 000 000 000</p><h2>Experience</h2><p><strong>Job title</strong> — Company <em>(2022 – present)</em></p><ul><li><p>Achievement with a measurable result</p></li><li><p>Responsibility that shows scope</p></li></ul><h2>Education</h2><p><strong>Degree</strong> — Institution <em>(Year)</em></p><h2>Skills</h2><p>Skill one · Skill two · Skill three</p>`) },
  { id: 'budget', kind: 'xlsx', name: 'Monthly budget', description: 'Income, expenses and savings with formulas', build: budget },
  { id: 'invoice', kind: 'xlsx', name: 'Invoice', description: 'Line items, tax and total due', build: invoice },
  { id: 'gradebook', kind: 'xlsx', name: 'Gradebook', description: 'Averages, pass/fail and highlights', build: gradebook },
]
