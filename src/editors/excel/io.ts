import * as XLSX from 'xlsx'
import { Evaluator, formatValue } from './formula'
import { newSheet, toAddr, type Sheet, type Workbook } from './model'

function sheetFromAoA(name: string, ws: XLSX.WorkSheet): Sheet {
  const s = newSheet(name)
  const ref = ws['!ref']
  if (!ref) return s
  const range = XLSX.utils.decode_range(ref)
  s.rows = Math.max(100, range.e.r + 20)
  s.cols = Math.max(26, range.e.c + 3)
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined
      if (!cell) continue
      let v = ''
      if (cell.f) v = '=' + cell.f
      else if (cell.t === 'n') v = String(cell.v)
      else if (cell.t === 'b') v = cell.v ? 'TRUE' : 'FALSE'
      else if (cell.t === 'd' && cell.v instanceof Date) v = cell.v.toISOString().slice(0, 10)
      else if (cell.v != null) v = String(cell.v)
      if (v !== '') s.cells[toAddr(c, r)] = { v }
    }
  }
  return s
}

export function workbookFromXlsx(buf: ArrayBuffer): Workbook {
  const wb = XLSX.read(buf, { type: 'array', cellFormula: true })
  const sheets = wb.SheetNames.map(n => sheetFromAoA(n, wb.Sheets[n]))
  return { sheets: sheets.length ? sheets : [newSheet('Sheet1')], active: 0 }
}
export function workbookFromCsv(text: string): Workbook {
  const wb = XLSX.read(text, { type: 'string', raw: true })
  return { sheets: [sheetFromAoA('Sheet1', wb.Sheets[wb.SheetNames[0]])], active: 0 }
}

function toXlsxBook(wb: Workbook): XLSX.WorkBook {
  const ev = new Evaluator(wb)
  const out = XLSX.utils.book_new()
  wb.sheets.forEach((s, si) => {
    const ws: XLSX.WorkSheet = {}
    let maxR = 0, maxC = 0
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.v === '') continue
      const pos = XLSX.utils.decode_cell(addr)
      maxR = Math.max(maxR, pos.r); maxC = Math.max(maxC, pos.c)
      const val = ev.get(si, addr)
      const isErr = typeof val === 'object' && val !== null
      if (cell.v.startsWith('=')) {
        ws[addr] = isErr ? { t: 'e', v: 0, f: cell.v.slice(1) } : typeof val === 'number' ? { t: 'n', v: val, f: cell.v.slice(1) } : typeof val === 'boolean' ? { t: 'b', v: val, f: cell.v.slice(1) } : { t: 's', v: String(val ?? ''), f: cell.v.slice(1) }
      } else if (typeof val === 'number') ws[addr] = { t: 'n', v: val }
      else if (typeof val === 'boolean') ws[addr] = { t: 'b', v: val }
      else ws[addr] = { t: 's', v: String(val ?? '') }
    }
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } })
    XLSX.utils.book_append_sheet(out, ws, s.name.slice(0, 31).replace(/[\\/?*[\]:]/g, '_') || `Sheet${si + 1}`)
  })
  return out
}
export function exportXlsx(wb: Workbook): Blob {
  const out = XLSX.write(toXlsxBook(wb), { type: 'array', bookType: 'xlsx' })
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}
export function exportCsv(wb: Workbook): string {
  const s = wb.sheets[wb.active]
  const ev = new Evaluator(wb)
  let maxR = -1, maxC = -1
  for (const k of Object.keys(s.cells)) { if (!s.cells[k].v) continue; const p = XLSX.utils.decode_cell(k); maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c) }
  const lines: string[] = []
  for (let r = 0; r <= maxR; r++) {
    const row: string[] = []
    for (let c = 0; c <= maxC; c++) {
      const a = toAddr(c, r)
      const t = formatValue(ev.get(wb.active, a), s.cells[a]?.s?.fmt)
      row.push(/[",\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t)
    }
    lines.push(row.join(','))
  }
  return lines.join('\n')
}
