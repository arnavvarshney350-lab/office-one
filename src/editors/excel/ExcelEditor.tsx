import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { AlignCenter, AlignLeft, AlignRight, ArrowDownAZ, ArrowDownToLine, ArrowDownZA, ArrowRightToLine, BarChart3, Bold, ChevronDown, ClipboardPaste, Copy, Grid2x2, Italic, PaintBucket, Palette, Plus, Redo2, Replace, Scissors, Search, Sigma, Snowflake, Strikethrough, Type, Underline as UnderlineIcon, Undo2, WrapText, X, ChevronUp, MoveHorizontal, Trash2, Pencil } from 'lucide-react'
import { Button, ColorMenu, ConfirmModal, ContextMenu, IconButton, Menu, PromptModal, Ribbon, RibbonGroup, cx, type MenuEntry } from '../../design-system'
import { FontPicker, SizeBox } from '../../design-system/FontPicker'
import { cssFamily, preloadFonts } from '../../services/fonts'
import { EditorHeader, type MenuItem } from '../../components/EditorHeader'
import { useAutosave } from '../../hooks/useAutosave'
import { useHistory } from '../../hooks/useHistory'
import { createFile, saveContent } from '../../file-system/fileService'
import { download, safeName } from '../../services/download'
import { useToast } from '../../contexts/ToastContext'
import { balanceFormula, FormulaError, formatValue, validateFormula, type Scalar } from './formula'
import { FUNCTION_GROUPS, FUNCTION_HINTS, FUNCTIONS } from './functions'
import { useEvaluator } from './useEvaluator'
import { colName, deleteCols, deleteRows, insertCols, insertRows, newSheet, parseAddr, shiftFormula, toAddr, type Cell, type CellStyle, type Chart, type CondKind, type CondRule, type Sheet, type Workbook } from './model'
import { cfContext, cfFor, chartData, fillRange, findCells, parseRange, rangeToA1, replaceInCell, sortRange, type Rect } from './sheetOps'
import { ChartView } from './ChartView'
import { CF_KINDS, ChartDialog, CondFormatDialog } from './Dialogs'
import { uid } from '../../utilities/id'
import type { OFile } from '../../types'

const FN_NAMES = Object.keys(FUNCTIONS).sort()
const activeIdxOf = (wb: Workbook) => Math.min(Math.max(0, wb.active), wb.sheets.length - 1)
const ROW_H = 24, HDR_W = 46, DEF_W = 100, OVERSCAN = 8
interface Pos { r: number; c: number }
interface Clip { cells: (Cell | undefined)[][]; origin: Pos; text: string }

function normalize(c: unknown): Workbook {
  if (typeof c === 'string') { return { sheets: [{ ...newSheet('Sheet1'), cells: Object.fromEntries(c.split(/\r?\n/).flatMap((l, r) => l.split(',').map((v, ci) => [toAddr(ci, r), { v }] as const).filter(([, x]) => x.v))) }], active: 0 } }
  const w = c as Workbook | null
  return w && Array.isArray(w.sheets) && w.sheets.length ? w : { sheets: [newSheet('Sheet1')], active: 0 }
}

interface TdProps { r: number; c: number; display: string; style?: CellStyle; isErr: boolean; isNum: boolean; spilled?: boolean; cf?: { bg?: string; color?: string } | null; stickyTop?: number; stickyLeft?: number }
const Td = memo(function Td({ r, c, display, style, isErr, isNum, spilled, cf, stickyTop, stickyLeft }: TdProps) {
  const sticky = stickyTop !== undefined || stickyLeft !== undefined
  return (
    <td data-r={r} data-c={c} title={isErr ? display : undefined}
      style={{
        height: ROW_H, color: isErr ? 'var(--danger)' : cf?.color ?? (spilled ? 'var(--spill)' : style?.color ?? ((style?.bg || cf?.bg) ? '#141824' : undefined)), background: cf?.bg ?? style?.bg ?? (sticky ? 'var(--surface)' : undefined),
        fontWeight: style?.b ? 700 : undefined, fontStyle: style?.i ? 'italic' : undefined,
        textDecoration: [style?.u && 'underline', style?.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
        fontFamily: style?.font ? cssFamily(style.font) : undefined, fontSize: style?.size ? style.size : undefined,
        textAlign: style?.align ?? (isErr ? 'center' : isNum ? 'right' : 'left'),
        boxShadow: style?.border ? 'inset 0 0 0 1px var(--text)' : undefined,
        whiteSpace: style?.wrap ? 'pre-wrap' : 'nowrap', overflowWrap: style?.wrap ? 'anywhere' : undefined,
        position: sticky ? 'sticky' : undefined, top: stickyTop, left: stickyLeft, zIndex: sticky ? (stickyTop !== undefined && stickyLeft !== undefined ? 4 : 3) : undefined,
      }}
      className="border-r border-b border-line px-1.5 overflow-hidden text-ellipsis text-[13px] scroll-mt-6 scroll-ml-12">{display}</td>
  )
})

export function ExcelEditor({ file }: { file: OFile }) {
  const nav = useNavigate()
  const toast = useToast()
  const hist = useHistory<Workbook>(normalize(file.content))
  const wb = hist.state
  const sheet: Sheet = wb.sheets[activeIdxOf(wb)]
  const activeIdx = wb.sheets.indexOf(sheet)
  const [version, setVersion] = useState(0)
  const [tab, setTab] = useState('Home')
  const [anchor, setAnchor] = useState<Pos>({ r: 0, c: 0 })
  const [focus, setFocus] = useState<Pos>({ r: 0, c: 0 })
  const [draft, setDraft] = useState<string | null>(null)
  const [nameBox, setNameBox] = useState<string | null>(null)
  const [renaming, setRenaming] = useState<number | null>(null)
  const [clip, setClip] = useState<Clip | null>(null)
  const [scroll, setScroll] = useState({ top: 0, h: 600 })
  const [ctx, setCtx] = useState<{ x: number; y: number; items: MenuEntry[] } | null>(null)
  const [confirmDel, setConfirmDel] = useState<number | null>(null)
  const [sugIdx, setSugIdx] = useState(0)
  const [sugOpen, setSugOpen] = useState(true)
  const [dialog, setDialog] = useState<null | { kind: 'cf'; initial?: CondKind } | { kind: 'chart'; edit?: Chart }>(null)
  const [selChart, setSelChart] = useState<string | null>(null)
  const [findOpen, setFindOpen] = useState(false)
  const [replaceOpen, setReplaceOpen] = useState(false)
  const [findQ, setFindQ] = useState(''), [replQ, setReplQ] = useState('')
  const [findCase, setFindCase] = useState(false), [findWhole, setFindWhole] = useState(false), [findFormulas, setFindFormulas] = useState(false)
  const [findIdx, setFindIdx] = useState(0)
  const chartGesture = useRef<{ id: string; mode: 'move' | 'size'; sx: number; sy: number; o: Chart; base: Workbook } | null>(null)
  const dragging = useRef(false)
  const fillRef = useRef<Rect | null>(null)
  const [fillTo, setFillTo] = useState<{ r: number; c: number } | null>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const [barFocused, setBarFocused] = useState(false)
  const importRef = useRef<HTMLInputElement>(null)
  const resize = useRef<{ c: number; x: number; w: number } | null>(null)

  const wbRef = useRef(wb); wbRef.current = wb
  const commit = useCallback((next: Workbook) => { hist.commit(next); setVersion(v => v + 1) }, [hist])
  const save = useCallback(async () => { await saveContent(file.id, wbRef.current) }, [file.id])
  const saveState = useAutosave(version, save)

  const ev = useEvaluator(wb)
  // Dynamic arrays must be evaluated before paint so the cells they spill into render their values.
  const spillTick = useMemo(() => { ev.invalidateVolatile(); ev.ensureFormulas(activeIdxOf(wb)); return version }, [ev, wb, version])
  const rect = { r1: Math.min(anchor.r, focus.r), r2: Math.max(anchor.r, focus.r), c1: Math.min(anchor.c, focus.c), c2: Math.max(anchor.c, focus.c) }
  const addr = toAddr(focus.c, focus.r)
  const activeCell = sheet.cells[addr]
  const editing = draft !== null

  const colW = (c: number) => sheet.colW[c] ?? DEF_W
  const colLeft = useMemo(() => { const a: number[] = [0]; for (let c = 0; c < sheet.cols; c++) a.push(a[c] + (sheet.colW[c] ?? DEF_W)); return a }, [sheet.colW, sheet.cols])

  const setSheet = (fn: (s: Sheet) => Sheet) => commit({ ...wb, sheets: wb.sheets.map((s, i) => (i === activeIdx ? fn(s) : s)) })
  const setCellValue = (cells: Record<string, Cell>, a: string, v: string) => {
    const prev = cells[a]
    if (v === '' && !prev?.s) delete cells[a]; else cells[a] = { ...prev, v }
  }

  function move(dr: number, dc: number, extend = false) {
    const next = { r: Math.max(0, Math.min(sheet.rows - 1, focus.r + dr)), c: Math.max(0, Math.min(sheet.cols - 1, focus.c + dc)) }
    setFocus(next); if (!extend) setAnchor(next)
    ensureVisible(next)
  }
  /** Scroll so a cell is on screen. Rows are virtualised, so this is computed rather than DOM-queried. */
  function ensureVisible(p: Pos) {
    const el = gridRef.current; if (!el) return
    const top = ROW_H + p.r * ROW_H, bottom = top + ROW_H
    if (top - ROW_H < el.scrollTop) el.scrollTop = top - ROW_H
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight
    const l = HDR_W + colLeft[p.c], r = HDR_W + colLeft[p.c + 1]
    if (l - HDR_W < el.scrollLeft) el.scrollLeft = l - HDR_W
    else if (r > el.scrollLeft + el.clientWidth) el.scrollLeft = r - el.clientWidth
  }
  function startEdit(initial?: string) { setDraft(initial ?? activeCell?.v ?? ''); setSugIdx(0); setSugOpen(true) }
  /** Finishes the edit; false when a formula was refused and the cell stays in edit mode. */
  function endEdit(save: boolean, dr = 0, dc = 0): boolean {
    if (draft === null) return true
    const text = save ? balanceFormula(draft) : draft
    const problem = save ? validateFormula(text) : null
    if (problem) { toast(problem, 'error'); return false } // like Excel: the cell stays in edit mode until the formula is fixed
    if (save && text !== (activeCell?.v ?? '')) setSheet(s => { const cells = { ...s.cells }; setCellValue(cells, addr, text); return { ...s, cells } })
    setDraft(null); setSugOpen(true); move(dr, dc); gridRef.current?.focus()
    return true
  }

  function forEachSel(fn: (a: string, r: number, c: number) => void) { for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) fn(toAddr(c, r), r, c) }
  function clearSel() { setSheet(s => { const cells = { ...s.cells }; forEachSel(a => { const p = cells[a]; if (p?.s) cells[a] = { v: '', s: p.s }; else delete cells[a] }); return { ...s, cells } }) }
  function applyStyle(patch: Partial<CellStyle>) {
    setSheet(s => { const cells = { ...s.cells }; forEachSel(a => { cells[a] = { v: cells[a]?.v ?? '', s: { ...cells[a]?.s, ...patch } } }); return { ...s, cells } })
  }
  const toggle = (k: 'b' | 'i' | 'u' | 'strike' | 'wrap' | 'border') => applyStyle({ [k]: !activeCell?.s?.[k] })
  const bumpDecimals = (d: 1 | -1) => {
    const st = activeCell?.s
    const v = ev.displayValue(activeIdx, addr)
    // start from what is shown now: the explicit setting, the format's default, or the digits the number has
    const shown = st?.dec ?? (st?.fmt === 'number' || st?.fmt === 'currency' ? 2 : st?.fmt === 'percent' ? 0 : typeof v === 'number' ? Math.min(10, (String(Number(v.toPrecision(12))).split('.')[1] ?? '').length) : 0)
    applyStyle({ dec: Math.max(0, Math.min(10, shown + d)) })
  }
  useEffect(() => { const used = new Set<string>(); for (const sh of wbRef.current.sheets) for (const c of Object.values(sh.cells)) if (c.s?.font) used.add(c.s.font); void preloadFonts(used) }, [file.id])

  // Fill handle: drag the square at the corner of the selection, release to fill.
  const fillToRef = useRef(fillTo); fillToRef.current = fillTo
  const sheetRef = useRef(sheet); sheetRef.current = sheet
  useEffect(() => {
    const up = () => {
      const src = fillRef.current, to = fillToRef.current
      fillRef.current = null; setFillTo(null)
      if (!src || !to) return
      const res = fillRange(sheetRef.current.cells, src, to)
      if (!res) return
      setSheet(s => ({ ...s, cells: res.cells }))
      setAnchor({ r: res.rect.r1, c: res.rect.c1 }); setFocus({ r: res.rect.r2, c: res.rect.c2 })
    }
    window.addEventListener('mouseup', up)
    return () => window.removeEventListener('mouseup', up)
  }, [])
  const fillPreview = (() => {
    if (!fillTo || !fillRef.current) return null
    const S = fillRef.current
    const dv = fillTo.r > S.r2 ? fillTo.r - S.r2 : fillTo.r < S.r1 ? fillTo.r - S.r1 : 0, dh = fillTo.c > S.c2 ? fillTo.c - S.c2 : fillTo.c < S.c1 ? fillTo.c - S.c1 : 0
    if (!dv && !dh) return null
    return Math.abs(dv) >= Math.abs(dh) ? { ...S, r1: dv < 0 ? S.r1 + dv : S.r1, r2: dv > 0 ? S.r2 + dv : S.r2 } : { ...S, c1: dh < 0 ? S.c1 + dh : S.c1, c2: dh > 0 ? S.c2 + dh : S.c2 }
  })()

  function fill(dir: 'down' | 'right') {
    setSheet(s => {
      const cells = { ...s.cells }
      if (dir === 'down') for (let c = rect.c1; c <= rect.c2; c++) { const src = s.cells[toAddr(c, rect.r1)]; for (let r = rect.r1 + 1; r <= rect.r2; r++) setFilled(cells, toAddr(c, r), src, 0, r - rect.r1) }
      else for (let r = rect.r1; r <= rect.r2; r++) { const src = s.cells[toAddr(rect.c1, r)]; for (let c = rect.c1 + 1; c <= rect.c2; c++) setFilled(cells, toAddr(c, r), src, c - rect.c1, 0) }
      return { ...s, cells }
    })
  }
  const setFilled = (cells: Record<string, Cell>, a: string, src: Cell | undefined, dc: number, dr: number) => {
    if (!src) { delete cells[a]; return }
    cells[a] = { ...src, v: src.v.startsWith('=') ? shiftFormula(src.v, dc, dr) : src.v }
  }
  function structure(kind: 'insRow' | 'delRow' | 'insCol' | 'delCol', where: 'before' | 'after' = 'before') {
    const n = kind.endsWith('Row') ? rect.r2 - rect.r1 + 1 : rect.c2 - rect.c1 + 1
    const startIdx = kind.endsWith('Row') ? rect.r1 : rect.c1
    const at = where === 'after' ? startIdx + n : startIdx
    const next = kind === 'insRow' ? insertRows(wb, activeIdx, at, n) : kind === 'delRow' ? deleteRows(wb, activeIdx, startIdx, n) : kind === 'insCol' ? insertCols(wb, activeIdx, at, n) : deleteCols(wb, activeIdx, startIdx, n)
    commit(next)
  }
  function showContext(e: React.MouseEvent, target: 'cell' | 'row' | 'col') {
    e.preventDefault()
    if (target === 'cell') { const p = cellFrom(e.target); if (p && (p.r < rect.r1 || p.r > rect.r2 || p.c < rect.c1 || p.c > rect.c2)) { setAnchor(p); setFocus(p) } }
    const items: MenuEntry[] = [
      { label: 'Cut', hint: '⌘X', onClick: () => doCopy(true) }, { label: 'Copy', hint: '⌘C', onClick: () => doCopy(false) },
      { label: 'Paste', hint: '⌘V', disabled: !clip, onClick: () => doPaste() }, 'sep',
      { label: 'Fill down', hint: '⌘D', onClick: () => fill('down') }, { label: 'Fill right', hint: '⌘R', onClick: () => fill('right') },
      { label: 'Clear contents', hint: 'Del', onClick: clearSel }, 'sep',
      { label: 'Insert row above', onClick: () => structure('insRow') }, { label: 'Insert row below', onClick: () => structure('insRow', 'after') }, { label: 'Delete row', danger: true, onClick: () => structure('delRow') },
      { label: 'Insert column left', onClick: () => structure('insCol') }, { label: 'Insert column right', onClick: () => structure('insCol', 'after') }, { label: 'Delete column', danger: true, onClick: () => structure('delCol') },
    ]
    setCtx({ x: e.clientX, y: e.clientY, items })
  }

  /* ---------- data tools ---------- */
  function doSort(asc: boolean) {
    const single = rect.r1 === rect.r2 && rect.c1 === rect.c2
    // a single cell sorts the contiguous block around it, as Excel does
    let r = { ...rect }
    if (single) {
      const has = (rr: number, cc: number) => !!sheet.cells[toAddr(cc, rr)]?.v
      let r1 = rect.r1, r2 = rect.r1, c1 = rect.c1, c2 = rect.c1
      let grew = true
      while (grew) {
        grew = false
        for (let cc = c1 - 1; cc <= c2 + 1; cc++) for (const rr of [r1 - 1, r2 + 1]) if (rr >= 0 && cc >= 0 && has(rr, cc)) { if (rr < r1) { r1 = rr; grew = true } else if (rr > r2) { r2 = rr; grew = true } }
        for (let rr = r1; rr <= r2; rr++) for (const cc of [c1 - 1, c2 + 1]) if (rr >= 0 && cc >= 0 && has(rr, cc)) { if (cc < c1) { c1 = cc; grew = true } else if (cc > c2) { c2 = cc; grew = true } }
      }
      r = { r1, r2, c1, c2 }
    }
    if (r.r1 === r.r2) { toast('Select a range with more than one row to sort', 'error'); return }
    // treat the first row as a header when it is text and the row below it holds numbers in that column
    const top = ev.displayValue(activeIdx, toAddr(Math.min(Math.max(focus.c, r.c1), r.c2), r.r1)), below = ev.displayValue(activeIdx, toAddr(Math.min(Math.max(focus.c, r.c1), r.c2), r.r1 + 1))
    const header = typeof top === 'string' && typeof below === 'number'
    commit(sortRange(wb, activeIdx, r, Math.min(Math.max(focus.c, r.c1), r.c2), asc, header))
    toast(`Sorted ${rangeToA1(r)} ${asc ? 'A to Z' : 'Z to A'}${header ? ' (header kept)' : ''}`)
  }
  function setFreeze(f: { rows: number; cols: number } | undefined) { setSheet(s => ({ ...s, freeze: f && (f.rows || f.cols) ? f : undefined })) }
  const addRule = (rule: CondRule) => setSheet(s => ({ ...s, cf: [...(s.cf ?? []), rule] }))
  const removeRule = (id: string) => setSheet(s => ({ ...s, cf: (s.cf ?? []).filter(x => x.id !== id) }))
  const addChart = (v: { type: Chart['type']; title: string; range: string }) => {
    // place the chart just right of its data so it never covers the cells it plots
    const rg = parseRange(v.range)
    const n = sheet.charts?.length ?? 0
    const c: Chart = { id: uid(), ...v, x: rg ? colLeft[Math.min(rg.c2 + 2, sheet.cols)] + n * 16 : 120, y: rg ? rg.r1 * ROW_H + n * 16 : 60, w: 460, h: 300 }
    setSheet(s => ({ ...s, charts: [...(s.charts ?? []), c] })); setSelChart(c.id)
  }
  const patchChart = (id: string, p: Partial<Chart>) => setSheet(s => ({ ...s, charts: (s.charts ?? []).map(c => (c.id === id ? { ...c, ...p } : c)) }))
  function startChartGesture(e: React.PointerEvent, c: Chart, mode: 'move' | 'size') {
    e.stopPropagation(); e.preventDefault()
    try { (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId) } catch { /* best effort */ }
    setSelChart(c.id); chartGesture.current = { id: c.id, mode, sx: e.clientX, sy: e.clientY, o: c, base: wbRef.current }
  }
  function moveChartGesture(e: React.PointerEvent) {
    const g = chartGesture.current; if (!g) return
    const dx = e.clientX - g.sx, dy = e.clientY - g.sy
    const next = g.mode === 'move' ? { x: Math.max(0, g.o.x + dx), y: Math.max(0, g.o.y + dy) } : { w: Math.max(200, g.o.w + dx), h: Math.max(140, g.o.h + dy) }
    hist.replace({ ...g.base, sheets: g.base.sheets.map((sh, i) => (i === activeIdx ? { ...sh, charts: (sh.charts ?? []).map(c => (c.id === g.id ? { ...c, ...next } : c)) } : sh)) })
  }
  function endChartGesture() { const g = chartGesture.current; chartGesture.current = null; if (g) { hist.record(g.base); setVersion(v => v + 1) } }

  /* ---------- find & replace ---------- */
  const findOpts = { matchCase: findCase, whole: findWhole, inFormulas: findFormulas }
  const hits = useMemo(() => (findOpen ? findCells(wb, activeIdx, findQ, findOpts) : []), [findOpen, wb, activeIdx, findQ, findCase, findWhole, findFormulas]) // eslint-disable-line react-hooks/exhaustive-deps
  function gotoHit(i: number) {
    if (!hits.length) return
    const n = (i + hits.length) % hits.length; setFindIdx(n)
    const p = parseAddr(hits[n].addr)!; setAnchor({ r: p.r, c: p.c }); setFocus({ r: p.r, c: p.c }); ensureVisible({ r: p.r, c: p.c })
  }
  function replaceHit(all: boolean) {
    if (!hits.length) { toast('Nothing to replace', 'error'); return }
    const targets = all ? hits : [hits[Math.min(findIdx, hits.length - 1)]]
    setSheet(s => {
      const cells = { ...s.cells }
      for (const h of targets) { const cur = cells[h.addr]; if (cur) cells[h.addr] = { ...cur, v: replaceInCell(cur.v, findQ, replQ, findOpts) } }
      return { ...s, cells }
    })
    toast(all ? `Replaced ${targets.length} ${targets.length === 1 ? 'cell' : 'cells'}` : 'Replaced')
  }
  const openFind = (replace: boolean) => { setFindOpen(true); setReplaceOpen(replace); setTimeout(() => document.getElementById('xl-find')?.focus(), 0) }
  // Find/Replace work from anywhere in the editor (formula bar, ribbon), not only while the grid is focused.
  const openFindRef = useRef(openFind); openFindRef.current = openFind
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey) || e.altKey) return
      const key = e.key.toLowerCase()
      // Cmd+H is "Hide" on macOS, so Replace is Ctrl+H there (as in Excel for Mac)
      if (key === 'f' && !e.shiftKey) { e.preventDefault(); openFindRef.current(false) }
      else if (key === 'h' && e.ctrlKey) { e.preventDefault(); openFindRef.current(true) }
    }
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [])

  function buildClip(): Clip {
    const cells: (Cell | undefined)[][] = []; const lines: string[] = []
    for (let r = rect.r1; r <= rect.r2; r++) {
      const row: (Cell | undefined)[] = []; const t: string[] = []
      for (let c = rect.c1; c <= rect.c2; c++) { const x = sheet.cells[toAddr(c, r)]; row.push(x); t.push(formatValue(ev.displayValue(activeIdx, toAddr(c, r)), 'general')) }
      cells.push(row); lines.push(t.join('\t'))
    }
    return { cells, origin: { r: rect.r1, c: rect.c1 }, text: lines.join('\n') }
  }
  function doCopy(cut: boolean, e?: React.ClipboardEvent) {
    const cl = buildClip(); setClip(cl)
    if (e) { e.preventDefault(); e.clipboardData.setData('text/plain', cl.text) } else void navigator.clipboard?.writeText(cl.text).catch(() => {})
    if (cut) clearSel()
  }
  function doPaste(text?: string) {
    const target = { r: rect.r1, c: rect.c1 }
    if (clip && (text === undefined || text === clip.text)) {
      const dr = target.r - clip.origin.r, dc = target.c - clip.origin.c
      setSheet(s => {
        const cells = { ...s.cells }
        clip.cells.forEach((row, i) => row.forEach((cell, j) => {
          const a = toAddr(target.c + j, target.r + i); if (target.c + j >= s.cols || target.r + i >= s.rows) return
          if (!cell) { delete cells[a]; return }
          cells[a] = { ...cell, v: cell.v.startsWith('=') ? shiftFormula(cell.v, dc, dr) : cell.v }
        }))
        return { ...s, cells }
      })
    } else if (text) {
      const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'))
      setSheet(s => {
        const cells = { ...s.cells }
        rows.forEach((row, i) => row.forEach((v, j) => { if (target.c + j < s.cols && target.r + i < s.rows) setCellValue(cells, toAddr(target.c + j, target.r + i), v) }))
        return { ...s, cells }
      })
    }
  }

  function onKeyDown(e: React.KeyboardEvent) {
    if (editing) return
    const mod = e.metaKey || e.ctrlKey
    if (mod) {
      const k = e.key.toLowerCase()
      if (k === 'z') { e.preventDefault(); if (e.shiftKey) hist.redo(); else hist.undo(); setVersion(v => v + 1); return }
      if (k === 'y') { e.preventDefault(); hist.redo(); setVersion(v => v + 1); return }
      if (k === 'b') { e.preventDefault(); toggle('b'); return }
      if (k === 'i') { e.preventDefault(); toggle('i'); return }
      if (k === 's') { e.preventDefault(); void save().then(() => toast('Saved')); return }
      if (k === 'u') { e.preventDefault(); toggle('u'); return }
      if (k === 'd') { e.preventDefault(); fill('down'); return }
      if (k === 'r') { e.preventDefault(); fill('right'); return }
      if (k === 'a') { e.preventDefault(); setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }); return }
      return
    }
    switch (e.key) {
      case 'ArrowDown': e.preventDefault(); move(1, 0, e.shiftKey); return
      case 'ArrowUp': e.preventDefault(); move(-1, 0, e.shiftKey); return
      case 'ArrowLeft': e.preventDefault(); move(0, -1, e.shiftKey); return
      case 'ArrowRight': e.preventDefault(); move(0, 1, e.shiftKey); return
      case 'Tab': e.preventDefault(); move(0, e.shiftKey ? -1 : 1); return
      case 'Enter': case 'F2': e.preventDefault(); startEdit(); return
      case 'Delete': case 'Backspace': e.preventDefault(); clearSel(); return
      default: if (e.key.length === 1 && !e.altKey) { e.preventDefault(); startEdit(e.key) }
    }
  }

  // mouse selection (delegated)
  const cellFrom = (t: EventTarget | null): Pos | null => { const td = (t as HTMLElement)?.closest?.('td[data-r]') as HTMLElement | null; return td ? { r: Number(td.dataset.r), c: Number(td.dataset.c) } : null }
  useEffect(() => {
    const el = gridRef.current; if (!el) return
    const ro = new ResizeObserver(() => setScroll(p => (p.h === el.clientHeight ? p : { ...p, h: el.clientHeight })))
    ro.observe(el); return () => ro.disconnect()
  }, [])
  // Window listeners are attached once; they read the latest state through a ref so nothing is re-bound per render.
  const live = useRef({ wb, activeIdx, hist, setVersion })
  live.current = { wb, activeIdx, hist, setVersion }
  useEffect(() => {
    let base: Workbook | null = null
    const mv = (e: MouseEvent) => {
      const rz = resize.current; if (!rz) return
      const L = live.current; base ??= L.wb
      const w = Math.max(30, rz.w + e.clientX - rz.x)
      L.hist.replace({ ...base, sheets: base.sheets.map((s, i) => (i === L.activeIdx ? { ...s, colW: { ...s.colW, [rz.c]: w } } : s)) })
    }
    const up = () => {
      dragging.current = false
      if (resize.current && base) { live.current.hist.record(base); live.current.setVersion(v => v + 1) } // one undo step per resize
      resize.current = null; base = null
    }
    window.addEventListener('mouseup', up); window.addEventListener('mousemove', mv)
    return () => { window.removeEventListener('mouseup', up); window.removeEventListener('mousemove', mv) }
  }, [])

  // sheet operations
  const stats = useMemo(() => {
    const nums: number[] = []; let count = 0
    for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) { const v = ev.displayValue(activeIdx, toAddr(c, r)); if (v !== null && v !== '') count++; if (typeof v === 'number') nums.push(v) }
    const sum = nums.reduce((a, b) => a + b, 0)
    return { count, sum, avg: nums.length ? sum / nums.length : 0, n: nums.length }
  }, [wb, ev, activeIdx, rect.r1, rect.r2, rect.c1, rect.c2])

  function addSheet() { let n = wb.sheets.length + 1; while (wb.sheets.some(s => s.name === `Sheet${n}`)) n++; commit({ sheets: [...wb.sheets, newSheet(`Sheet${n}`)], active: wb.sheets.length }); resetSel() }
  function deleteSheet(i: number) {
    if (wb.sheets.length < 2) { toast('A workbook needs at least one sheet', 'error'); return }
    commit({ sheets: wb.sheets.filter((_, j) => j !== i), active: Math.max(0, Math.min(wb.active, wb.sheets.length - 2)) }); resetSel()
  }
  function renameSheet(i: number, name: string) {
    const nm = name.trim(); const old = wb.sheets[i].name
    if (!nm || nm === old) return
    if (wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === nm.toLowerCase())) { toast('A sheet with that name already exists', 'error'); return }
    const esc = old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const re = new RegExp(`(?:'${esc}'|\\b${esc})!`, 'g')
    const q = /^[A-Za-z_][\w.]*$/.test(nm) ? nm : `'${nm}'`
    commit({ ...wb, sheets: wb.sheets.map((s, j) => ({ ...s, name: j === i ? nm : s.name, cells: Object.fromEntries(Object.entries(s.cells).map(([a, c]) => [a, c.v.startsWith('=') ? { ...c, v: c.v.replace(re, `${q}!`) } : c])) })) })
  }
  function resetSel() { setAnchor({ r: 0, c: 0 }); setFocus({ r: 0, c: 0 }); setDraft(null) }
  function selectSheet(i: number) { if (i === wb.active) return; hist.replace({ ...wb, active: i }); setVersion(v => v + 1); resetSel() }

  function insertFn(fn: string) {
    const multi = rect.r1 !== rect.r2 || rect.c1 !== rect.c2
    if (!multi) { startEdit(`=${fn}(`); return }
    const rng = `${toAddr(rect.c1, rect.r1)}:${toAddr(rect.c2, rect.r2)}`
    const tr = Math.min(rect.r2 + 1, sheet.rows - 1)
    setSheet(s => ({ ...s, cells: { ...s.cells, [toAddr(rect.c1, tr)]: { ...s.cells[toAddr(rect.c1, tr)], v: `=${fn}(${rng})` } } }))
    setAnchor({ r: tr, c: rect.c1 }); setFocus({ r: tr, c: rect.c1 })
  }

  async function importFile(f: File | undefined) {
    if (!f) return
    try {
      const io = await import('./io')
      const next = /\.csv$/i.test(f.name) ? io.workbookFromCsv(await f.text()) : io.workbookFromXlsx(await f.arrayBuffer())
      commit(next); resetSel(); toast(`Imported ${f.name}`)
    } catch (e) { toast(`Import failed: ${(e as Error).message}`, 'error') }
  }
  const exp = async (k: 'xlsx' | 'csv') => {
    try {
      const io = await import('./io')
      if (k === 'xlsx') download(io.exportXlsx(wb), `${safeName(file.name)}.xlsx`); else download(io.exportCsv(wb), `${safeName(file.name)}.csv`, 'text/csv;charset=utf-8')
      toast(`Exported ${k.toUpperCase()}`)
    } catch (e) { toast(`Export failed: ${(e as Error).message}`, 'error') }
  }
  const fileMenu: MenuItem[] = [
    { label: 'New spreadsheet', onClick: async () => { const n = await createFile('xlsx'); nav(`/excel/${n.id}`) } },
    { label: 'Import CSV / XLSX…', onClick: () => importRef.current?.click() },
    'sep',
    { label: 'Save now', onClick: () => { void save().then(() => toast('Saved')) } },
    { label: 'Export as XLSX', onClick: () => exp('xlsx') },
    { label: 'Export as CSV (current sheet)', onClick: () => exp('csv') },
  ]

  function submitNameBox(v: string) {
    const [a, b] = v.toUpperCase().split(':'); const p = parseAddr(a); const q = b ? parseAddr(b) : p
    setNameBox(null); if (!p || !q) { toast('Enter a cell like B3 or a range like A1:C5', 'error'); return }
    setAnchor({ r: Math.min(p.r, sheet.rows - 1), c: Math.min(p.c, sheet.cols - 1) }); setFocus({ r: Math.min(q.r, sheet.rows - 1), c: Math.min(q.c, sheet.cols - 1) }); gridRef.current?.focus()
  }

  /** Function-name autocomplete for the token being typed at the end of a formula. */
  const suggestions = useMemo(() => {
    if (draft === null || !draft.startsWith('=') || !sugOpen) return []
    const m = /([A-Za-z][A-Za-z.\d]*)$/.exec(draft)
    if (!m) return []
    const q = m[1].toUpperCase()
    return FN_NAMES.filter(n => n.startsWith(q)).slice(0, 8)
  }, [draft, sugOpen])
  const acceptSuggestion = (name: string) => {
    setDraft(d => (d ?? '').replace(/([A-Za-z][A-Za-z.\d]*)$/, name + '('))
    setSugOpen(false)
  }
  /** Keys the suggestion list consumes while it is open. Returns true when handled. */
  function suggestionKey(e: React.KeyboardEvent): boolean {
    if (!suggestions.length) return false
    if (e.key === 'ArrowDown') { e.preventDefault(); setSugIdx(i => (i + 1) % suggestions.length); return true }
    if (e.key === 'ArrowUp') { e.preventDefault(); setSugIdx(i => (i - 1 + suggestions.length) % suggestions.length); return true }
    if (e.key === 'Tab' || e.key === 'Enter') { e.preventDefault(); acceptSuggestion(suggestions[Math.min(sugIdx, suggestions.length - 1)]); return true }
    if (e.key === 'Escape') { e.preventDefault(); setSugOpen(false); return true }
    return false
  }
  const SuggestList = suggestions.length > 0 ? (
    <div role="listbox" aria-label="Function suggestions" data-testid="fn-suggest"
      className="absolute z-[60] mt-0.5 min-w-[280px] max-w-[380px] bg-surface border border-line rounded-lg shadow-soft overflow-hidden">
      {suggestions.map((n, i) => (
        <button key={n} role="option" aria-selected={i === Math.min(sugIdx, suggestions.length - 1)} onMouseDown={e => { e.preventDefault(); acceptSuggestion(n) }}
          className={`w-full flex items-baseline gap-2 px-3 h-7 text-left text-[12.5px] ${i === Math.min(sugIdx, suggestions.length - 1) ? 'bg-accent-soft' : 'hover:bg-panel2'}`}>
          <span className="font-semibold text-accent">{n}</span>
          <span className="text-muted truncate">{FUNCTION_HINTS[n] ?? ''}</span>
        </button>
      ))}
    </div>
  ) : null

  const fmt = activeCell?.s?.fmt ?? 'general'
  const left = HDR_W + colLeft[rect.c1], top = ROW_H + rect.r1 * ROW_H
  const width = colLeft[rect.c2 + 1] - colLeft[rect.c1], height = (rect.r2 - rect.r1 + 1) * ROW_H
  const multi = rect.r1 !== rect.r2 || rect.c1 !== rect.c2
  const first = Math.max(0, Math.floor(scroll.top / ROW_H) - OVERSCAN)
  const last = Math.min(sheet.rows - 1, Math.ceil((scroll.top + scroll.h) / ROW_H) + OVERSCAN)
  const rows = useMemo(() => Array.from({ length: Math.max(0, last - first + 1) }, (_, i) => first + i), [first, last])
  const cols = useMemo(() => Array.from({ length: sheet.cols }, (_, c) => c), [sheet.cols])
  const sumDisplay = (n: number) => Number(n.toPrecision(12)).toLocaleString()
  const cf = useMemo(() => (sheet.cf?.length ? cfContext(ev, activeIdx) : null), [ev, activeIdx, sheet.cf, wb]) // eslint-disable-line react-hooks/exhaustive-deps
  const frozenRows = sheet.freeze?.rows ?? 0, frozenCols = sheet.freeze?.cols ?? 0
  // frozen rows are always rendered (sticky), even when scrolled far past them
  const renderRows = useMemo(() => {
    const set = new Set<number>(); for (let r = 0; r < frozenRows; r++) set.add(r)
    for (const r of rows) set.add(r)
    return [...set].sort((a, b) => a - b)
  }, [rows, frozenRows])
  const ribbonTabs = ['Home', 'Insert', 'Data', 'Formulas', 'View']

  return (
    <div className="h-full flex flex-col min-h-0">
      <EditorHeader file={file} color="var(--color-excel)" appName="Excel" save={saveState} fileMenu={fileMenu} />
      <Ribbon tabs={ribbonTabs} tab={tab} onTab={setTab}>
        {tab === 'Home' && (<>
          <RibbonGroup label="Undo">
            <IconButton label="Undo (⌘Z)" disabled={!hist.canUndo} onClick={() => { hist.undo(); setVersion(v => v + 1) }}><Undo2 size={16} /></IconButton>
            <IconButton label="Redo (⇧⌘Z)" disabled={!hist.canRedo} onClick={() => { hist.redo(); setVersion(v => v + 1) }}><Redo2 size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Clipboard">
            <IconButton label="Cut" onClick={() => doCopy(true)}><Scissors size={16} /></IconButton>
            <IconButton label="Copy" onClick={() => doCopy(false)}><Copy size={16} /></IconButton>
            <IconButton label="Paste" disabled={!clip} onClick={() => doPaste()}><ClipboardPaste size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Font">
            <FontPicker value={activeCell?.s?.font ?? ''} onPick={name => applyStyle({ font: name || undefined })} width={140} />
            <SizeBox value={activeCell?.s?.size ?? 13} min={6} max={72} onPick={n => applyStyle({ size: n })} />
            <IconButton label="Bold (⌘B)" active={!!activeCell?.s?.b} onClick={() => toggle('b')}><Bold size={16} /></IconButton>
            <IconButton label="Italic (⌘I)" active={!!activeCell?.s?.i} onClick={() => toggle('i')}><Italic size={16} /></IconButton>
            <IconButton label="Underline (⌘U)" active={!!activeCell?.s?.u} onClick={() => toggle('u')}><UnderlineIcon size={16} /></IconButton>
            <IconButton label="Strikethrough" active={!!activeCell?.s?.strike} onClick={() => toggle('strike')}><Strikethrough size={16} /></IconButton>
            <ColorMenu label="Font colour" icon={<Type size={16} />} allowNone onPick={c => applyStyle({ color: c ?? undefined })} />
            <ColorMenu label="Fill colour" icon={<PaintBucket size={16} />} allowNone onPick={c => applyStyle({ bg: c ?? undefined })} />
            <IconButton label="Borders" active={!!activeCell?.s?.border} onClick={() => toggle('border')}><Grid2x2 size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Alignment">
            <IconButton label="Align left" active={activeCell?.s?.align === 'left'} onClick={() => applyStyle({ align: 'left' })}><AlignLeft size={16} /></IconButton>
            <IconButton label="Centre" active={activeCell?.s?.align === 'center'} onClick={() => applyStyle({ align: 'center' })}><AlignCenter size={16} /></IconButton>
            <IconButton label="Align right" active={activeCell?.s?.align === 'right'} onClick={() => applyStyle({ align: 'right' })}><AlignRight size={16} /></IconButton>
            <IconButton label="Wrap text" active={!!activeCell?.s?.wrap} onClick={() => toggle('wrap')}><WrapText size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Number">
            <select aria-label="Number format" value={fmt} onChange={e => applyStyle({ fmt: e.target.value as CellStyle['fmt'], dec: undefined })} className="h-8 rounded-md border border-line bg-surface px-1.5 text-[13px]">
              <option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percent</option><option value="date">Date</option><option value="time">Time</option><option value="scientific">Scientific</option>
            </select>
            <IconButton label="Decrease decimal places" onClick={() => bumpDecimals(-1)}><span className="text-[11px] font-semibold">.0←</span></IconButton>
            <IconButton label="Increase decimal places" onClick={() => bumpDecimals(1)}><span className="text-[11px] font-semibold">.00→</span></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Styles">
            <Menu title="Conditional formatting" trigger={<><Palette size={15} />Conditional formatting<ChevronDown size={12} /></>} items={[
              ...CF_KINDS.map(k => ({ label: k.label, onClick: () => setDialog({ kind: 'cf', initial: k.id }) })),
              'sep' as const,
              ...(sheet.cf ?? []).map(r => ({ label: `Remove: ${CF_KINDS.find(k => k.id === r.kind)?.label} on ${r.range}${r.a ? ` (${r.a}${r.b ? `–${r.b}` : ''})` : ''}`, danger: true, onClick: () => removeRule(r.id) })),
              { label: `Clear all rules${sheet.cf?.length ? ` (${sheet.cf.length})` : ''}`, danger: true, disabled: !sheet.cf?.length, onClick: () => setSheet(s => ({ ...s, cf: undefined })) },
            ]} />
          </RibbonGroup>
          <RibbonGroup label="Editing">
            <Button variant="ghost" onClick={() => openFind(false)}><Search size={15} />Find</Button>
            <Button variant="ghost" onClick={() => openFind(true)}><Replace size={15} />Replace</Button>
            <IconButton label="Sort A to Z" onClick={() => doSort(true)}><ArrowDownAZ size={16} /></IconButton>
            <IconButton label="Sort Z to A" onClick={() => doSort(false)}><ArrowDownZA size={16} /></IconButton>
          </RibbonGroup>
        </>)}

        {tab === 'Insert' && (<>
          <RibbonGroup label="Charts">
            <Button variant="ghost" onClick={() => setDialog({ kind: 'chart' })}><BarChart3 size={16} />Chart</Button>
            <span className="text-[12px] text-muted leading-[34px] px-1">{sheet.charts?.length ?? 0} on this sheet</span>
          </RibbonGroup>
          <RibbonGroup label="Cells">
            <Menu trigger={<>Insert<ChevronDown size={12} /></>} items={[{ label: 'Insert row above', onClick: () => structure('insRow') }, { label: 'Insert column left', onClick: () => structure('insCol') }]} />
            <Menu trigger={<>Delete<ChevronDown size={12} /></>} items={[{ label: 'Delete row', danger: true, onClick: () => structure('delRow') }, { label: 'Delete column', danger: true, onClick: () => structure('delCol') }]} />
          </RibbonGroup>
        </>)}

        {tab === 'Data' && (<>
          <RibbonGroup label="Sort">
            <Button variant="ghost" onClick={() => doSort(true)}><ArrowDownAZ size={16} />A → Z</Button>
            <Button variant="ghost" onClick={() => doSort(false)}><ArrowDownZA size={16} />Z → A</Button>
          </RibbonGroup>
          <RibbonGroup label="Window">
            <Menu title="Freeze panes" trigger={<><Snowflake size={15} />Freeze panes{(frozenRows || frozenCols) ? ' ✓' : ''}<ChevronDown size={12} /></>} items={[
              { label: 'Freeze top row', onClick: () => setFreeze({ rows: 1, cols: frozenCols }) },
              { label: 'Freeze first column', onClick: () => setFreeze({ rows: frozenRows, cols: 1 }) },
              { label: 'Freeze panes at selection', onClick: () => setFreeze({ rows: rect.r1, cols: rect.c1 }) },
              'sep',
              { label: 'Unfreeze panes', disabled: !(frozenRows || frozenCols), onClick: () => setFreeze(undefined) },
            ]} />
          </RibbonGroup>
          <RibbonGroup label="Find">
            <Button variant="ghost" onClick={() => openFind(false)}><Search size={15} />Find</Button>
            <Button variant="ghost" onClick={() => openFind(true)}><Replace size={15} />Replace</Button>
          </RibbonGroup>
        </>)}

        {tab === 'Formulas' && (<>
          {Object.entries(FUNCTION_GROUPS).map(([label, fns]) => (
            <RibbonGroup key={label} label={label}>
              <Menu title={`${label} functions`} trigger={<><Sigma size={15} />{label}<ChevronDown size={12} /></>} items={fns.map(fn => ({ label: `${fn}(…)`, onClick: () => insertFn(fn) }))} />
            </RibbonGroup>
          ))}
          <RibbonGroup label="Fill">
            <IconButton label="Fill down (⌘D)" onClick={() => fill('down')}><ArrowDownToLine size={16} /></IconButton>
            <IconButton label="Fill right (⌘R)" onClick={() => fill('right')}><ArrowRightToLine size={16} /></IconButton>
          </RibbonGroup>
        </>)}

        {tab === 'View' && (<>
          <RibbonGroup label="Freeze">
            <Button variant="ghost" onClick={() => setFreeze({ rows: 1, cols: frozenCols })}>Top row</Button>
            <Button variant="ghost" onClick={() => setFreeze({ rows: frozenRows, cols: 1 })}>First column</Button>
            <Button variant="ghost" disabled={!(frozenRows || frozenCols)} onClick={() => setFreeze(undefined)}>Unfreeze</Button>
          </RibbonGroup>
          <RibbonGroup label="Column width">
            <Button variant="ghost" onClick={() => setSheet(s => { const colW = { ...s.colW }; for (let c = rect.c1; c <= rect.c2; c++) { let w = 40; for (let r = 0; r < Math.min(s.rows, 2000); r++) { const v = ev.displayValue(activeIdx, toAddr(c, r)); if (v !== null) w = Math.max(w, formatValue(v, s.cells[toAddr(c, r)]?.s?.fmt, s.cells[toAddr(c, r)]?.s?.dec).length * 8 + 22) } colW[c] = Math.min(w, 420) } return { ...s, colW } })}><MoveHorizontal size={15} />AutoFit</Button>
          </RibbonGroup>
          <RibbonGroup label="Rules">
            <span className="text-[12px] text-muted leading-[34px] px-1">{sheet.cf?.length ?? 0} conditional format{sheet.cf?.length === 1 ? '' : 's'}</span>
          </RibbonGroup>
        </>)}
      </Ribbon>

      {findOpen && (
        <div className="bg-panel border-b border-line px-4 py-1.5 flex flex-col gap-1.5 shrink-0" data-testid="xl-find-bar">
          <div className="flex items-center gap-2">
            <Search size={14} className="text-muted shrink-0" />
            <input id="xl-find" aria-label="Find in sheet" value={findQ} onChange={e => { setFindQ(e.target.value); setFindIdx(0) }}
              onKeyDown={e => { if (e.key === 'Enter') gotoHit(findIdx + (e.shiftKey ? -1 : 1)); if (e.key === 'Escape') { setFindOpen(false); gridRef.current?.focus() } }}
              placeholder="Find" className="h-7 w-56 px-2 rounded border border-line bg-surface" />
            <span className="text-[12px] text-muted w-24" data-testid="xl-find-count">{findQ ? (hits.length ? `${Math.min(findIdx, hits.length - 1) + 1} of ${hits.length}` : 'No results') : ''}</span>
            <IconButton label="Previous match" onClick={() => gotoHit(findIdx - 1)}><ChevronUp size={15} /></IconButton>
            <IconButton label="Next match" onClick={() => gotoHit(findIdx + 1)}><ChevronDown size={15} /></IconButton>
            <IconButton label="Match case" active={findCase} onClick={() => setFindCase(v => !v)}><span className="text-[12px] font-semibold">Aa</span></IconButton>
            <IconButton label="Match entire cell" active={findWhole} onClick={() => setFindWhole(v => !v)}><span className="text-[12px] font-semibold">ab|</span></IconButton>
            <IconButton label="Search formulas" active={findFormulas} onClick={() => setFindFormulas(v => !v)}><span className="text-[12px] font-semibold">fx</span></IconButton>
            <IconButton label={replaceOpen ? 'Hide replace' : 'Show replace'} active={replaceOpen} onClick={() => setReplaceOpen(v => !v)}><Replace size={15} /></IconButton>
            <div className="flex-1" />
            <IconButton label="Close find" onClick={() => { setFindOpen(false); gridRef.current?.focus() }}><X size={15} /></IconButton>
          </div>
          {replaceOpen && (
            <div className="flex items-center gap-2 pl-6">
              <input aria-label="Replace with" value={replQ} onChange={e => setReplQ(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') replaceHit(false) }} placeholder="Replace with" className="h-7 w-56 px-2 rounded border border-line bg-surface" />
              <Button disabled={!hits.length} onClick={() => replaceHit(false)}>Replace</Button>
              <Button disabled={!hits.length} onClick={() => replaceHit(true)}>Replace all</Button>
            </div>
          )}
        </div>
      )}

      <div className="flex items-center gap-2 h-9 px-3 bg-surface border-b border-line shrink-0">
        <input aria-label="Name box" value={nameBox ?? (multi ? `${toAddr(rect.c1, rect.r1)}:${toAddr(rect.c2, rect.r2)}` : addr)}
          onFocus={e => { setNameBox(e.target.value); e.target.select() }} onChange={e => setNameBox(e.target.value)} onBlur={() => setNameBox(null)}
          onKeyDown={e => { if (e.key === 'Enter') submitNameBox((e.target as HTMLInputElement).value); if (e.key === 'Escape') { setNameBox(null); gridRef.current?.focus() } }}
          className="w-24 h-7 px-2 text-center rounded border border-line bg-panel text-[13px]" />
        <span className="text-muted italic font-serif px-1">fx</span>
        <div className="flex-1 relative">
          <input aria-label="Formula bar" value={draft ?? activeCell?.v ?? ''} onFocus={() => { setBarFocused(true); if (draft === null) setDraft(activeCell?.v ?? '') }} onBlur={() => setBarFocused(false)}
            onChange={e => { setDraft(e.target.value); setSugIdx(0); setSugOpen(true) }}
            onKeyDown={e => { if (suggestionKey(e)) return; if (e.key === 'Enter') { e.preventDefault(); endEdit(true, 1, 0) } if (e.key === 'Escape') endEdit(false) }}
            className="w-full h-7 px-2 rounded border border-line bg-surface text-[13px] font-mono" />
          {barFocused && SuggestList}
        </div>
      </div>

      <div ref={gridRef} tabIndex={0} onKeyDown={onKeyDown} onCopy={e => { if (!editing) doCopy(false, e) }} onCut={e => { if (!editing) doCopy(true, e) }}
        onPaste={e => { if (!editing) { e.preventDefault(); doPaste(e.clipboardData.getData('text/plain')) } }}
        className="flex-1 overflow-auto relative bg-surface outline-none" data-testid="grid"
        onScroll={e => { const t = e.currentTarget; setScroll(p => (Math.abs(p.top - t.scrollTop) < ROW_H / 2 && p.h === t.clientHeight ? p : { top: t.scrollTop, h: t.clientHeight })) }}
        onContextMenu={e => showContext(e, 'cell')}
        onMouseDown={e => {
          if (editing && (e.target as HTMLElement).tagName === 'INPUT') return
          const p = cellFrom(e.target); if (!p) return
          if (editing && !endEdit(true)) { e.preventDefault(); return }
          e.preventDefault() // stop the browser's own text selection from painting over neighbouring cells while dragging
          gridRef.current?.focus(); dragging.current = true
          if (e.shiftKey) setFocus(p); else { setAnchor(p); setFocus(p) }
        }}
        onMouseOver={e => { if (fillRef.current) { const p = cellFrom(e.target); if (p) setFillTo(p); return } if (!dragging.current) return; const p = cellFrom(e.target); if (p) setFocus(p) }}
        onDoubleClick={e => { if (cellFrom(e.target)) startEdit() }}>
        <div style={{ width: HDR_W + colLeft[sheet.cols], position: 'relative' }}>
          <table className="select-none" style={{ tableLayout: 'fixed', width: HDR_W + colLeft[sheet.cols], borderCollapse: 'separate', borderSpacing: 0 }}>
            <colgroup><col style={{ width: HDR_W }} />{cols.map(c => <col key={c} style={{ width: colW(c) }} />)}</colgroup>
            <thead className="sticky top-0 z-10">
              <tr>
                <th className="sticky left-0 z-20 bg-panel border-r border-b border-line" style={{ height: ROW_H }} onClick={() => { setAnchor({ r: 0, c: 0 }); setFocus({ r: sheet.rows - 1, c: sheet.cols - 1 }) }} />
                {cols.map(c => (
                  <th key={c} className={`relative bg-panel border-r border-b border-line text-[12px] font-medium select-none ${c >= rect.c1 && c <= rect.c2 ? 'bg-accent-soft text-accent' : 'text-muted'}`} style={{ height: ROW_H, ...(c < frozenCols ? { position: 'sticky', left: HDR_W + colLeft[c], zIndex: 15 } : {}) }}
                    onMouseDown={e => { if ((e.target as HTMLElement).dataset.grip) return; gridRef.current?.focus(); setAnchor({ r: 0, c }); setFocus({ r: sheet.rows - 1, c }) }}>
                    {colName(c)}
                    <span data-grip="1" onMouseDown={e => { e.preventDefault(); e.stopPropagation(); resize.current = { c, x: e.clientX, w: colW(c) } }} className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize" />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {first > 0 && <tr aria-hidden style={{ height: first * ROW_H }}><td colSpan={sheet.cols + 1} /></tr>}
              {renderRows.map(r => (
                <tr key={r}>
                  <th className={`sticky left-0 z-[5] border-r border-b border-line text-[12px] font-medium select-none ${r >= rect.r1 && r <= rect.r2 ? 'bg-accent-soft text-accent' : 'bg-panel text-muted'}`}
                    style={r < frozenRows ? { top: ROW_H * (r + 1), zIndex: 6 } : undefined}
                    onMouseDown={() => { gridRef.current?.focus(); setAnchor({ r, c: 0 }); setFocus({ r, c: sheet.cols - 1 }) }}>{r + 1}</th>
                  {cols.map(c => {
                    const a = toAddr(c, r); const cell = sheet.cells[a]
                    const v: Scalar = cell || spillTick >= 0 ? ev.displayValue(activeIdx, a) : null
                    const stTop = r < frozenRows ? ROW_H * (r + 1) : undefined
                    const stLeft = c < frozenCols ? HDR_W + colLeft[c] : undefined
                    const fx = cf ? cfFor(sheet.cf, a, v, cf) : null
                    if (!cell && v === null && !fx && stTop === undefined && stLeft === undefined) return <Td key={c} r={r} c={c} display="" isErr={false} isNum={false} />
                    return <Td key={c} r={r} c={c} display={formatValue(v, cell?.s?.fmt, cell?.s?.dec)} style={cell?.s} isErr={v instanceof FormulaError} isNum={typeof v === 'number'} spilled={!cell && v !== null} cf={fx} stickyTop={stTop} stickyLeft={stLeft} />
                  })}
                </tr>
              ))}
              {last < sheet.rows - 1 && <tr aria-hidden style={{ height: (sheet.rows - 1 - last) * ROW_H }}><td colSpan={sheet.cols + 1} /></tr>}
            </tbody>
          </table>
          {(sheet.charts ?? []).map(ch => (
            <div key={ch.id} data-testid="chart" aria-label={ch.title || `${ch.type} chart`}
              onPointerMove={moveChartGesture} onPointerUp={endChartGesture} onPointerCancel={endChartGesture}
              onMouseDown={e => e.stopPropagation()}
              style={{ position: 'absolute', left: HDR_W + ch.x, top: ROW_H + ch.y, width: ch.w, height: ch.h, zIndex: 7 }}
              className={cx('bg-surface rounded-lg shadow-soft border', selChart === ch.id ? 'border-accent' : 'border-line')}>
              <div onPointerDown={e => startChartGesture(e, ch, 'move')} className="h-6 cursor-move flex items-center justify-end gap-0.5 px-1 rounded-t-lg bg-panel border-b border-line">
                <button aria-label="Edit chart" onPointerDown={e => e.stopPropagation()} onClick={() => setDialog({ kind: 'chart', edit: ch })} className="w-5 h-5 rounded hover:bg-panel2 flex items-center justify-center"><Pencil size={11} /></button>
                <button aria-label="Delete chart" onPointerDown={e => e.stopPropagation()} onClick={() => { setSheet(s => ({ ...s, charts: (s.charts ?? []).filter(c => c.id !== ch.id) })); setSelChart(null) }} className="w-5 h-5 rounded hover:bg-panel2 text-danger flex items-center justify-center"><Trash2 size={11} /></button>
              </div>
              <ChartView type={ch.type} data={chartData(ev, activeIdx, ch)} title={ch.title} width={ch.w - 2} height={ch.h - 26} />
              <div aria-label="Resize chart" onPointerDown={e => startChartGesture(e, ch, 'size')} className="absolute right-0 bottom-0 w-3.5 h-3.5 cursor-nwse-resize" style={{ background: 'linear-gradient(135deg, transparent 50%, var(--muted) 50%)' }} />
            </div>
          ))}
          <div aria-hidden className="absolute pointer-events-none border-2 border-accent z-[6]" style={{ left, top, width, height, background: multi ? 'color-mix(in srgb, var(--accent) 10%, transparent)' : undefined }} />
          {!editing && (
            <div data-testid="fill-handle" aria-label="Fill handle" title="Drag to fill" onMouseDown={e => { e.preventDefault(); e.stopPropagation(); fillRef.current = { ...rect }; setFillTo({ r: rect.r2, c: rect.c2 }) }}
              className="absolute z-[9] w-[9px] h-[9px] bg-accent border border-surface cursor-crosshair" style={{ left: left + width - 5, top: top + height - 5 }} />
          )}
          {fillPreview && <div aria-hidden data-testid="fill-preview" className="absolute pointer-events-none border-2 border-dashed border-accent z-[6]" style={{ left: HDR_W + colLeft[fillPreview.c1], top: ROW_H + fillPreview.r1 * ROW_H, width: colLeft[fillPreview.c2 + 1] - colLeft[fillPreview.c1], height: (fillPreview.r2 - fillPreview.r1 + 1) * ROW_H }} />}
          {editing && !barFocused && (
            <div className="absolute z-[8]" style={{ left: HDR_W + colLeft[focus.c], top: ROW_H + focus.r * ROW_H }}>
              <input ref={el => { if (el && document.activeElement !== el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length) } }} aria-label="Cell editor" value={draft ?? ''}
                onChange={e => { setDraft(e.target.value); setSugIdx(0); setSugOpen(true) }}
                onKeyDown={e => { e.stopPropagation(); if (suggestionKey(e)) return; if (e.key === 'Enter') { e.preventDefault(); endEdit(true, e.shiftKey ? -1 : 1, 0) } else if (e.key === 'Tab') { e.preventDefault(); endEdit(true, 0, e.shiftKey ? -1 : 1) } else if (e.key === 'Escape') endEdit(false) }}
                className="px-1.5 text-[13px] bg-surface border-2 border-accent outline-none block select-text"
                style={{ minWidth: colW(focus.c), height: ROW_H }} />
              {SuggestList}
            </div>
          )}
        </div>
      </div>

      <div className="h-9 shrink-0 bg-panel border-t border-line flex items-center gap-1 px-2 overflow-x-auto">
        <IconButton label="Add worksheet" onClick={addSheet}><Plus size={16} /></IconButton>
        {wb.sheets.map((s, i) => (
          <div key={s.id} className={`flex items-center rounded-t-md border border-b-0 ${i === activeIdx ? 'bg-surface border-line font-semibold text-excel' : 'border-transparent text-muted hover:bg-panel2'}`}>
            <button onClick={() => selectSheet(i)} onDoubleClick={() => setRenaming(i)} className="px-3 h-7 text-[13px]" aria-label={`Sheet ${s.name}`}>{s.name}</button>
            {i === activeIdx && <Menu title="Sheet options" trigger={<ChevronDown size={13} />} items={[
              { label: 'Rename', onClick: () => setRenaming(i) }, { label: 'Delete', danger: true, disabled: wb.sheets.length < 2, onClick: () => setConfirmDel(i) }]} />}
          </div>
        ))}
        <div className="flex-1" />
        <div className="text-[12px] text-muted flex gap-4 pr-3 whitespace-nowrap">
          {stats.count > 1 && <><span>Average: {sumDisplay(stats.avg)}</span><span>Count: {stats.count}</span><span>Sum: {sumDisplay(stats.sum)}</span></>}
          {stats.count <= 1 && <span>Ready</span>}
        </div>
      </div>
      <input ref={importRef} type="file" accept=".csv,.xlsx,.xls" hidden data-testid="excel-import-input" onChange={e => { void importFile(e.target.files?.[0]); e.target.value = '' }} />
      {confirmDel !== null && <ConfirmModal title="Delete sheet?" message={<>Delete <b>{wb.sheets[confirmDel]?.name}</b> and all of its data? Formulas that refer to it will show #REF!.</>} onConfirm={() => deleteSheet(confirmDel)} onClose={() => setConfirmDel(null)} />}
      {dialog?.kind === 'cf' && <CondFormatDialog range={rangeToA1(rect)} initial={dialog.initial} onSave={addRule} onClose={() => setDialog(null)} />}
      {dialog?.kind === 'chart' && <ChartDialog range={rangeToA1(rect)} initial={dialog.edit ? { type: dialog.edit.type, title: dialog.edit.title, range: dialog.edit.range } : undefined}
        onSave={v => (dialog.edit ? patchChart(dialog.edit.id, v) : addChart(v))} onClose={() => setDialog(null)} />}
      {ctx && <ContextMenu x={ctx.x} y={ctx.y} items={ctx.items} onClose={() => setCtx(null)} />}
      {renaming !== null && <PromptModal title="Rename sheet" label="Sheet name" initial={wb.sheets[renaming]?.name} confirm="Rename" onClose={() => setRenaming(null)} onSubmit={v => renameSheet(renaming, v)} />}
    </div>
  )
}
