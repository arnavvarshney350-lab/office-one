import { useMemo } from 'react'
import { FileIcon } from '../design-system'
import { Evaluator, formatValue } from '../editors/excel/formula'
import { toAddr, type Workbook } from '../editors/excel/model'
import { SlideView } from '../editors/powerpoint/SlideView'
import { SLIDE_H, SLIDE_W, type Deck } from '../editors/powerpoint/model'
import type { OFile } from '../types'

/** A small, honest picture of what is inside a file: its words, its cells, its first slide. */
export default function FilePreview({ file, width }: { file: OFile; width: number }) {
  const height = Math.round(width * 0.75)
  return (
    <div style={{ width, height }} className="bg-canvas rounded-lg overflow-hidden flex items-center justify-center relative" aria-hidden>
      {file.kind === 'docx' && <DocPreview file={file} h={height} />}
      {file.kind === 'xlsx' && <SheetPreview file={file} w={width} h={height} />}
      {file.kind === 'pptx' && <DeckPreview file={file} w={width} h={height} />}
      {(file.kind === 'pdf' || file.kind === 'txt' || file.kind === 'csv' || file.kind === 'folder') && <FileIcon kind={file.kind} size={34} />}
    </div>
  )
}

function DocPreview({ file, h }: { file: OFile; h: number }) {
  const lines = useMemo(() => {
    const html = typeof file.content === 'string' ? `<p>${file.content}</p>` : (file.content as { html?: string } | null)?.html ?? ''
    const d = new DOMParser().parseFromString(html, 'text/html')
    return [...d.body.querySelectorAll('h1,h2,h3,p,li')].map(el => ({ t: (el.textContent ?? '').trim(), h: el.tagName.startsWith('H') })).filter(l => l.t).slice(0, 9)
  }, [file.content])
  const pageH = h - 22
  return (
    <div className="bg-page shadow-soft rounded-[2px] overflow-hidden px-3.5 pt-3.5" style={{ width: pageH * 0.77, height: pageH }}>
      {lines.length === 0 && <div className="text-[7px] text-muted">Empty document</div>}
      {lines.map((l, i) => l.h
        ? <div key={i} className="font-display text-[10px] leading-tight mb-1 truncate" style={{ fontWeight: 600 }}>{l.t}</div>
        : <div key={i} className="text-[5.5px] leading-[1.5] text-muted mb-[2px] line-clamp-2">{l.t}</div>)}
    </div>
  )
}

function SheetPreview({ file, w, h }: { file: OFile; w: number; h: number }) {
  const grid = useMemo(() => {
    const wb = file.content as Workbook | null
    if (!wb?.sheets?.length) return null
    const si = Math.min(Math.max(0, wb.active), wb.sheets.length - 1)
    const ev = new Evaluator(wb)
    ev.ensureFormulas(si)
    const cols = 5, rows = 8
    return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => {
      const a = toAddr(c, r), cell = wb.sheets[si].cells[a]
      return { text: formatValue(ev.displayValue(si, a), cell?.s?.fmt, cell?.s?.dec), b: !!cell?.s?.b, bg: cell?.s?.bg, num: typeof ev.displayValue(si, a) === 'number' }
    }))
  }, [file.content])
  if (!grid) return <FileIcon kind="xlsx" size={34} />
  return (
    <div className="bg-page shadow-soft rounded-[3px] overflow-hidden" style={{ width: w - 24, height: h - 22 }}>
      {grid.map((row, r) => (
        <div key={r} className="flex" style={{ height: `${100 / grid.length}%` }}>
          {row.map((c, i) => (
            <div key={i} className="flex-1 min-w-0 border-r border-b border-line px-[3px] text-[5.5px] leading-none flex items-center truncate" style={{ background: c.bg ?? (r === 0 ? 'color-mix(in srgb, var(--c-excel) 10%, transparent)' : undefined), color: c.bg ? '#141824' : undefined, fontWeight: c.b || r === 0 ? 600 : 400, justifyContent: c.num ? 'flex-end' : 'flex-start' }}>{c.text}</div>
          ))}
        </div>
      ))}
    </div>
  )
}

function DeckPreview({ file, w, h }: { file: OFile; w: number; h: number }) {
  const deck = file.content as Deck | null
  if (!deck?.slides?.length) return <FileIcon kind="pptx" size={34} />
  const scale = Math.min((w - 24) / SLIDE_W, (h - 22) / SLIDE_H)
  return <div className="shadow-soft rounded-[3px] overflow-hidden" style={{ width: SLIDE_W * scale, height: SLIDE_H * scale }}><SlideView slide={deck.slides[0]} deckTheme={deck.theme} scale={scale} /></div>
}
