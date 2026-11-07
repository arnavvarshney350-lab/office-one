import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { useNavigate } from 'react-router-dom'
import { CornerDownLeft, FileText, HardDrive, Home, Keyboard, Moon, Presentation, Search, Settings, Sheet, Sun } from 'lucide-react'
import { FileIcon, cx } from '../design-system'
import { OPEN_PALETTE } from './TopBar'
import { useFiles } from '../hooks/useFiles'
import { useOpenFile } from '../hooks/useOpenFile'
import { useTheme } from '../contexts/ThemeContext'
import { createFile, fullName, routeFor } from '../file-system/fileService'
import { formatDate } from '../utilities/id'
import type { FileKind } from '../types'

interface Cmd { id: string; label: string; hint?: string; icon: React.ReactNode; run: () => void; keywords?: string }

/** ⌘K: jump anywhere, start anything, find any file, by typing. */
export function CommandPalette({ onShortcuts }: { onShortcuts: () => void }) {
  const [open, setOpen] = useState(false)
  const [q, setQ] = useState('')
  const [idx, setIdx] = useState(0)
  const nav = useNavigate()
  const openFile = useOpenFile()
  const files = useFiles()
  const { theme, toggle } = useTheme()
  const input = useRef<HTMLInputElement>(null)
  const list = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); setOpen(o => !o); setQ(''); setIdx(0) }
    }
    const show = () => { setOpen(true); setQ(''); setIdx(0) }
    window.addEventListener('keydown', k); window.addEventListener(OPEN_PALETTE, show)
    return () => { window.removeEventListener('keydown', k); window.removeEventListener(OPEN_PALETTE, show) }
  }, [])
  useEffect(() => { if (open) input.current?.focus() }, [open])

  const create = (kind: FileKind) => async () => { const f = await createFile(kind); nav(routeFor(f)!) }
  const commands: Cmd[] = useMemo(() => [
    { id: 'new-doc', label: 'New document', icon: <FileText size={16} style={{ color: 'var(--c-word)' }} />, run: create('docx'), keywords: 'word write create' },
    { id: 'new-sheet', label: 'New spreadsheet', icon: <Sheet size={16} style={{ color: 'var(--c-excel)' }} />, run: create('xlsx'), keywords: 'excel workbook create' },
    { id: 'new-deck', label: 'New presentation', icon: <Presentation size={16} style={{ color: 'var(--c-ppt)' }} />, run: create('pptx'), keywords: 'powerpoint slides deck create' },
    { id: 'home', label: 'Go to Home', icon: <Home size={16} />, run: () => nav('/') },
    { id: 'drive', label: 'Go to Drive', icon: <HardDrive size={16} />, run: () => nav('/files'), keywords: 'files folders trash favourites' },
    { id: 'word', label: 'Go to Word', icon: <FileText size={16} />, run: () => nav('/word') },
    { id: 'excel', label: 'Go to Excel', icon: <Sheet size={16} />, run: () => nav('/excel') },
    { id: 'ppt', label: 'Go to PowerPoint', icon: <Presentation size={16} />, run: () => nav('/powerpoint') },
    { id: 'theme', label: theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme', icon: theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />, run: toggle, keywords: 'appearance mode' },
    { id: 'keys', label: 'Keyboard shortcuts', hint: '?', icon: <Keyboard size={16} />, run: onShortcuts },
    { id: 'settings', label: 'Settings', icon: <Settings size={16} />, run: () => nav('/settings') },
  ], [theme]) // eslint-disable-line react-hooks/exhaustive-deps

  const needle = q.trim().toLowerCase()
  const cmdHits = commands.filter(c => !needle || needle.split(/\s+/).every(w => `${c.label} ${c.keywords ?? ''}`.toLowerCase().includes(w)))
  const fileHits = needle && files ? files.filter(f => !f.trashed && f.kind !== 'folder' && needle.split(/\s+/).every(w => fullName(f).toLowerCase().includes(w))).sort((a, b) => b.updatedAt - a.updatedAt).slice(0, 6) : []
  const recent = !needle && files ? files.filter(f => !f.trashed && f.kind !== 'folder' && f.openedAt).sort((a, b) => (b.openedAt ?? 0) - (a.openedAt ?? 0)).slice(0, 4) : []
  const rows = [
    ...[...recent, ...fileHits].map(f => ({ key: f.id, group: needle ? 'Files' : 'Recent', label: fullName(f), meta: formatDate(f.updatedAt), icon: <FileIcon kind={f.kind} size={16} />, run: () => openFile(f) })),
    ...cmdHits.map(c => ({ key: c.id, group: 'Actions', label: c.label, meta: c.hint, icon: c.icon, run: c.run })),
  ]
  useEffect(() => { setIdx(0) }, [q])
  useEffect(() => { list.current?.querySelector('[aria-selected=true]')?.scrollIntoView({ block: 'nearest' }) }, [idx])
  const close = () => setOpen(false)
  const choose = (i: number) => { const r = rows[i]; if (!r) return; close(); r.run() }

  if (!open) return null
  return createPortal(
    <div className="fixed inset-0 z-[200] bg-[#0b1020]/40 backdrop-in flex justify-center pt-[14vh] px-4" onMouseDown={e => { if (e.target === e.currentTarget) close() }}>
      <div role="dialog" aria-label="Command palette" aria-modal="true" className="w-[560px] max-w-full h-fit max-h-[62vh] flex flex-col bg-surface rounded-2xl shadow-pop dialog-in overflow-hidden">
        <div className="flex items-center gap-3 px-4 h-14 border-b border-line">
          <Search size={18} className="text-muted shrink-0" />
          <input ref={input} aria-label="Search commands and files" role="combobox" aria-expanded aria-controls="palette-list" value={q} onChange={e => setQ(e.target.value)}
            placeholder="Search files or type a command"
            onKeyDown={e => {
              if (e.key === 'ArrowDown') { e.preventDefault(); setIdx(i => Math.min(rows.length - 1, i + 1)) }
              else if (e.key === 'ArrowUp') { e.preventDefault(); setIdx(i => Math.max(0, i - 1)) }
              else if (e.key === 'Enter') { e.preventDefault(); choose(idx) }
              else if (e.key === 'Escape') close()
            }}
            className="flex-1 bg-transparent outline-none text-[15px] placeholder:text-muted" />
          <kbd className="text-[11px] text-muted border border-line rounded px-1.5 py-0.5">esc</kbd>
        </div>
        <div ref={list} id="palette-list" role="listbox" aria-label="Results" className="overflow-y-auto p-1.5">
          {rows.length === 0 && <div className="px-3 py-8 text-center text-[13px] text-muted">Nothing matches “{q}”. Try a file name or an action like “new document”.</div>}
          {rows.map((r, i) => (
            <div key={r.key}>
              {(i === 0 || rows[i - 1].group !== r.group) && <div className="px-3 pt-2.5 pb-1 text-[12px] text-muted">{r.group}</div>}
              <button role="option" aria-selected={i === idx} onMouseMove={() => setIdx(i)} onClick={() => choose(i)}
                className={cx('w-full flex items-center gap-3 px-3 h-10 rounded-lg text-left text-[13.5px]', i === idx ? 'bg-accent-soft' : '')}>
                <span className="w-5 flex justify-center shrink-0">{r.icon}</span>
                <span className="flex-1 truncate">{r.label}</span>
                {r.meta && <span className="text-[12px] text-muted tabular">{r.meta}</span>}
                {i === idx && <CornerDownLeft size={13} className="text-muted" />}
              </button>
            </div>
          ))}
        </div>
      </div>
    </div>, document.body)
}
