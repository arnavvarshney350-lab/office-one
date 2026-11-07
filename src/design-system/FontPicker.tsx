import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { Check, ChevronDown, Cloud, CloudOff, Loader2, Search } from 'lucide-react'
import { cssFamily, useFonts, type FontCategory, type FontDef } from '../services/fonts'
import { cx } from './index'

const ROW = 34
const VIEW = 340
const OVERSCAN = 6
const RECENT_KEY = 'oo-recent-fonts'
const PREVIEW_BUDGET = 400

const CATS: { id: FontCategory | 'all'; label: string }[] = [
  { id: 'all', label: 'All' }, { id: 'sans', label: 'Sans' }, { id: 'serif', label: 'Serif' },
  { id: 'display', label: 'Display' }, { id: 'handwriting', label: 'Script' }, { id: 'mono', label: 'Mono' },
]

function readRecent(): string[] {
  try { const v = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]'); return Array.isArray(v) ? v.slice(0, 8) : [] } catch { return [] }
}
function pushRecent(name: string) {
  try { localStorage.setItem(RECENT_KEY, JSON.stringify([name, ...readRecent().filter(n => n !== name)].slice(0, 8))) } catch { /* private mode */ }
}

/**
 * Searchable font picker over the system fonts and the whole Google Fonts library.
 * The list is virtualised and previews load only for the rows actually on screen.
 */
export function FontPicker({ value, onPick, width = 150, label = 'Font family', disabled }: {
  value: string; onPick: (font: string) => void; width?: number; label?: string; disabled?: boolean
}) {
  const { fonts, systemCount, googleCount, online, stateOf, ensure } = useFonts()
  const [open, setOpen] = useState(false)
  const [q, setQ] = useState('')
  const [cat, setCat] = useState<FontCategory | 'all'>('all')
  const [idx, setIdx] = useState(0)
  const [scroll, setScroll] = useState(0)
  const [pos, setPos] = useState({ top: 0, left: 0 })
  const [recent, setRecent] = useState<string[]>(readRecent)
  const btn = useRef<HTMLButtonElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  const search = useRef<HTMLInputElement>(null)
  const loaded = useRef(new Set<string>())

  // the selected font must render in the trigger even before the list is opened
  useEffect(() => { if (value) void ensure(value) }, [value, ensure])

  const rows = useMemo(() => {
    const needle = q.trim().toLowerCase()
    let list = fonts
    if (cat !== 'all') list = list.filter(f => f.category === cat)
    if (needle) {
      const starts: FontDef[] = [], contains: FontDef[] = []
      for (const f of list) {
        const n = f.name.toLowerCase()
        if (n.startsWith(needle)) starts.push(f)
        else if (n.includes(needle)) contains.push(f)
      }
      return [...starts, ...contains]
    }
    if (cat !== 'all') return list
    const pinned = recent.map(n => fonts.find(f => f.name === n)).filter(Boolean) as FontDef[]
    const pinnedNames = new Set(pinned.map(f => f.name))
    return [...pinned, ...list.filter(f => !pinnedNames.has(f.name))]
  }, [fonts, q, cat, recent])

  const first = Math.max(0, Math.floor(scroll / ROW) - OVERSCAN)
  const last = Math.min(rows.length - 1, Math.ceil((scroll + VIEW) / ROW) + OVERSCAN)
  const visible = rows.slice(first, last + 1)

  // Load previews for the rows actually on screen, shortly after scrolling settles.
  useEffect(() => {
    if (!open) return
    const t = setTimeout(() => {
      for (const f of visible) {
        if (loaded.current.size > PREVIEW_BUDGET) break
        if (f.source === 'google' && !loaded.current.has(f.name)) { loaded.current.add(f.name); void ensure(f.name) }
      }
    }, 180)
    return () => clearTimeout(t)
  }, [open, first, last, rows, ensure]) // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => { setIdx(0); if (listRef.current) listRef.current.scrollTop = 0; setScroll(0) }, [q, cat])
  useLayoutEffect(() => { if (open) search.current?.focus() }, [open])
  useEffect(() => {
    if (!open) return
    const away = (e: MouseEvent) => { const t = e.target as Node; if (!btn.current?.contains(t) && !panel.current?.contains(t)) setOpen(false) }
    const follow = () => {
      const r = btn.current?.getBoundingClientRect()
      if (r) setPos({ top: Math.min(r.bottom + 4, window.innerHeight - VIEW - 110), left: Math.min(r.left, window.innerWidth - 300) })
    }
    document.addEventListener('mousedown', away); window.addEventListener('resize', follow); window.addEventListener('scroll', follow, true)
    return () => { document.removeEventListener('mousedown', away); window.removeEventListener('resize', follow); window.removeEventListener('scroll', follow, true) }
  }, [open])

  const pick = useCallback(async (name: string) => {
    setOpen(false); setQ('')
    if (name) { pushRecent(name); setRecent(readRecent()); await ensure(name) }
    onPick(name)
  }, [ensure, onPick])

  const scrollTo = (i: number) => {
    const el = listRef.current; if (!el) return
    const top = i * ROW
    if (top < el.scrollTop) el.scrollTop = top
    else if (top + ROW > el.scrollTop + el.clientHeight) el.scrollTop = top + ROW - el.clientHeight
  }
  const move = (d: number) => { const n = Math.max(0, Math.min(rows.length - 1, idx + d)); setIdx(n); scrollTo(n) }

  const toggle = () => {
    if (disabled) return
    if (!open && btn.current) {
      const r = btn.current.getBoundingClientRect()
      setPos({ top: Math.min(r.bottom + 4, window.innerHeight - VIEW - 110), left: Math.min(r.left, window.innerWidth - 300) })
    }
    setOpen(o => !o)
  }

  return (
    <>
      <button ref={btn} type="button" aria-label={label} aria-haspopup="listbox" aria-expanded={open} disabled={disabled}
        onMouseDown={e => e.preventDefault()} onClick={toggle} style={{ width }}
        className="h-8 px-2.5 rounded-md bg-surface border border-line hover:bg-panel2 flex items-center justify-between gap-1 text-[13px]">
        <span className="truncate" style={{ fontFamily: value ? cssFamily(value) : undefined }}>{value || 'Default'}</span>
        <ChevronDown size={13} className="shrink-0 text-muted" />
      </button>

      {open && createPortal(
        <div ref={panel} style={{ position: 'fixed', top: pos.top, left: pos.left }}
          className="z-[120] w-[300px] bg-surface rounded-xl shadow-pop overflow-hidden flex flex-col menu-in">
          <div className="p-2 border-b border-line">
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted" />
              <input ref={search} aria-label="Search fonts" value={q} onChange={e => setQ(e.target.value)} placeholder={`Search ${rows.length.toLocaleString()} fonts`}
                onKeyDown={e => {
                  if (e.key === 'ArrowDown') { e.preventDefault(); move(1) }
                  else if (e.key === 'ArrowUp') { e.preventDefault(); move(-1) }
                  else if (e.key === 'Enter') { e.preventDefault(); if (rows[idx]) void pick(rows[idx].name) }
                  else if (e.key === 'Escape') { e.preventDefault(); setOpen(false) }
                }}
                className="w-full h-8 pl-8 pr-2 rounded-md border border-line bg-panel" />
            </div>
            <div className="flex gap-1 mt-2" role="tablist" aria-label="Font categories">
              {CATS.map(c => (
                <button key={c.id} role="tab" aria-selected={cat === c.id} onClick={() => setCat(c.id)}
                  className={cx('px-1.5 h-6 rounded text-[11px] flex-1', cat === c.id ? 'bg-accent-soft text-accent font-semibold' : 'text-muted hover:bg-panel2')}>{c.label}</button>
              ))}
            </div>
          </div>

          <div ref={listRef} role="listbox" aria-label="Fonts" style={{ height: VIEW }} className="overflow-y-auto"
            onScroll={e => setScroll(e.currentTarget.scrollTop)}>
            {value && !q && (
              <button role="option" aria-selected={false} onMouseDown={e => e.preventDefault()} onClick={() => void pick('')}
                className="w-full text-left px-3 h-8 text-[12.5px] text-muted hover:bg-panel2 border-b border-line">Use the theme default</button>
            )}
            <div style={{ height: rows.length * ROW, position: 'relative' }}>
              {visible.map((f, i) => {
                const n = first + i
                const st = stateOf(f.name)
                const usable = f.source === 'system' || st === 'ready'
                return (
                  <button key={f.name} role="option" aria-selected={f.name === value}
                    onMouseDown={e => e.preventDefault()} onMouseEnter={() => setIdx(n)} onClick={() => void pick(f.name)}
                    style={{ position: 'absolute', top: n * ROW, height: ROW, left: 0, right: 0, fontFamily: usable ? cssFamily(f.name) : undefined }}
                    className={cx('w-full px-3 flex items-center gap-2 text-left text-[15px]', n === idx ? 'bg-accent-soft' : 'hover:bg-panel2')}>
                    <span className="truncate flex-1">{f.name}</span>
                    {st === 'failed' && <span title="This font could not be downloaded" className="text-[10px] text-muted shrink-0">unavailable</span>}
                    {f.name === value && <Check size={14} className="text-accent shrink-0" />}
                  </button>
                )
              })}
            </div>
            {rows.length === 0 && <div className="px-3 py-4 text-[13px] text-muted">No fonts match “{q}”.</div>}
          </div>

          <div className="px-3 py-1.5 border-t border-line text-[11px] text-muted flex items-center gap-1.5">
            {online === null ? <><Loader2 size={11} className="animate-spin" />Checking Google Fonts…</>
              : online ? <><Cloud size={11} />{systemCount} installed · {googleCount.toLocaleString()} Google Fonts</>
                : <><CloudOff size={11} />Offline — showing {systemCount} installed fonts</>}
          </div>
        </div>, document.body)}
    </>
  )
}

export const SIZE_PRESETS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72, 96]

/** Font size box: type any value, or pick one of the presets. */
export function SizeBox({ value, onPick, min = 1, max = 409, disabled, label = 'Font size' }: {
  value: number | ''; onPick: (n: number) => void; min?: number; max?: number; disabled?: boolean; label?: string
}) {
  const [draft, setDraft] = useState<string | null>(null)
  const id = useRef(`sizes-${Math.random().toString(36).slice(2)}`).current
  const commit = () => {
    if (draft === null) return
    const n = Math.round(Number(draft))
    setDraft(null)
    if (Number.isFinite(n) && n >= min && n <= max) onPick(n)
  }
  return (
    <>
      <input aria-label={label} list={id} disabled={disabled} inputMode="numeric" value={draft ?? String(value)} placeholder="—"
        onFocus={e => e.target.select()}
        onChange={e => { setDraft(e.target.value); const n = Number(e.target.value); if (SIZE_PRESETS.includes(n)) { setDraft(null); onPick(n) } }}
        onBlur={commit}
        onKeyDown={e => { if (e.key === 'Enter') { commit(); (e.target as HTMLInputElement).blur() } if (e.key === 'Escape') { setDraft(null); (e.target as HTMLInputElement).blur() } }}
        className="w-[62px] h-8 px-2 rounded-md border border-line bg-surface text-[13px] text-center" />
      <datalist id={id}>{SIZE_PRESETS.map(n => <option key={n} value={n} />)}</datalist>
    </>
  )
}
