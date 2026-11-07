import { useEffect, useRef, useState, type ButtonHTMLAttributes, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { Check, CloudOff, FileSpreadsheet, FileText, FileType2, Folder, Presentation, FileType } from 'lucide-react'
import type { FileKind, SaveState } from '../types'

export function cx(...a: (string | false | null | undefined)[]) { return a.filter(Boolean).join(' ') }

export function Button({ variant = 'default', className, ...p }: ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'default' | 'primary' | 'ghost' | 'danger' }) {
  const v = {
    default: 'bg-surface border-line hover:bg-panel2 hover:border-line',
    primary: 'bg-accent text-accent-fg border-transparent hover:brightness-110 active:brightness-95',
    ghost: 'border-transparent hover:bg-panel2',
    danger: 'bg-danger text-white border-transparent hover:brightness-110',
  }[variant]
  return <button type="button" {...p} className={cx('inline-flex items-center gap-1.5 px-3 h-8 rounded-md border text-[13px] font-medium transition-colors duration-100 whitespace-nowrap select-none', v, className)} />
}

export function IconButton({ active, label, className, children, ...p }: ButtonHTMLAttributes<HTMLButtonElement> & { active?: boolean; label: string }) {
  return (
    <button type="button" title={label} aria-label={label} aria-pressed={active} {...p}
      onMouseDown={e => { e.preventDefault(); p.onMouseDown?.(e) }}
      className={cx('inline-flex items-center justify-center h-8 min-w-8 px-1.5 rounded-md transition-colors duration-100 select-none', active ? 'bg-accent-soft text-accent' : 'text-ink hover:bg-panel2', className)}>
      {children}
    </button>
  )
}

export function Divider() { return <span aria-hidden className="w-px h-5 bg-line mx-1.5 shrink-0" /> }

/**
 * Toolbar group. The caption is kept for assistive technology and tooltips but not drawn: a thin
 * divider already says "these belong together", and dropping captions halves the ribbon's height.
 */
export function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div role="group" aria-label={label} title={label}
      className="flex items-center gap-0.5 shrink-0 px-2 first:pl-0 relative before:absolute before:left-0 before:top-1/2 before:-translate-y-1/2 before:h-5 before:w-px before:bg-line first:before:hidden">
      {children}
    </div>
  )
}
export function Ribbon({ tabs, tab, onTab, children }: { tabs: string[]; tab: string; onTab: (t: string) => void; children: ReactNode }) {
  return (
    <div className="bg-surface shrink-0">
      <div className="flex gap-1 px-4 border-b border-line" role="tablist">
        {tabs.map(t => (
          <button key={t} role="tab" aria-selected={t === tab} onClick={() => onTab(t)}
            className={cx('relative px-3 h-9 text-[13px] transition-colors', t === tab ? 'text-ink font-medium' : 'text-muted hover:text-ink')}>
            {t}
            {t === tab && <span aria-hidden className="absolute left-2 right-2 -bottom-px h-[2px] rounded-full bg-accent" />}
          </button>
        ))}
      </div>
      <div role="toolbar" aria-label={`${tab} tools`} className="bg-panel border-b border-line flex items-center overflow-x-auto px-4 min-h-[46px]">{children}</div>
    </div>
  )
}

/** Popover rendered in a portal so ribbon overflow can never clip it. */
function usePopover() {
  const [open, setOpen] = useState(false)
  const [pos, setPos] = useState<{ top: number; left: number; right: number }>({ top: 0, left: 0, right: 0 })
  const btn = useRef<HTMLDivElement>(null)
  const panel = useRef<HTMLDivElement>(null)
  const toggle = () => {
    if (!open && btn.current) { const r = btn.current.getBoundingClientRect(); setPos({ top: r.bottom + 4, left: r.left, right: window.innerWidth - r.right }) }
    setOpen(o => !o)
  }
  useEffect(() => {
    if (!open) return
    const h = (e: MouseEvent) => { const t = e.target as Node; if (!btn.current?.contains(t) && !panel.current?.contains(t)) setOpen(false) }
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') setOpen(false) }
    // the popover follows its trigger when something scrolls or the window resizes, instead of vanishing
    const follow = () => { const r = btn.current?.getBoundingClientRect(); if (r) setPos({ top: r.bottom + 4, left: r.left, right: window.innerWidth - r.right }) }
    document.addEventListener('mousedown', h); document.addEventListener('keydown', k); window.addEventListener('resize', follow); window.addEventListener('scroll', follow, true)
    return () => { document.removeEventListener('mousedown', h); document.removeEventListener('keydown', k); window.removeEventListener('resize', follow); window.removeEventListener('scroll', follow, true) }
  }, [open])
  return { open, setOpen, pos, btn, panel, toggle }
}

export function Menu({ trigger, items, align = 'left', title }: {
  trigger: ReactNode; title?: string; align?: 'left' | 'right'
  items: ({ label: string; onClick: () => void; icon?: ReactNode; danger?: boolean; disabled?: boolean } | 'sep')[]
}) {
  const { open, setOpen, pos, btn, panel, toggle } = usePopover()
  return (
    <div ref={btn} className="relative inline-block">
      <button type="button" title={title} aria-label={title} aria-haspopup="menu" aria-expanded={open} onMouseDown={e => e.preventDefault()} onClick={toggle}
        className="inline-flex items-center gap-1 h-8 px-2 rounded-md hover:bg-panel2 aria-expanded:bg-panel2 text-[13px] select-none">{trigger}</button>
      {open && createPortal(
        <div ref={panel} role="menu" style={{ position: 'fixed', top: pos.top, ...(align === 'right' ? { right: pos.right } : { left: pos.left }) }}
          className="z-[120] min-w-[200px] p-1 bg-surface rounded-xl shadow-pop menu-in">
          {items.map((it, i) => it === 'sep' ? <div key={i} className="my-1 border-t border-line" /> : (
            <button key={i} role="menuitem" disabled={it.disabled} onMouseDown={e => e.preventDefault()}
              onClick={() => { setOpen(false); it.onClick() }}
              className={cx('w-full flex items-center gap-2 px-2.5 h-8 rounded-md text-left text-[13px] hover:bg-panel2', it.danger && 'text-danger')}>
              {it.icon}{it.label}
            </button>
          ))}
        </div>, document.body)}
    </div>
  )
}

export function Modal({ title, onClose, children, footer, width = 420 }: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode; width?: number }) {
  useEffect(() => {
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', k); return () => document.removeEventListener('keydown', k)
  }, [onClose])
  return createPortal(
    <div className="fixed inset-0 z-[90] bg-[#0b1020]/40 flex items-center justify-center p-4 backdrop-in" onMouseDown={e => { if (e.target === e.currentTarget) onClose() }}>
      <div role="dialog" aria-label={title} aria-modal="true" style={{ width }} className="max-w-full bg-surface rounded-2xl shadow-pop dialog-in">
        <div className="px-6 pt-5 pb-1 text-[17px] font-semibold tracking-tight">{title}</div>
        <div className="px-6 py-3">{children}</div>
        {footer && <div className="px-6 py-4 flex justify-end gap-2 mt-1">{footer}</div>}
      </div>
    </div>, document.body)
}

export function PromptModal({ title, label, initial = '', confirm = 'OK', onSubmit, onClose }: { title: string; label: string; initial?: string; confirm?: string; onSubmit: (v: string) => void; onClose: () => void }) {
  const [v, setV] = useState(initial)
  const go = () => { onSubmit(v); onClose() }
  return (
    <Modal title={title} onClose={onClose} footer={<><Button onClick={onClose}>Cancel</Button><Button variant="primary" onClick={go}>{confirm}</Button></>}>
      <label className="block text-[13px] text-muted mb-1">{label}</label>
      <input autoFocus value={v} onChange={e => setV(e.target.value)} onFocus={e => e.target.select()} onKeyDown={e => { if (e.key === 'Enter') go() }}
        className="w-full h-10 px-3 rounded-lg border border-line bg-panel focus:bg-surface" />
    </Modal>
  )
}

export const SWATCHES = ['#000000', '#434343', '#7f7f7f', '#d9d9d9', '#ffffff', '#c00000', '#ff0000', '#ff9900', '#ffff00', '#00b050', '#00b0f0', '#2563eb', '#7030a0', '#ff66cc', '#8b5a2b']
export function ColorMenu({ icon, label, onPick, allowNone }: { icon: ReactNode; label: string; onPick: (c: string | null) => void; allowNone?: boolean }) {
  const { open, setOpen, pos, btn, panel, toggle } = usePopover()
  return (
    <div ref={btn} className="relative inline-block">
      <IconButton label={label} onClick={toggle}>{icon}</IconButton>
      {open && createPortal(
        <div ref={panel} style={{ position: 'fixed', top: pos.top, left: pos.left }} className="z-[120] p-2.5 bg-surface rounded-xl shadow-pop w-[140px] menu-in">
          <div className="grid grid-cols-5 gap-1.5">
            {SWATCHES.map(c => <button key={c} aria-label={c} title={c} onMouseDown={e => e.preventDefault()} onClick={() => { onPick(c); setOpen(false) }} style={{ background: c }} className="w-[18px] h-[18px] rounded-full ring-1 ring-inset ring-black/10 hover:scale-110 transition-transform" />)}
          </div>
          {allowNone && <button onMouseDown={e => e.preventDefault()} onClick={() => { onPick(null); setOpen(false) }} className="mt-2.5 w-full text-[12px] h-7 rounded-md hover:bg-panel2">No colour</button>}
        </div>, document.body)}
    </div>
  )
}

export function SaveBadge({ state }: { state: SaveState }) {
  const base = 'inline-flex items-center gap-1.5 text-[12px] tabular'
  if (state === 'idle') return <span className={cx(base, 'text-muted')}>Ready</span>
  if (state === 'saving') return <span className={cx(base, 'text-muted')} role="status"><span className="w-1.5 h-1.5 rounded-full bg-accent animate-pulse" />Saving…</span>
  if (state === 'saved') return <span className={cx(base, 'text-muted')} role="status"><Check size={13} />Saved</span>
  return <span className={cx(base, 'text-danger font-medium')} role="alert"><CloudOff size={13} />Save failed. Retrying…</span>
}

export const KIND_STYLE: Record<FileKind, { color: string; Icon: typeof FileText }> = {
  docx: { color: 'var(--color-word)', Icon: FileText },
  xlsx: { color: 'var(--color-excel)', Icon: FileSpreadsheet },
  pptx: { color: 'var(--color-ppt)', Icon: Presentation },
  pdf: { color: '#d92d20', Icon: FileType2 },
  csv: { color: '#0e9384', Icon: FileSpreadsheet },
  txt: { color: '#667085', Icon: FileType },
  folder: { color: '#d4a017', Icon: Folder },
}
export function FileIcon({ kind, size = 20 }: { kind: FileKind; size?: number }) {
  const { color, Icon } = KIND_STYLE[kind]
  return <Icon size={size} style={{ color }} aria-hidden />
}

/** Brand mark: a ring with a single point in the middle, "O" and "one". */
export function AppMark({ size = 28 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden>
      <rect width="32" height="32" rx="9" fill="var(--accent)" />
      <circle cx="16" cy="16" r="7.2" fill="none" stroke="var(--accent-fg)" strokeWidth="3" />
      <circle cx="16" cy="16" r="2.1" fill="var(--accent-fg)" />
    </svg>
  )
}

export type MenuEntry = { label: string; onClick: () => void; icon?: ReactNode; danger?: boolean; disabled?: boolean; hint?: string } | 'sep'

/** Right-click menu rendered at viewport coordinates. */
export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuEntry[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ left: x, top: y })
  useEffect(() => {
    const r = ref.current?.getBoundingClientRect()
    if (r) setPos({ left: Math.max(4, Math.min(x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) })
    const h = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) onClose() }
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('mousedown', h); document.addEventListener('keydown', k); window.addEventListener('resize', onClose); window.addEventListener('blur', onClose)
    return () => { document.removeEventListener('mousedown', h); document.removeEventListener('keydown', k); window.removeEventListener('resize', onClose); window.removeEventListener('blur', onClose) }
  }, [x, y, onClose])
  return createPortal(
    <div ref={ref} role="menu" style={{ position: 'fixed', ...pos }} className="z-[130] min-w-[220px] p-1 bg-surface rounded-xl shadow-pop menu-in" onContextMenu={e => e.preventDefault()}>
      {items.map((it, i) => it === 'sep' ? <div key={i} className="my-1 border-t border-line" /> : (
        <button key={i} role="menuitem" disabled={it.disabled} onMouseDown={e => e.preventDefault()} onClick={() => { onClose(); it.onClick() }}
          className={cx('w-full flex items-center gap-2 px-2.5 h-8 rounded-md text-left text-[13px] hover:bg-panel2', it.danger && 'text-danger')}>
          <span className="w-4 shrink-0">{it.icon}</span><span className="flex-1">{it.label}</span>{it.hint && <span className="text-[11px] text-muted">{it.hint}</span>}
        </button>
      ))}
    </div>, document.body)
}

export function ConfirmModal({ title, message, confirm = 'Delete', danger = true, onConfirm, onClose }: { title: string; message: ReactNode; confirm?: string; danger?: boolean; onConfirm: () => void; onClose: () => void }) {
  return (
    <Modal title={title} onClose={onClose} footer={<><Button onClick={onClose}>Cancel</Button><Button variant={danger ? 'danger' : 'primary'} autoFocus onClick={() => { onConfirm(); onClose() }}>{confirm}</Button></>}>
      <div className="text-[13.5px] text-muted leading-relaxed">{message}</div>
    </Modal>
  )
}
