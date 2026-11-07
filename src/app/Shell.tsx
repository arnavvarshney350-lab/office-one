import { useEffect, useState, type ReactNode } from 'react'
import { NavLink, useLocation } from 'react-router-dom'
import { FileText, HardDrive, Home, Keyboard, Presentation, Settings, Sheet } from 'lucide-react'
import { ShortcutsDialog } from '../components/ShortcutsDialog'
import { CommandPalette } from '../components/CommandPalette'
import { DropImport } from '../components/DropImport'
import { AppMark, cx } from '../design-system'

const NAV = [
  { to: '/', label: 'Home', Icon: Home, color: 'var(--accent)' },
  { to: '/word', label: 'Word', Icon: FileText, color: 'var(--c-word)' },
  { to: '/excel', label: 'Excel', Icon: Sheet, color: 'var(--c-excel)' },
  { to: '/powerpoint', label: 'PowerPoint', Icon: Presentation, color: 'var(--c-ppt)' },
  { to: '/files', label: 'Drive', Icon: HardDrive, color: 'var(--accent)' },
]
const TITLES: Record<string, string> = { '/': 'Home', '/files': 'Drive', '/settings': 'Settings', '/word': 'Word', '/excel': 'Excel', '/powerpoint': 'PowerPoint' }

const item = (active: boolean) => cx('relative w-12 py-2 rounded-xl flex flex-col items-center gap-1 text-[10.5px] leading-none transition-colors', active ? 'bg-surface text-ink font-medium shadow-soft' : 'text-muted hover:bg-panel2 hover:text-ink')

export function Shell({ children }: { children: ReactNode }) {
  const [help, setHelp] = useState(false)
  const { pathname } = useLocation()
  // editors set their own title to the document name; list pages get a per-route title
  useEffect(() => { const t = TITLES[pathname]; if (t) document.title = `${t} · Office One` }, [pathname])
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement
      if (e.key === '?' && !e.metaKey && !e.ctrlKey && !['INPUT', 'TEXTAREA', 'SELECT'].includes(t.tagName) && !t.isContentEditable) { e.preventDefault(); setHelp(true) }
    }
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [])
  return (
    <div className="h-full flex">
      <a href="#main" className="sr-only focus:not-sr-only focus:absolute focus:z-[300] focus:top-2 focus:left-2 focus:px-3 focus:py-2 focus:rounded-md focus:bg-accent focus:text-accent-fg">Skip to content</a>
      <nav aria-label="Applications" className="w-[64px] shrink-0 flex flex-col items-center py-3 gap-1">
        <NavLink to="/" aria-label="Office One home" className="mb-4 mt-0.5 rounded-xl"><AppMark size={32} /></NavLink>
        {NAV.map(({ to, label, Icon, color }) => (
          <NavLink key={to} to={to} end={to === '/'} title={label} className={({ isActive }) => item(isActive)}>
            {({ isActive }) => (<><Icon size={20} strokeWidth={isActive ? 2 : 1.75} style={{ color: isActive ? color : undefined }} />{label}</>)}
          </NavLink>
        ))}
        <div className="flex-1" />
        <button onClick={() => setHelp(true)} title="Keyboard shortcuts (?)" aria-label="Keyboard shortcuts" className={item(false)}><Keyboard size={20} strokeWidth={1.75} />Shortcuts</button>
        <NavLink to="/settings" title="Settings" className={({ isActive }) => item(isActive)}><Settings size={20} strokeWidth={1.75} />Settings</NavLink>
      </nav>
      {/* the working surface is a lifted sheet against the desk-coloured rail */}
      <main id="main" tabIndex={-1} className="flex-1 min-w-0 min-h-0 flex flex-col outline-none bg-bg">
        <div className="flex-1 min-h-0 flex flex-col bg-surface rounded-tl-2xl border-l border-t border-line overflow-hidden">{children}</div>
      </main>
      {help && <ShortcutsDialog onClose={() => setHelp(false)} />}
      <CommandPalette onShortcuts={() => setHelp(true)} />
      <DropImport />
    </div>
  )
}
