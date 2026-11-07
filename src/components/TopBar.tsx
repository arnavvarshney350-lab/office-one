import { type ReactNode } from 'react'
import { Moon, Search, Sun } from 'lucide-react'
import { IconButton } from '../design-system'
import { useTheme } from '../contexts/ThemeContext'

export function ThemeToggle() {
  const { theme, toggle } = useTheme()
  return (
    <IconButton label={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'} onClick={toggle}>
      {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
    </IconButton>
  )
}

export const OPEN_PALETTE = 'oo:open-palette'
const MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform)

/** Page header. The search field is the front door to the command palette, so there is one search, not two. */
export function TopBar({ contained, children }: { contained?: boolean; children?: ReactNode }) {
  return (
    <header className="h-16 shrink-0">
      <div className={`h-full flex items-center gap-4 ${contained ? 'max-w-[1080px] mx-auto px-10' : 'px-8'}`}>
      <button type="button" aria-label="Search files and commands" onClick={() => window.dispatchEvent(new Event(OPEN_PALETTE))}
        className="flex-1 max-w-xl h-10 pl-3.5 pr-3 rounded-xl bg-panel border border-transparent hover:border-line flex items-center gap-2.5 text-left text-[13.5px] text-muted transition-colors">
        <Search size={16} className="shrink-0" /><span className="flex-1">Search files or jump to anything</span>
        <kbd aria-hidden className="text-[11px] border border-line rounded px-1.5 py-0.5 bg-surface">{MAC ? '⌘K' : 'Ctrl K'}</kbd>
      </button>
      <div className="flex-1" />
      {children}
      <ThemeToggle />
      </div>
    </header>
  )
}
