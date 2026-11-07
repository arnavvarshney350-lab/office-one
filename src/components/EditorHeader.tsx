import { useEffect, useState, type ReactNode } from 'react'
import { Link } from 'react-router-dom'
import { ChevronLeft, Star } from 'lucide-react'
import { IconButton, Menu, SaveBadge } from '../design-system'
import { ThemeToggle } from './TopBar'
import type { SaveState } from '../types'
import { renameFile, toggleFavourite } from '../file-system/fileService'
import type { OFile } from '../types'
import { useLiveQuery } from 'dexie-react-hooks'
import { db } from '../storage/db'

export type MenuItem = Parameters<typeof Menu>[0]['items'][number]

export function EditorHeader({ file, color, appName, save, fileMenu, children }: {
  file: OFile; color: string; appName: string; save: SaveState; fileMenu: MenuItem[]; children?: ReactNode
}) {
  const live = useLiveQuery(() => db.files.get(file.id), [file.id])
  const [name, setName] = useState(file.name)
  useEffect(() => { if (live && live.name !== name && document.activeElement?.id !== 'doc-title') setName(live.name) }, [live]) // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { document.title = `${name || 'Untitled'} · ${appName} · Office One` }, [name, appName])
  const commit = () => { const n = name.trim() || 'Untitled'; setName(n); if (n !== live?.name) void renameFile(file.id, n) }
  return (
    <div className="h-14 shrink-0 bg-surface flex items-center gap-1.5 pl-3 pr-4">
      <Link to="/" aria-label="Back to home" title="Back to home" className="h-9 w-9 flex items-center justify-center rounded-lg text-muted hover:bg-panel2 hover:text-ink"><ChevronLeft size={20} /></Link>
      <span className="flex items-center gap-1.5 text-[13px] font-medium pr-1" style={{ color }}><span aria-hidden className="w-2 h-2 rounded-full" style={{ background: color }} />{appName}</span>
      <Menu trigger="File" items={fileMenu} />
      <span aria-hidden className="w-px h-5 bg-line mx-1" />
      <input id="doc-title" aria-label="Document title" value={name} onChange={e => setName(e.target.value)} onBlur={commit}
        onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
        style={{ width: `${Math.min(44, Math.max(10, name.length + 2))}ch` }}
        className="h-9 px-2 -mx-1 rounded-lg border border-transparent hover:border-line focus:border-accent focus:bg-panel bg-transparent font-display text-[19px] min-w-0 max-w-[46vw] outline-none" />
      <IconButton label={live?.favourite ? 'Remove from favourites' : 'Add to favourites'} active={!!live?.favourite} onClick={() => live && toggleFavourite(live)}><Star size={16} className={live?.favourite ? 'fill-current' : ''} /></IconButton>
      <div className="flex-1" />
      <SaveBadge state={save} />
      {children}
      <ThemeToggle />
    </div>
  )
}
