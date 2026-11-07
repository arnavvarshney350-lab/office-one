import { lazy, Suspense, useMemo, useRef, useState } from 'react'
import { useSearchParams } from 'react-router-dom'
import { ChevronRight, Clock, Download, FolderOpen, FolderPlus, MoreHorizontal, Plus, Star, Trash2, Upload, ArrowDownUp, LayoutGrid, List, Search } from 'lucide-react'
import { TopBar } from '../../components/TopBar'
import { Button, ConfirmModal, FileIcon, IconButton, Menu, Modal, PromptModal, cx, type MenuEntry } from '../../design-system'
import { useFiles } from '../../hooks/useFiles'
import { useOpenFile } from '../../hooks/useOpenFile'
import { createFile, deleteForever, duplicateFile, emptyTrash, fullName, renameFile, restoreRecursive, routeFor, toggleFavourite, trashRecursive } from '../../file-system/fileService'
import { importFile, IMPORT_ACCEPT } from '../../services/importers'
import { download, safeName } from '../../services/download'
import { db } from '../../storage/db'
import { useToast } from '../../contexts/ToastContext'
import { formatBytes, formatDate } from '../../utilities/id'
import { KIND_LABEL } from '../../file-system/templates'
import type { FileKind, OFile } from '../../types'

const FilePreview = lazy(() => import('../../components/FilePreview'))

type View = 'recent' | 'mine' | 'favourites' | 'trash'
type Sort = 'name' | 'updated' | 'size' | 'type'

export function FilesPage() {
  const [params, setParams] = useSearchParams()
  const view = (params.get('view') as View) || 'mine'
  const folderId = params.get('folder')
  const q = params.get('q') ?? ''
  const files = useFiles()
  const open = useOpenFile()
  const toast = useToast()
  const [sort, setSort] = useState<Sort>('updated')
  const [dir, setDir] = useState<1 | -1>(-1)
  const [prompt, setPrompt] = useState<null | { kind: 'folder' } | { kind: 'rename'; f: OFile }>(null)
  const [moving, setMoving] = useState<OFile | null>(null)
  const input = useRef<HTMLInputElement>(null)

  const setView = (v: View, folder?: string | null) => {
    const n = new URLSearchParams(); n.set('view', v); if (folder) n.set('folder', folder); if (q) n.set('q', q); setParams(n)
  }
  const setQ = (v: string) => { const n = new URLSearchParams(params); if (v) n.set('q', v); else n.delete('q'); setParams(n, { replace: true }) }

  const byId = useMemo(() => new Map((files ?? []).map(f => [f.id, f])), [files])
  const rows = useMemo(() => {
    let list = (files ?? []).filter(f => (view === 'trash' ? f.trashed === 1 : f.trashed === 0))
    if (view === 'recent') list = list.filter(f => f.kind !== 'folder' && f.openedAt).sort((a, b) => (b.openedAt ?? 0) - (a.openedAt ?? 0)).slice(0, 50)
    if (view === 'favourites') list = list.filter(f => f.favourite)
    if (view === 'mine' && !q) list = list.filter(f => f.parentId === folderId)
    if (view === 'trash') list = list.filter(f => !f.parentId || !byId.get(f.parentId)?.trashed)
    if (q) list = list.filter(f => fullName(f).toLowerCase().includes(q.toLowerCase()))
    if (view === 'recent' && !q) return list
    const cmp: Record<Sort, (a: OFile, b: OFile) => number> = {
      name: (a, b) => a.name.localeCompare(b.name), updated: (a, b) => a.updatedAt - b.updatedAt,
      size: (a, b) => a.size - b.size, type: (a, b) => a.kind.localeCompare(b.kind),
    }
    return [...list].sort((a, b) => (a.kind === 'folder' ? 0 : 1) - (b.kind === 'folder' ? 0 : 1) || cmp[sort](a, b) * dir)
  }, [files, view, folderId, q, sort, dir, byId])

  const crumbs: OFile[] = []
  for (let c = folderId ? byId.get(folderId) : undefined; c; c = c.parentId ? byId.get(c.parentId) : undefined) crumbs.unshift(c)

  async function newOf(kind: FileKind) { const f = await createFile(kind, { parentId: view === 'mine' ? folderId : null }); const r = routeFor(f); if (r) open(f) }
  async function onUpload(list: FileList | null) {
    if (!list) return
    for (const file of Array.from(list)) { try { const f = await importFile(file, view === 'mine' ? folderId : null); toast(`Imported ${fullName(f)}`) } catch (e) { toast((e as Error).message, 'error') } }
  }
  function openRow(f: OFile) {
    if (f.trashed) return
    if (f.kind === 'folder') { setView('mine', f.id); return }
    void db.files.update(f.id, { openedAt: Date.now() }); open(f)
  }
  async function exportFile(f: OFile) {
    const name = safeName(f.name)
    if (f.blob) return download(f.blob, name + '.' + f.kind)
    if (f.kind === 'docx') { const [{ exportDocx }, { readDoc }] = await Promise.all([import('../../editors/word/io'), import('../../editors/word/page')]); const d = readDoc(f.content); return download(await exportDocx(f.name, d.html, d.page), name + '.docx') }
    if (f.kind === 'xlsx') { const { exportXlsx } = await import('../../editors/excel/io'); return download(exportXlsx(f.content as never), name + '.xlsx') }
    if (f.kind === 'pptx') { const { exportPptx } = await import('../../editors/powerpoint/io'); return exportPptx(f.name, f.content as never) }
  }

  const NAVS: [View, string, typeof Clock][] = [['recent', 'Recent', Clock], ['mine', 'My files', FolderOpen], ['favourites', 'Favourites', Star], ['trash', 'Trash', Trash2]]
  const title = { recent: 'Recent', mine: 'My files', favourites: 'Favourites', trash: 'Trash' }[view]
  const count = (v: View) => (files ?? []).filter(f => (v === 'trash' ? f.trashed === 1 && (!f.parentId || !byId.get(f.parentId)?.trashed) : f.trashed === 0 && (v === 'favourites' ? f.favourite === 1 : v === 'recent' ? f.kind !== 'folder' && !!f.openedAt : f.parentId === null))).length
  const [layout, setLayoutState] = useState<'grid' | 'list'>(() => { try { return localStorage.getItem('oo-drive-layout') === 'list' ? 'list' : 'grid' } catch { return 'grid' } })
  const setLayout = (l: 'grid' | 'list') => { setLayoutState(l); try { localStorage.setItem('oo-drive-layout', l) } catch { /* private mode */ } }
  const [confirm, setConfirm] = useState<null | { kind: 'forever'; f: OFile } | { kind: 'empty' }>(null)

  const trash = async (f: OFile) => {
    await trashRecursive(f)
    toast(`Moved “${f.name}” to Trash`, 'info', { label: 'Undo', onClick: () => void restoreRecursive(f) })
  }
  const menuFor = (f: OFile): MenuEntry[] => f.trashed ? [
    { label: 'Restore', onClick: () => restoreRecursive(f) },
    { label: 'Delete forever…', danger: true, onClick: () => setConfirm({ kind: 'forever', f }) },
  ] : [
    { label: 'Open', onClick: () => openRow(f) },
    { label: f.favourite ? 'Remove from favourites' : 'Add to favourites', onClick: () => toggleFavourite(f), disabled: f.kind === 'folder' },
    { label: 'Rename', onClick: () => setPrompt({ kind: 'rename', f }) },
    { label: 'Duplicate', onClick: async () => { await duplicateFile(f); toast('Duplicated') } },
    { label: 'Move to…', onClick: () => setMoving(f) },
    { label: 'Download', icon: <Download size={14} />, onClick: () => exportFile(f), disabled: f.kind === 'folder' },
    'sep',
    { label: 'Move to trash', danger: true, onClick: () => trash(f) },
  ]
  const empty = q ? `Nothing matches “${q}”. Try a shorter search.`
    : view === 'trash' ? 'Trash is empty. Deleted files wait here until you empty it.'
      : view === 'favourites' ? 'Star a file and it will be kept here for quick access.'
        : view === 'recent' ? 'Files you open will show up here.' : 'This folder is empty. Create a file, or drop one here.'

  return (
    <div className="h-full flex flex-col min-h-0">
      <TopBar />
      <div className="flex-1 min-h-0 flex">
        <aside aria-label="Drive sections" className="w-52 shrink-0 pl-6 pr-3 pt-2 space-y-0.5 hidden sm:block">
          {NAVS.map(([v, label, Icon]) => (
            <button key={v} onClick={() => setView(v)} aria-current={view === v ? 'page' : undefined}
              className={cx('w-full flex items-center gap-2.5 px-3 h-9 rounded-lg text-[13.5px] transition-colors', view === v ? 'bg-accent-soft text-accent font-medium' : 'text-muted hover:bg-panel2 hover:text-ink')}>
              <Icon size={16} /><span className="flex-1 text-left">{label}</span><span className="text-[12px] tabular opacity-70">{count(v) || ''}</span>
            </button>
          ))}
        </aside>
        <section className="flex-1 min-w-0 flex flex-col pr-8">
          <div className="flex items-center gap-2 h-12 flex-wrap shrink-0">
            <h1 className="flex items-center gap-1 font-display text-[24px] min-w-0">
              <button onClick={() => view === 'mine' ? setView('mine') : undefined} className="hover:underline">{title}</button>
              {view === 'mine' && crumbs.map(c => <span key={c.id} className="flex items-center gap-1"><ChevronRight size={16} className="text-muted" /><button onClick={() => setView('mine', c.id)} className="hover:underline truncate">{c.name}</button></span>)}
            </h1>
            <div className="flex-1" />
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-muted" />
              <input aria-label="Filter files" value={q} onChange={e => setQ(e.target.value)} placeholder="Filter" className="h-8 w-40 pl-8 pr-2 rounded-lg bg-panel border border-transparent focus:border-accent focus:bg-surface outline-none" />
            </div>
            <label className="flex items-center gap-1 text-[13px] text-muted"><ArrowDownUp size={14} />
              <select aria-label="Sort by" value={sort} onChange={e => setSort(e.target.value as Sort)} className="h-8 rounded-lg bg-panel px-1.5 border border-transparent">
                <option value="updated">Modified</option><option value="name">Name</option><option value="size">Size</option><option value="type">Type</option>
              </select>
            </label>
            <Button variant="ghost" onClick={() => setDir(d => (d === 1 ? -1 : 1))} aria-label="Toggle sort direction">{dir === 1 ? 'Asc' : 'Desc'}</Button>
            <div role="group" aria-label="Layout" className="flex rounded-lg bg-panel p-0.5">
              <IconButton label="Grid view" active={layout === 'grid'} onClick={() => setLayout('grid')} className="h-7 min-w-7"><LayoutGrid size={15} /></IconButton>
              <IconButton label="List view" active={layout === 'list'} onClick={() => setLayout('list')} className="h-7 min-w-7"><List size={15} /></IconButton>
            </div>
            {view === 'trash' ? <Button variant="danger" disabled={!rows.length} onClick={() => setConfirm({ kind: 'empty' })}>Empty trash</Button> : (<>
              <Button onClick={() => setPrompt({ kind: 'folder' })}><FolderPlus size={15} />New folder</Button>
              <Menu trigger={<><Plus size={15} />New</>} align="right" items={[
                { label: 'Document', onClick: () => newOf('docx') }, { label: 'Spreadsheet', onClick: () => newOf('xlsx') }, { label: 'Presentation', onClick: () => newOf('pptx') }]} />
              <Button variant="primary" onClick={() => input.current?.click()}><Upload size={15} />Upload</Button>
              <input ref={input} type="file" multiple hidden accept={IMPORT_ACCEPT} data-testid="files-upload" onChange={e => { void onUpload(e.target.files); e.target.value = '' }} />
            </>)}
          </div>

          <div className="flex-1 overflow-auto pb-10 -mr-8 pr-8">
            {files && rows.length === 0 && (
              <div className="mt-6 rounded-2xl border border-dashed border-line px-6 py-14 text-center">
                <p className="font-display text-[22px]">{q ? 'No matches' : view === 'trash' ? 'Trash is empty' : 'Nothing here yet'}</p>
                <p className="text-muted text-[13.5px] mt-1.5 max-w-sm mx-auto">{empty}</p>
              </div>
            )}
            {layout === 'grid' ? (
              <div className="mt-3 flex flex-wrap gap-x-6 gap-y-7">
                {rows.map(f => (
                  <div key={f.id} data-testid="file-row" className="group relative w-[200px]" onDoubleClick={() => openRow(f)}>
                    <button onClick={() => openRow(f)} aria-label={`Open ${fullName(f)}`} className="block w-full text-left focus-visible:outline-offset-4">
                      <div className="rounded-xl overflow-hidden ring-1 ring-transparent group-hover:ring-line transition-shadow">
                        <Suspense fallback={<div className="bg-canvas" style={{ width: 200, height: 150 }} />}><FilePreview file={f} width={200} /></Suspense>
                      </div>
                    </button>
                    <div className="mt-2 flex items-center gap-1.5 min-w-0 pr-7">
                      <FileIcon kind={f.kind} size={15} /><span className="text-[13.5px] font-medium truncate">{fullName(f)}</span>
                      {f.favourite === 1 && <Star size={13} className="text-amber-500 fill-amber-500 shrink-0" aria-label="Favourite" />}
                    </div>
                    <div className="text-[12px] text-muted tabular">{formatDate(f.updatedAt)}{f.kind !== 'folder' && ` · ${formatBytes(f.size)}`}</div>
                    <div className="absolute right-0 bottom-0"><Menu title="More actions" align="right" trigger={<MoreHorizontal size={16} />} items={menuFor(f)} /></div>
                  </div>
                ))}
              </div>
            ) : (
              <div role="table" aria-label="Files" className="mt-1">
                <div role="row" className="grid grid-cols-[1fr_120px_90px_110px_40px] px-3 h-8 items-center text-[12px] text-muted border-b border-line sticky top-0 bg-surface z-10">
                  <div>Name</div><div>Modified</div><div>Size</div><div>Type</div><div />
                </div>
                {rows.map(f => (
                  <div key={f.id} role="row" onDoubleClick={() => openRow(f)} className="grid grid-cols-[1fr_120px_90px_110px_40px] px-3 h-12 items-center border-b border-line/70 hover:bg-panel rounded-md text-[13px]" data-testid="file-row">
                    <button onClick={() => openRow(f)} className="flex items-center gap-3 min-w-0 text-left">
                      <FileIcon kind={f.kind} size={18} />
                      <span className="truncate">{fullName(f)}</span>{f.favourite === 1 && <Star size={13} className="text-amber-500 fill-amber-500 shrink-0" aria-label="Favourite" />}
                    </button>
                    <div className="text-muted tabular">{formatDate(f.updatedAt)}</div>
                    <div className="text-muted tabular">{f.kind === 'folder' ? '—' : formatBytes(f.size)}</div>
                    <div className="text-muted">{KIND_LABEL[f.kind]}</div>
                    <Menu title="More actions" align="right" trigger={<MoreHorizontal size={16} />} items={menuFor(f)} />
                  </div>
                ))}
              </div>
            )}
          </div>
        </section>
      </div>
      {prompt?.kind === 'folder' && <PromptModal title="New folder" label="Folder name" initial="New folder" confirm="Create" onClose={() => setPrompt(null)} onSubmit={v => { void createFile('folder', { name: v.trim() || 'New folder', parentId: view === 'mine' ? folderId : null }) }} />}
      {prompt?.kind === 'rename' && <PromptModal title="Rename" label="Name" initial={prompt.f.name} confirm="Rename" onClose={() => setPrompt(null)} onSubmit={v => { void renameFile(prompt.f.id, v) }} />}
      {moving && <MoveDialog file={moving} folders={(files ?? []).filter(f => f.kind === 'folder' && !f.trashed && f.id !== moving.id)} onClose={() => setMoving(null)} />}
      {confirm?.kind === 'forever' && <ConfirmModal title="Delete forever?" confirm="Delete forever" message={<>“{confirm.f.name}” will be permanently deleted. This can’t be undone.</>}
        onConfirm={async () => { await deleteForever(confirm.f); toast('Deleted permanently') }} onClose={() => setConfirm(null)} />}
      {confirm?.kind === 'empty' && <ConfirmModal title="Empty Trash?" confirm="Empty Trash" message="Everything in Trash will be permanently deleted. This can’t be undone."
        onConfirm={async () => { await emptyTrash(); toast('Trash emptied') }} onClose={() => setConfirm(null)} />}
    </div>
  )
}

function MoveDialog({ file, folders, onClose }: { file: OFile; folders: OFile[]; onClose: () => void }) {
  const [target, setTarget] = useState<string>(file.parentId ?? '')
  const toast = useToast()
  // prevent moving a folder into its own descendant
  const forbidden = new Set<string>()
  if (file.kind === 'folder') { const walk = (id: string) => { forbidden.add(id); folders.filter(f => f.parentId === id).forEach(c => walk(c.id)) }; walk(file.id) }
  return (
    <Modal title={`Move “${file.name}”`} onClose={onClose} footer={<><Button onClick={onClose}>Cancel</Button><Button variant="primary" onClick={async () => { await db.files.update(file.id, { parentId: target || null }); toast('Moved'); onClose() }}>Move here</Button></>}>
      <select aria-label="Destination folder" value={target} onChange={e => setTarget(e.target.value)} className="w-full h-9 rounded-md border border-line bg-panel px-2">
        <option value="">My files (top level)</option>
        {folders.filter(f => !forbidden.has(f.id)).map(f => <option key={f.id} value={f.id}>{f.name}</option>)}
      </select>
    </Modal>
  )
}
