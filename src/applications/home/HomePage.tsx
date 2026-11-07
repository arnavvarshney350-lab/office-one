import { lazy, Suspense, useRef } from 'react'
import { useNavigate } from 'react-router-dom'
import { FileText, Presentation, Sheet, Star, Upload } from 'lucide-react'
import { TopBar } from '../../components/TopBar'
import { FileIcon, cx } from '../../design-system'
import { useFiles } from '../../hooks/useFiles'
import { useOpenFile } from '../../hooks/useOpenFile'
import { createFile, fullName, routeFor } from '../../file-system/fileService'
import { STARTER_TEMPLATES } from '../../file-system/starterTemplates'
import { importFile, IMPORT_ACCEPT } from '../../services/importers'
import { useToast } from '../../contexts/ToastContext'
import { formatDate } from '../../utilities/id'
import { KIND_LABEL } from '../../file-system/templates'
import type { FileKind, OFile } from '../../types'

const FilePreview = lazy(() => import('../../components/FilePreview'))

const greeting = () => { const h = new Date().getHours(); return h < 5 ? 'Working late' : h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening' }

/* The start tiles are drawn as the thing they create: a page, a grid, a slide. */
function PageArt() {
  return (
    <div className="bg-page shadow-soft rounded-[2px] w-[92px] h-[118px] px-3 pt-3.5 flex flex-col gap-[5px]">
      <div className="h-[6px] w-[60%] rounded-full bg-ink/80" />
      {[100, 92, 100, 70, 96, 84].map((w, i) => <div key={i} className="h-[3px] rounded-full bg-ink/15" style={{ width: `${w}%` }} />)}
    </div>
  )
}
function GridArt() {
  return (
    <div className="bg-page shadow-soft rounded-[3px] w-[128px] h-[96px] overflow-hidden grid grid-cols-4 grid-rows-5">
      {Array.from({ length: 20 }, (_, i) => (
        <div key={i} className="border-r border-b border-line" style={{ background: i < 4 ? 'color-mix(in srgb, var(--c-excel) 16%, transparent)' : i === 9 || i === 14 ? 'color-mix(in srgb, var(--c-excel) 8%, transparent)' : undefined }}>
          {i >= 4 && i % 4 !== 0 && <div className="h-[3px] mt-[9px] mr-1.5 ml-auto rounded-full bg-ink/20" style={{ width: `${40 + ((i * 17) % 40)}%` }} />}
        </div>
      ))}
    </div>
  )
}
function SlideArt() {
  return (
    <div className="shadow-soft rounded-[3px] w-[136px] h-[76px] overflow-hidden relative" style={{ background: 'var(--page)' }}>
      <div className="absolute left-3.5 top-4 h-[7px] w-[58%] rounded-full bg-ink/85" />
      <div className="absolute left-3.5 top-[30px] h-[3px] w-[40%] rounded-full bg-ink/20" />
      <div className="absolute left-3.5 top-[38px] h-[3px] w-[32%] rounded-full bg-ink/20" />
      <div className="absolute right-3 bottom-3 w-10 h-10 rounded-full" style={{ background: 'var(--c-ppt)' }} />
      <div className="absolute right-8 bottom-6 w-7 h-7 rounded-[4px]" style={{ background: 'color-mix(in srgb, var(--c-ppt) 35%, transparent)' }} />
    </div>
  )
}

function StartTile({ label, hint, color, onClick, children }: { label: string; hint: string; color: string; onClick: () => void; children: React.ReactNode }) {
  return (
    <button onClick={onClick} aria-label={label} className="group text-left w-[212px] focus-visible:outline-offset-4">
      <div className="h-[156px] rounded-2xl bg-canvas flex items-center justify-center transition-all duration-150 group-hover:-translate-y-0.5 group-hover:shadow-soft group-active:translate-y-0">{children}</div>
      <div className="mt-3 flex items-center gap-2 text-[14px] font-medium"><span aria-hidden className="w-2 h-2 rounded-full" style={{ background: color }} />{label}</div>
      <div className="text-[12.5px] text-muted mt-0.5">{hint}</div>
    </button>
  )
}

function FileTile({ f, i, onOpen }: { f: OFile; i: number; onOpen: () => void }) {
  return (
    <button onClick={onOpen} data-testid="recent-tile" aria-label={`Open ${fullName(f)}`} style={{ ['--i' as string]: i }}
      className="settle group text-left w-[212px] focus-visible:outline-offset-4">
      <div className="rounded-xl overflow-hidden ring-1 ring-transparent group-hover:ring-line transition-shadow">
        <Suspense fallback={<div className="bg-canvas rounded-lg" style={{ width: 212, height: 159 }} />}><FilePreview file={f} width={212} /></Suspense>
      </div>
      <div className="mt-2.5 flex items-center gap-1.5 min-w-0">
        <FileIcon kind={f.kind} size={15} />
        <span className="text-[13.5px] font-medium truncate flex-1">{f.name}</span>
        {f.favourite === 1 && <Star size={13} className="text-amber-500 fill-amber-500 shrink-0" aria-label="Favourite" />}
      </div>
      <div className="text-[12px] text-muted mt-0.5 tabular">Edited {formatDate(f.updatedAt)} · {KIND_LABEL[f.kind]}</div>
    </button>
  )
}

export function HomePage() {
  const nav = useNavigate()
  const files = useFiles()
  const open = useOpenFile()
  const toast = useToast()
  const input = useRef<HTMLInputElement>(null)
  const live = (files ?? []).filter(f => !f.trashed && f.kind !== 'folder')
  const pinned = live.filter(f => f.favourite)
  const recent = [...live].filter(f => f.openedAt && !f.favourite).sort((a, b) => (b.openedAt ?? 0) - (a.openedAt ?? 0)).slice(0, 8)

  async function create(kind: FileKind) { const f = await createFile(kind); nav(routeFor(f)!) }
  async function onUpload(list: FileList | null) {
    if (!list) return
    const picked = Array.from(list) // the input is reset right after this call, which empties the live FileList
    for (const file of picked) {
      try { const f = await importFile(file); toast(`Imported ${fullName(f)}`); if (picked.length === 1) open(f) } catch (e) { toast((e as Error).message, 'error') }
    }
  }

  return (
    <div className="h-full flex flex-col min-h-0">
      <TopBar contained />
      <div className="flex-1 overflow-auto">
        <div className="max-w-[1080px] mx-auto px-10 pt-6 pb-16">
          <h1 className="font-display text-[44px] leading-[1.05]">{greeting()}</h1>
          <p className="text-muted mt-2 text-[15px]">Start something new, or pick up where you left off.</p>

          <div className="mt-9 flex flex-wrap gap-x-7 gap-y-8">
            <StartTile label="New document" hint="Word" color="var(--c-word)" onClick={() => create('docx')}><PageArt /></StartTile>
            <StartTile label="New spreadsheet" hint="Excel" color="var(--c-excel)" onClick={() => create('xlsx')}><GridArt /></StartTile>
            <StartTile label="New presentation" hint="PowerPoint" color="var(--c-ppt)" onClick={() => create('pptx')}><SlideArt /></StartTile>
            <button onClick={() => input.current?.click()} className="group text-left w-[212px] focus-visible:outline-offset-4">
              <div className="h-[156px] rounded-2xl border border-dashed border-line flex flex-col items-center justify-center gap-2 text-muted transition-colors group-hover:border-muted group-hover:text-ink">
                <Upload size={22} strokeWidth={1.75} /><span className="text-[12.5px]">or drop a file anywhere</span>
              </div>
              <div className="mt-3 text-[14px] font-medium">Open a file</div>
              <div className="text-[12.5px] text-muted mt-0.5">DOCX, XLSX, CSV, PPTX, PDF</div>
            </button>
            <input ref={input} type="file" multiple accept={IMPORT_ACCEPT} hidden data-testid="home-upload" onChange={e => { void onUpload(e.target.files); e.target.value = '' }} />
          </div>

          <section className="mt-12" aria-labelledby="tpl">
            <h2 id="tpl" className="text-[15px] font-semibold">Start from a template</h2>
            <div className="mt-3 flex flex-wrap gap-2">
              {STARTER_TEMPLATES.map(t => {
                const Icon = t.kind === 'docx' ? FileText : t.kind === 'xlsx' ? Sheet : Presentation
                return (
                  <button key={t.id} aria-label={`Template ${t.name}`} title={t.description}
                    onClick={async () => { const f = await createFile(t.kind, { name: t.name, content: t.build() }); nav(routeFor(f)!) }}
                    className="h-9 pl-3 pr-3.5 rounded-full border border-line hover:bg-panel2 text-[13px] inline-flex items-center gap-2 transition-colors">
                    <Icon size={15} style={{ color: t.kind === 'docx' ? 'var(--c-word)' : 'var(--c-excel)' }} />{t.name}
                  </button>
                )
              })}
            </div>
          </section>

          {pinned.length > 0 && (
            <section className="mt-12" aria-labelledby="pinned">
              <h2 id="pinned" className="text-[15px] font-semibold">Pinned</h2>
              <div className="mt-4 flex flex-wrap gap-x-7 gap-y-8">{pinned.map((f, i) => <FileTile key={f.id} f={f} i={i} onOpen={() => open(f)} />)}</div>
            </section>
          )}

          <section className="mt-12" aria-labelledby="recent">
            <div className="flex items-baseline justify-between">
              <h2 id="recent" className="text-[15px] font-semibold">Recent</h2>
              {live.length > 0 && <button onClick={() => nav('/files')} className="text-[13px] text-accent hover:underline">Open Drive</button>}
            </div>
            {files !== undefined && recent.length === 0 && pinned.length === 0 ? (
              <div className="mt-4 rounded-2xl border border-dashed border-line px-6 py-12 text-center">
                <p className="font-display text-[22px]">Nothing here yet</p>
                <p className="text-muted text-[13.5px] mt-1.5">Create a document above, or drop a file anywhere to add it. Everything you make stays on this device.</p>
              </div>
            ) : (
              <div className={cx('mt-4 flex flex-wrap gap-x-7 gap-y-8')}>{recent.map((f, i) => <FileTile key={f.id} f={f} i={i} onOpen={() => open(f)} />)}</div>
            )}
          </section>
        </div>
      </div>
    </div>
  )
}
