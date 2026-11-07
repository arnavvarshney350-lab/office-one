import { useMemo, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { Plus } from 'lucide-react'
import { TopBar } from '../../components/TopBar'
import { FileIcon } from '../../design-system'
import { useFiles } from '../../hooks/useFiles'
import { createFile, fullName, routeFor } from '../../file-system/fileService'
import { STARTER_TEMPLATES } from '../../file-system/starterTemplates'
import { formatDate } from '../../utilities/id'
import { layoutObjects, newDeck, THEMES, type Slide, type ThemeMood } from '../../editors/powerpoint/model'
import { MoodFilter } from '../../editors/powerpoint/ThemeGallery'
import { ThemeThumb } from '../../editors/powerpoint/ThemeThumb'
import { uid } from '../../utilities/id'
import type { FileKind } from '../../types'

const META: Record<string, { name: string; color: string; blank: string; noun: string; lead: string }> = {
  docx: { name: 'Word', color: 'var(--c-word)', blank: 'document', noun: 'documents', lead: 'Write, format and export documents.' },
  xlsx: { name: 'Excel', color: 'var(--c-excel)', blank: 'workbook', noun: 'workbooks', lead: 'Spreadsheets with formulas, charts and sorting.' },
  pptx: { name: 'PowerPoint', color: 'var(--c-ppt)', blank: 'presentation', noun: 'presentations', lead: 'Design slides and present them full screen.' },
}

export function AppLanding({ kind }: { kind: FileKind }) {
  const nav = useNavigate()
  const files = useFiles()
  const m = META[kind]
  const mine = (files ?? []).filter(f => !f.trashed && f.kind === kind).sort((a, b) => b.updatedAt - a.updatedAt)
  const [mood, setMood] = useState<ThemeMood | 'all'>('all')
  const themes = THEMES.filter(t => mood === 'all' || t.mood === mood)
  // One title slide per theme, drawn by the real renderer.
  const previews = useMemo(() => new Map(THEMES.map(t => [t.id, { id: uid(), bg: t.bg, layout: 'title', objects: layoutObjects('title', t) } as Slide])), [])
  const starters = STARTER_TEMPLATES.filter(t => t.kind === kind)
  async function create(themeId?: string) {
    const f = await createFile(kind, themeId ? { content: newDeck(themeId), name: `${THEMES.find(t => t.id === themeId)?.name} presentation` } : {})
    nav(routeFor(f)!)
  }
  const tileArt = 'h-[95px] rounded-xl flex items-center justify-center transition-all duration-150 group-hover:-translate-y-0.5 group-hover:shadow-soft'
  return (
    <div className="h-full flex flex-col min-h-0">
      <TopBar contained />
      <div className="flex-1 overflow-auto">
        <div className="max-w-[1080px] mx-auto px-10 pt-6 pb-16">
          <div className="flex items-center gap-3"><span aria-hidden className="w-2.5 h-2.5 rounded-full" style={{ background: m.color }} /><h1 className="font-display text-[40px] leading-none">{m.name}</h1></div>
          <p className="text-muted mt-2.5 text-[15px]">{m.lead}</p>

          <section className="mt-9" aria-label={`Start a ${m.blank}`}>
            {kind === 'pptx' && <div className="mb-5 flex items-center gap-3"><h2 className="text-[15px] font-semibold">Themes</h2><MoodFilter value={mood} onChange={setMood} /></div>}
            <div className="flex flex-wrap items-start gap-x-6 gap-y-7">
              <button onClick={() => create()} aria-label={`New blank ${m.blank}`} className="group text-left w-[168px] focus-visible:outline-offset-4">
                <div className={`${tileArt} bg-canvas`}><Plus size={30} strokeWidth={1.5} style={{ color: m.color }} /></div>
                <div className="mt-2.5 text-[13.5px] font-medium">Blank {m.blank}</div>
              </button>
              {kind === 'pptx' && themes.map(t => (
                <button key={t.id} onClick={() => create(t.id)} aria-label={`Template ${t.name}`} title={t.blurb} className="group text-left w-[168px] focus-visible:outline-offset-4">
                  <div className="rounded-xl overflow-hidden transition-all duration-150 motion-reduce:transition-none group-hover:-translate-y-0.5 motion-reduce:group-hover:translate-y-0 group-hover:shadow-soft" style={{ boxShadow: 'inset 0 0 0 1px var(--border)' }}>
                    <ThemeThumb source={previews.get(t.id)!} themeId={t.id} width={168} />
                  </div>
                  <div className="mt-2.5 text-[13.5px] font-medium">{t.name}</div>
                  <div className="text-[12px] text-muted leading-snug truncate">{t.blurb}</div>
                </button>
              ))}
              {starters.map(t => (
                <button key={t.id} onClick={async () => { const f = await createFile(t.kind, { name: t.name, content: t.build() }); nav(routeFor(f)!) }} aria-label={`Template ${t.name}`} className="group text-left w-[168px] focus-visible:outline-offset-4">
                  <div className={`${tileArt} bg-canvas`}><FileIcon kind={t.kind} size={30} /></div>
                  <div className="mt-2.5 text-[13.5px] font-medium">{t.name}</div>
                  <div className="text-[12px] text-muted leading-snug">{t.description}</div>
                </button>
              ))}
            </div>
          </section>

          <section className="mt-12" aria-labelledby="mine">
            <h2 id="mine" className="text-[15px] font-semibold">Your {m.noun}</h2>
            {mine.length === 0 ? (
              <p className="mt-3 text-[13.5px] text-muted">No {m.noun} yet. Start with a blank one above.</p>
            ) : (
              <ul className="mt-3 divide-y divide-line/70">
                {mine.map(f => (
                  <li key={f.id}>
                    <button onClick={() => nav(routeFor(f)!)} className="w-full flex items-center gap-3 px-2 -mx-2 h-12 text-left rounded-lg hover:bg-panel transition-colors">
                      <FileIcon kind={f.kind} size={18} /><span className="flex-1 truncate text-[13.5px]">{fullName(f)}</span><span className="text-[12px] text-muted tabular">{formatDate(f.updatedAt)}</span>
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </div>
      </div>
    </div>
  )
}
