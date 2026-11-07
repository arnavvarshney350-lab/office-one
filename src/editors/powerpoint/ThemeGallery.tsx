import { useDeferredValue, useEffect, useRef, useState } from 'react'
import { Check } from 'lucide-react'
import { MOODS, THEMES, type Slide, type ThemeMood } from './model'
import { ThemeThumb } from './ThemeThumb'

const cx = (...c: (string | false | undefined)[]) => c.filter(Boolean).join(' ')

/** Filter chips shared by the gallery and the PowerPoint landing page. */
export function MoodFilter({ value, onChange }: { value: ThemeMood | 'all'; onChange: (v: ThemeMood | 'all') => void }) {
  return (
    <div role="group" aria-label="Filter themes" className="flex flex-wrap gap-1">
      {([{ id: 'all', name: 'All' }, ...MOODS] as { id: ThemeMood | 'all'; name: string }[]).map(m => (
        <button key={m.id} aria-pressed={value === m.id} onClick={() => onChange(m.id)}
          className={cx('h-6 px-2.5 rounded-full border text-[12px] transition-colors motion-reduce:transition-none', value === m.id ? 'bg-accent-soft border-accent text-accent font-semibold' : 'border-line bg-surface text-muted hover:text-ink hover:bg-panel2')}>
          {m.name}
        </button>
      ))}
    </div>
  )
}

/** Scrollable theme library: each tile is the current slide drawn in that theme by the real renderer. */
export function ThemeGallery({ slide, current, onPick }: { slide: Slide; current: string; onPick: (id: string) => void }) {
  const [mood, setMood] = useState<ThemeMood | 'all'>('all')
  const deferred = useDeferredValue(slide)
  const strip = useRef<HTMLDivElement>(null)
  const list = THEMES.filter(t => mood === 'all' || t.mood === mood)
  useEffect(() => { strip.current?.querySelector<HTMLElement>('[aria-pressed="true"]')?.scrollIntoView({ inline: 'center', block: 'nearest' }) }, [mood])
  return (
    <div className="flex flex-col gap-1.5 py-1.5 w-[min(760px,58vw)]">
      <MoodFilter value={mood} onChange={setMood} />
      <div ref={strip} role="group" aria-label="Themes" className="flex gap-2.5 overflow-x-auto pb-1 pr-2" data-testid="theme-gallery">
        {list.map(t => {
          const on = t.id === current
          return (
            <button key={t.id} title={`${t.name}: ${t.blurb}`} aria-label={`Theme ${t.name}`} aria-pressed={on} onClick={() => onPick(t.id)} className="shrink-0 text-left group rounded-md focus-visible:outline-offset-2">
              <div className={cx('relative rounded-md overflow-hidden border-2 transition-colors motion-reduce:transition-none', on ? 'border-accent' : 'border-line group-hover:border-muted')}>
                <ThemeThumb source={deferred} from={current} themeId={t.id} width={116} />
                {on && <span aria-hidden className="absolute right-1 top-1 w-4 h-4 rounded-full bg-accent text-accent-fg flex items-center justify-center"><Check size={11} strokeWidth={3} /></span>}
              </div>
              <div className={cx('mt-1 text-[11.5px] leading-none truncate w-[120px]', on ? 'text-accent font-semibold' : 'text-muted')}>{t.name}</div>
            </button>
          )
        })}
      </div>
    </div>
  )
}
