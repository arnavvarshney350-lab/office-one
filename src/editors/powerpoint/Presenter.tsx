import { useCallback, useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { Clock, Monitor, RotateCcw } from 'lucide-react'
import { SlideView } from './SlideView'
import { TransitionStage } from './TransitionPlayer'
import { SLIDE_H, SLIDE_W, type Deck } from './model'

interface Anim { id: number; from: number; to: number; reverse: boolean }
const mmss = (s: number) => `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`

/**
 * Full-screen slideshow.
 * The transition of the slide being entered plays going forward; going back plays the reverse of the
 * slide being left. Press V for presenter view (notes, next slide and a timer).
 */
export function Presenter({ deck, start, onExit }: { deck: Deck; start: number; onExit: () => void }) {
  const [cur, setCur] = useState(start)
  const [anim, setAnim] = useState<Anim | null>(null)
  const [view, setView] = useState<'audience' | 'presenter'>('audience')
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  const [elapsed, setElapsed] = useState(0)
  const root = useRef<HTMLDivElement>(null)
  const seq = useRef(0)
  const n = deck.slides.length

  const go = useCallback((delta: number, absolute?: number) => {
    const base = anim ? anim.to : cur // a key press during a transition finishes it instantly, then moves on
    const next = Math.max(0, Math.min(n - 1, absolute ?? base + delta))
    if (next === base) { if (anim) { setCur(base); setAnim(null) } return }
    const forward = next > base
    const t = forward ? deck.slides[next].transition : deck.slides[base].transition
    if (!t || t.type === 'none' || Math.abs(next - base) > 1) { setAnim(null); setCur(next); return }
    setCur(base); setAnim({ id: ++seq.current, from: base, to: next, reverse: !forward })
  }, [anim, cur, deck.slides, n])

  useEffect(() => {
    const el = root.current
    el?.focus()
    void el?.requestFullscreen?.().catch(() => { /* refused: the overlay still fills the window */ })
    const onFs = () => { if (!document.fullscreenElement) onExit() }
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight })
    document.addEventListener('fullscreenchange', onFs); window.addEventListener('resize', onResize)
    return () => { document.removeEventListener('fullscreenchange', onFs); window.removeEventListener('resize', onResize); if (document.fullscreenElement) void document.exitFullscreen() }
  }, [onExit])
  useEffect(() => { const t = setInterval(() => setElapsed(e => e + 1), 1000); return () => clearInterval(t) }, [])

  const presenter = view === 'presenter'
  const mainBox = presenter ? { w: size.w * 0.62 - 32, h: size.h - 120 } : size
  const scale = Math.min(mainBox.w / SLIDE_W, mainBox.h / SLIDE_H)
  const shown = anim ? anim.to : cur
  const finish = useCallback(() => setAnim(a => { if (a) setCur(a.to); return null }), [])
  const nextSlide = deck.slides[Math.min(n - 1, shown + 1)]

  const stage = anim
    ? <TransitionStage key={anim.id} from={deck.slides[anim.from]} to={deck.slides[anim.to]} transition={(anim.reverse ? deck.slides[anim.from] : deck.slides[anim.to]).transition!} reverse={anim.reverse} scale={scale} deckTheme={deck.theme} onDone={finish} />
    : <SlideView slide={deck.slides[cur]} deckTheme={deck.theme} scale={scale} />

  return createPortal(
    <div ref={root} tabIndex={0} role="dialog" aria-label="Slide show" data-testid="presenter" data-view={view}
      className={`fixed inset-0 z-[200] outline-none select-none ${presenter ? 'bg-neutral-900 text-white' : 'bg-black cursor-pointer'} flex items-center justify-center`}
      onClick={() => { if (!presenter) go(1) }}
      onKeyDown={e => {
        if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n', 'N'].includes(e.key)) { e.preventDefault(); go(1) }
        else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(e.key)) { e.preventDefault(); go(-1) }
        else if (e.key === 'v' || e.key === 'V') { e.preventDefault(); setView(v => (v === 'audience' ? 'presenter' : 'audience')) }
        else if (e.key === 'Escape') onExit()
        else if (e.key === 'Home') { e.preventDefault(); go(0, 0) } else if (e.key === 'End') { e.preventDefault(); go(0, n - 1) }
      }}>
      {!presenter ? (
        <>
          {stage}
          <div className="absolute bottom-3 right-4 text-white/60 text-xs" data-testid="present-counter">{shown + 1} / {n}</div>
          <button onClick={e => { e.stopPropagation(); setView('presenter') }} className="absolute top-3 left-4 text-white/60 hover:text-white text-xs border border-white/30 rounded px-2 py-1 flex items-center gap-1"><Monitor size={12} />Presenter view (V)</button>
          <button onClick={e => { e.stopPropagation(); onExit() }} className="absolute top-3 right-4 text-white/60 hover:text-white text-xs border border-white/30 rounded px-2 py-1">Exit (Esc)</button>
        </>
      ) : (
        <div className="w-full h-full flex gap-6 p-6" data-testid="presenter-view">
          <div className="flex flex-col items-center justify-center" style={{ width: '62%' }}>
            <div className="shadow-2xl">{stage}</div>
            <div className="mt-4 flex items-center gap-3 text-sm text-white/80">
              <button onClick={() => go(-1)} disabled={shown === 0} aria-label="Previous slide" className="px-3 h-8 rounded border border-white/30 hover:bg-white/10 disabled:opacity-40">‹ Prev</button>
              <span data-testid="present-counter" className="tabular-nums">{shown + 1} / {n}</span>
              <button onClick={() => go(1)} disabled={shown === n - 1} aria-label="Next slide" className="px-3 h-8 rounded border border-white/30 hover:bg-white/10 disabled:opacity-40">Next ›</button>
            </div>
          </div>
          <div className="flex-1 min-w-0 flex flex-col gap-4">
            <div className="flex items-center gap-3 text-white/80">
              <Clock size={16} /><span className="text-2xl tabular-nums font-semibold" data-testid="present-timer">{mmss(elapsed)}</span>
              <button onClick={() => setElapsed(0)} aria-label="Reset timer" className="ml-1 p-1.5 rounded hover:bg-white/10"><RotateCcw size={14} /></button>
              <div className="flex-1" />
              <button onClick={() => setView('audience')} className="text-xs border border-white/30 rounded px-2 py-1 hover:bg-white/10">Audience view (V)</button>
              <button onClick={onExit} className="text-xs border border-white/30 rounded px-2 py-1 hover:bg-white/10">Exit</button>
            </div>
            <div>
              <div className="text-xs uppercase tracking-wide text-white/50 mb-1">Next</div>
              {shown < n - 1
                ? <div className="inline-block border border-white/20"><SlideView slide={nextSlide} deckTheme={deck.theme} scale={Math.min(0.32, (size.w * 0.3) / SLIDE_W)} /></div>
                : <div className="text-white/50 text-sm">End of presentation</div>}
            </div>
            <div className="flex-1 min-h-0 flex flex-col">
              <div className="text-xs uppercase tracking-wide text-white/50 mb-1">Notes</div>
              <div data-testid="present-notes" className="flex-1 overflow-y-auto rounded-lg bg-white/5 border border-white/10 p-3 text-[17px] leading-relaxed whitespace-pre-wrap">
                {deck.slides[shown].notes?.trim() || <span className="text-white/40">No notes for this slide.</span>}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>, document.body)
}
