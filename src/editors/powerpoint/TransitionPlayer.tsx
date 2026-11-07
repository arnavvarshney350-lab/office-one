import { useLayoutEffect, useRef } from 'react'
import { SlideView } from './SlideView'
import { SLIDE_H, SLIDE_W, type Slide, type Transition, type TransitionDir } from './model'

const OPPOSITE: Record<TransitionDir, TransitionDir> = { left: 'right', right: 'left', up: 'down', down: 'up', horizontal: 'horizontal', vertical: 'vertical', in: 'out', out: 'in' }
const EASE = 'cubic-bezier(0.4, 0, 0.2, 1)'

/** Where the incoming slide starts, as a translate() in percent, for a "from <side>" direction. */
const offset = (d: TransitionDir): [number, number] => (d === 'left' ? [-100, 0] : d === 'right' ? [100, 0] : d === 'up' ? [0, -100] : [0, 100])
const tr = (x: number, y: number) => `translate(${x}%, ${y}%)`
const wipeFrom = (d: TransitionDir) => (d === 'left' ? 'inset(0 100% 0 0)' : d === 'right' ? 'inset(0 0 0 100%)' : d === 'up' ? 'inset(0 0 100% 0)' : 'inset(100% 0 0 0)')

/** Build the animations for one transition. `top` is the layer that must be painted above the other. */
export function buildTransition(t: Transition, reverse: boolean, from: HTMLElement, to: HTMLElement): { anims: Animation[]; toOnTop: boolean } {
  const dir = reverse ? OPPOSITE[t.direction] : t.direction
  const opts: KeyframeAnimationOptions = { duration: Math.max(1, t.duration), easing: EASE, fill: 'both' }
  const go = (el: HTMLElement, kf: Keyframe[], o: KeyframeAnimationOptions = opts) => el.animate(kf, o)
  switch (t.type) {
    case 'fade': return { anims: [go(to, [{ opacity: 0 }, { opacity: 1 }])], toOnTop: true }
    case 'dissolve': return { anims: [go(to, [{ opacity: 0, filter: 'blur(10px)' }, { opacity: 1, filter: 'blur(0px)' }], { ...opts, easing: 'steps(10, end)' })], toOnTop: true }
    case 'push': { const [x, y] = offset(dir); return { anims: [go(to, [{ transform: tr(x, y) }, { transform: tr(0, 0) }]), go(from, [{ transform: tr(0, 0) }, { transform: tr(-x, -y) }])], toOnTop: true } }
    case 'cover': { const [x, y] = offset(dir); return { anims: [go(to, [{ transform: tr(x, y) }, { transform: tr(0, 0) }])], toOnTop: true } }
    case 'uncover': { const [x, y] = offset(dir); return { anims: [go(from, [{ transform: tr(0, 0) }, { transform: tr(-x, -y) }])], toOnTop: false } }
    case 'wipe': return { anims: [go(to, [{ clipPath: wipeFrom(dir) }, { clipPath: 'inset(0 0 0 0)' }])], toOnTop: true }
    case 'split': {
      const start = dir === 'vertical' ? 'inset(50% 0 50% 0)' : 'inset(0 50% 0 50%)'
      return { anims: [go(to, [{ clipPath: start }, { clipPath: 'inset(0 0 0 0)' }])], toOnTop: true }
    }
    case 'reveal': return { anims: [go(to, [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }])], toOnTop: true }
    case 'zoom': {
      return dir === 'in'
        ? { anims: [go(to, [{ transform: 'scale(0.25)', opacity: 0 }, { transform: 'scale(1)', opacity: 1 }])], toOnTop: true }
        : { anims: [go(from, [{ transform: 'scale(1)', opacity: 1 }, { transform: 'scale(1.8)', opacity: 0 }])], toOnTop: false }
    }
    default: return { anims: [], toOnTop: true }
  }
}

/**
 * Plays `from` -> `to` with the given transition and calls onDone when the animation has finished.
 * Both slides are real renderings; nothing is faked with labels or timers.
 */
export function TransitionStage({ from, to, transition, reverse, scale, deckTheme, onDone }: {
  from: Slide; to: Slide; transition: Transition; reverse: boolean; scale: number; deckTheme: string; onDone: () => void
}) {
  const fromRef = useRef<HTMLDivElement>(null), toRef = useRef<HTMLDivElement>(null)
  const done = useRef(onDone); done.current = onDone
  useLayoutEffect(() => {
    const f = fromRef.current, t = toRef.current
    if (!f || !t) return
    const { anims, toOnTop } = buildTransition(transition, reverse, f, t)
    f.style.zIndex = toOnTop ? '1' : '2'; t.style.zIndex = toOnTop ? '2' : '1'
    let cancelled = false
    void Promise.all(anims.map(a => a.finished)).then(() => { if (!cancelled) done.current() }).catch(() => { /* cancelled by a newer transition */ })
    return () => { cancelled = true; anims.forEach(a => a.cancel()) }
  }, [transition, reverse])
  const w = SLIDE_W * scale, h = SLIDE_H * scale
  const layer = { position: 'absolute' as const, left: 0, top: 0, width: w, height: h, willChange: 'transform, opacity, clip-path' }
  return (
    <div data-testid="transition-stage" data-transition={transition.type} style={{ position: 'relative', width: w, height: h, overflow: 'hidden', background: '#000' }}>
      <div ref={fromRef} style={layer}><SlideView slide={from} deckTheme={deckTheme} scale={scale} /></div>
      <div ref={toRef} style={layer}><SlideView slide={to} deckTheme={deckTheme} scale={scale} /></div>
    </div>
  )
}
