import { memo, useMemo, type CSSProperties, type PointerEventHandler, type ReactNode } from 'react'
import { SLIDE_H, SLIDE_W, fontFor, objHtml, slideBgCss, themeOf, type Slide, type SlideObj } from './model'
import { cssFamily } from '../../services/fonts'
import { safeHtml } from './sanitize'

export function ShapeSvg({ o }: { o: SlideObj }) {
  const f = o.fill ?? '#2563eb'
  return (
    <svg opacity={o.opacity} width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none" style={{ display: 'block' }}>
      {o.shape === 'ellipse' ? <ellipse cx="50" cy="50" rx="50" ry="50" fill={f} />
        : o.shape === 'triangle' ? <polygon points="50,0 100,100 0,100" fill={f} />
        : o.shape === 'arrow' ? <polygon points="0,30 62,30 62,0 100,50 62,100 62,70 0,70" fill={f} />
        : <rect width="100" height="100" fill={f} />}
    </svg>
  )
}

/** Box-level typography. Per-run formatting lives inside the HTML. */
export const boxStyle = (o: SlideObj, deckFont: string): CSSProperties => ({
  fontFamily: cssFamily(o.fontFamily || deckFont), fontSize: o.fontSize ?? 24, fontWeight: o.bold ? 700 : 400, fontStyle: o.italic ? 'italic' : 'normal',
  color: o.color ?? '#000', textAlign: o.align ?? 'left', lineHeight: o.lineHeight ?? 1.2, letterSpacing: o.letterSpacing ? `${o.letterSpacing}px` : undefined,
  whiteSpace: 'pre-wrap', overflowWrap: 'anywhere', padding: 8,
})

const RichText = memo(function RichText({ o, font }: { o: SlideObj; font: string }) {
  const html = useMemo(() => safeHtml(objHtml(o)), [o])
  return <div className="slide-text" style={{ ...boxStyle(o, font), width: '100%', height: '100%', overflow: 'hidden' }} dangerouslySetInnerHTML={{ __html: html }} />
})

export function ObjView({ o, font, children, onPointerDown, selected }: { o: SlideObj; font: string; children?: ReactNode; onPointerDown?: PointerEventHandler; selected?: boolean }) {
  return (
    <div data-obj={o.id} onPointerDown={onPointerDown} style={{ position: 'absolute', left: o.x, top: o.y, width: o.w, height: o.h, outline: selected ? '2px solid var(--accent)' : undefined, touchAction: 'none' }}>
      {o.type === 'shape' && <ShapeSvg o={o} />}
      {o.type === 'image' && <img src={o.src} alt="" draggable={false} style={{ width: '100%', height: '100%', objectFit: 'fill', display: 'block', pointerEvents: 'none' }} />}
      {o.type === 'text' && (children ?? <RichText o={o} font={font} />)}
    </div>
  )
}

/** Static, scaled rendering of one slide (thumbnails, presenter, transitions, PDF export). */
export const SlideView = memo(function SlideView({ slide, deckTheme, scale }: { slide: Slide; deckTheme: string; scale: number }) {
  const theme = themeOf(deckTheme)
  return (
    <div style={{ width: SLIDE_W * scale, height: SLIDE_H * scale, overflow: 'hidden', position: 'relative', background: slide.bg }}>
      <div style={{ width: SLIDE_W, height: SLIDE_H, transform: `scale(${scale})`, transformOrigin: 'top left', position: 'absolute', left: 0, top: 0, ...slideBgCss(slide, theme) }}>
        {slide.objects.map(o => <ObjView key={o.id} o={o} font={fontFor(o, theme)} />)}
      </div>
    </div>
  )
})
