import { memo, useEffect, useMemo, useRef, useState } from 'react'
import { SlideView } from './SlideView'
import { buildCatalogue, loadFont } from '../../services/fonts'
import { applyTheme, SLIDE_H, SLIDE_W, themeFonts, themeOf, type Slide } from './model'

/**
 * A theme preview drawn by the real slide renderer. Nothing is mounted (and no font is fetched)
 * until the tile scrolls near the viewport, so a long gallery stays cheap.
 */
export const ThemeThumb = memo(function ThemeThumb({ source, from, themeId, width }: { source: Slide; /** Theme the source slide currently uses; omit when it already uses `themeId`. */ from?: string; themeId: string; width: number }) {
  const ref = useRef<HTMLDivElement>(null)
  const [near, setNear] = useState(typeof IntersectionObserver === 'undefined')
  useEffect(() => {
    const el = ref.current
    if (!el || near) return
    const io = new IntersectionObserver(es => { if (es.some(e => e.isIntersecting)) { setNear(true); io.disconnect() } }, { rootMargin: '240px' })
    io.observe(el)
    return () => io.disconnect()
  }, [near])
  useEffect(() => {
    if (!near) return
    buildCatalogue()
    for (const f of themeFonts(themeOf(themeId))) void loadFont(f)
  }, [near, themeId])
  const scale = width / SLIDE_W
  const slide = useMemo(() => (near && from && from !== themeId ? applyTheme({ theme: from, slides: [source] }, themeId).slides[0] : source), [near, from, themeId, source])
  return (
    <div ref={ref} aria-hidden style={{ width, height: SLIDE_H * scale, background: themeOf(themeId).bg }}>
      {near && <SlideView slide={slide} deckTheme={themeId} scale={scale} />}
    </div>
  )
})
