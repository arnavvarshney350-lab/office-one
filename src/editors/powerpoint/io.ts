import { createElement } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { download, safeName } from '../../services/download'
import { SlideView } from './SlideView'
import { SLIDE_H, SLIDE_W, deckFontNames, fontFor, objHtml, themeOf, type Deck, type Slide, type Theme } from './model'
import { htmlToRuns } from './richText'
import { preloadFonts } from '../../services/fonts'

const hex = (c: string | undefined, d = '000000') => (c && /^#[0-9a-f]{6}$/i.test(c) ? c.slice(1).toUpperCase() : d)

/**
 * A theme's gradient or pattern background as an image. PPTX has no portable pattern fill and
 * pptxgenjs cannot write gradient fills, so these backgrounds are exported as a picture.
 */
export function rasterBackground(t: Theme, w = 1600): string | null {
  const b = t.bgStyle
  if (b.kind === 'solid') return null
  const k = w / SLIDE_W, h = Math.round(SLIDE_H * k)
  const c = document.createElement('canvas'); c.width = w; c.height = h
  const g = c.getContext('2d'); if (!g) return null
  g.fillStyle = t.bg; g.fillRect(0, 0, w, h)
  if (b.kind === 'gradient') {
    // Same geometry as CSS linear-gradient(<angle>deg, from, to)
    const a = (b.angle * Math.PI) / 180, dx = Math.sin(a), dy = -Math.cos(a)
    const len = Math.abs(w * dx) + Math.abs(h * dy)
    const gr = g.createLinearGradient(w / 2 - (dx * len) / 2, h / 2 - (dy * len) / 2, w / 2 + (dx * len) / 2, h / 2 + (dy * len) / 2)
    gr.addColorStop(0, t.bg); gr.addColorStop(1, b.to)
    g.fillStyle = gr; g.fillRect(0, 0, w, h)
  } else {
    const [r, gg, bb] = [1, 3, 5].map(i => parseInt(b.ink.slice(i, i + 2), 16)), step = b.size * k
    g.fillStyle = `rgba(${r},${gg},${bb},${b.alpha})`; g.strokeStyle = g.fillStyle; g.lineWidth = Math.max(1, k)
    if (b.pattern === 'dots') for (let x = step / 2; x < w; x += step) for (let y = step / 2; y < h; y += step) { g.beginPath(); g.arc(x, y, 1.7 * k, 0, 7); g.fill() }
    else {
      g.beginPath()
      if (b.pattern === 'grid') for (let x = 0.5; x < w; x += step) { g.moveTo(x, 0); g.lineTo(x, h) }
      for (let y = 0.5; y < h; y += step) { g.moveTo(0, y); g.lineTo(w, y) }
      g.stroke()
    }
  }
  return c.toDataURL('image/jpeg', 0.92)
}

/**
 * PPTX export. Text keeps per-run formatting (font, size, colour, bold/italic/underline/strike, highlight).
 * Theme fonts become the file's heading/body theme fonts and are named on every run; PowerPoint substitutes a
 * font that is not installed on the viewing machine. Theme colours are written as explicit colours on each object.
 * Gradient and pattern backgrounds are exported as a background picture; translucent decoration keeps its transparency.
 */
export async function exportPptx(name: string, deck: Deck): Promise<void> {
  const { default: PptxGenJS } = await import('pptxgenjs')
  const p = new PptxGenJS()
  p.layout = 'LAYOUT_WIDE' // 13.33 x 7.5 in = 72 units per inch, so 1 slide unit = 1 pt
  p.title = name; p.author = 'Office One'
  const theme = themeOf(deck.theme)
  p.theme = { headFontFace: theme.headingFont, bodyFontFace: theme.font }
  const px = (n: number) => n / 72
  const themedBg = rasterBackground(theme)
  for (const s of deck.slides) {
    const slide = p.addSlide()
    slide.background = themedBg && s.bg.toLowerCase() === theme.bg.toLowerCase() ? { data: themedBg } : { color: hex(s.bg, 'FFFFFF') }
    if (s.notes?.trim()) slide.addNotes(s.notes)
    for (const o of s.objects) {
      const pos = { x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h) }
      if (o.type === 'text') {
        const runs = htmlToRuns(objHtml(o), { fontFace: fontFor(o, theme), fontSize: o.fontSize ?? 24, color: hex(o.color), bold: o.bold, italic: o.italic, align: o.align ?? 'left' })
        slide.addText(runs as never, { ...pos, valign: 'top', margin: 6, fit: 'none', lineSpacingMultiple: o.lineHeight ? o.lineHeight / 1.2 : undefined, charSpacing: o.letterSpacing || undefined })
      } else if (o.type === 'shape') {
        const type = { rect: p.ShapeType.rect, ellipse: p.ShapeType.ellipse, triangle: p.ShapeType.triangle, arrow: p.ShapeType.rightArrow }[o.shape ?? 'rect']
        slide.addShape(type, { ...pos, fill: { color: hex(o.fill, '2563EB'), transparency: o.opacity != null ? Math.round((1 - o.opacity) * 100) : 0 }, line: { color: hex(o.fill, '2563EB'), width: 0 } })
      } else if (o.type === 'image' && o.src) slide.addImage({ data: o.src, ...pos })
      // PowerPoint transitions are not written: pptxgenjs has no transition API, so they stay in Office One.
    }
  }
  const out = (await p.write({ outputType: 'blob' })) as Blob
  download(out, `${safeName(name)}.pptx`)
}

/** Rasterise one slide with the same renderer the editor uses, so fonts and rich text match exactly. */
async function slideToCanvas(slide: Slide, deckTheme: string, scale: number): Promise<HTMLCanvasElement> {
  const { default: html2canvas } = await import('html2canvas')
  const host = document.createElement('div')
  host.style.cssText = `position:fixed;left:-20000px;top:0;width:${SLIDE_W}px;height:${SLIDE_H}px;`
  document.body.appendChild(host)
  const root = createRoot(host)
  try {
    flushSync(() => root.render(createElement(SlideView, { slide, deckTheme, scale: 1 })))
    await document.fonts.ready
    return await html2canvas(host, { scale, width: SLIDE_W, height: SLIDE_H, backgroundColor: slide.bg, logging: false, useCORS: true })
  } finally { root.unmount(); host.remove() }
}

/** PDF export: each slide is rendered exactly as shown on screen (text becomes part of the page image). */
export async function exportPdf(name: string, deck: Deck): Promise<void> {
  await preloadFonts(deckFontNames(deck))
  const { jsPDF } = await import('jspdf')
  const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: [SLIDE_W, SLIDE_H], compress: true })
  doc.setProperties({ title: name })
  for (let i = 0; i < deck.slides.length; i++) {
    const canvas = await slideToCanvas(deck.slides[i], deck.theme, 2)
    if (i > 0) doc.addPage([SLIDE_W, SLIDE_H], 'landscape')
    doc.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, SLIDE_W, SLIDE_H)
  }
  download(doc.output('blob'), `${safeName(name)}.pdf`)
}
