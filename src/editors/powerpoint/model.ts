import { uid } from '../../utilities/id'
import { THEMES } from './themes'

export { THEMES }

export const SLIDE_W = 960
export const SLIDE_H = 540

export type ShapeKind = 'rect' | 'ellipse' | 'triangle' | 'arrow'
export interface SlideObj {
  id: string
  type: 'text' | 'shape' | 'image'
  role?: 'title' | 'body'
  /** Objects sharing a group id are selected, moved and resized together. */
  group?: string
  x: number; y: number; w: number; h: number
  text?: string
  /** Rich text (HTML produced by the in-box editor). When absent, `text` is rendered as plain paragraphs. */
  html?: string
  fontFamily?: string
  fontSize?: number; bold?: boolean; italic?: boolean; color?: string
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Line height multiplier (e.g. 1.2) and letter spacing in slide px. */
  lineHeight?: number; letterSpacing?: number
  shape?: ShapeKind
  fill?: string
  /** Fill opacity of a shape (0-1). */
  opacity?: number
  /** Theme-owned decoration: replaced when the theme changes, never part of the user's content. */
  deco?: boolean
  src?: string
}
export type LayoutId = 'title' | 'content' | 'two' | 'section' | 'blank'
export type TransitionType = 'none' | 'fade' | 'push' | 'wipe' | 'split' | 'reveal' | 'cover' | 'uncover' | 'zoom' | 'dissolve'
export type TransitionDir = 'left' | 'right' | 'up' | 'down' | 'horizontal' | 'vertical' | 'in' | 'out'
export interface Transition { type: TransitionType; duration: number; direction: TransitionDir }
export interface Slide { id: string; bg: string; layout: LayoutId; objects: SlideObj[]; transition?: Transition; /** Speaker notes (plain text). */ notes?: string }

export const TRANSITIONS: { id: TransitionType; name: string; directions: TransitionDir[]; hint: string }[] = [
  { id: 'none', name: 'None', directions: [], hint: 'Cut straight to the next slide' },
  { id: 'fade', name: 'Fade', directions: [], hint: 'Cross-fade between slides' },
  { id: 'push', name: 'Push', directions: ['right', 'left', 'up', 'down'], hint: 'The new slide pushes the old one out' },
  { id: 'wipe', name: 'Wipe', directions: ['left', 'right', 'up', 'down'], hint: 'The new slide is wiped in from an edge' },
  { id: 'split', name: 'Split', directions: ['horizontal', 'vertical'], hint: 'The new slide opens from the centre' },
  { id: 'reveal', name: 'Reveal', directions: [], hint: 'The new slide is revealed through an expanding iris' },
  { id: 'cover', name: 'Cover', directions: ['right', 'left', 'up', 'down'], hint: 'The new slide slides over the old one' },
  { id: 'uncover', name: 'Uncover', directions: ['right', 'left', 'up', 'down'], hint: 'The old slide slides away to expose the new one' },
  { id: 'zoom', name: 'Zoom', directions: ['in', 'out'], hint: 'Zoom into the new slide (or out of the old one)' },
  { id: 'dissolve', name: 'Dissolve', directions: [], hint: 'A stepped, blurred dissolve' },
]
export const DIR_LABEL: Record<TransitionDir, string> = { left: 'From left', right: 'From right', up: 'From top', down: 'From bottom', horizontal: 'Horizontal', vertical: 'Vertical', in: 'Zoom in', out: 'Zoom out' }
export const defaultTransition = (type: TransitionType): Transition => ({ type, duration: 700, direction: TRANSITIONS.find(t => t.id === type)?.directions[0] ?? 'right' })

const escapeHtml = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
export const plainToHtml = (t: string) => t.split('\n').map(l => `<p>${escapeHtml(l)}</p>`).join('')
/** HTML content of a text object (legacy decks only have plain `text`). */
export const objHtml = (o: SlideObj): string => o.html ?? plainToHtml(o.text ?? '')
export const objPlain = (o: SlideObj): string => {
  if (o.html == null) return o.text ?? ''
  const d = new DOMParser().parseFromString(o.html, 'text/html')
  return [...d.body.querySelectorAll('p,li')].map(p => p.textContent ?? '').join('\n')
}
export interface Deck { slides: Slide[]; theme: string }

export type ThemeMood = 'light' | 'dark' | 'colour'
export type ThemeUse = 'business' | 'pitch' | 'education' | 'creative' | 'minimal' | 'editorial' | 'technical'
export type BgStyle =
  | { kind: 'solid' }
  | { kind: 'gradient'; to: string; angle: number }
  | { kind: 'pattern'; pattern: 'grid' | 'dots' | 'lines'; ink: string; alpha: number; size: number }
export type DecoStyle = 'bar' | 'column' | 'sidebar' | 'ribbon' | 'orbs' | 'rules' | 'tick' | 'blocks'

export interface Theme {
  id: string; name: string; blurb: string
  mood: ThemeMood; uses: ThemeUse[]
  /** Surface colour (the gradient's start colour for gradient themes). */
  bg: string; bgStyle: BgStyle
  title: string; body: string
  accent: string; accent2: string
  /** Text colour for text placed on an accent fill. */
  onAccent: string
  /** Body face (also the deck default). Older code only knows this one. */
  font: string
  /** Title face. */
  headingFont: string
  headingBold: boolean
  deco: DecoStyle
  titleAlign: 'center' | 'left'
}
export const themeOf = (id: string): Theme => THEMES.find(t => t.id === id) ?? THEMES[0]
export const MOODS: { id: ThemeMood; name: string }[] = [{ id: 'light', name: 'Light' }, { id: 'dark', name: 'Dark' }, { id: 'colour', name: 'Colourful' }]
/** The font a text object renders in: its own choice, else the theme's title or body face. */
export const fontFor = (o: SlideObj, t: Theme): string => o.fontFamily || (o.role === 'title' ? t.headingFont : t.font)
export const themeFonts = (t: Theme): string[] => [...new Set([t.font, t.headingFont])]

/** Every font a deck needs: both theme faces plus any face a text object chose itself. */
export const deckFontNames = (d: Deck): string[] => [...new Set([...themeFonts(themeOf(d.theme)), ...d.slides.flatMap(s => s.objects.map(o => o.fontFamily ?? ''))])].filter(Boolean)

/* ------------------------------------------------- colour helpers */
const rgbOf = (h: string) => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16))
const toHex = (c: number[]) => '#' + c.map(v => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0')).join('')
const lum = (h: string) => { const [r, g, b] = rgbOf(h).map(v => v / 255).map(v => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)); return 0.2126 * r + 0.7152 * g + 0.0722 * b }
/** WCAG contrast ratio between two #rrggbb colours. */
export const contrast = (a: string, b: string): number => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05) }
/** `top` painted over `base` at the given opacity. */
export const blend = (base: string, top: string, a: number): string => { const x = rgbOf(base), y = rgbOf(top); return toHex(x.map((v, i) => v * (1 - a) + y[i] * a)) }
const rgba = (h: string, a: number) => `rgba(${rgbOf(h).join(',')},${a})`

/** CSS background for a theme's surface (solid, gradient or subtle pattern). */
export function themeBgCss(t: Theme): { backgroundColor: string; backgroundImage?: string; backgroundSize?: string } {
  const b = t.bgStyle
  if (b.kind === 'gradient') return { backgroundColor: t.bg, backgroundImage: `linear-gradient(${b.angle}deg, ${t.bg}, ${b.to})` }
  if (b.kind === 'pattern') {
    const c = rgba(b.ink, b.alpha), s = `${b.size}px ${b.size}px`
    if (b.pattern === 'grid') return { backgroundColor: t.bg, backgroundImage: `linear-gradient(${c} 1px, transparent 1px), linear-gradient(90deg, ${c} 1px, transparent 1px)`, backgroundSize: s }
    if (b.pattern === 'dots') return { backgroundColor: t.bg, backgroundImage: `radial-gradient(${c} 1.6px, transparent 1.8px)`, backgroundSize: s }
    return { backgroundColor: t.bg, backgroundImage: `repeating-linear-gradient(0deg, ${c} 0 1px, transparent 1px ${b.size}px)` }
  }
  return { backgroundColor: t.bg }
}
/** A slide shows its theme's background style until the user picks a different colour for it. */
export const slideBgCss = (slide: Slide, t: Theme) =>
  slide.bg.toLowerCase() === t.bg.toLowerCase() ? themeBgCss(t) : { backgroundColor: slide.bg }

export const LAYOUTS: { id: LayoutId; name: string }[] = [
  { id: 'title', name: 'Title slide' }, { id: 'content', name: 'Title and content' },
  { id: 'two', name: 'Two content' }, { id: 'section', name: 'Section header' }, { id: 'blank', name: 'Blank' },
]

/** Section headers on these styles are a full-bleed accent field, so their text uses the on-accent colour. */
export const sectionFilled = (t: Theme) => ['column', 'sidebar', 'orbs', 'blocks'].includes(t.deco)
export const roleColor = (role: 'title' | 'body', layout: LayoutId, t: Theme) =>
  layout === 'section' && sectionFilled(t) ? t.onAccent : role === 'title' ? t.title : t.body

const tb = (role: 'title' | 'body', t: Theme, layout: LayoutId, x: number, y: number, w: number, h: number, text: string, fontSize: number, extra: Partial<SlideObj> = {}): SlideObj =>
  ({ id: uid(), type: 'text', role, x, y, w, h, text, fontSize, bold: role === 'title' ? t.headingBold : false, color: roleColor(role, layout, t), align: 'left', ...extra })

const R = (x: number, y: number, w: number, h: number, fill: string, opacity?: number): SlideObj => ({ id: uid(), type: 'shape', shape: 'rect', deco: true, x, y, w, h, fill, ...(opacity != null ? { opacity } : {}) })
const E = (x: number, y: number, w: number, h: number, fill: string, opacity?: number): SlideObj => ({ id: uid(), type: 'shape', shape: 'ellipse', deco: true, x, y, w, h, fill, ...(opacity != null ? { opacity } : {}) })

/** Theme decoration for a layout, drawn behind the text. Solid shapes stay out of the text boxes; translucent ones may overlap. */
export function decor(layout: LayoutId, t: Theme): SlideObj[] {
  const { accent: a, accent2: b, onAccent: on } = t
  const left = t.titleAlign === 'left'
  // Translucent shapes on an accent field are tinted away from the text colour, so they can only raise contrast.
  const lift = lum(on) < 0.4 ? '#ffffff' : '#000000'
  const fill = layout === 'section' && sectionFilled(t) ? [R(0, 0, SLIDE_W, SLIDE_H, a)] : []
  const content = layout === 'content' || layout === 'two'
  switch (t.deco) {
    case 'bar':
      if (layout === 'title') return left ? [R(80, 316, 160, 5, a)] : [R(400, 288, 160, 5, a)]
      if (layout === 'section') return [R(0, 0, 28, SLIDE_H, a)]
      return []
    case 'column':
      if (layout === 'title') return [R(700, 0, 260, SLIDE_H, a), E(720, 360, 150, 150, b)]
      if (content) return [R(0, 0, SLIDE_W, 12, a), R(60, 126, 72, 4, b)]
      if (layout === 'section') return [...fill, E(700, 300, 360, 360, lift, 0.12)]
      return []
    case 'sidebar':
      if (layout === 'title') return [R(0, 0, 36, SLIDE_H, a), R(36, 0, 8, SLIDE_H, b)]
      if (content) return [R(0, 0, 18, SLIDE_H, a)]
      if (layout === 'section') return [...fill, R(98, 326, 96, 5, on)]
      return []
    case 'ribbon':
      if (layout === 'title') return [R(0, 0, SLIDE_W, 16, a), R(0, 524, SLIDE_W, 16, b)]
      if (content) return [R(0, 0, SLIDE_W, 10, a), R(0, 10, SLIDE_W, 4, b)]
      if (layout === 'section') return [R(0, 0, 28, SLIDE_H, a), R(28, 0, 8, SLIDE_H, b)]
      return []
    case 'orbs':
      if (layout === 'title') return [E(560, -140, 520, 520, a, 0.2), E(700, 250, 380, 380, b, 0.18)]
      if (content) return [E(780, -100, 260, 260, a, 0.16), E(-80, 430, 220, 220, b, 0.14)]
      if (layout === 'section') return [...fill, E(560, -140, 520, 520, lift, 0.12), E(-100, 360, 300, 300, lift, 0.1)]
      return []
    case 'rules':
      if (layout === 'title') return [R(80, 166, 590, 4, a), R(80, 418, 590, 2, b)]
      if (content) return [R(60, 124, 840, 3, a), R(60, 500, 840, 1, b)]
      if (layout === 'section') return [R(90, 170, 110, 5, a)]
      return []
    case 'tick':
      if (layout === 'title') return [R(80, 164, 48, 6, b), R(80, 430, 590, 2, a)]
      if (content) return [R(60, 28, 40, 5, b), R(60, 512, 840, 2, a)]
      if (layout === 'section') return [R(90, 170, 56, 6, b)]
      return []
    case 'blocks':
      if (layout === 'title') return [R(0, 430, SLIDE_W, 110, a), R(0, 430, 260, 110, b)]
      if (content) return [R(0, 516, SLIDE_W, 24, a), R(0, 516, 200, 24, b)]
      if (layout === 'section') return [...fill, R(0, 430, 330, 110, b)]
      return []
  }
}

/** Decoration sits behind the text, except the original centred accent bar, which has always come after it. */
const withDecor = (layout: LayoutId, t: Theme, d: SlideObj[], rest: SlideObj[]): SlideObj[] => (layout === 'title' && t.deco === 'bar' ? [...rest, ...d] : [...d, ...rest])

export function layoutObjects(layout: LayoutId, t: Theme): SlideObj[] {
  const d = decor(layout, t), L = layout
  switch (layout) {
    case 'title': return withDecor(L, t, d, t.titleAlign === 'left' ? [
      tb('title', t, L, 80, 190, 590, 130, 'Presentation title', 50),
      tb('body', t, L, 80, 326, 590, 70, 'Subtitle', 24),
    ] : [
      tb('title', t, L, 80, 170, 800, 110, 'Presentation title', 54, { align: 'center' }),
      tb('body', t, L, 80, 310, 800, 60, 'Subtitle', 26, { align: 'center' }),
    ])
    case 'content': return [
      ...d,
      tb('title', t, L, 60, 40, 840, 80, 'Slide title', 40),
      tb('body', t, L, 60, 140, 840, 340, '• First point\n• Second point\n• Third point', 26),
    ]
    case 'two': return [
      ...d,
      tb('title', t, L, 60, 40, 840, 80, 'Slide title', 40),
      tb('body', t, L, 60, 140, 410, 340, '• Left point\n• Left point', 24),
      tb('body', t, L, 490, 140, 410, 340, '• Right point\n• Right point', 24),
    ]
    case 'section': return [...d, tb('title', t, L, 90, 200, 780, 110, 'Section title', 48)]
    default: return []
  }
}
export const makeSlide = (layout: LayoutId, themeId: string): Slide => {
  const t = themeOf(themeId)
  return { id: uid(), bg: t.bg, layout, objects: layoutObjects(layout, t) }
}
export const newDeck = (themeId = 'classic'): Deck => ({ theme: themeId, slides: [makeSlide('title', themeId)] })

/** Is this object decoration owned by the theme (flagged, or one of the original accent bars of older decks)? */
export function isDecoObj(o: SlideObj, layout: LayoutId, t: Theme): boolean {
  if (o.deco) return true
  if (o.type !== 'shape' || o.fill?.toLowerCase() !== t.accent.toLowerCase()) return false
  return (layout === 'title' && o.x === 400 && o.y === 288 && o.w === 160 && o.h === 5) || (layout === 'section' && o.x === 0 && o.y === 0 && o.w === 28 && o.h === SLIDE_H)
}

const colourMap = (a: Theme, b: Theme) => {
  const m = new Map<string, string>()
  for (const [x, y] of [[a.accent, b.accent], [a.accent2, b.accent2], [a.title, b.title], [a.body, b.body], [a.onAccent, b.onAccent], [a.bg, b.bg]]) if (!m.has(x.toLowerCase())) m.set(x.toLowerCase(), y)
  return m
}
const sameBox = (a: SlideObj, b: SlideObj) => a.x === b.x && a.y === b.y && a.w === b.w && a.h === b.h

/**
 * Switch a deck to another theme. Text, images and positions the user changed are kept; theme colours
 * and fonts are swapped everywhere (including custom shapes and inline rich-text colours); theme
 * decoration is replaced; untouched placeholders snap to the new theme's arrangement.
 */
export function applyTheme(deck: Deck, themeId: string): Deck {
  const t = themeOf(themeId), old = themeOf(deck.theme)
  const map = colourMap(old, t)
  const re = (c?: string) => (c ? map.get(c.toLowerCase()) ?? c : c)
  const reHtml = (h: string) => h
    .replace(/#[0-9a-fA-F]{6}\b/g, m => map.get(m.toLowerCase()) ?? m)
    .replace(/rgb\(\s*(\d+),\s*(\d+),\s*(\d+)\s*\)/g, (m, r, g, b) => map.get(toHex([+r, +g, +b])) ?? m)
  const fontMap = (f?: string) => (f === old.font ? t.font : f === old.headingFont ? t.headingFont : f)
  return {
    theme: t.id,
    slides: deck.slides.map(s => {
      const oldDef = layoutObjects(s.layout, old).filter(o => o.role), newDef = layoutObjects(s.layout, t).filter(o => o.role)
      let ri = 0
      const objects = s.objects.filter(o => !isDecoObj(o, s.layout, old)).map((o): SlideObj => {
        if (o.type === 'text') {
          const n: SlideObj = { ...o, html: o.html != null ? reHtml(o.html) : o.html }
          if (o.role) {
            const od = oldDef[ri], nd = newDef[ri]; ri++
            n.color = !o.color || o.color.toLowerCase() === roleColor(o.role, s.layout, old).toLowerCase() ? roleColor(o.role, s.layout, t) : re(o.color)
            if (o.fontFamily === old.font || o.fontFamily === old.headingFont) delete n.fontFamily
            if (od && nd) {
              if (sameBox(o, od)) Object.assign(n, { x: nd.x, y: nd.y, w: nd.w, h: nd.h })
              if (o.align === od.align) n.align = nd.align
              if (o.fontSize === od.fontSize) n.fontSize = nd.fontSize
            }
            if (o.role === 'title' && o.bold === old.headingBold) n.bold = t.headingBold
          } else {
            n.color = re(o.color)
            n.fontFamily = fontMap(o.fontFamily)
          }
          return n
        }
        return o.type === 'shape' ? { ...o, fill: re(o.fill) } : o
      })
      return { ...s, bg: t.bg, objects: withDecor(s.layout, t, decor(s.layout, t), objects) }
    }),
  }
}
