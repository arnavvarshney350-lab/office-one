import { describe, expect, it } from 'vitest'
import { GOOGLE_FONTS } from '../services/googleFonts'
import {
  applyTheme, blend, contrast, decor, deckFontNames, isDecoObj, layoutObjects, makeSlide, newDeck, sectionFilled, slideBgCss, themeBgCss, themeOf, THEMES,
  SLIDE_H, SLIDE_W, type Deck, type LayoutId, type Slide, type SlideObj, type Theme,
} from '../editors/powerpoint/model'

const HEX = /^#[0-9a-f]{6}$/i
const LAYOUT_IDS: LayoutId[] = ['title', 'content', 'two', 'section']
const surfaces = (t: Theme) => (t.bgStyle.kind === 'gradient' ? [t.bg, t.bgStyle.to] : [t.bg])
const GOOGLE = new Set(GOOGLE_FONTS.map(f => f.name))
const SYSTEM = new Set(['Arial', 'Georgia', 'Helvetica'])

describe('theme library', () => {
  it('has the original five ids plus at least 16 new themes', () => {
    expect(THEMES.length).toBeGreaterThanOrEqual(21)
    for (const id of ['classic', 'midnight', 'sunrise', 'forest', 'mono']) expect(THEMES.find(t => t.id === id)).toBeTruthy()
    expect(new Set(THEMES.map(t => t.id)).size).toBe(THEMES.length)
    expect(new Set(THEMES.map(t => t.name)).size).toBe(THEMES.length)
  })
  it('covers light, dark and colourful moods and a range of uses', () => {
    for (const m of ['light', 'dark', 'colour']) expect(THEMES.filter(t => t.mood === m).length).toBeGreaterThanOrEqual(4)
    const uses = new Set(THEMES.flatMap(t => t.uses))
    for (const u of ['business', 'pitch', 'education', 'creative', 'minimal', 'editorial', 'technical']) expect(uses.has(u as never)).toBe(true)
  })
  it('falls back to the first theme for an unknown id (old or foreign decks)', () => expect(themeOf('nope').id).toBe(THEMES[0].id))

  describe.each(THEMES.map(t => [t.id, t] as const))('%s', (_id, t) => {
    it('has valid colours', () => {
      for (const c of [t.bg, t.title, t.body, t.accent, t.accent2, t.onAccent]) expect(c).toMatch(HEX)
      if (t.bgStyle.kind === 'gradient') expect(t.bgStyle.to).toMatch(HEX)
      if (t.bgStyle.kind === 'pattern') { expect(t.bgStyle.ink).toMatch(HEX); expect(t.bgStyle.alpha).toBeLessThan(0.2) }
    })
    it('title and body text reach 4.5:1 on every surface colour', () => {
      for (const s of surfaces(t)) { expect(contrast(t.title, s)).toBeGreaterThanOrEqual(4.5); expect(contrast(t.body, s)).toBeGreaterThanOrEqual(4.5) }
    })
    it('text on accent reaches 4.5:1 and the accent is visible on the surface', () => {
      expect(contrast(t.onAccent, t.accent)).toBeGreaterThanOrEqual(4.5)
      for (const s of surfaces(t)) expect(contrast(t.accent, s)).toBeGreaterThanOrEqual(t.id === 'sunrise' || t.id === 'primer' ? 2.5 : 3)
    })
    it('uses fonts that exist (Google catalogue or a classic system face)', () => {
      for (const f of [t.font, t.headingFont]) expect(GOOGLE.has(f) || SYSTEM.has(f)).toBe(true)
    })
    it('keeps text readable over translucent decoration', () => {
      for (const l of LAYOUT_IDS) {
        const filled = l === 'section' && sectionFilled(t)
        for (const o of decor(l, t).filter(d => d.opacity != null)) {
          for (const s of filled ? [t.accent] : surfaces(t)) {
            const under = blend(s, o.fill!, o.opacity!)
            const [a, b] = filled ? [t.onAccent, t.onAccent] : [t.title, t.body]
            expect(contrast(a, under)).toBeGreaterThanOrEqual(4.5); expect(contrast(b, under)).toBeGreaterThanOrEqual(4.5)
          }
        }
      }
    })
    it('keeps solid decoration out of the text boxes', () => {
      const hit = (a: SlideObj, b: SlideObj) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h
      for (const l of LAYOUT_IDS) {
        const objs = layoutObjects(l, t), text = objs.filter(o => o.role)
        for (const d of objs.filter(o => o.deco && o.opacity == null && !(o.w >= SLIDE_W && o.h >= SLIDE_H))) for (const x of text) expect(hit(d, x), `${l}: ${d.x},${d.y}`).toBe(false)
      }
    })
    it('lays out every layout with themed text colours', () => {
      for (const l of LAYOUT_IDS) {
        for (const o of layoutObjects(l, t).filter(x => x.role)) {
          expect(o.color).toBe(l === 'section' && sectionFilled(t) ? t.onAccent : o.role === 'title' ? t.title : t.body)
          expect(o.x).toBeGreaterThanOrEqual(0); expect(o.x + o.w).toBeLessThanOrEqual(SLIDE_W); expect(o.y + o.h).toBeLessThanOrEqual(SLIDE_H)
        }
      }
    })
    it('produces a CSS background', () => expect(themeBgCss(t).backgroundColor).toBe(t.bg))
  })
})

describe('applyTheme', () => {
  const custom = (over: Partial<SlideObj>): SlideObj => ({ id: 'c' + Math.random(), type: 'text', x: 100, y: 100, w: 200, h: 50, text: 'Keep me', fontSize: 20, ...over })
  const build = (themeId: string): Deck => {
    const old = themeOf(themeId)
    const title = makeSlide('title', themeId), content = makeSlide('content', themeId), section = makeSlide('section', themeId)
    title.objects.push(
      custom({ id: 'note', text: 'Custom note', color: old.body, fontFamily: old.font }),
      { id: 'badge', type: 'shape', shape: 'ellipse', x: 10, y: 10, w: 30, h: 30, fill: old.accent },
      { id: 'img', type: 'image', x: 0, y: 0, w: 10, h: 10, src: 'data:image/png;base64,AAAA' },
      custom({ id: 'rich', html: `<p><span style="color: ${old.accent}">Accent word</span> plain</p>`, text: 'Accent word plain', color: '#123456' }),
    )
    content.objects[1].text = 'Edited body the user typed'
    content.objects[0].x = 33 // user moved this title
    return { theme: themeId, slides: [title, content, section] }
  }

  it('keeps every piece of user content', () => {
    const d = build('classic'), out = applyTheme(d, 'lumen')
    expect(out.theme).toBe('lumen'); expect(out.slides).toHaveLength(3)
    const t = out.slides[0].objects
    expect(t.find(o => o.id === 'note')?.text).toBe('Custom note')
    expect(t.find(o => o.id === 'img')?.src).toBe('data:image/png;base64,AAAA')
    expect(t.find(o => o.id === 'rich')?.text).toBe('Accent word plain')
    expect(t.filter(o => o.role).map(o => o.text)).toEqual(['Presentation title', 'Subtitle'])
    expect(out.slides[1].objects.find(o => o.role === 'body')?.text).toBe('Edited body the user typed')
    expect(out.slides.map(s => s.id)).toEqual(d.slides.map(s => s.id))
  })
  it('recolours and re-fonts text, shapes and inline colours', () => {
    const lumen = themeOf('lumen'), out = applyTheme(build('classic'), 'lumen')
    for (const s of out.slides) expect(s.bg).toBe(lumen.bg)
    const t = out.slides[0].objects
    expect(t.find(o => o.id === 'note')).toMatchObject({ color: lumen.body, fontFamily: lumen.font })
    expect(t.find(o => o.id === 'badge')?.fill).toBe(lumen.accent)
    expect(t.find(o => o.id === 'rich')?.html).toContain(lumen.accent)
    expect(t.find(o => o.id === 'rich')?.color).toBe('#123456') // a colour the user chose stays
    expect(t.find(o => o.role === 'title')).toMatchObject({ color: lumen.title })
    expect(t.find(o => o.role === 'title')?.fontFamily).toBeUndefined() // follows the theme's heading face
  })
  it('replaces theme decoration instead of stacking it', () => {
    let d = build('classic')
    for (const id of ['harbour', 'orchid', 'classic', 'kiln', 'blueprint']) d = applyTheme(d, id)
    for (const s of d.slides) {
      const t = themeOf(d.theme)
      expect(s.objects.filter(o => o.deco).length).toBe(decor(s.layout, t).length)
    }
  })
  it('migrates the unflagged accent bars of decks saved before this feature', () => {
    const legacy: Deck = { theme: 'classic', slides: [{ id: 's', bg: '#ffffff', layout: 'title', objects: [
      { id: 'a', type: 'text', role: 'title', x: 80, y: 170, w: 800, h: 110, text: 'T', fontSize: 54, bold: true, color: '#16325c', align: 'center' },
      { id: 'bar', type: 'shape', shape: 'rect', x: 400, y: 288, w: 160, h: 5, fill: '#2563eb' }] }] }
    expect(isDecoObj(legacy.slides[0].objects[1], 'title', themeOf('classic'))).toBe(true)
    const out = applyTheme(legacy, 'midnight')
    expect(out.slides[0].objects.find(o => o.id === 'bar')).toBeUndefined()
    expect(out.slides[0].objects.filter(o => o.deco)).toHaveLength(decor('title', themeOf('midnight')).length)
    expect(out.slides[0].objects.find(o => o.id === 'a')?.text).toBe('T')
  })
  it('leaves a title the user moved where it is, but snaps untouched ones to the new arrangement', () => {
    const out = applyTheme(build('classic'), 'harbour') // centred -> left-aligned title slide
    expect(out.slides[0].objects.find(o => o.role === 'title')).toMatchObject({ x: 80, w: 590, align: 'left' })
    expect(out.slides[1].objects.find(o => o.role === 'title')?.x).toBe(33)
  })
  it('round-trips through every theme without losing slides or text', () => {
    let d = build('forest'); const texts = (x: Deck) => x.slides.flatMap(s => s.objects.filter(o => o.type === 'text').map(o => o.text)).sort()
    const before = texts(d)
    for (const t of THEMES) { d = applyTheme(d, t.id); expect(texts(d)).toEqual(before) }
    expect(d.slides).toHaveLength(3)
  })
  it('unknown theme ids render with the default theme', () => {
    const d: Deck = { ...newDeck('classic'), theme: 'from-the-future' }
    expect(applyTheme(d, 'classic').theme).toBe('classic')
  })
})

describe('backgrounds and fonts', () => {
  it('shows the theme background style until a slide has its own colour', () => {
    const t = themeOf('blueprint'), s: Slide = { id: 'x', layout: 'blank', objects: [], bg: t.bg }
    expect(slideBgCss(s, t).backgroundImage).toContain('linear-gradient')
    expect(slideBgCss({ ...s, bg: '#ffffff' }, t)).toEqual({ backgroundColor: '#ffffff' })
    expect(themeBgCss(themeOf('lumen')).backgroundImage).toContain('135deg')
  })
  it('lists both theme fonts for preloading', () => {
    const d = newDeck('plum')
    d.slides[0].objects.push({ id: 'z', type: 'text', x: 0, y: 0, w: 1, h: 1, text: 'x', fontFamily: 'Lora' })
    expect(deckFontNames(d)).toEqual(['Figtree', 'Gloock', 'Lora'])
  })
})
