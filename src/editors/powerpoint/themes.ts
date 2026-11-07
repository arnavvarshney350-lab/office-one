import type { Theme } from './model'

/**
 * The theme library. Ids of the first five are stored in saved decks, so they never change.
 * Fonts are Google Fonts families (or classic system fonts for the original five) so a theme looks
 * the same on every machine. Every text/background pair is checked at >= 4.5:1 by the unit tests.
 *
 * Tokens: bg / bgStyle (surface), title / body (text), accent / accent2 (decoration),
 * onAccent (text that sits on an accent fill), font = body face, headingFont = title face.
 */
const solid = { kind: 'solid' } as const

export const THEMES: Theme[] = [
  // ---- original five (unchanged look) ----
  { id: 'classic', name: 'Classic Blue', mood: 'light', uses: ['business'], blurb: 'Clean navy on white', bg: '#ffffff', bgStyle: solid, title: '#16325c', body: '#344054', accent: '#2563eb', accent2: '#f59e0b', onAccent: '#ffffff', font: 'Arial', headingFont: 'Arial', headingBold: true, deco: 'bar', titleAlign: 'center' },
  { id: 'midnight', name: 'Midnight', mood: 'dark', uses: ['business'], blurb: 'Deep navy with a sky accent', bg: '#0f172a', bgStyle: solid, title: '#f8fafc', body: '#cbd5e1', accent: '#38bdf8', accent2: '#a78bfa', onAccent: '#0f172a', font: 'Arial', headingFont: 'Arial', headingBold: true, deco: 'bar', titleAlign: 'center' },
  { id: 'sunrise', name: 'Sunrise', mood: 'light', uses: ['creative'], blurb: 'Warm and friendly serif', bg: '#fff7ed', bgStyle: solid, title: '#9a3412', body: '#57534e', accent: '#f97316', accent2: '#facc15', onAccent: '#431407', font: 'Georgia', headingFont: 'Georgia', headingBold: true, deco: 'bar', titleAlign: 'center' },
  { id: 'forest', name: 'Forest', mood: 'light', uses: ['education'], blurb: 'Calm greens, serif type', bg: '#f0fdf4', bgStyle: solid, title: '#14532d', body: '#3f4f45', accent: '#16a34a', accent2: '#ca8a04', onAccent: '#052e16', font: 'Georgia', headingFont: 'Georgia', headingBold: true, deco: 'bar', titleAlign: 'center' },
  { id: 'mono', name: 'Mono Minimal', mood: 'light', uses: ['minimal'], blurb: 'Black on stone, nothing extra', bg: '#f5f5f4', bgStyle: solid, title: '#111111', body: '#444444', accent: '#111111', accent2: '#a8a29e', onAccent: '#ffffff', font: 'Helvetica', headingFont: 'Helvetica', headingBold: true, deco: 'bar', titleAlign: 'center' },

  // ---- light ----
  { id: 'harbour', name: 'Harbour', mood: 'light', uses: ['business'], blurb: 'Teal and amber on cool paper', bg: '#f3f7fa', bgStyle: solid, title: '#0b2e4a', body: '#35495c', accent: '#0a7c86', accent2: '#f2a541', onAccent: '#ffffff', font: 'Inter', headingFont: 'Sora', headingBold: true, deco: 'column', titleAlign: 'left' },
  { id: 'quarterly', name: 'Quarterly', mood: 'light', uses: ['business'], blurb: 'Report-style, red signal colour', bg: '#ffffff', bgStyle: solid, title: '#14141a', body: '#3d3d46', accent: '#d62839', accent2: '#14141a', onAccent: '#ffffff', font: 'IBM Plex Sans', headingFont: 'Archivo', headingBold: true, deco: 'blocks', titleAlign: 'left' },
  { id: 'almanac', name: 'Almanac', mood: 'light', uses: ['editorial'], blurb: 'Ink-blue serif with a vermilion rule', bg: '#eceef6', bgStyle: solid, title: '#1b2559', body: '#3a4163', accent: '#c93d12', accent2: '#1b2559', onAccent: '#ffffff', font: 'Newsreader', headingFont: 'Fraunces', headingBold: true, deco: 'rules', titleAlign: 'left' },
  { id: 'gallery', name: 'Gallery', mood: 'light', uses: ['minimal', 'editorial'], blurb: 'White walls, a thin red tick', bg: '#ffffff', bgStyle: solid, title: '#0a0a0a', body: '#4a4a4a', accent: '#0a0a0a', accent2: '#e11d48', onAccent: '#ffffff', font: 'Inter', headingFont: 'Instrument Serif', headingBold: false, deco: 'tick', titleAlign: 'left' },
  { id: 'primer', name: 'Primer', mood: 'light', uses: ['education'], blurb: 'Rounded type, coral and aqua', bg: '#fcfcff', bgStyle: solid, title: '#2a2766', body: '#3f3d66', accent: '#ff6b57', accent2: '#2ec4b6', onAccent: '#2a2766', font: 'Nunito', headingFont: 'Baloo 2', headingBold: true, deco: 'orbs', titleAlign: 'center' },
  { id: 'orchid', name: 'Orchid', mood: 'light', uses: ['creative'], blurb: 'Lilac wash, violet and fuchsia', bg: '#f7ecff', bgStyle: solid, title: '#3b1566', body: '#4f3a6b', accent: '#a21caf', accent2: '#6d28d9', onAccent: '#ffffff', font: 'DM Sans', headingFont: 'Syne', headingBold: true, deco: 'orbs', titleAlign: 'left' },

  // ---- dark ----
  { id: 'blueprint', name: 'Blueprint', mood: 'dark', uses: ['technical'], blurb: 'Drafting-grid blue, mono headings', bg: '#0b2a4a', bgStyle: { kind: 'pattern', pattern: 'grid', ink: '#5cc8ff', alpha: 0.13, size: 32 }, title: '#ffffff', body: '#bfd8f2', accent: '#5cc8ff', accent2: '#ffd166', onAccent: '#0b2a4a', font: 'IBM Plex Sans', headingFont: 'JetBrains Mono', headingBold: true, deco: 'tick', titleAlign: 'left' },
  { id: 'graphite', name: 'Graphite', mood: 'dark', uses: ['technical', 'business'], blurb: 'Dotted graphite, amber and teal', bg: '#1b1e23', bgStyle: { kind: 'pattern', pattern: 'dots', ink: '#ffffff', alpha: 0.07, size: 24 }, title: '#f5f5f4', body: '#b9bec6', accent: '#ff9f1c', accent2: '#2ec4b6', onAccent: '#1b1e23', font: 'IBM Plex Sans', headingFont: 'Space Grotesk', headingBold: true, deco: 'sidebar', titleAlign: 'left' },
  { id: 'lumen', name: 'Lumen', mood: 'dark', uses: ['pitch'], blurb: 'Indigo to violet glow for pitch decks', bg: '#1a1446', bgStyle: { kind: 'gradient', to: '#5b2a86', angle: 135 }, title: '#ffffff', body: '#e2d9f5', accent: '#ff6fb5', accent2: '#8b6cff', onAccent: '#1a1446', font: 'Hanken Grotesk', headingFont: 'Bricolage Grotesque', headingBold: true, deco: 'orbs', titleAlign: 'left' },
  { id: 'chalkboard', name: 'Chalkboard', mood: 'dark', uses: ['education'], blurb: 'Slate green, chalk lettering', bg: '#1f3d36', bgStyle: solid, title: '#f7f5ec', body: '#d6e4dd', accent: '#f6c453', accent2: '#8fd3c1', onAccent: '#1f3d36', font: 'Nunito', headingFont: 'Caveat', headingBold: true, deco: 'ribbon', titleAlign: 'center' },
  { id: 'plum', name: 'Plum Noir', mood: 'dark', uses: ['editorial', 'creative'], blurb: 'Deep plum, rose and gold', bg: '#2a1230', bgStyle: solid, title: '#fbeff7', body: '#e3c9db', accent: '#ff8fb1', accent2: '#f5c26b', onAccent: '#2a1230', font: 'Figtree', headingFont: 'Gloock', headingBold: false, deco: 'rules', titleAlign: 'left' },
  { id: 'basalt', name: 'Basalt', mood: 'dark', uses: ['minimal', 'business'], blurb: 'Volcanic grey with a gold line', bg: '#15171b', bgStyle: solid, title: '#f4f4f5', body: '#a8adb5', accent: '#d4af37', accent2: '#8a8f98', onAccent: '#15171b', font: 'Inter', headingFont: 'Cormorant Garamond', headingBold: true, deco: 'bar', titleAlign: 'center' },

  // ---- colourful ----
  { id: 'sorbet', name: 'Sorbet', mood: 'colour', uses: ['creative'], blurb: 'Raspberry on pink, playful serif', bg: '#ffe3ec', bgStyle: solid, title: '#7a1240', body: '#5a2a40', accent: '#d6125f', accent2: '#f59e0b', onAccent: '#ffffff', font: 'Figtree', headingFont: 'Young Serif', headingBold: false, deco: 'blocks', titleAlign: 'left' },
  { id: 'kiln', name: 'Kiln', mood: 'colour', uses: ['pitch', 'creative'], blurb: 'Fired orange, poster capitals', bg: '#ff8a3d', bgStyle: solid, title: '#1d0e06', body: '#3b1d0d', accent: '#1d0e06', accent2: '#fff3e6', onAccent: '#fff3e6', font: 'Work Sans', headingFont: 'Bebas Neue', headingBold: false, deco: 'blocks', titleAlign: 'left' },
  { id: 'lagoon', name: 'Lagoon', mood: 'colour', uses: ['pitch', 'education'], blurb: 'Deep teal water, sun yellow', bg: '#0b6f79', bgStyle: { kind: 'gradient', to: '#084f5c', angle: 160 }, title: '#ffffff', body: '#ecfbf9', accent: '#ffd447', accent2: '#ff7a6b', onAccent: '#0a3a42', font: 'Figtree', headingFont: 'Fraunces', headingBold: true, deco: 'column', titleAlign: 'left' },
  { id: 'cobalt', name: 'Cobalt', mood: 'colour', uses: ['pitch'], blurb: 'Electric blue with a yellow hit', bg: '#1f3fff', bgStyle: { kind: 'gradient', to: '#3a1fd1', angle: 150 }, title: '#ffffff', body: '#dfe5ff', accent: '#ffe14d', accent2: '#ff5d8f', onAccent: '#0f1a66', font: 'Inter', headingFont: 'Outfit', headingBold: true, deco: 'ribbon', titleAlign: 'left' },
  { id: 'marigold', name: 'Marigold', mood: 'colour', uses: ['pitch', 'education'], blurb: 'Sunny yellow, black type, red pop', bg: '#ffd23f', bgStyle: solid, title: '#1a1a1a', body: '#2f2a10', accent: '#1a1a1a', accent2: '#ee4266', onAccent: '#ffd23f', font: 'Karla', headingFont: 'Rubik', headingBold: true, deco: 'sidebar', titleAlign: 'left' },
]
