// Regenerates src/services/googleFonts.ts from Google's public font metadata.
//   npm run fonts:update
// The catalogue is ordered by popularity and stored compactly; no font files are downloaded.
import { writeFileSync } from 'node:fs'

const res = await fetch('https://fonts.google.com/metadata/fonts')
if (!res.ok) throw new Error(`Google Fonts metadata request failed: ${res.status}`)
const { familyMetadataList: families } = await res.json()

const CATS = { 'Sans Serif': 0, Serif: 1, Display: 2, Handwriting: 3, Monospace: 4 }
const rows = [...families]
  .sort((a, b) => (a.popularity ?? 99999) - (b.popularity ?? 99999))
  .filter(f => !f.family.includes('|'))
  .map(f => {
    let mask = 0, italic = 0
    for (const key of Object.keys(f.fonts ?? {})) {
      let k = key
      if (k.endsWith('i')) { italic = 1; k = k.slice(0, -1) }
      const w = Number(k)
      if (w >= 100 && w <= 900 && w % 100 === 0) mask |= 1 << (w / 100 - 1)
    }
    for (const ax of f.axes ?? []) {
      if (ax.tag === 'wght') for (let w = 100; w <= 900; w += 100) if (w >= ax.min && w <= ax.max) mask |= 1 << (w / 100 - 1)
      if (ax.tag === 'ital') italic = 1
    }
    return `${f.family}|${CATS[f.category] ?? 0}|${mask || 8}|${italic}`
  })

const out = `/**
 * Google Fonts catalogue (${rows.length} families), generated from fonts.google.com/metadata/fonts
 * and ordered by popularity. Each record is \`family|category|weightMask|hasItalic\`, where
 * weightMask has bit n set when weight (n+1)*100 is published.
 *
 * Nothing here is bundled as font data: a family's CSS is fetched from Google Fonts only when
 * it is first previewed or applied, and \`services/fonts.ts\` verifies it actually renders before
 * offering it. Regenerate with \`npm run fonts:update\`.
 */
export const GF_CATEGORIES = ['sans', 'serif', 'display', 'handwriting', 'mono'] as const
export type GFCategory = (typeof GF_CATEGORIES)[number]

const PACKED = \`${rows.join('\n')}\`

export interface GoogleFont { name: string; category: GFCategory; weights: number[]; italic: boolean; rank: number }

export const GOOGLE_FONTS: GoogleFont[] = PACKED.split('\\n').map((line, rank) => {
  const [name, cat, mask, ital] = line.split('|')
  const bits = Number(mask)
  const weights: number[] = []
  for (let i = 0; i < 9; i++) if (bits & (1 << i)) weights.push((i + 1) * 100)
  return { name, category: GF_CATEGORIES[Number(cat)], weights: weights.length ? weights : [400], italic: ital === '1', rank }
})
`
writeFileSync(new URL('../src/services/googleFonts.ts', import.meta.url), out)
console.log(`Wrote ${rows.length} families`)
