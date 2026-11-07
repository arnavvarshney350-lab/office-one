/** Page setup for the document canvas, print layout and DOCX export. */
export type PaperId = 'letter' | 'a4' | 'legal' | 'a5' | 'tabloid'
export type MarginId = 'normal' | 'narrow' | 'moderate' | 'wide'

export interface PageSetup {
  paper: PaperId
  landscape: boolean
  margin: MarginId
  columns: 1 | 2 | 3
}
export const DEFAULT_PAGE: PageSetup = { paper: 'letter', landscape: false, margin: 'normal', columns: 1 }

/** Paper sizes in CSS pixels at 96dpi. */
export const PAPERS: Record<PaperId, { name: string; w: number; h: number }> = {
  letter: { name: 'Letter (8.5 × 11 in)', w: 816, h: 1056 },
  a4: { name: 'A4 (210 × 297 mm)', w: 794, h: 1123 },
  legal: { name: 'Legal (8.5 × 14 in)', w: 816, h: 1344 },
  a5: { name: 'A5 (148 × 210 mm)', w: 559, h: 794 },
  tabloid: { name: 'Tabloid (11 × 17 in)', w: 1056, h: 1632 },
}
/** Margins in CSS pixels at 96dpi: [top, right, bottom, left]. */
export const MARGINS: Record<MarginId, { name: string; box: [number, number, number, number] }> = {
  normal: { name: 'Normal (1 in)', box: [96, 96, 96, 96] },
  narrow: { name: 'Narrow (0.5 in)', box: [48, 48, 48, 48] },
  moderate: { name: 'Moderate (1 × 0.75 in)', box: [96, 72, 96, 72] },
  wide: { name: 'Wide (1 × 2 in)', box: [96, 192, 96, 192] },
}

export interface PageMetrics { width: number; height: number; margin: [number, number, number, number]; contentWidth: number }
export function pageMetrics(p: PageSetup): PageMetrics {
  const paper = PAPERS[p.paper]
  const width = p.landscape ? paper.h : paper.w
  const height = p.landscape ? paper.w : paper.h
  const margin = MARGINS[p.margin].box
  return { width, height, margin, contentWidth: width - margin[1] - margin[3] }
}

/** Document payload. Older documents stored only `{ html }`. */
export interface WordDoc { html: string; page?: PageSetup }
export function readDoc(content: unknown): { html: string; page: PageSetup } {
  if (typeof content === 'string') return { html: content, page: { ...DEFAULT_PAGE } }
  const d = content as WordDoc | null
  return { html: d?.html ?? '<p></p>', page: { ...DEFAULT_PAGE, ...d?.page } }
}

/** DOCX page size and margins, in twips (1/1440 in). */
export const toTwips = (px: number) => Math.round((px / 96) * 1440)
