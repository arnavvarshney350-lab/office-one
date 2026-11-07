import { Editor, type ChainedCommands, type Extensions } from '@tiptap/core'
import { StarterKit } from '@tiptap/starter-kit'
import { TextStyleKit } from '@tiptap/extension-text-style'
import Highlight from '@tiptap/extension-highlight'
import TextAlign from '@tiptap/extension-text-align'

/** Text-box editor schema: paragraphs, lists and inline formatting only. */
export const richExtensions = (): Extensions => [
  StarterKit.configure({ heading: false, blockquote: false, codeBlock: false, code: false, horizontalRule: false, link: false }),
  TextStyleKit, Highlight.configure({ multicolor: true }), TextAlign.configure({ types: ['paragraph'] }),
]

/** Apply a formatting command to an entire text box without opening it for editing. */
export function applyToAllHtml(html: string, fn: (c: ChainedCommands) => ChainedCommands): string {
  const ed = new Editor({ extensions: richExtensions(), content: html })
  try { fn(ed.chain().selectAll()).run(); return ed.getHTML() } finally { ed.destroy() }
}

export { safeHtml } from './sanitize'

/* ---------- PPTX run conversion ---------- */
export interface Run { text: string; options: Record<string, unknown> }
const hex = (c: string | undefined): string | undefined => {
  if (!c) return
  const m = /^#([0-9a-f]{6})$/i.exec(c.trim()); if (m) return m[1].toUpperCase()
  const r = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (r) return [r[1], r[2], r[3]].map(n => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase()
}
interface Fmt { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; color?: string; highlight?: string; fontFace?: string; fontSize?: number }

/** Convert stored HTML into pptxgenjs text runs, preserving per-run formatting and paragraph breaks. */
export function htmlToRuns(html: string, base: { fontFace: string; fontSize: number; color: string; bold?: boolean; italic?: boolean; align: string }): Run[] {
  const dom = new DOMParser().parseFromString(html, 'text/html')
  const runs: Run[] = []
  const inline = (node: Node, f: Fmt, out: Run[]) => {
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { if (n.textContent) out.push({ text: n.textContent, options: { bold: f.bold ?? base.bold, italic: f.italic ?? base.italic, underline: f.underline ? { style: 'sng' } : undefined, strike: f.strike ? 'sngStrike' : undefined, color: f.color ?? base.color, fontFace: f.fontFace ?? base.fontFace, fontSize: f.fontSize ?? base.fontSize, highlight: f.highlight } }); return }
      if (n.nodeType !== 1) return
      const el = n as HTMLElement, tag = el.tagName.toLowerCase(), g: Fmt = { ...f }
      if (tag === 'strong' || tag === 'b') g.bold = true
      if (tag === 'em' || tag === 'i') g.italic = true
      if (tag === 'u') g.underline = true
      if (tag === 's' || tag === 'strike') g.strike = true
      if (tag === 'mark') g.highlight = hex(el.style.backgroundColor || el.getAttribute('data-color') || '') ?? 'FFFF00'
      if (el.style.color) g.color = hex(el.style.color) ?? g.color
      if (el.style.backgroundColor && tag !== 'mark') g.highlight = hex(el.style.backgroundColor)
      if (el.style.fontFamily) g.fontFace = el.style.fontFamily.split(',')[0].replace(/["']/g, '').trim()
      if (el.style.fontSize) { const px = parseFloat(el.style.fontSize); if (px) g.fontSize = px }
      if (tag === 'br') { out.push({ text: '\n', options: {} }); return }
      inline(el, g, out)
    })
  }
  const paras: { el: HTMLElement; bullet?: 'bullet' | 'number' }[] = []
  dom.body.childNodes.forEach(n => {
    if (n.nodeType !== 1) return
    const el = n as HTMLElement
    if (el.tagName === 'UL' || el.tagName === 'OL') el.querySelectorAll(':scope > li').forEach(li => paras.push({ el: li as HTMLElement, bullet: el.tagName === 'UL' ? 'bullet' : 'number' }))
    else paras.push({ el })
  })
  paras.forEach((p, i) => {
    const out: Run[] = []
    inline(p.el, {}, out)
    if (!out.length) out.push({ text: ' ', options: { fontFace: base.fontFace, fontSize: base.fontSize } })
    const align = ((p.el.querySelector(':scope > p') as HTMLElement | null) ?? p.el).style.textAlign || base.align
    out[0].options.align = align
    if (p.bullet) out[0].options.bullet = p.bullet === 'number' ? { type: 'number' } : true
    if (i < paras.length - 1) out[out.length - 1].options.breakLine = true
    runs.push(...out)
  })
  return runs
}
