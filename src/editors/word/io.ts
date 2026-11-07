import {
  AlignmentType, Document, ExternalHyperlink, HeadingLevel, ImageRun, LevelFormat, Packer, PageOrientation, Paragraph, Table, TableCell, TableRow, TextRun, WidthType,
  type ParagraphChild, type FileChild,
} from 'docx'
import { DEFAULT_PAGE, pageMetrics, toTwips, type PageSetup } from './page'

const ALIGN: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED,
}
interface Fmt { bold?: boolean; italics?: boolean; underline?: boolean; strike?: boolean; color?: string; highlight?: string; size?: number; font?: string }

const hex = (c: string | undefined): string | undefined => {
  if (!c) return
  const m = /^#([0-9a-f]{6})$/i.exec(c.trim()); if (m) return m[1].toUpperCase()
  const r = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (r) return [r[1], r[2], r[3]].map(n => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function dataUrlToBytes(src: string): { data: Uint8Array; type: 'png' | 'jpg' | 'gif' } | null {
  const m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/.exec(src); if (!m) return null
  const bin = atob(m[2]); const data = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i)
  return { data, type: m[1] === 'png' ? 'png' : m[1] === 'gif' ? 'gif' : 'jpg' }
}

function inlines(node: Node, fmt: Fmt, out: ParagraphChild[]) {
  node.childNodes.forEach(n => {
    if (n.nodeType === 3) {
      const t = n.textContent ?? ''
      if (t) out.push(new TextRun({ text: t, bold: fmt.bold, italics: fmt.italics, underline: fmt.underline ? {} : undefined, strike: fmt.strike, color: fmt.color, size: fmt.size, font: fmt.font, shading: fmt.highlight ? { fill: fmt.highlight, type: 'clear', color: 'auto' } : undefined }))
      return
    }
    if (n.nodeType !== 1) return
    const el = n as HTMLElement
    const tag = el.tagName.toLowerCase()
    const f: Fmt = { ...fmt }
    if (tag === 'strong' || tag === 'b') f.bold = true
    if (tag === 'em' || tag === 'i') f.italics = true
    if (tag === 'u') f.underline = true
    if (tag === 's' || tag === 'strike') f.strike = true
    if (tag === 'mark') f.highlight = hex(el.style.backgroundColor || el.getAttribute('data-color') || '') ?? 'FFFF00'
    if (el.style.color) f.color = hex(el.style.color)
    if (el.style.backgroundColor && tag !== 'mark') f.highlight = hex(el.style.backgroundColor)
    if (el.style.fontSize) { const px = parseFloat(el.style.fontSize); if (px) f.size = Math.round(px * 1.5) } // px -> half-points
    if (el.style.fontFamily) f.font = el.style.fontFamily.split(',')[0].replace(/["']/g, '').trim()
    if (tag === 'br') { out.push(new TextRun({ break: 1 })); return }
    if (tag === 'img') {
      const img = dataUrlToBytes(el.getAttribute('src') ?? '')
      if (img) {
        const w = Math.min(parseInt(el.getAttribute('width') ?? '') || (el as HTMLImageElement).naturalWidth || 400, 560)
        const nat = (el as HTMLImageElement).naturalWidth, natH = (el as HTMLImageElement).naturalHeight
        out.push(new ImageRun({ type: img.type, data: img.data, transformation: { width: w, height: nat && natH ? Math.round(w * natH / nat) : Math.round(w * 0.6) } }))
      }
      return
    }
    if (tag === 'a') {
      const kids: ParagraphChild[] = []; inlines(el, { ...f, color: '0563C1', underline: true }, kids)
      out.push(new ExternalHyperlink({ link: el.getAttribute('href') ?? '', children: kids as TextRun[] })); return
    }
    inlines(el, f, out)
  })
}

function paragraph(el: HTMLElement, extra: ConstructorParameters<typeof Paragraph>[0] = {}): Paragraph {
  const kids: ParagraphChild[] = []; inlines(el, {}, kids)
  const align = ALIGN[el.style.textAlign]
  return new Paragraph({ children: kids, alignment: align, ...extra as object })
}

function blocks(parent: Element, out: FileChild[], listLevel = -1) {
  parent.childNodes.forEach(n => {
    if (n.nodeType !== 1) return
    const el = n as HTMLElement
    switch (el.tagName.toLowerCase()) {
      case 'h1': out.push(paragraph(el, { heading: HeadingLevel.HEADING_1 })); break
      case 'h2': out.push(paragraph(el, { heading: HeadingLevel.HEADING_2 })); break
      case 'h3': out.push(paragraph(el, { heading: HeadingLevel.HEADING_3 })); break
      case 'p': out.push(paragraph(el)); break
      case 'blockquote': blocks(el, out); break
      case 'pre': out.push(paragraph(el)); break
      case 'ul': case 'ol': {
        const isOrdered = el.tagName.toLowerCase() === 'ol'
        el.childNodes.forEach(li => {
          if (li.nodeType !== 1) return
          const lie = li as HTMLElement
          const first = lie.querySelector(':scope > p') as HTMLElement | null
          const para = first ?? lie
          const kids: ParagraphChild[] = []
          if (first) inlines(first, {}, kids); else lie.childNodes.forEach(c => { if (c.nodeType === 3 || !['ul', 'ol'].includes((c as HTMLElement).tagName?.toLowerCase())) inlines({ childNodes: [c] } as unknown as Node, {}, kids) })
          const lvl = Math.max(listLevel + 1, 0)
          out.push(new Paragraph({ children: kids, alignment: ALIGN[(para as HTMLElement).style?.textAlign], ...(isOrdered ? { numbering: { reference: 'ordered', level: Math.min(lvl, 3) } } : { bullet: { level: Math.min(lvl, 3) } }) }))
          lie.querySelectorAll(':scope > ul, :scope > ol').forEach(sub => blocks({ childNodes: [sub] } as unknown as Element, out, lvl))
        })
        break
      }
      case 'table': {
        const rows: TableRow[] = []
        el.querySelectorAll('tr').forEach(tr => {
          const cells: TableCell[] = []
          tr.querySelectorAll(':scope > td, :scope > th').forEach(td => {
            const inner: FileChild[] = []; blocks(td as Element, inner)
            if (!inner.length) inner.push(new Paragraph(''))
            cells.push(new TableCell({ children: inner as Paragraph[], columnSpan: Number(td.getAttribute('colspan')) || undefined, width: { size: 100 / Math.max(tr.children.length, 1), type: WidthType.PERCENTAGE } }))
          })
          rows.push(new TableRow({ children: cells }))
        })
        if (rows.length) out.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }))
        out.push(new Paragraph(''))
        break
      }
      case 'img': { const kids: ParagraphChild[] = []; inlines({ childNodes: [el] } as unknown as Node, {}, kids); out.push(new Paragraph({ children: kids })); break }
      case 'hr': out.push(new Paragraph({ border: { bottom: { style: 'single', size: 6, color: '999999', space: 1 } } })); break
      default: blocks(el, out)
    }
  })
}

export async function exportDocx(title: string, html: string, page: PageSetup = DEFAULT_PAGE): Promise<Blob> {
  const dom = new DOMParser().parseFromString(html, 'text/html')
  const children: FileChild[] = []
  blocks(dom.body, children)
  if (!children.length) children.push(new Paragraph(''))
  const m = pageMetrics(page)
  const [top, right, bottom, left] = m.margin
  const doc = new Document({
    creator: 'Office One', title,
    numbering: { config: [{ reference: 'ordered', levels: [0, 1, 2, 3].map(level => ({ level, format: LevelFormat.DECIMAL, text: `%${level + 1}.`, alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } } })) }] },
    sections: [{
      properties: {
        page: {
          size: { width: toTwips(page.landscape ? m.height : m.width), height: toTwips(page.landscape ? m.width : m.height), orientation: page.landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
          margin: { top: toTwips(top), right: toTwips(right), bottom: toTwips(bottom), left: toTwips(left) },
        },
        column: page.columns > 1 ? { count: page.columns, space: 420 } : undefined,
      },
      children,
    }],
  })
  return Packer.toBlob(doc)
}

export async function exportPdf(title: string, html: string, page: PageSetup = DEFAULT_PAGE): Promise<Blob> {
  const { jsPDF } = await import('jspdf')
  const m = pageMetrics(page)
  const holder = document.createElement('div')
  holder.style.cssText = `position:fixed;left:-10000px;top:0;width:${m.width}px;background:#fff;color:#000`
  const body = document.createElement('div')
  body.className = 'ProseMirror'
  body.style.cssText = `--border:#999;--panel:#f1f1f1;--accent:#0563c1;--accent-soft:#e7efff;background:#fff;color:#000;padding:0;min-height:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;width:${m.contentWidth}px`
  body.innerHTML = html
  holder.appendChild(body)
  document.body.appendChild(holder)
  try {
    // page sizes are in CSS pixels at 96dpi; jsPDF points are 1/72 in
    const pt = (px: number) => (px / 96) * 72
    const doc = new jsPDF({ unit: 'pt', format: [pt(m.width), pt(m.height)], orientation: page.landscape ? 'landscape' : 'portrait' })
    doc.setProperties({ title })
    await doc.html(body, { x: pt(m.margin[3]), y: pt(m.margin[0]), width: pt(m.contentWidth), windowWidth: m.contentWidth, autoPaging: 'text', margin: [pt(m.margin[0]), pt(m.margin[1]), pt(m.margin[2]), pt(m.margin[3])] })
    return doc.output('blob')
  } finally { holder.remove() }
}
