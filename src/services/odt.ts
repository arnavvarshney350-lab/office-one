/**
 * OpenDocument Text (.odt) import. An .odt is a zip whose content.xml holds the text. It is converted to the
 * small HTML subset the Word editor understands: headings, paragraphs, bold/italic/underline/strike,
 * sub/superscript, links, lists, tables and embedded pictures.
 */
const T = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
const TB = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
const ST = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
const FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
const DR = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'
const XL = 'http://www.w3.org/1999/xlink'

interface Fmt { b?: boolean; i?: boolean; u?: boolean; s?: boolean; sub?: boolean; sup?: boolean; align?: string }
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp' }

const kids = (el: Element) => Array.from(el.childNodes)
const local = (n: Node) => (n as Element).localName
const isEl = (n: Node): n is Element => n.nodeType === 1

function readStyles(...docs: (Document | null)[]): Map<string, Fmt> {
  const out = new Map<string, Fmt>()
  for (const d of docs) {
    if (!d) continue
    for (const st of Array.from(d.getElementsByTagNameNS(ST, 'style'))) {
      const name = st.getAttributeNS(ST, 'name'); if (!name) continue
      const f: Fmt = {}
      const tp = st.getElementsByTagNameNS(ST, 'text-properties')[0], pp = st.getElementsByTagNameNS(ST, 'paragraph-properties')[0]
      if (tp) {
        f.b = /bold|[6-9]00/.test(tp.getAttributeNS(FO, 'font-weight') ?? '') || undefined
        f.i = ['italic', 'oblique'].includes(tp.getAttributeNS(FO, 'font-style') ?? '') || undefined
        const u = tp.getAttributeNS(ST, 'text-underline-style'); f.u = u && u !== 'none' ? true : undefined
        const s = tp.getAttributeNS(ST, 'text-line-through-style'); f.s = s && s !== 'none' ? true : undefined
        const pos = tp.getAttributeNS(ST, 'text-position') ?? ''
        if (pos.startsWith('super')) f.sup = true; else if (pos.startsWith('sub')) f.sub = true
      }
      const al = pp?.getAttributeNS(FO, 'text-align'); if (al === 'center' || al === 'end' || al === 'right' || al === 'justify') f.align = al === 'end' ? 'right' : al
      out.set(name, f)
    }
  }
  // Named styles inherit from a parent ("Heading_20_1" -> "Heading"); resolve one level so common bold headings keep their look.
  for (const d of docs) if (d) for (const st of Array.from(d.getElementsByTagNameNS(ST, 'style'))) {
    const name = st.getAttributeNS(ST, 'name'), parent = st.getAttributeNS(ST, 'parent-style-name')
    if (name && parent && out.has(parent)) out.set(name, { ...out.get(parent), ...out.get(name) })
  }
  return out
}

export async function odtToHtml(buf: ArrayBuffer): Promise<string> {
  const { default: JSZip } = await import('jszip')
  let zip: Awaited<ReturnType<typeof JSZip.loadAsync>>
  try { zip = await JSZip.loadAsync(buf) } catch { throw new Error('This is not a valid .odt file') }
  const contentXml = await zip.file('content.xml')?.async('string')
  if (!contentXml) throw new Error('This is not a valid .odt file (content.xml is missing)')
  const parse = (x: string) => new DOMParser().parseFromString(x, 'application/xml')
  const doc = parse(contentXml)
  if (doc.getElementsByTagName('parsererror').length) throw new Error('The .odt file could not be read')
  const stylesXml = await zip.file('styles.xml')?.async('string')
  const styles = readStyles(stylesXml ? parse(stylesXml) : null, doc)

  // Embedded pictures become data URIs so the document stays self-contained.
  const pics = new Map<string, string>()
  for (const path of Object.keys(zip.files)) {
    const ext = path.split('.').pop()?.toLowerCase() ?? ''
    if (/^pictures\//i.test(path) && MIME[ext]) pics.set(path, `data:${MIME[ext]};base64,${await zip.files[path].async('base64')}`)
  }
  const ordered = new Set<string>()
  for (const ls of Array.from(doc.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'list-style'))) {
    const n = ls.getAttributeNS(ST, 'name'); if (n && ls.getElementsByTagNameNS(T, 'list-level-style-number').length) ordered.add(n)
  }

  const out = document.createElement('div')
  const wrap = (parent: Node, tag: string): HTMLElement => parent.appendChild(document.createElement(tag))

  function inline(node: Node, parent: Node, fmt: Fmt) {
    if (node.nodeType === 3) {
      let tgt: Node = parent
      for (const [on, tag] of [[fmt.b, 'strong'], [fmt.i, 'em'], [fmt.u, 'u'], [fmt.s, 's'], [fmt.sub, 'sub'], [fmt.sup, 'sup']] as const) if (on && node.textContent?.trim()) tgt = wrap(tgt, tag)
      tgt.appendChild(document.createTextNode(node.textContent ?? ''))
      return
    }
    if (!isEl(node)) return
    switch (local(node)) {
      case 'span': inline2(node, parent, { ...fmt, ...styles.get(node.getAttributeNS(T, 'style-name') ?? '') }); return
      case 'a': { const href = node.getAttributeNS(XL, 'href'); const a: Node = /^(https?:|mailto:)/i.test(href ?? '') ? wrap(parent, 'a') : parent; if (a !== parent) (a as HTMLElement).setAttribute('href', href!); inline2(node, a, fmt); return }
      case 'line-break': wrap(parent, 'br'); return
      case 'tab': parent.appendChild(document.createTextNode('    ')); return
      case 's': parent.appendChild(document.createTextNode(' '.repeat(Number(node.getAttributeNS(T, 'c') || 1)))); return
      case 'frame': case 'image': {
        const img = local(node) === 'image' ? node : node.getElementsByTagNameNS(DR, 'image')[0]
        const src = img && pics.get((img.getAttributeNS(XL, 'href') ?? '').replace(/^\.\//, ''))
        if (src) wrap(parent, 'img').setAttribute('src', src)
        return
      }
      case 'bookmark': case 'bookmark-start': case 'bookmark-end': case 'soft-page-break': return
      default: inline2(node, parent, fmt)
    }
  }
  function inline2(el: Element, parent: Node, fmt: Fmt) { for (const k of kids(el)) inline(k, parent, fmt) }

  function block(el: Element, parent: Node) {
    switch (local(el)) {
      case 'h': {
        const lvl = Math.min(6, Math.max(1, Number(el.getAttributeNS(T, 'outline-level') || 1)))
        const h = wrap(parent, `h${lvl}`); inline2(el, h, {}); return
      }
      case 'p': {
        const f = styles.get(el.getAttributeNS(T, 'style-name') ?? '') ?? {}
        const p = wrap(parent, 'p') as HTMLElement
        if (f.align) p.style.textAlign = f.align
        inline2(el, p, { b: f.b, i: f.i, u: f.u }); return
      }
      case 'list': {
        const ol = ordered.has(el.getAttributeNS(T, 'style-name') ?? '')
        const list = wrap(parent, ol ? 'ol' : 'ul')
        for (const item of kids(el)) if (isEl(item) && local(item) === 'list-item') {
          const li = wrap(list, 'li')
          for (const c of kids(item)) if (isEl(c)) {
            if (local(c) === 'p' || local(c) === 'h') { const p = wrap(li, 'p'); inline2(c, p, {}) } else block(c, li)
          }
        }
        return
      }
      case 'table': {
        const table = wrap(parent, 'table'), body = wrap(table, 'tbody')
        const owner = (r: Element) => { let p = r.parentElement; while (p && !(p.localName === 'table' && p.namespaceURI === TB)) p = p.parentElement; return p }
        const rows = Array.from(el.getElementsByTagNameNS(TB, 'table-row')).filter(r => owner(r) === el) // not rows of nested tables
        for (const r of rows) {
          const tr = wrap(body, 'tr')
          for (const c of kids(r)) if (isEl(c) && local(c) === 'table-cell') {
            const td = wrap(tr, 'td') as HTMLTableCellElement
            const span = Number(c.getAttributeNS(TB, 'number-columns-spanned') || 1); if (span > 1) td.colSpan = span
            const inner = kids(c).filter(isEl)
            if (!inner.length) wrap(td, 'p')
            for (const b of inner) block(b, td)
            if (!td.firstChild) wrap(td, 'p')
          }
        }
        return
      }
      case 'section': case 'table-of-content': case 'index-body': for (const k of kids(el)) if (isEl(k)) block(k, parent); return
      case 'index-title': return
      default: return
    }
  }
  const body = doc.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'text')[0]
  if (body) for (const k of kids(body)) if (isEl(k)) block(k, out)
  return out.innerHTML || '<p></p>'
}
