import { createFile } from '../file-system/fileService'
import type { FileKind, OFile } from '../types'
import { odtToHtml } from './odt'

const escapeHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
export const textToHtml = (t: string) => t.split(/\r?\n/).map(l => `<p>${escapeHtml(l)}</p>`).join('') || '<p></p>'

export { odtToHtml }
export async function docxToHtml(buf: ArrayBuffer): Promise<string> {
  const mammoth = await import('mammoth')
  const r = await mammoth.convertToHtml({ arrayBuffer: buf })
  return r.value || '<p></p>'
}

/** Import a local file into the shared file system. */
export async function importFile(file: File, parentId: string | null = null): Promise<OFile> {
  const m = /^(.*?)(?:\.([A-Za-z0-9]+))?$/.exec(file.name)!
  const base = m[1] || 'Untitled'
  const ext = (m[2] || '').toLowerCase()
  switch (ext) {
    case 'docx': return createFile('docx', { name: base, parentId, content: { html: await docxToHtml(await file.arrayBuffer()) } })
    case 'odt': return createFile('docx', { name: base, parentId, content: { html: await odtToHtml(await file.arrayBuffer()) } })
    case 'txt': case 'md': return createFile('docx', { name: base, parentId, content: { html: textToHtml(await file.text()) } })
    case 'xlsx': case 'xls': {
      const { workbookFromXlsx } = await import('../editors/excel/io')
      return createFile('xlsx', { name: base, parentId, content: workbookFromXlsx(await file.arrayBuffer()) })
    }
    case 'csv': {
      const { workbookFromCsv } = await import('../editors/excel/io')
      return createFile('xlsx', { name: base, parentId, content: workbookFromCsv(await file.text()) })
    }
    case 'pdf': case 'pptx': return createFile(ext as FileKind, { name: base, parentId, content: null, blob: file })
    default: throw new Error(`Unsupported file type: .${ext || 'unknown'}`)
  }
}
export const IMPORT_ACCEPT = '.docx,.odt,.xlsx,.xls,.csv,.txt,.md,.pdf,.pptx'
