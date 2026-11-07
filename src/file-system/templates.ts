import type { FileKind } from '../types'
import { emptyWorkbook } from '../editors/excel/model'
import { newDeck } from '../editors/powerpoint/model'

export function defaultContent(kind: FileKind): unknown {
  switch (kind) {
    case 'docx': return { html: '<p></p>' }
    case 'xlsx': return emptyWorkbook()
    case 'pptx': return newDeck()
    default: return ''
  }
}
export const KIND_LABEL: Record<FileKind, string> = {
  docx: 'Document', xlsx: 'Spreadsheet', pptx: 'Presentation', pdf: 'PDF', csv: 'CSV', txt: 'Text', folder: 'Folder',
}
export const KIND_DEFAULT_NAME: Record<FileKind, string> = {
  docx: 'Untitled document', xlsx: 'Untitled spreadsheet', pptx: 'Untitled presentation', pdf: 'Untitled', csv: 'Untitled', txt: 'Untitled', folder: 'New folder',
}
