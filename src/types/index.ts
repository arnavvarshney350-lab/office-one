export type FileKind = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'csv' | 'txt' | 'folder'

export interface OFile {
  id: string
  name: string
  kind: FileKind
  parentId: string | null
  /** Editable payload: {html} for docx, Workbook for xlsx, Deck for pptx, string for txt/csv */
  content: unknown
  /** Binary payload for opaque files (pdf, imported pptx) */
  blob?: Blob
  size: number
  createdAt: number
  updatedAt: number
  openedAt: number | null
  favourite: 0 | 1
  trashed: 0 | 1
}

export type SaveState = 'idle' | 'saving' | 'saved' | 'error'
