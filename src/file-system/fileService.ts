import { db } from '../storage/db'
import type { FileKind, OFile } from '../types'
import { uid } from '../utilities/id'
import { defaultContent, KIND_DEFAULT_NAME } from './templates'

function sizeOf(content: unknown, blob?: Blob): number {
  if (blob) return blob.size
  if (content == null) return 0
  try { return new Blob([typeof content === 'string' ? content : JSON.stringify(content)]).size } catch { return 0 }
}

export async function createFile(kind: FileKind, opts: { name?: string; parentId?: string | null; content?: unknown; blob?: Blob } = {}): Promise<OFile> {
  const now = Date.now()
  const content = opts.content ?? defaultContent(kind)
  const f: OFile = {
    id: uid(), name: opts.name ?? KIND_DEFAULT_NAME[kind], kind, parentId: opts.parentId ?? null,
    content, blob: opts.blob, size: kind === 'folder' ? 0 : sizeOf(content, opts.blob),
    createdAt: now, updatedAt: now, openedAt: kind === 'folder' ? null : now, favourite: 0, trashed: 0,
  }
  await db.files.add(f)
  return f
}

export const getFile = (id: string) => db.files.get(id)

export async function saveContent(id: string, content: unknown): Promise<void> {
  const n = await db.files.update(id, { content, size: sizeOf(content), updatedAt: Date.now() })
  if (!n) throw new Error('File no longer exists')
}
export const renameFile = (id: string, name: string) => db.files.update(id, { name: name.trim() || 'Untitled', updatedAt: Date.now() })
export const touchOpened = (id: string) => db.files.update(id, { openedAt: Date.now() })
export const toggleFavourite = async (f: OFile) => { await db.files.update(f.id, { favourite: f.favourite ? 0 : 1 }) }
export const trashFile = (id: string) => db.files.update(id, { trashed: 1 })
export const restoreFile = (id: string) => db.files.update(id, { trashed: 0 })

async function descendants(id: string): Promise<string[]> {
  const kids = await db.files.where('parentId').equals(id).toArray()
  const out: string[] = []
  for (const k of kids) { out.push(k.id); if (k.kind === 'folder') out.push(...(await descendants(k.id))) }
  return out
}
export async function trashRecursive(f: OFile) {
  const ids = f.kind === 'folder' ? [f.id, ...(await descendants(f.id))] : [f.id]
  await db.files.where('id').anyOf(ids).modify({ trashed: 1 })
}
/** Undo of trashRecursive: bring a folder back together with everything that was inside it. */
export async function restoreRecursive(f: OFile) {
  const ids = f.kind === 'folder' ? [f.id, ...(await descendants(f.id))] : [f.id]
  await db.files.where('id').anyOf(ids).modify({ trashed: 0 })
}
export async function deleteForever(f: OFile) {
  const ids = f.kind === 'folder' ? [f.id, ...(await descendants(f.id))] : [f.id]
  await db.files.bulkDelete(ids)
}
export async function emptyTrash() {
  await db.files.where('trashed').equals(1).delete()
}
export async function duplicateFile(f: OFile): Promise<OFile> {
  if (f.kind === 'folder') return createFile('folder', { name: `${f.name} copy`, parentId: f.parentId })
  const copy = await createFile(f.kind, { name: `${f.name} copy`, parentId: f.parentId, content: structuredClone(f.content), blob: f.blob })
  return copy
}
export function routeFor(f: OFile): string | null {
  switch (f.kind) {
    case 'docx': case 'txt': return `/word/${f.id}`
    case 'xlsx': case 'csv': return `/excel/${f.id}`
    case 'pptx': return `/powerpoint/${f.id}`
    default: return null
  }
}
export const EXT: Record<FileKind, string> = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx', pdf: '.pdf', csv: '.csv', txt: '.txt', folder: '' }
export const fullName = (f: OFile) => f.name + EXT[f.kind]
