import Dexie, { type EntityTable } from 'dexie'
import type { OFile } from '../types'

export interface Meta { key: string; value: unknown }

class OfficeDB extends Dexie {
  files!: EntityTable<OFile, 'id'>
  meta!: EntityTable<Meta, 'key'>
  constructor() {
    super('office-one')
    this.version(1).stores({
      files: 'id, parentId, kind, trashed, favourite, updatedAt, openedAt',
      emails: 'id, folder, date',
      events: 'id, date',
      meta: 'key',
    })
    // v2: mail and calendar were removed from the product; drop their tables from existing profiles.
    this.version(2).stores({ files: 'id, parentId, kind, trashed, favourite, updatedAt, openedAt', emails: null, events: null, meta: 'key' })
  }
}
export const db = new OfficeDB()
