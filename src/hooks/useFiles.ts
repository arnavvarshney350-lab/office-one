import { useLiveQuery } from 'dexie-react-hooks'
import { db } from '../storage/db'
import type { OFile } from '../types'

/** Live list of every file record; undefined while loading. */
export function useFiles(): OFile[] | undefined {
  return useLiveQuery(() => db.files.toArray(), [])
}
