import { useEffect, useState } from 'react'
import { getFile, touchOpened } from '../file-system/fileService'
import type { OFile } from '../types'

/** Loads a file once by id (editors own their in-memory state afterwards). */
export function useFileLoader(id: string | undefined): OFile | null | 'missing' {
  const [file, setFile] = useState<OFile | null | 'missing'>(null)
  useEffect(() => {
    let cancelled = false
    setFile(null)
    if (!id) { setFile('missing'); return }
    getFile(id).then(f => {
      if (cancelled) return
      if (!f || f.trashed) { setFile('missing'); return }
      void touchOpened(id); setFile(f)
    }).catch(() => !cancelled && setFile('missing'))
    return () => { cancelled = true }
  }, [id])
  return file
}
