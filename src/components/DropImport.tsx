import { useEffect, useRef, useState } from 'react'
import { UploadCloud } from 'lucide-react'
import { importFile } from '../services/importers'
import { useOpenFile } from '../hooks/useOpenFile'
import { useToast } from '../contexts/ToastContext'
import { fullName } from '../file-system/fileService'

/** Drop files anywhere in the app to import them into Drive. A single file opens straight away. */
export function DropImport() {
  const [over, setOver] = useState(false)
  const depth = useRef(0)
  const open = useOpenFile()
  const toast = useToast()
  const openRef = useRef(open); openRef.current = open
  const toastRef = useRef(toast); toastRef.current = toast

  useEffect(() => {
    const hasFiles = (e: DragEvent) => !!e.dataTransfer && [...e.dataTransfer.types].includes('Files')
    const enter = (e: DragEvent) => { if (!hasFiles(e)) return; e.preventDefault(); depth.current++; setOver(true) }
    const overEv = (e: DragEvent) => { if (hasFiles(e)) e.preventDefault() }
    const leave = (e: DragEvent) => { if (!hasFiles(e)) return; depth.current = Math.max(0, depth.current - 1); if (depth.current === 0) setOver(false) }
    const drop = async (e: DragEvent) => {
      if (!hasFiles(e)) return
      e.preventDefault(); depth.current = 0; setOver(false)
      const files = [...(e.dataTransfer?.files ?? [])]
      for (const f of files) {
        try {
          const imported = await importFile(f)
          toastRef.current(`Imported ${fullName(imported)}`)
          if (files.length === 1) openRef.current(imported)
        } catch (err) { toastRef.current((err as Error).message, 'error') }
      }
    }
    window.addEventListener('dragenter', enter); window.addEventListener('dragover', overEv); window.addEventListener('dragleave', leave); window.addEventListener('drop', drop)
    return () => { window.removeEventListener('dragenter', enter); window.removeEventListener('dragover', overEv); window.removeEventListener('dragleave', leave); window.removeEventListener('drop', drop) }
  }, [])

  if (!over) return null
  return (
    <div aria-hidden data-testid="drop-overlay" className="fixed inset-0 z-[250] bg-[#0b1020]/45 flex items-center justify-center p-8 backdrop-in pointer-events-none">
      <div className="border-2 border-dashed border-white/70 rounded-3xl px-14 py-12 text-white text-center">
        <UploadCloud size={36} className="mx-auto" />
        <div className="mt-3 text-xl font-display">Drop to add to Drive</div>
        <div className="mt-1 text-[13px] text-white/75">Word, Excel, CSV, text, PowerPoint and PDF files</div>
      </div>
    </div>
  )
}
