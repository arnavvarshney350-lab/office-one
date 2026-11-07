import { useNavigate } from 'react-router-dom'
import { routeFor } from '../file-system/fileService'
import type { OFile } from '../types'

export function useOpenFile() {
  const nav = useNavigate()
  return (f: OFile) => {
    const r = routeFor(f)
    if (r) { nav(r); return }
    if (f.blob) { const url = URL.createObjectURL(f.blob); window.open(url, '_blank'); setTimeout(() => URL.revokeObjectURL(url), 60000) }
  }
}
