import { useEffect, useRef, useState } from 'react'
import type { SaveState } from '../types'

const RETRY_MS = [1500, 4000, 10000]

/**
 * Debounced autosave. `version` changes whenever the user edits; the first render is never saved.
 * - A failed save is retried with backoff before the badge settles on "Save failed".
 * - Pending work is flushed on unmount/pagehide, and closing the tab while a save is pending or has
 *   failed raises the browser's "leave site?" prompt so edits are never silently lost.
 */
export function useAutosave(version: number, save: () => Promise<void>, delay = 700): SaveState {
  const [state, setState] = useState<SaveState>('idle')
  const saveRef = useRef(save)
  saveRef.current = save
  const pending = useRef(false)
  const failed = useRef(false)
  const first = useRef(true)

  useEffect(() => {
    if (first.current) { first.current = false; return }
    pending.current = true
    setState('saving')
    let cancelled = false
    let retry: ReturnType<typeof setTimeout> | undefined
    const attempt = async (n: number) => {
      try { await saveRef.current(); if (cancelled) return; pending.current = false; failed.current = false; setState('saved') }
      catch {
        if (cancelled) return
        if (n < RETRY_MS.length) { retry = setTimeout(() => void attempt(n + 1), RETRY_MS[n]); return }
        failed.current = true; setState('error')
      }
    }
    const t = setTimeout(() => void attempt(0), delay)
    return () => { cancelled = true; clearTimeout(t); clearTimeout(retry) }
  }, [version, delay])

  useEffect(() => {
    const flush = () => { if (pending.current) { pending.current = false; void saveRef.current().catch(() => { failed.current = true }) } }
    const guard = (e: BeforeUnloadEvent) => { if (pending.current || failed.current) { e.preventDefault(); e.returnValue = '' } }
    window.addEventListener('pagehide', flush)
    window.addEventListener('beforeunload', guard)
    return () => { window.removeEventListener('pagehide', flush); window.removeEventListener('beforeunload', guard); flush() }
  }, [])
  return state
}
