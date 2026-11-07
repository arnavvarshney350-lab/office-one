import { useCallback, useRef, useState } from 'react'

/** Immutable-state undo/redo. `commit` records history, `replace` does not. */
export function useHistory<T>(initial: T, limit = 100) {
  const [state, setState] = useState<T>(initial)
  const ref = useRef(state); ref.current = state
  const past = useRef<T[]>([]), future = useRef<T[]>([])
  const [, tick] = useState(0)
  const commit = useCallback((next: T) => {
    past.current.push(ref.current); if (past.current.length > limit) past.current.shift()
    future.current = []; ref.current = next; setState(next)
  }, [limit])
  /** Record `prev` as an undo step without changing current state (end of a live drag). */
  const record = useCallback((prev: T) => { past.current.push(prev); if (past.current.length > limit) past.current.shift(); future.current = []; tick(n => n + 1) }, [limit])
  const replace = useCallback((next: T) => { ref.current = next; setState(next) }, [])
  const undo = useCallback(() => {
    const p = past.current.pop(); if (p === undefined) return false
    future.current.push(ref.current); ref.current = p; setState(p); tick(n => n + 1); return true
  }, [])
  const redo = useCallback(() => {
    const f = future.current.pop(); if (f === undefined) return false
    past.current.push(ref.current); ref.current = f; setState(f); tick(n => n + 1); return true
  }, [])
  return { state, commit, replace, record, undo, redo, canUndo: past.current.length > 0, canRedo: future.current.length > 0 }
}
