import { useMemo, useRef } from 'react'
import { Evaluator } from './formula'
import type { Cell, Workbook } from './model'

/**
 * Keeps one Evaluator alive across edits. When only cell contents changed, just the
 * edited cells (and their dependents) are invalidated; any structural change rebuilds it.
 */
export function useEvaluator(wb: Workbook): Evaluator {
  const ref = useRef<{ ev: Evaluator; wb: Workbook } | null>(null)
  return useMemo(() => {
    const prev = ref.current
    const sameShape = prev && prev.wb.sheets.length === wb.sheets.length && prev.wb.sheets.every((s, i) => s.id === wb.sheets[i].id && s.name === wb.sheets[i].name && s.rows === wb.sheets[i].rows)
    if (prev && sameShape) {
      prev.ev.setWorkbook(wb)
      wb.sheets.forEach((s, si) => {
        const old = prev.wb.sheets[si].cells
        if (old === s.cells) return
        const addrs = new Set([...Object.keys(old), ...Object.keys(s.cells)])
        for (const a of addrs) { const x: Cell | undefined = old[a], y: Cell | undefined = s.cells[a]; if (x !== y && x?.v !== y?.v) prev.ev.invalidate(si, a) }
      })
      prev.wb = wb
      return prev.ev
    }
    const ev = new Evaluator(wb)
    ref.current = { ev, wb }
    return ev
  }, [wb])
}
