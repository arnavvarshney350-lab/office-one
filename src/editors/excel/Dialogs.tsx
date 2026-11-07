import { useState } from 'react'
import { Button, Modal } from '../../design-system'
import type { ChartType, CondKind, CondRule } from './model'
import { parseRange } from './sheetOps'
import { uid } from '../../utilities/id'

const inp = 'mt-1 w-full h-9 px-2.5 rounded-md border border-line bg-panel text-[13px]'

export const CF_KINDS: { id: CondKind; label: string; fields: 0 | 1 | 2; hint?: string }[] = [
  { id: 'gt', label: 'Greater than', fields: 1 }, { id: 'lt', label: 'Less than', fields: 1 }, { id: 'eq', label: 'Equal to', fields: 1 },
  { id: 'between', label: 'Between', fields: 2 }, { id: 'text', label: 'Text contains', fields: 1 },
  { id: 'dup', label: 'Duplicate values', fields: 0 }, { id: 'top', label: 'Top N values', fields: 1, hint: 'How many?' },
  { id: 'bottom', label: 'Bottom N values', fields: 1, hint: 'How many?' }, { id: 'scale', label: 'Colour scale', fields: 0 },
]
const FILLS = [['#fecaca', 'Light red'], ['#fde68a', 'Yellow'], ['#bbf7d0', 'Green'], ['#bfdbfe', 'Blue'], ['#e9d5ff', 'Purple'], ['#63be7b', 'Scale green']] as const

export function CondFormatDialog({ range, initial, onSave, onClose }: { range: string; initial?: CondKind; onSave: (r: CondRule) => void; onClose: () => void }) {
  const [kind, setKind] = useState<CondKind>(initial ?? 'gt')
  const [a, setA] = useState(''), [b, setB] = useState('')
  const [rng, setRng] = useState(range)
  const [bg, setBg] = useState<string>(initial === 'scale' ? '#63be7b' : '#fecaca')
  const def = CF_KINDS.find(k => k.id === kind)!
  const valid = !!parseRange(rng) && (def.fields === 0 || a.trim() !== '') && (def.fields < 2 || b.trim() !== '')
  return (
    <Modal title="Conditional formatting" onClose={onClose} width={400} footer={<><Button onClick={onClose}>Cancel</Button>
      <Button variant="primary" disabled={!valid} onClick={() => { onSave({ id: uid(), range: rng.toUpperCase(), kind, a: a.trim(), b: b.trim(), bg }); onClose() }}>Apply</Button></>}>
      <div className="space-y-3 text-[13px]">
        <label className="block">Apply to range<input aria-label="Range" value={rng} onChange={e => setRng(e.target.value)} className={inp} /></label>
        <label className="block">Rule<select aria-label="Rule" value={kind} onChange={e => { const k = e.target.value as CondKind; setKind(k); if (k === 'scale') setBg('#63be7b') }} className={inp}>{CF_KINDS.map(k => <option key={k.id} value={k.id}>{k.label}</option>)}</select></label>
        {def.fields >= 1 && <label className="block">{def.hint ?? 'Value'}<input aria-label="Value" autoFocus value={a} onChange={e => setA(e.target.value)} className={inp} /></label>}
        {def.fields === 2 && <label className="block">and<input aria-label="Second value" value={b} onChange={e => setB(e.target.value)} className={inp} /></label>}
        <div>
          <div className="mb-1">{kind === 'scale' ? 'Highest-value colour' : 'Highlight with'}</div>
          <div className="flex gap-2 flex-wrap">{FILLS.map(([c, n]) => (
            <button key={c} type="button" title={n} aria-label={n} aria-pressed={bg === c} onClick={() => setBg(c)} style={{ background: c }} className={`w-8 h-8 rounded-md border-2 ${bg === c ? 'border-accent' : 'border-line'}`} />
          ))}</div>
        </div>
      </div>
    </Modal>
  )
}

export function ChartDialog({ range, initial, onSave, onClose }: {
  range: string; initial?: { type: ChartType; title: string; range: string }; onSave: (v: { type: ChartType; title: string; range: string }) => void; onClose: () => void
}) {
  const [type, setType] = useState<ChartType>(initial?.type ?? 'column')
  const [title, setTitle] = useState(initial?.title ?? '')
  const [rng, setRng] = useState(initial?.range ?? range)
  const TYPES: [ChartType, string][] = [['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie']]
  return (
    <Modal title={initial ? 'Edit chart' : 'Insert chart'} onClose={onClose} width={400} footer={<><Button onClick={onClose}>Cancel</Button>
      <Button variant="primary" disabled={!parseRange(rng)} onClick={() => { onSave({ type, title: title.trim(), range: rng.toUpperCase() }); onClose() }}>{initial ? 'Save' : 'Insert'}</Button></>}>
      <div className="space-y-3 text-[13px]">
        <label className="block">Data range<input aria-label="Chart data range" autoFocus value={rng} onChange={e => setRng(e.target.value)} className={inp} />
          <span className="text-[12px] text-muted">First row = series names, first column = categories.</span></label>
        <label className="block">Title<input aria-label="Chart title" value={title} onChange={e => setTitle(e.target.value)} className={inp} /></label>
        <div role="radiogroup" aria-label="Chart type" className="flex gap-2 flex-wrap">
          {TYPES.map(([t, n]) => <button key={t} role="radio" aria-checked={type === t} onClick={() => setType(t)} className={`h-8 px-3 rounded-md border text-[13px] ${type === t ? 'bg-accent-soft border-accent text-accent font-semibold' : 'border-line bg-surface hover:bg-panel2'}`}>{n}</button>)}
        </div>
      </div>
    </Modal>
  )
}
