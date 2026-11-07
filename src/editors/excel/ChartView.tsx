import { memo, useMemo } from 'react'
import type { ChartData } from './sheetOps'
import type { ChartType } from './model'

export const CHART_COLORS = ['#2563eb', '#16a34a', '#f59e0b', '#dc2626', '#7c3aed', '#0891b2', '#db2777', '#65a30d']
const fmt = (n: number) => (Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : String(Number(n.toPrecision(4))))

/** Pick "nice" axis ticks covering [min, max]. */
function ticks(min: number, max: number, n = 5): number[] {
  if (min === max) { max = min + 1 }
  const span = max - min, raw = span / n
  const mag = Math.pow(10, Math.floor(Math.log10(raw)))
  const step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => s >= raw) ?? raw
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step
  const out: number[] = []
  for (let v = lo; v <= hi + step / 2; v += step) out.push(Number(v.toPrecision(12)))
  return out
}

/** SVG chart. Pure rendering of already-evaluated data, so it is trivially themeable and printable. */
export const ChartView = memo(function ChartView({ type, data, title, width, height }: { type: ChartType; data: ChartData; title: string; width: number; height: number }) {
  const m = { t: title ? 34 : 14, r: 14, b: 44, l: 48 }
  const w = Math.max(40, width - m.l - m.r), h = Math.max(40, height - m.t - m.b)
  const all = data.series.flatMap(s => s.values)
  const stacked = false
  const { min, max } = useMemo(() => ({ min: Math.min(0, ...all), max: Math.max(0, ...all) }), [all])
  const ty = useMemo(() => ticks(min, max), [min, max])
  const lo = ty[0], hi = ty[ty.length - 1]
  const y = (v: number) => m.t + h - ((v - lo) / (hi - lo || 1)) * h
  const nCat = Math.max(1, data.categories.length), band = w / nCat
  const empty = data.series.length === 0 || nCat === 0

  return (
    <svg width={width} height={height} role="img" aria-label={title || `${type} chart`} style={{ display: 'block', fontFamily: 'inherit' }}>
      {title && <text x={width / 2} y={20} textAnchor="middle" fontSize={14} fontWeight={600} fill="var(--text)">{title}</text>}
      {empty ? <text x={width / 2} y={height / 2} textAnchor="middle" fontSize={12} fill="var(--muted)">No data in range</text> : type === 'pie' ? (
        <Pie data={data} cx={width / 2} cy={m.t + h / 2} r={Math.min(w, h) / 2 - 4} />
      ) : (
        <>
          {ty.map(t => (
            <g key={t}>
              <line x1={m.l} x2={m.l + w} y1={y(t)} y2={y(t)} stroke="var(--border)" strokeWidth={1} />
              <text x={m.l - 6} y={y(t) + 4} textAnchor="end" fontSize={10.5} fill="var(--muted)">{fmt(t)}</text>
            </g>
          ))}
          {data.categories.map((c, i) => (
            <text key={i} x={m.l + band * i + band / 2} y={m.t + h + 16} textAnchor="middle" fontSize={10.5} fill="var(--muted)">{c.length > 10 ? c.slice(0, 9) + '…' : c}</text>
          ))}
          {(type === 'column' || type === 'bar') && data.series.map((s, si) => {
            const bw = Math.max(2, (band * 0.72) / data.series.length)
            return s.values.map((v, i) => {
              const x = m.l + band * i + band * 0.14 + bw * si
              const y0 = y(0), y1 = y(v)
              return <rect key={`${si}-${i}`} x={x} y={Math.min(y0, y1)} width={bw - 1} height={Math.max(1, Math.abs(y0 - y1))} rx={2} fill={CHART_COLORS[si % CHART_COLORS.length]}><title>{`${s.name} · ${data.categories[i]}: ${fmt(v)}`}</title></rect>
            })
          })}
          {(type === 'line' || type === 'area') && data.series.map((s, si) => {
            const pts = s.values.map((v, i) => [m.l + band * i + band / 2, y(v)] as const)
            const d = pts.map(([px, py], i) => `${i ? 'L' : 'M'}${px},${py}`).join(' ')
            const col = CHART_COLORS[si % CHART_COLORS.length]
            return (
              <g key={si}>
                {type === 'area' && pts.length > 0 && <path d={`${d} L${pts[pts.length - 1][0]},${y(0)} L${pts[0][0]},${y(0)} Z`} fill={col} opacity={0.18} />}
                <path d={d} fill="none" stroke={col} strokeWidth={2.2} strokeLinejoin="round" />
                {pts.map(([px, py], i) => <circle key={i} cx={px} cy={py} r={3.2} fill={col}><title>{`${s.name} · ${data.categories[i]}: ${fmt(s.values[i])}`}</title></circle>)}
              </g>
            )
          })}
          <line x1={m.l} x2={m.l + w} y1={y(0)} y2={y(0)} stroke="var(--muted)" strokeWidth={1} />
        </>
      )}
      {!empty && data.series.length > 1 && type !== 'pie' && (
        <g transform={`translate(${m.l}, ${height - 12})`}>
          {data.series.slice(0, 6).map((s, i) => (
            <g key={i} transform={`translate(${i * 90}, 0)`}>
              <rect width={10} height={10} y={-9} rx={2} fill={CHART_COLORS[i % CHART_COLORS.length]} />
              <text x={14} fontSize={10.5} fill="var(--muted)">{s.name.length > 10 ? s.name.slice(0, 9) + '…' : s.name}</text>
            </g>
          ))}
        </g>
      )}
      {stacked && null}
    </svg>
  )
})

function Pie({ data, cx, cy, r }: { data: ChartData; cx: number; cy: number; r: number }) {
  const vals = (data.series[0]?.values ?? []).map(v => Math.max(0, v))
  const total = vals.reduce((a, b) => a + b, 0)
  if (!total) return <text x={cx} y={cy} textAnchor="middle" fontSize={12} fill="var(--muted)">No positive values</text>
  let a0 = -Math.PI / 2
  return (
    <g>
      {vals.map((v, i) => {
        const a1 = a0 + (v / total) * Math.PI * 2
        const big = a1 - a0 > Math.PI ? 1 : 0
        const p = (a: number) => [cx + r * Math.cos(a), cy + r * Math.sin(a)]
        const [x0, y0] = p(a0), [x1, y1] = p(a1)
        const mid = (a0 + a1) / 2
        const d = vals.length === 1 ? `M${cx - r},${cy} a${r},${r} 0 1,0 ${r * 2},0 a${r},${r} 0 1,0 ${-r * 2},0` : `M${cx},${cy} L${x0},${y0} A${r},${r} 0 ${big} 1 ${x1},${y1} Z`
        const res = (
          <g key={i}>
            <path d={d} fill={CHART_COLORS[i % CHART_COLORS.length]} stroke="var(--surface)" strokeWidth={2}><title>{`${data.categories[i]}: ${fmt(vals[i])} (${((v / total) * 100).toFixed(1)}%)`}</title></path>
            {v / total > 0.06 && <text x={cx + r * 0.62 * Math.cos(mid)} y={cy + r * 0.62 * Math.sin(mid) + 4} textAnchor="middle" fontSize={11} fill="#fff" fontWeight={600}>{`${Math.round((v / total) * 100)}%`}</text>}
          </g>
        )
        a0 = a1
        return res
      })}
    </g>
  )
}
