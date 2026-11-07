/**
 * Font catalogue: fonts installed on this device plus the full Google Fonts library.
 *
 * Google families are never bundled. A family's stylesheet is requested from Google Fonts the
 * first time it is previewed or applied, and a font is only reported as usable once it has been
 * observed to render (its text metrics differ from the generic fallback). If Google Fonts cannot
 * be reached, those families are hidden rather than offered and silently substituted.
 */
import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from 'react'
import { GOOGLE_FONTS, type GFCategory, type GoogleFont } from './googleFonts'

export type FontCategory = GFCategory
export interface FontDef { name: string; category: FontCategory; source: 'system' | 'google'; weights: number[]; italic: boolean; rank: number }

/** Fonts expected to be installed locally (Office and OS defaults). Each is verified before use. */
const SYSTEM_CANDIDATES: [string, FontCategory][] = [
  ['Aptos', 'sans'], ['Aptos Display', 'sans'], ['Calibri', 'sans'], ['Arial', 'sans'], ['Arial Black', 'display'],
  ['Helvetica', 'sans'], ['Helvetica Neue', 'sans'], ['Verdana', 'sans'], ['Tahoma', 'sans'], ['Trebuchet MS', 'sans'],
  ['Segoe UI', 'sans'], ['Avenir', 'sans'], ['Futura', 'sans'], ['Gill Sans', 'sans'], ['Optima', 'sans'],
  ['Times New Roman', 'serif'], ['Georgia', 'serif'], ['Garamond', 'serif'], ['Palatino', 'serif'], ['Baskerville', 'serif'],
  ['Cambria', 'serif'], ['Book Antiqua', 'serif'], ['Didot', 'serif'], ['Hoefler Text', 'serif'], ['Rockwell', 'serif'],
  ['Courier New', 'mono'], ['Menlo', 'mono'], ['Monaco', 'mono'], ['Consolas', 'mono'], ['SF Mono', 'mono'],
  ['Comic Sans MS', 'handwriting'], ['Bradley Hand', 'handwriting'], ['Brush Script MT', 'handwriting'], ['Snell Roundhand', 'handwriting'],
  ['Impact', 'display'], ['Copperplate', 'display'], ['Papyrus', 'display'], ['Chalkduster', 'display'], ['Luminari', 'display'],
]

const GENERIC: Record<FontCategory, string> = { sans: 'sans-serif', serif: 'serif', display: 'sans-serif', handwriting: 'cursive', mono: 'monospace' }
const byName = new Map<string, FontDef>()

/** CSS font-family value with a generic fallback appropriate to the family's category. */
export function cssFamily(name: string): string {
  if (!name) return ''
  const generic = byName.get(name)?.category ?? 'sans'
  return `"${name}", ${GENERIC[generic]}`
}

/* ------------------------------------------------- render detection */
let ctx: CanvasRenderingContext2D | null = null
function textWidth(font: string): number {
  ctx ??= document.createElement('canvas').getContext('2d')
  if (!ctx) return 0
  ctx.font = `72px ${font}`
  return ctx.measureText('mmmmmmmmmmlliWQ@#0123456789wi').width
}
/** A font renders when substituting it changes metrics against every generic fallback. */
export function isRenderable(name: string): boolean {
  const q = `"${name}"`
  return (['monospace', 'serif', 'sans-serif', 'cursive'] as const).some(g => Math.abs(textWidth(`${q}, ${g}`) - textWidth(g)) > 0.01)
}

/* ------------------------------------------------- Google Fonts loading */
type LoadState = 'idle' | 'loading' | 'ready' | 'failed'
const loadState = new Map<string, LoadState>()
const listeners = new Set<() => void>()
let networkOk: boolean | null = null
const notify = () => listeners.forEach(l => l())

function gfHref(f: GoogleFont): string {
  const weights = [...new Set([400, 700, ...f.weights])].filter(w => f.weights.includes(w)).sort((a, b) => a - b)
  const list = weights.length ? weights : [400]
  const spec = f.italic
    ? `ital,wght@${list.map(w => `0,${w}`).join(';')};${list.map(w => `1,${w}`).join(';')}`
    : `wght@${list.join(';')}`
  return `https://fonts.googleapis.com/css2?family=${encodeURIComponent(f.name).replace(/%20/g, '+')}:${spec}&display=swap`
}

/** Request a Google family's stylesheet once and resolve when its faces are usable. */
export function loadFont(name: string): Promise<boolean> {
  const def = byName.get(name)
  if (!def) return Promise.resolve(false)
  if (def.source === 'system') return Promise.resolve(isRenderable(name))
  const state = loadState.get(name)
  if (state === 'ready') return Promise.resolve(true)
  if (state === 'failed') return Promise.resolve(false)
  if (state === 'loading') return new Promise(res => { const t = setInterval(() => { const s = loadState.get(name); if (s === 'ready' || s === 'failed') { clearInterval(t); res(s === 'ready') } }, 80) })

  loadState.set(name, 'loading'); notify()
  const gf = GOOGLE_FONTS.find(g => g.name === name)!
  const id = 'gf-' + name.replace(/\W+/g, '-')
  return new Promise<boolean>(resolve => {
    let settled = false
    const finish = async (ok: boolean) => {
      if (settled) return
      settled = true
      if (ok) {
        try { await Promise.all([document.fonts.load(`16px "${name}"`), document.fonts.load(`700 16px "${name}"`)]) } catch { /* face list may be unavailable */ }
      }
      const good = ok && isRenderable(name)
      loadState.set(name, good ? 'ready' : 'failed')
      if (good) networkOk = true
      notify()
      resolve(good)
    }
    let link = document.getElementById(id) as HTMLLinkElement | null
    if (link) { void finish(true); return }
    link = document.createElement('link')
    link.id = id; link.rel = 'stylesheet'; link.href = gfHref(gf)
    link.onload = () => void finish(true)
    link.onerror = () => void finish(false)
    document.head.appendChild(link)
    setTimeout(() => void finish(document.fonts.check(`16px "${name}"`)), 8000) // offline or blocked
  })
}

/** Probe whether Google Fonts is reachable at all, so an offline session hides those families. */
let probe: Promise<boolean> | null = null
export function googleFontsReachable(): Promise<boolean> {
  probe ??= loadFont('Inter').then(ok => { networkOk = ok; notify(); return ok })
  return probe
}

/* ------------------------------------------------- catalogue */
let catalogue: FontDef[] | null = null
let systemFonts: FontDef[] = []

/** Build the catalogue once: verified system fonts first, then the Google library by popularity. */
export function buildCatalogue(): FontDef[] {
  if (catalogue) return catalogue
  systemFonts = SYSTEM_CANDIDATES
    .filter(([n]) => isRenderable(n))
    .map(([name, category], rank) => ({ name, category, source: 'system' as const, weights: [400, 700], italic: true, rank }))
  const sysNames = new Set(systemFonts.map(f => f.name))
  const google = GOOGLE_FONTS
    .filter(g => !sysNames.has(g.name))
    .map(g => ({ name: g.name, category: g.category, source: 'google' as const, weights: g.weights, italic: g.italic, rank: g.rank }))
  catalogue = [...systemFonts, ...google]
  byName.clear()
  for (const f of catalogue) byName.set(f.name, f)
  return catalogue
}

const subscribe = (cb: () => void) => { listeners.add(cb); return () => { listeners.delete(cb) } }
const snapshot = () => `${networkOk}|${loadState.size}`

export interface FontsView {
  fonts: FontDef[]
  systemCount: number
  googleCount: number
  /** null while the Google Fonts probe is still running. */
  online: boolean | null
  stateOf: (name: string) => LoadState
  ensure: (name: string) => Promise<boolean>
}

/** Font catalogue for pickers: system fonts always, Google families once reachable. */
export function useFonts(): FontsView {
  useSyncExternalStore(subscribe, snapshot, snapshot)
  const all = useMemo(() => buildCatalogue(), [])
  const [, force] = useState(0)
  useEffect(() => { void googleFontsReachable().then(() => force(n => n + 1)) }, [])
  const ensure = useCallback((name: string) => loadFont(name), [])
  const fonts = useMemo(() => (networkOk === false ? all.filter(f => f.source === 'system') : all), [all])
  return {
    fonts,
    systemCount: systemFonts.length,
    googleCount: networkOk === false ? 0 : all.length - systemFonts.length,
    online: networkOk,
    stateOf: name => loadState.get(name) ?? 'idle',
    ensure,
  }
}

/** Load every font a document references, so it renders correctly straight after opening. */
export async function preloadFonts(names: Iterable<string>) {
  buildCatalogue()
  const seen = new Set<string>()
  await Promise.all([...names].filter(n => n && !seen.has(n) && seen.add(n)).map(n => loadFont(n).catch(() => false)))
}
