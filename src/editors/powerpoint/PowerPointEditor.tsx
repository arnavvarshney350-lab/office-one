import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import type { ChainedCommands, Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import {
  AlignCenter, AlignEndHorizontal, AlignEndVertical, AlignHorizontalJustifyCenter, AlignJustify, AlignLeft, AlignRight, AlignStartHorizontal, AlignStartVertical, AlignVerticalJustifyCenter,
  ArrowDown, ArrowUp, Bold, Group, Ungroup, StickyNote, ChevronDown, Circle, Copy, Grid3x3, Highlighter, ImagePlus, Italic, MoveRight, Play, Plus, Redo2, Sparkles, Square, Strikethrough, Trash2, Triangle, Type, Underline as UnderlineIcon, Undo2,
} from 'lucide-react'
import { Button, ColorMenu, ContextMenu, Divider, IconButton, Menu, Ribbon, RibbonGroup, type MenuEntry } from '../../design-system'
import { FontPicker, SizeBox } from '../../design-system/FontPicker'
import { EditorHeader, type MenuItem } from '../../components/EditorHeader'
import { useAutosave } from '../../hooks/useAutosave'
import { useHistory } from '../../hooks/useHistory'
import { createFile, saveContent } from '../../file-system/fileService'
import { useToast } from '../../contexts/ToastContext'
import { cssFamily, preloadFonts } from '../../services/fonts'
import { uid } from '../../utilities/id'
import { ObjView, SlideView } from './SlideView'
import { Presenter } from './Presenter'
import { ThemeGallery } from './ThemeGallery'
import { TextBoxEditor } from './TextBoxEditor'
import { TransitionStage } from './TransitionPlayer'
import { applyToAllHtml } from './richText'
import {
  applyTheme, deckFontNames, defaultTransition, DIR_LABEL, fontFor, isDecoObj, LAYOUTS, layoutObjects, makeSlide, newDeck, objHtml, SLIDE_H, SLIDE_W, slideBgCss, THEMES, themeOf, TRANSITIONS,
  type Deck, type LayoutId, type ShapeKind, type Slide, type SlideObj, type Transition, type TransitionType,
} from './model'
import type { OFile } from '../../types'

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const HPOS: Record<Handle, [number, number]> = { nw: [0, 0], n: [.5, 0], ne: [1, 0], e: [1, .5], se: [1, 1], s: [.5, 1], sw: [0, 1], w: [0, .5] }
interface Box { x: number; y: number; w: number; h: number }
type Gesture = { kind: 'move' | 'resize'; handle?: Handle; ids: string[]; origs: Map<string, SlideObj>; bbox: Box; sx: number; sy: number; base: Deck; moved: boolean }
const bboxOf = (os: SlideObj[]): Box => {
  const x = Math.min(...os.map(o => o.x)), y = Math.min(...os.map(o => o.y))
  return { x, y, w: Math.max(...os.map(o => o.x + o.w)) - x, h: Math.max(...os.map(o => o.y + o.h)) - y }
}
const GRID = 10
const LINE_SPACINGS = [1, 1.2, 1.5, 2, 2.5, 3]
const CHAR_SPACINGS: [string, number][] = [['Very tight', -2], ['Tight', -1], ['Normal', 0], ['Loose', 2], ['Very loose', 5]]

const asDeck = (c: unknown): Deck => { const d = c as Deck | null; return d && Array.isArray(d.slides) && d.slides.length ? d : newDeck() }
const plainFont = (v: string | undefined) => (v ?? '').split(',')[0].replace(/["']/g, '').trim()

export function PowerPointEditor({ file }: { file: OFile }) {
  const nav = useNavigate()
  const toast = useToast()
  const hist = useHistory<Deck>(asDeck(file.content))
  const deck = hist.state
  const deckRef = useRef(deck); deckRef.current = deck
  const [cur, setCur] = useState(0)
  const [sel, setSel] = useState<string | null>(null)
  const [extra, setExtra] = useState<string[]>([])
  const [marquee, setMarquee] = useState<Box | null>(null)
  const [notesOpen, setNotesOpen] = useState(false)
  const notesBase = useRef<Deck | null>(null)
  const marqueeRef = useRef<{ sx: number; sy: number; add: boolean; keep: string[] } | null>(null)
  const [editing, setEditing] = useState<string | null>(null)
  const [editor, setEditor] = useState<Editor | null>(null)
  const [tab, setTab] = useState('Home')
  const [version, setVersion] = useState(0)
  const [present, setPresent] = useState(false)
  const [preview, setPreview] = useState(false)
  const [fit, setFit] = useState(0.8)
  const [zoom, setZoom] = useState(0)
  const [snap, setSnap] = useState(true)
  const [showGrid, setShowGrid] = useState(false)
  const [dragIdx, setDragIdx] = useState<number | null>(null)
  const [ctx, setCtx] = useState<{ x: number; y: number; items: MenuEntry[] } | null>(null)
  const stageBox = useRef<HTMLDivElement>(null)
  const gesture = useRef<Gesture | null>(null)
  const editBase = useRef<Deck | null>(null)
  const clipObj = useRef<SlideObj[] | null>(null)
  const imgInput = useRef<HTMLInputElement>(null)
  const stageRef = useRef<HTMLDivElement>(null)

  const slideIdx = Math.min(cur, deck.slides.length - 1)
  const slide = deck.slides[slideIdx]
  const theme = themeOf(deck.theme)
  const selIds = useMemo(() => (sel ? [sel, ...extra] : []), [sel, extra])
  const selected = slide.objects.filter(o => selIds.includes(o.id))
  const multi = selected.length > 1
  /** The single selected object, used by the formatting controls; null for none or several. */
  const selObj = selected.length === 1 ? selected[0] : null
  const selBox = selected.length ? bboxOf(selected) : null
  const grouped = selected.some(o => o.group)
  const clearSel = () => { setSel(null); setExtra([]) }
  /** Select an object (and its whole group). `add` toggles it in or out of the current selection. */
  const selectObj = (o: SlideObj, add = false) => {
    const ids = o.group ? slide.objects.filter(x => x.group === o.group).map(x => x.id) : [o.id]
    if (!add) { setSel(ids[0]); setExtra(ids.slice(1)); return }
    const cur = new Set(selIds)
    const allIn = ids.every(i => cur.has(i))
    ids.forEach(i => (allIn ? cur.delete(i) : cur.add(i)))
    const arr = [...cur]; setSel(arr[0] ?? null); setExtra(arr.slice(1))
  }
  const scale = zoom ? zoom / 100 : fit

  const commit = useCallback((next: Deck) => { hist.commit(next); setVersion(v => v + 1) }, [hist])
  const save = useCallback(async () => { await saveContent(file.id, deckRef.current) }, [file.id])
  const saveState = useAutosave(version, save)
  const undo = () => { if (hist.undo()) setVersion(v => v + 1) }
  const redo = () => { if (hist.redo()) setVersion(v => v + 1) }

  const mapSlide = (fn: (s: Slide) => Slide, base: Deck = deck, idx = slideIdx): Deck => ({ ...base, slides: base.slides.map((s, i) => (i === idx ? fn(s) : s)) })
  const patchObj = (id: string, patch: Partial<SlideObj>) => commit(mapSlide(s => ({ ...s, objects: s.objects.map(o => (o.id === id ? { ...o, ...patch } : o)) })))

  // Fonts a deck already uses must be fetched before first paint, or text reflows after opening.
  useEffect(() => { void preloadFonts(deckFontNames(asDeck(file.content))) }, [file.content])
  useEffect(() => { void preloadFonts(deckFontNames(deckRef.current)) }, [deck.theme])

  useLayoutEffect(() => {
    const el = stageBox.current; if (!el) return
    const f = () => setFit(Math.max(0.2, Math.min((el.clientWidth - 48) / SLIDE_W, (el.clientHeight - 48) / SLIDE_H)))
    f(); const ro = new ResizeObserver(f); ro.observe(el); return () => ro.disconnect()
  }, [])

  /* ---------------- text editing session ---------------- */
  function startEditing(o: SlideObj) {
    if (o.type !== 'text') return
    editBase.current = deckRef.current
    // move box-level bold/italic into the content so the toolbar toggles behave per selection
    if (o.bold || o.italic) {
      const html = applyToAllHtml(objHtml(o), c => { let x = c; if (o.bold) x = x.setBold(); if (o.italic) x = x.setItalic(); return x })
      hist.replace(mapSlide(s => ({ ...s, objects: s.objects.map(x => (x.id === o.id ? { ...x, html, bold: false, italic: false } : x)) }), deckRef.current))
    }
    setSel(o.id); setEditing(o.id)
  }
  const stopEditing = useCallback(() => {
    const base = editBase.current; editBase.current = null
    setEditing(null); setEditor(null)
    if (base && base !== deckRef.current) { hist.record(base); setVersion(v => v + 1) } // the whole editing session is one undo step
  }, [hist])
  function onTextChange(id: string, html: string) {
    hist.replace(mapSlide(s => ({ ...s, objects: s.objects.map(o => (o.id === id ? { ...o, html } : o)) }), deckRef.current)); setVersion(v => v + 1)
  }

  const es = useEditorState({
    editor,
    selector: ({ editor: e }) => !e ? null : ({
      bold: e.isActive('bold'), italic: e.isActive('italic'), underline: e.isActive('underline'), strike: e.isActive('strike'),
      font: plainFont(e.getAttributes('textStyle').fontFamily as string | undefined), size: Number(String(e.getAttributes('textStyle').fontSize ?? '').replace('px', '')) || 0,
      aL: e.isActive({ textAlign: 'left' }), aC: e.isActive({ textAlign: 'center' }), aR: e.isActive({ textAlign: 'right' }), aJ: e.isActive({ textAlign: 'justify' }),
    }),
  })
  const isText = selObj?.type === 'text'
  const live = editing === selObj?.id && es ? es : null
  const shown = {
    font: live?.font || (selObj ? fontFor(selObj, theme) : theme.font), size: live?.size || selObj?.fontSize || 24,
    bold: live ? live.bold : !!selObj?.bold, italic: live ? live.italic : !!selObj?.italic, underline: !!live?.underline, strike: !!live?.strike,
    align: live ? (live.aC ? 'center' : live.aR ? 'right' : live.aJ ? 'justify' : live.aL ? 'left' : selObj?.align ?? 'left') : selObj?.align ?? 'left',
  }

  /** Apply formatting to the editor selection, or to the whole box when it is merely selected. */
  function format(fn: (c: ChainedCommands) => ChainedCommands, boxPatch: Partial<SlideObj> = {}) {
    if (editing && editor) { fn(editor.chain().focus()).run(); return }
    if (selObj?.type !== 'text') return
    patchObj(selObj.id, { html: applyToAllHtml(objHtml(selObj), fn), ...boxPatch })
  }
  const setFont = (name: string) => format(c => (name ? c.setFontFamily(cssFamily(name)) : c.unsetFontFamily()), { fontFamily: name || undefined })
  const setSize = (n: number) => format(c => c.setFontSize(`${n}px`), { fontSize: n })
  const toggleMark = (m: 'bold' | 'italic' | 'underline' | 'strike') => {
    if (!editing && (m === 'bold' || m === 'italic') && selObj) { patchObj(selObj.id, { [m]: !selObj[m] }); return }
    format(c => (m === 'bold' ? c.toggleBold() : m === 'italic' ? c.toggleItalic() : m === 'underline' ? c.toggleUnderline() : c.toggleStrike()))
  }
  const setAlign = (a: 'left' | 'center' | 'right' | 'justify') => format(c => c.setTextAlign(a), { align: a })

  /* ---------------- slide operations ---------------- */
  function addSlide(layout: LayoutId = 'content', at = slideIdx + 1) { const s = makeSlide(layout, deck.theme); const slides = [...deck.slides]; slides.splice(at, 0, s); commit({ ...deck, slides }); setCur(at); clearSel() }
  function duplicateSlide(i = slideIdx) {
    const src = deck.slides[i]; const copy: Slide = { ...structuredClone(src), id: uid() }; copy.objects = copy.objects.map(o => ({ ...o, id: uid() }))
    const slides = [...deck.slides]; slides.splice(i + 1, 0, copy); commit({ ...deck, slides }); setCur(i + 1)
  }
  function deleteSlide(i = slideIdx) {
    if (deck.slides.length < 2) { toast('A presentation needs at least one slide', 'error'); return }
    commit({ ...deck, slides: deck.slides.filter((_, j) => j !== i) }); setCur(Math.max(0, Math.min(slideIdx, deck.slides.length - 2))); clearSel()
  }
  function moveSlide(from: number, to: number) {
    if (to < 0 || to >= deck.slides.length || from === to) return
    const slides = [...deck.slides]; const [m] = slides.splice(from, 1); slides.splice(to, 0, m); commit({ ...deck, slides }); setCur(to)
  }
  function setLayout(l: LayoutId) {
    const kept = slide.objects.filter(o => !o.role && !isDecoObj(o, slide.layout, theme))
    commit(mapSlide(s => ({ ...s, layout: l, objects: [...layoutObjects(l, theme), ...kept] }))); clearSel()
  }
  function setTransition(t: Transition | undefined, all = false) {
    commit({ ...deck, slides: deck.slides.map((s, i) => (all || i === slideIdx ? { ...s, transition: t && t.type !== 'none' ? t : undefined } : s)) })
  }
  const trans = slide.transition ?? defaultTransition('none')
  const transDef = TRANSITIONS.find(t => t.id === trans.type)!
  const pickTransition = (type: TransitionType) => { const d = defaultTransition(type); setTransition(type === 'none' ? undefined : { ...d, duration: slide.transition?.duration ?? d.duration }) }

  /* ---------------- object operations ---------------- */
  const newObj = (o: Omit<SlideObj, 'id'>) => { const full = { ...o, id: uid() } as SlideObj; commit(mapSlide(s => ({ ...s, objects: [...s.objects, full] }))); setSel(full.id); return full }
  function addText() { const o = newObj({ type: 'text', x: 300, y: 220, w: 360, h: 80, text: 'Text box', fontSize: 28, color: theme.body, align: 'left' }); startEditing(o) }
  function addShape(shape: ShapeKind) { newObj({ type: 'shape', shape, x: 380, y: 200, w: 200, h: shape === 'arrow' ? 120 : 140, fill: theme.accent }) }
  function addImage(f: File | undefined) {
    if (!f) return
    const r = new FileReader()
    r.onload = () => {
      const src = String(r.result); const im = new Image()
      im.onload = () => { const k = Math.min(1, 520 / im.naturalWidth, 380 / im.naturalHeight); const w = Math.round(im.naturalWidth * k), h = Math.round(im.naturalHeight * k)
        const o: SlideObj = { id: uid(), type: 'image', src, x: Math.round((SLIDE_W - w) / 2), y: Math.round((SLIDE_H - h) / 2), w, h }
        commit(mapSlide(s => ({ ...s, objects: [...s.objects, o] }), deckRef.current)); setSel(o.id) }
      im.onerror = () => toast('That image could not be read', 'error')
      im.src = src
    }
    r.readAsDataURL(f)
  }
  const patchMany = (ids: string[], fn: (o: SlideObj) => Partial<SlideObj>) =>
    commit(mapSlide(sl => ({ ...sl, objects: sl.objects.map(o => (ids.includes(o.id) ? { ...o, ...fn(o) } : o)) })))
  const deleteObj = () => { if (!selected.length) return; setEditing(null); commit(mapSlide(sl => ({ ...sl, objects: sl.objects.filter(o => !selIds.includes(o.id)) }))); clearSel() }
  const duplicateObj = () => {
    if (!selected.length) return
    const groups = new Map<string, string>()
    const copies = selected.map(o => ({ ...o, id: uid(), x: o.x + 20, y: o.y + 20, group: o.group ? (groups.get(o.group) ?? (groups.set(o.group, uid()), groups.get(o.group))) : undefined }))
    commit(mapSlide(sl => ({ ...sl, objects: [...sl.objects, ...copies] })))
    setSel(copies[0].id); setExtra(copies.slice(1).map(c => c.id))
  }
  function layer(kind: 'front' | 'forward' | 'backward' | 'back') {
    if (!selected.length) return
    commit(mapSlide(sl => {
      const rest = sl.objects.filter(o => !selIds.includes(o.id)), block = sl.objects.filter(o => selIds.includes(o.id))
      const first = sl.objects.findIndex(o => selIds.includes(o.id))
      const at = kind === 'front' ? rest.length : kind === 'back' ? 0 : kind === 'forward' ? Math.min(rest.length, first + 1) : Math.max(0, first - 1)
      return { ...sl, objects: [...rest.slice(0, at), ...block, ...rest.slice(at)] }
    }))
  }
  /** Align to the slide (one object) or to each other (several objects). */
  function alignTo(a: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') {
    if (!selected.length || !selBox) return
    const ref: Box = multi ? selBox : { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H }
    patchMany(selIds, o => a === 'left' ? { x: ref.x } : a === 'center' ? { x: Math.round(ref.x + (ref.w - o.w) / 2) } : a === 'right' ? { x: ref.x + ref.w - o.w } : a === 'top' ? { y: ref.y } : a === 'middle' ? { y: Math.round(ref.y + (ref.h - o.h) / 2) } : { y: ref.y + ref.h - o.h })
  }
  function distribute(axis: 'h' | 'v') {
    if (selected.length < 3) { toast('Select three or more objects to distribute', 'error'); return }
    const sorted = [...selected].sort((p, q) => (axis === 'h' ? p.x - q.x : p.y - q.y))
    const first = sorted[0], last = sorted[sorted.length - 1]
    const span = axis === 'h' ? last.x + last.w - first.x : last.y + last.h - first.y
    const total = sorted.reduce((t, o) => t + (axis === 'h' ? o.w : o.h), 0)
    const gap = (span - total) / (sorted.length - 1)
    let pos = axis === 'h' ? first.x : first.y
    const place = new Map<string, number>()
    for (const o of sorted) { place.set(o.id, Math.round(pos)); pos += (axis === 'h' ? o.w : o.h) + gap }
    patchMany(selIds, o => (axis === 'h' ? { x: place.get(o.id)! } : { y: place.get(o.id)! }))
  }
  function groupSel() { if (selected.length < 2) return; const g = uid(); patchMany(selIds, () => ({ group: g })); toast('Grouped') }
  function ungroupSel() { patchMany(selIds, () => ({ group: undefined })); toast('Ungrouped') }
  function selectAll() { const ids = slide.objects.map(o => o.id); setSel(ids[0] ?? null); setExtra(ids.slice(1)) }

  /* ---------------- pointer gestures ---------------- */
  const snapV = (n: number, alt: boolean) => (snap && !alt ? Math.round(n / GRID) * GRID : Math.round(n))
  function startGesture(e: React.PointerEvent, o: SlideObj, kind: Gesture['kind'], handle?: Handle) {
    if (editing === o.id) return
    e.stopPropagation()
    if (editing) stopEditing()
    if (kind === 'move' && (e.shiftKey || e.metaKey || e.ctrlKey)) { selectObj(o, true); return } // modifier-click toggles membership
    try { (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId) } catch { /* capture is best-effort */ }
    // dragging an unselected object selects it (with its group); dragging a selected one moves the whole selection
    let ids = selIds
    if (kind === 'move' && !selIds.includes(o.id)) { ids = o.group ? slide.objects.filter(x => x.group === o.group).map(x => x.id) : [o.id]; setSel(ids[0]); setExtra(ids.slice(1)) }
    const objs = slide.objects.filter(x => ids.includes(x.id))
    gesture.current = { kind, handle, ids, origs: new Map(objs.map(x => [x.id, x])), bbox: bboxOf(objs), sx: e.clientX, sy: e.clientY, base: deckRef.current, moved: false }
  }
  function onPointerMove(e: React.PointerEvent) {
    const m = marqueeRef.current
    if (m) { updateMarquee(e); return }
    const g = gesture.current; if (!g) return
    const dx = (e.clientX - g.sx) / scale, dy = (e.clientY - g.sy) / scale
    if (!g.moved && Math.abs(dx) + Math.abs(dy) < 3 / scale) return
    g.moved = true
    const b = g.bbox
    let nx = b.x, ny = b.y, nw = b.w, nh = b.h
    if (g.kind === 'move') { nx = snapV(b.x + dx, e.altKey); ny = snapV(b.y + dy, e.altKey) } else {
      const hd = g.handle!
      let l = b.x, t = b.y, r = b.x + b.w, bt = b.y + b.h
      if (hd.includes('e')) r = snapV(b.x + b.w + dx, e.altKey)
      if (hd.includes('w')) l = snapV(b.x + dx, e.altKey)
      if (hd.includes('s')) bt = snapV(b.y + b.h + dy, e.altKey)
      if (hd.includes('n')) t = snapV(b.y + dy, e.altKey)
      if (r - l < 20) { if (hd.includes('w')) l = r - 20; else r = l + 20 }
      if (bt - t < 20) { if (hd.includes('n')) t = bt - 20; else bt = t + 20 }
      nx = l; ny = t; nw = r - l; nh = bt - t
    }
    const fx = nw / (b.w || 1), fy = nh / (b.h || 1)
    hist.replace(mapSlide(sl => ({
      ...sl, objects: sl.objects.map(ob => {
        const o0 = g.origs.get(ob.id); if (!o0) return ob
        return g.kind === 'move'
          ? { ...ob, x: Math.round(o0.x + (nx - b.x)), y: Math.round(o0.y + (ny - b.y)) }
          : { ...ob, x: Math.round(nx + (o0.x - b.x) * fx), y: Math.round(ny + (o0.y - b.y) * fy), w: Math.max(4, Math.round(o0.w * fx)), h: Math.max(4, Math.round(o0.h * fy)) }
      }),
    }), g.base))
  }
  function endGesture() {
    if (marqueeRef.current) { marqueeRef.current = null; setMarquee(null); return }
    const g = gesture.current; gesture.current = null
    if (g?.moved) { hist.record(g.base); setVersion(v => v + 1) }
  }
  /** Rubber-band selection started on empty slide space. */
  function startMarquee(e: React.PointerEvent) {
    if (e.button !== 0) return
    if (editing) stopEditing()
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const sx = (e.clientX - rect.left) / scale, sy = (e.clientY - rect.top) / scale
    const add = e.shiftKey || e.metaKey || e.ctrlKey
    marqueeRef.current = { sx, sy, add, keep: add ? selIds : [] }
    if (!add) clearSel()
    setMarquee({ x: sx, y: sy, w: 0, h: 0 })
    try { (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId) } catch { /* best effort */ }
  }
  function updateMarquee(e: React.PointerEvent) {
    const m = marqueeRef.current; if (!m) return
    const host = stageRef.current; if (!host) return
    const r = host.getBoundingClientRect()
    const x = (e.clientX - r.left) / scale, y = (e.clientY - r.top) / scale
    const box = { x: Math.min(m.sx, x), y: Math.min(m.sy, y), w: Math.abs(x - m.sx), h: Math.abs(y - m.sy) }
    setMarquee(box)
    const hit = slide.objects.filter(o => o.x < box.x + box.w && o.x + o.w > box.x && o.y < box.y + box.h && o.y + o.h > box.y)
    const ids = new Set(m.keep)
    for (const o of hit) (o.group ? slide.objects.filter(q => q.group === o.group) : [o]).forEach(q => ids.add(q.id))
    const arr = [...ids]; setSel(arr[0] ?? null); setExtra(arr.slice(1))
  }

  /* ---------------- keyboard (attached once, reads latest handler) ---------------- */
  const onKey = useRef<(e: KeyboardEvent) => void>(() => {})
  onKey.current = (e: KeyboardEvent) => {
    const t = e.target as HTMLElement
    if (present || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable) return
    const mod = e.metaKey || e.ctrlKey, k = e.key.toLowerCase()
    if (mod && k === 'z') { e.preventDefault(); if (e.shiftKey) redo(); else undo() }
    else if (mod && k === 'y') { e.preventDefault(); redo() }
    else if (mod && k === 's') { e.preventDefault(); void save().then(() => toast('Saved')) }
    else if (mod && k === 'a') { e.preventDefault(); selectAll() }
    else if (mod && k === 'g') { e.preventDefault(); if (e.shiftKey) ungroupSel(); else groupSel() }
    else if (mod && k === 'c' && selected.length) clipObj.current = selected
    else if (mod && k === 'v' && clipObj.current?.length) {
      const groups = new Map<string, string>()
      const copies = clipObj.current.map(o => ({ ...o, id: uid(), x: o.x + 20, y: o.y + 20, group: o.group ? (groups.get(o.group) ?? (groups.set(o.group, uid()), groups.get(o.group))) : undefined }))
      commit(mapSlide(s => ({ ...s, objects: [...s.objects, ...copies] }))); setSel(copies[0].id); setExtra(copies.slice(1).map(c => c.id)); clipObj.current = copies
    }
    else if (mod && k === 'd') { e.preventDefault(); if (selected.length) duplicateObj(); else duplicateSlide() }
    else if ((e.key === 'Delete' || e.key === 'Backspace') && selected.length) { e.preventDefault(); deleteObj() }
    else if (e.key === 'Enter' && selObj?.type === 'text') { e.preventDefault(); startEditing(selObj) }
    else if (e.key === 'Escape') clearSel()
    else if (selected.length && e.key.startsWith('Arrow')) { e.preventDefault(); const d = e.shiftKey ? GRID : 1; const dx = e.key === 'ArrowRight' ? d : e.key === 'ArrowLeft' ? -d : 0, dy = e.key === 'ArrowDown' ? d : e.key === 'ArrowUp' ? -d : 0; patchMany(selIds, o => ({ x: o.x + dx, y: o.y + dy })) }
    else if (!selected.length && (e.key === 'PageDown' || e.key === 'ArrowDown' || e.key === 'ArrowRight')) { e.preventDefault(); setCur(c => Math.min(deck.slides.length - 1, c + 1)) }
    else if (!selected.length && (e.key === 'PageUp' || e.key === 'ArrowUp' || e.key === 'ArrowLeft')) { e.preventDefault(); setCur(c => Math.max(0, c - 1)) }
  }
  useEffect(() => { const h = (e: KeyboardEvent) => onKey.current(e); window.addEventListener('keydown', h); return () => window.removeEventListener('keydown', h) }, [])

  async function doExport(kind: 'pptx' | 'pdf') {
    try { const io = await import('./io'); toast(`Preparing ${kind.toUpperCase()}…`); if (kind === 'pptx') await io.exportPptx(file.name, deckRef.current); else await io.exportPdf(file.name, deckRef.current); toast(`Exported ${kind.toUpperCase()}`) }
    catch (e) { toast(`Export failed: ${(e as Error).message}`, 'error') }
  }
  const fileMenu: MenuItem[] = [
    { label: 'New presentation', onClick: async () => { const n = await createFile('pptx'); nav(`/powerpoint/${n.id}`) } },
    { label: 'Save now', onClick: () => { void save().then(() => toast('Saved')) } },
    { label: 'Export as PPTX', onClick: () => doExport('pptx') },
    { label: 'Export as PDF', onClick: () => doExport('pdf') },
  ]
  const exitPresent = useCallback(() => setPresent(false), [])
  const slideMenu = (i: number): MenuEntry[] => [
    { label: 'New slide', onClick: () => addSlide('content', i + 1) }, { label: 'Duplicate slide', onClick: () => duplicateSlide(i) },
    { label: 'Move up', disabled: i === 0, onClick: () => moveSlide(i, i - 1) }, { label: 'Move down', disabled: i === deck.slides.length - 1, onClick: () => moveSlide(i, i + 1) },
    'sep', { label: 'Delete slide', danger: true, onClick: () => deleteSlide(i) },
  ]
  const objMenu = (o: SlideObj): MenuEntry[] => [
    ...(o.type === 'text' ? [{ label: 'Edit text', onClick: () => startEditing(o) } as MenuEntry, 'sep' as const] : []),
    { label: 'Duplicate', hint: '⌘D', onClick: duplicateObj }, { label: 'Bring to front', onClick: () => layer('front') }, { label: 'Send to back', onClick: () => layer('back') },
    ...(multi ? [{ label: 'Group', hint: '⌘G', onClick: groupSel } as MenuEntry] : []), ...(grouped ? [{ label: 'Ungroup', hint: '⇧⌘G', onClick: ungroupSel } as MenuEntry] : []), 'sep',
    { label: 'Delete', danger: true, hint: 'Del', onClick: deleteObj },
  ]

  const num = (v: number, set: (n: number) => void, label: string) => (
    <label className="flex items-center gap-1 text-[12px] text-muted">{label}
      <input aria-label={label} type="number" value={Math.round(v)} onChange={e => set(Number(e.target.value))} className="w-16 h-7 px-1.5 rounded border border-line bg-surface text-ink" /></label>)
  const textDisabled = !isText
  const FormatControls = (
    <>
      <FontPicker value={isText ? shown.font : ''} onPick={setFont} width={140} disabled={textDisabled} />
      <SizeBox value={isText ? shown.size : ''} disabled={textDisabled} onPick={setSize} />
      <IconButton label="Bold (⌘B)" disabled={textDisabled} active={shown.bold} onClick={() => toggleMark('bold')}><Bold size={16} /></IconButton>
      <IconButton label="Italic (⌘I)" disabled={textDisabled} active={shown.italic} onClick={() => toggleMark('italic')}><Italic size={16} /></IconButton>
      <IconButton label="Underline (⌘U)" disabled={textDisabled} active={shown.underline} onClick={() => toggleMark('underline')}><UnderlineIcon size={16} /></IconButton>
      <IconButton label="Strikethrough" disabled={textDisabled} active={shown.strike} onClick={() => toggleMark('strike')}><Strikethrough size={16} /></IconButton>
      <ColorMenu label="Font colour" icon={<Type size={16} />} onPick={c => c && format(x => x.setColor(c), { color: c })} />
      <ColorMenu label="Highlight colour" icon={<Highlighter size={16} />} allowNone onPick={c => format(x => (c ? x.setHighlight({ color: c }) : x.unsetHighlight()))} />
    </>
  )

  return (
    <div className="h-full flex flex-col min-h-0">
      <EditorHeader file={file} color="var(--color-ppt)" appName="PowerPoint" save={saveState} fileMenu={fileMenu}>
        <Button variant="primary" onClick={() => { if (editing) stopEditing(); setPresent(true) }} aria-label="Present from current slide"><Play size={14} />Present</Button>
      </EditorHeader>
      <Ribbon tabs={['Home', 'Insert', 'Design', 'Transitions', 'View']} tab={tab} onTab={setTab}>
        {tab === 'Home' && (<>
          <RibbonGroup label="Undo">
            <IconButton label="Undo (⌘Z)" disabled={!hist.canUndo} onClick={undo}><Undo2 size={16} /></IconButton>
            <IconButton label="Redo (⇧⌘Z)" disabled={!hist.canRedo} onClick={redo}><Redo2 size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Slides">
            <Menu trigger={<><Plus size={15} />New slide<ChevronDown size={12} /></>} items={LAYOUTS.map(l => ({ label: l.name, onClick: () => addSlide(l.id) }))} />
            <IconButton label="Duplicate slide" onClick={() => duplicateSlide()}><Copy size={16} /></IconButton>
            <IconButton label="Delete slide" onClick={() => deleteSlide()}><Trash2 size={16} /></IconButton>
            <IconButton label="Move slide up" disabled={slideIdx === 0} onClick={() => moveSlide(slideIdx, slideIdx - 1)}><ArrowUp size={16} /></IconButton>
            <IconButton label="Move slide down" disabled={slideIdx === deck.slides.length - 1} onClick={() => moveSlide(slideIdx, slideIdx + 1)}><ArrowDown size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Font">{FormatControls}</RibbonGroup>
          <RibbonGroup label="Paragraph">
            <IconButton label="Align left" disabled={textDisabled} active={isText && shown.align === 'left'} onClick={() => setAlign('left')}><AlignLeft size={16} /></IconButton>
            <IconButton label="Centre" disabled={textDisabled} active={isText && shown.align === 'center'} onClick={() => setAlign('center')}><AlignCenter size={16} /></IconButton>
            <IconButton label="Align right" disabled={textDisabled} active={isText && shown.align === 'right'} onClick={() => setAlign('right')}><AlignRight size={16} /></IconButton>
            <IconButton label="Justify" disabled={textDisabled} active={isText && shown.align === 'justify'} onClick={() => setAlign('justify')}><AlignJustify size={16} /></IconButton>
            <Menu title="Line spacing" trigger={<>Line<ChevronDown size={12} /></>} items={LINE_SPACINGS.map(v => ({ label: `${v.toFixed(v % 1 ? 2 : 1)}${selObj?.lineHeight === v || (!selObj?.lineHeight && v === 1.2) ? '  ✓' : ''}`, disabled: textDisabled, onClick: () => selObj && patchObj(selObj.id, { lineHeight: v }) }))} />
            <Menu title="Character spacing" trigger={<>Spacing<ChevronDown size={12} /></>} items={CHAR_SPACINGS.map(([label, v]) => ({ label: `${label}${(selObj?.letterSpacing ?? 0) === v ? '  ✓' : ''}`, disabled: textDisabled, onClick: () => selObj && patchObj(selObj.id, { letterSpacing: v || undefined }) }))} />
          </RibbonGroup>
          <RibbonGroup label="Arrange">
            <Menu title="Order" trigger={<>Order<ChevronDown size={12} /></>} items={[{ label: 'Bring to front', onClick: () => layer('front'), disabled: !selected.length }, { label: 'Bring forward', onClick: () => layer('forward'), disabled: !selected.length }, { label: 'Send backward', onClick: () => layer('backward'), disabled: !selected.length }, { label: 'Send to back', onClick: () => layer('back'), disabled: !selected.length }]} />
            <Menu title={multi ? 'Align objects' : 'Align to slide'} trigger={<>Align<ChevronDown size={12} /></>} items={[
              { label: 'Left', icon: <AlignStartVertical size={14} />, onClick: () => alignTo('left'), disabled: !selected.length }, { label: 'Centre', icon: <AlignHorizontalJustifyCenter size={14} />, onClick: () => alignTo('center'), disabled: !selected.length }, { label: 'Right', icon: <AlignEndVertical size={14} />, onClick: () => alignTo('right'), disabled: !selected.length },
              { label: 'Top', icon: <AlignStartHorizontal size={14} />, onClick: () => alignTo('top'), disabled: !selected.length }, { label: 'Middle', icon: <AlignVerticalJustifyCenter size={14} />, onClick: () => alignTo('middle'), disabled: !selected.length }, { label: 'Bottom', icon: <AlignEndHorizontal size={14} />, onClick: () => alignTo('bottom'), disabled: !selected.length },
              'sep', { label: 'Distribute horizontally', onClick: () => distribute('h'), disabled: selected.length < 3 }, { label: 'Distribute vertically', onClick: () => distribute('v'), disabled: selected.length < 3 }]} />
            <IconButton label="Group (⌘G)" disabled={!multi} onClick={groupSel}><Group size={16} /></IconButton>
            <IconButton label="Ungroup (⇧⌘G)" disabled={!grouped} onClick={ungroupSel}><Ungroup size={16} /></IconButton>
            <IconButton label="Duplicate object (⌘D)" disabled={!selected.length} onClick={duplicateObj}><Copy size={16} /></IconButton>
            <IconButton label="Delete object" disabled={!selected.length} onClick={deleteObj}><Trash2 size={16} /></IconButton>
          </RibbonGroup>
        </>)}
        {tab === 'Insert' && (<>
          <RibbonGroup label="Text"><Button variant="ghost" onClick={addText}><Type size={15} />Text box</Button></RibbonGroup>
          <RibbonGroup label="Shapes">
            <IconButton label="Rectangle" onClick={() => addShape('rect')}><Square size={16} /></IconButton>
            <IconButton label="Ellipse" onClick={() => addShape('ellipse')}><Circle size={16} /></IconButton>
            <IconButton label="Triangle" onClick={() => addShape('triangle')}><Triangle size={16} /></IconButton>
            <IconButton label="Arrow" onClick={() => addShape('arrow')}><MoveRight size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Images"><Button variant="ghost" onClick={() => imgInput.current?.click()}><ImagePlus size={15} />Picture</Button></RibbonGroup>
        </>)}
        {tab === 'Design' && (<>
          <RibbonGroup label="Themes">
            <ThemeGallery slide={slide} current={deck.theme} onPick={id => commit(applyTheme(deck, id))} />
          </RibbonGroup>
          <RibbonGroup label="Background">
            <ColorMenu label="Slide background" icon={<span className="text-[13px] px-1">This slide</span>} onPick={c => c && commit(mapSlide(s => ({ ...s, bg: c })))} />
            <ColorMenu label="Background for all slides" icon={<span className="text-[13px] px-1">All slides</span>} onPick={c => c && commit({ ...deck, slides: deck.slides.map(s => ({ ...s, bg: c })) })} />
          </RibbonGroup>
        </>)}
        {tab === 'Transitions' && (<>
          <RibbonGroup label="Transition to this slide">
            <div className="flex flex-wrap gap-1 max-w-[520px]" role="radiogroup" aria-label="Transition">
              {TRANSITIONS.map(t => (
                <button key={t.id} role="radio" aria-checked={trans.type === t.id} title={t.hint} onClick={() => pickTransition(t.id)}
                  className={`h-8 px-2.5 rounded-md border text-[12.5px] ${trans.type === t.id ? 'bg-accent-soft border-accent text-accent font-semibold' : 'border-line bg-surface hover:bg-panel2'}`}>{t.name}</button>
              ))}
            </div>
          </RibbonGroup>
          <RibbonGroup label="Timing">
            {transDef.directions.length > 0 && (
              <select aria-label="Transition direction" value={trans.direction} onChange={e => setTransition({ ...trans, direction: e.target.value as Transition['direction'] })} className="h-8 rounded-md border border-line bg-surface px-1.5 text-[13px]">
                {transDef.directions.map(d => <option key={d} value={d}>{DIR_LABEL[d]}</option>)}
              </select>)}
            <label className="flex items-center gap-1.5 text-[12.5px] text-muted">Duration
              <input aria-label="Transition duration (seconds)" type="number" min={0.1} max={5} step={0.1} disabled={trans.type === 'none'} value={(trans.duration / 1000).toFixed(1)}
                onChange={e => { const s = Math.min(5, Math.max(0.1, Number(e.target.value) || 0.1)); setTransition({ ...trans, duration: Math.round(s * 1000) }) }} className="w-[64px] h-8 px-2 rounded-md border border-line bg-surface text-ink" />s</label>
          </RibbonGroup>
          <RibbonGroup label="Apply">
            <Button variant="ghost" disabled={trans.type === 'none'} onClick={() => setPreview(true)} title="Preview this transition"><Sparkles size={15} />Preview</Button>
            <Button variant="ghost" disabled={trans.type === 'none'} onClick={() => { setTransition(trans, true); toast('Transition applied to all slides') }}>Apply to all</Button>
            <Button variant="ghost" disabled={trans.type === 'none'} onClick={() => setTransition(undefined)}>Remove</Button>
          </RibbonGroup>
        </>)}
        {tab === 'View' && (<>
          <RibbonGroup label="Show">
            <Button variant="ghost" aria-pressed={showGrid} onClick={() => setShowGrid(g => !g)} className={showGrid ? 'bg-accent-soft text-accent' : ''}><Grid3x3 size={15} />Grid</Button>
            <Button variant="ghost" aria-pressed={snap} onClick={() => setSnap(s => !s)} className={snap ? 'bg-accent-soft text-accent' : ''}>Snap to grid</Button>
          </RibbonGroup>
          <RibbonGroup label="Zoom">
            <select aria-label="Zoom" value={zoom} onChange={e => setZoom(Number(e.target.value))} className="h-8 rounded-md border border-line bg-surface px-1.5 text-[13px]">
              <option value={0}>Fit to window</option>{[50, 75, 100, 125, 150].map(z => <option key={z} value={z}>{z}%</option>)}
            </select>
          </RibbonGroup>
        </>)}
      </Ribbon>

      <div className="flex-1 min-h-0 flex">
        <aside aria-label="Slides" className="w-44 shrink-0 bg-panel border-r border-line overflow-y-auto p-3 space-y-3" data-testid="thumbs">
          {deck.slides.map((s, i) => (
            <div key={s.id} draggable onDragStart={() => setDragIdx(i)} onDragOver={e => e.preventDefault()} onDrop={() => { if (dragIdx !== null) moveSlide(dragIdx, i); setDragIdx(null) }}
              onContextMenu={e => { e.preventDefault(); setCur(i); setCtx({ x: e.clientX, y: e.clientY, items: slideMenu(i) }) }}
              className="flex gap-1.5 items-start cursor-pointer" onClick={() => { if (editing) stopEditing(); setCur(i); clearSel() }} data-testid="thumb" aria-label={`Slide ${i + 1}`}>
              <span className="text-[11px] text-muted w-4 pt-1">{i + 1}</span>
              <div className={`relative border-2 rounded overflow-hidden ${i === slideIdx ? 'border-accent shadow-soft' : 'border-line'}`}>
                <SlideView slide={s} deckTheme={deck.theme} scale={0.14} />
                {s.transition && <span data-testid="thumb-transition" title={`Transition: ${s.transition.type}`} className="absolute left-1 bottom-1 bg-accent text-accent-fg rounded px-1 py-px flex items-center gap-0.5 text-[9px]"><Sparkles size={9} />{s.transition.type}</span>}
              </div>
            </div>
          ))}
          <button onClick={() => addSlide('content', deck.slides.length)} className="w-full h-8 rounded border border-dashed border-line text-[12px] text-muted hover:bg-panel2 flex items-center justify-center gap-1"><Plus size={13} />Add slide</button>
        </aside>

        <div ref={stageBox} className="flex-1 min-w-0 overflow-auto flex" style={{ background: 'var(--canvas)' }} onPointerDown={() => { if (editing) stopEditing(); clearSel() }}>
          <div className="m-auto p-6">
            <div ref={stageRef} className="shadow-soft relative" style={{ width: SLIDE_W * scale, height: SLIDE_H * scale, background: slide.bg, overflow: 'hidden' }} data-testid="stage">
              <div style={{ width: SLIDE_W, height: SLIDE_H, transform: `scale(${scale})`, transformOrigin: 'top left', position: 'absolute', ...slideBgCss(slide, theme) }}
                onPointerDown={e => { e.stopPropagation(); startMarquee(e) }} onPointerMove={onPointerMove} onPointerUp={endGesture} onPointerCancel={endGesture}>
                {showGrid && <div aria-hidden style={{ position: 'absolute', inset: 0, pointerEvents: 'none', backgroundImage: 'linear-gradient(to right, rgba(120,120,140,.25) 1px, transparent 1px), linear-gradient(to bottom, rgba(120,120,140,.25) 1px, transparent 1px)', backgroundSize: `${GRID * 4}px ${GRID * 4}px` }} />}
                {slide.objects.map(o => (
                  <ObjViewEditable key={o.id} o={o} font={fontFor(o, theme)} selected={selIds.includes(o.id) && editing !== o.id} editing={editing === o.id}
                    onPointerDown={e => startGesture(e, o, 'move')} onDoubleClick={() => startEditing(o)}
                    onContextMenu={e => { e.preventDefault(); e.stopPropagation(); if (!selIds.includes(o.id)) selectObj(o); setCtx({ x: e.clientX, y: e.clientY, items: objMenu(o) }) }}
                    onTextChange={html => onTextChange(o.id, html)} onEditor={setEditor} onExit={stopEditing} />
                ))}
                {selBox && !editing && HANDLES.map(h => (
                  <div key={h} aria-label={`Resize ${h}`} onPointerDown={e => { e.stopPropagation(); startGesture(e, selected[0], 'resize', h) }}
                    style={{ position: 'absolute', left: selBox.x + HPOS[h][0] * selBox.w - 6, top: selBox.y + HPOS[h][1] * selBox.h - 6, width: 12, height: 12, background: '#fff', border: '2px solid var(--accent)', borderRadius: 3, cursor: `${h}-resize`, touchAction: 'none' }} />
                ))}
                {multi && selBox && !editing && <div aria-hidden data-testid="selection-box" style={{ position: 'absolute', left: selBox.x, top: selBox.y, width: selBox.w, height: selBox.h, border: '1.5px dashed var(--accent)', pointerEvents: 'none' }} />}
                {marquee && marquee.w + marquee.h > 2 && <div aria-hidden data-testid="marquee" style={{ position: 'absolute', left: marquee.x, top: marquee.y, width: marquee.w, height: marquee.h, border: '1px solid var(--accent)', background: 'color-mix(in srgb, var(--accent) 14%, transparent)', pointerEvents: 'none' }} />}
              </div>
              {preview && slide.transition && (
                <div className="absolute inset-0" data-testid="transition-preview">
                  <TransitionStage from={slideIdx > 0 ? deck.slides[slideIdx - 1] : { ...slide, bg: '#000000', objects: [] }} to={slide} transition={slide.transition} reverse={false} scale={scale} deckTheme={deck.theme} onDone={() => setPreview(false)} />
                </div>
              )}
            </div>
          </div>
        </div>

        <aside aria-label="Properties" className="hidden lg:block w-64 shrink-0 bg-panel border-l border-line overflow-y-auto p-4 space-y-4 text-[13px]">
          {multi ? (<>
            <div className="font-semibold" data-testid="multi-title">{selected.length} objects selected{grouped ? ' (grouped)' : ''}</div>
            <div className="flex gap-1 flex-wrap">
              <IconButton label="Align left edge" onClick={() => alignTo('left')}><AlignStartVertical size={15} /></IconButton><IconButton label="Align centre" onClick={() => alignTo('center')}><AlignHorizontalJustifyCenter size={15} /></IconButton><IconButton label="Align right edge" onClick={() => alignTo('right')}><AlignEndVertical size={15} /></IconButton>
              <IconButton label="Align top edge" onClick={() => alignTo('top')}><AlignStartHorizontal size={15} /></IconButton><IconButton label="Align middle" onClick={() => alignTo('middle')}><AlignVerticalJustifyCenter size={15} /></IconButton><IconButton label="Align bottom edge" onClick={() => alignTo('bottom')}><AlignEndHorizontal size={15} /></IconButton>
            </div>
            <div className="flex gap-2 flex-wrap">
              <Button onClick={() => distribute('h')} disabled={selected.length < 3}>Distribute H</Button><Button onClick={() => distribute('v')} disabled={selected.length < 3}>Distribute V</Button>
              <Button onClick={groupSel}><Group size={14} />Group</Button><Button onClick={ungroupSel} disabled={!grouped}><Ungroup size={14} />Ungroup</Button>
              <Button onClick={duplicateObj}><Copy size={14} />Duplicate</Button><Button variant="danger" onClick={deleteObj}><Trash2 size={14} />Delete</Button>
            </div>
            <p className="text-muted text-[12px]">Drag any selected object to move them together, or use the handles to resize the whole selection.</p>
          </>) : !selObj ? (<>
            <div className="font-semibold">Slide {slideIdx + 1}</div>
            <label className="block">Layout<select aria-label="Layout" value={slide.layout} onChange={e => setLayout(e.target.value as LayoutId)} className="mt-1 w-full h-8 rounded border border-line bg-surface px-1.5">{LAYOUTS.map(l => <option key={l.id} value={l.id}>{l.name}</option>)}</select></label>
            <label className="block">Theme<select aria-label="Theme" value={deck.theme} onChange={e => commit(applyTheme(deck, e.target.value))} className="mt-1 w-full h-8 rounded border border-line bg-surface px-1.5">{THEMES.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}</select></label>
            <div className="flex items-center justify-between">Background<ColorMenu label="Slide background colour" icon={<span className="w-5 h-5 rounded border border-line" style={{ background: slide.bg }} />} onPick={c => c && commit(mapSlide(s => ({ ...s, bg: c })))} /></div>
            <Button className="w-full justify-center" onClick={() => commit({ ...deck, slides: deck.slides.map(s => ({ ...s, bg: slide.bg })) })}>Apply background to all</Button>
            <div className="text-muted text-[12px]">Transition: <b className="text-ink">{slide.transition ? `${TRANSITIONS.find(t => t.id === slide.transition!.type)?.name}, ${(slide.transition.duration / 1000).toFixed(1)}s` : 'None'}</b></div>
            <p className="text-muted text-[12px]">Select an object to edit it. Double-click text (or press Enter) to type.</p>
          </>) : (<>
            <div className="font-semibold capitalize">{selObj.type === 'text' ? 'Text box' : selObj.type}</div>
            {selObj.type === 'text' && (
              <div className="space-y-2.5">
                <div className="flex items-center gap-1.5 flex-wrap">{FormatControls}</div>
                <div className="flex items-center gap-1">
                  <IconButton label="Align left" active={shown.align === 'left'} onClick={() => setAlign('left')}><AlignLeft size={15} /></IconButton>
                  <IconButton label="Centre" active={shown.align === 'center'} onClick={() => setAlign('center')}><AlignCenter size={15} /></IconButton>
                  <IconButton label="Align right" active={shown.align === 'right'} onClick={() => setAlign('right')}><AlignRight size={15} /></IconButton>
                  <IconButton label="Justify" active={shown.align === 'justify'} onClick={() => setAlign('justify')}><AlignJustify size={15} /></IconButton>
                </div>
                <label className="flex items-center justify-between">Line spacing
                  <select aria-label="Line spacing" value={selObj.lineHeight ?? 1.2} onChange={e => patchObj(selObj.id, { lineHeight: Number(e.target.value) })} className="h-8 w-24 rounded border border-line bg-surface px-1.5">{LINE_SPACINGS.map(v => <option key={v} value={v}>{v}</option>)}</select></label>
                <label className="flex items-center justify-between">Character spacing
                  <select aria-label="Character spacing" value={selObj.letterSpacing ?? 0} onChange={e => patchObj(selObj.id, { letterSpacing: Number(e.target.value) || undefined })} className="h-8 w-24 rounded border border-line bg-surface px-1.5">{CHAR_SPACINGS.map(([l, v]) => <option key={l} value={v}>{l}</option>)}</select></label>
                <Button onClick={() => startEditing(selObj)}>Edit text</Button>
              </div>)}
            {selObj.type === 'shape' && <div className="flex items-center justify-between">Fill<ColorMenu label="Fill colour" icon={<span className="w-5 h-5 rounded border border-line" style={{ background: selObj.fill }} />} onPick={c => c && patchObj(selObj.id, { fill: c })} /></div>}
            <Divider />
            <div className="grid grid-cols-2 gap-2">
              {num(selObj.x, v => patchObj(selObj.id, { x: v }), 'X')}{num(selObj.y, v => patchObj(selObj.id, { y: v }), 'Y')}
              {num(selObj.w, v => patchObj(selObj.id, { w: Math.max(10, v) }), 'W')}{num(selObj.h, v => patchObj(selObj.id, { h: Math.max(10, v) }), 'H')}
            </div>
            <div className="flex gap-1 flex-wrap">
              <IconButton label="Align left edge" onClick={() => alignTo('left')}><AlignStartVertical size={15} /></IconButton><IconButton label="Align centre" onClick={() => alignTo('center')}><AlignHorizontalJustifyCenter size={15} /></IconButton><IconButton label="Align right edge" onClick={() => alignTo('right')}><AlignEndVertical size={15} /></IconButton>
              <IconButton label="Align top edge" onClick={() => alignTo('top')}><AlignStartHorizontal size={15} /></IconButton><IconButton label="Align middle" onClick={() => alignTo('middle')}><AlignVerticalJustifyCenter size={15} /></IconButton><IconButton label="Align bottom edge" onClick={() => alignTo('bottom')}><AlignEndHorizontal size={15} /></IconButton>
            </div>
            <div className="flex gap-2 flex-wrap">
              <Button onClick={() => layer('front')}>Bring to front</Button><Button onClick={() => layer('back')}>Send to back</Button>
              <Button onClick={duplicateObj}><Copy size={14} />Duplicate</Button><Button variant="danger" onClick={deleteObj}><Trash2 size={14} />Delete</Button>
            </div>
          </>)}
        </aside>
      </div>
      {notesOpen && (
        <div className="shrink-0 border-t border-line bg-surface px-4 py-2" data-testid="notes-pane">
          <label className="block text-[12px] text-muted mb-1" htmlFor="slide-notes">Speaker notes for slide {slideIdx + 1}</label>
          <textarea id="slide-notes" aria-label="Speaker notes" rows={3} value={slide.notes ?? ''} placeholder="Click to add notes. They appear in presenter view and are exported to PPTX."
            onFocus={() => { notesBase.current = deckRef.current }}
            onChange={e => { hist.replace(mapSlide(sl => ({ ...sl, notes: e.target.value }), deckRef.current)); setVersion(v => v + 1) }}
            onBlur={() => { const b = notesBase.current; notesBase.current = null; if (b && b !== deckRef.current) { hist.record(b); setVersion(v => v + 1) } }}
            className="w-full resize-y min-h-[56px] max-h-[220px] px-3 py-2 rounded-lg border border-line bg-panel text-[13px]" />
        </div>
      )}
      <div className="h-7 shrink-0 bg-panel border-t border-line flex items-center gap-4 px-4 text-[12px] text-muted">
        <span>Slide {slideIdx + 1} of {deck.slides.length}</span><span>{theme.name}</span>{snap && <span>Snap to grid</span>}
        <button onClick={() => setNotesOpen(o => !o)} aria-pressed={notesOpen} aria-label="Show notes pane" className={`flex items-center gap-1 px-2 h-5 rounded ${notesOpen ? 'bg-accent-soft text-accent' : 'hover:bg-panel2'}`}><StickyNote size={12} />Notes{slide.notes?.trim() ? ' •' : ''}</button>
        <div className="flex-1" />
        <IconButton label="Previous slide" disabled={slideIdx === 0} onClick={() => setCur(slideIdx - 1)}>‹</IconButton>
        <IconButton label="Next slide" disabled={slideIdx === deck.slides.length - 1} onClick={() => setCur(slideIdx + 1)}>›</IconButton>
      </div>
      <input ref={imgInput} type="file" accept="image/*" hidden data-testid="ppt-image-input" onChange={e => { addImage(e.target.files?.[0]); e.target.value = '' }} />
      {ctx && <ContextMenu x={ctx.x} y={ctx.y} items={ctx.items} onClose={() => setCtx(null)} />}
      {present && <Presenter deck={deck} start={slideIdx} onExit={exitPresent} />}
    </div>
  )
}

function ObjViewEditable({ o, font, selected, editing, onPointerDown, onDoubleClick, onContextMenu, onTextChange, onEditor, onExit }: {
  o: SlideObj; font: string; selected: boolean; editing: boolean
  onPointerDown: (e: React.PointerEvent) => void; onDoubleClick: () => void; onContextMenu: (e: React.MouseEvent) => void
  onTextChange: (html: string) => void; onEditor: (e: Editor | null) => void; onExit: () => void
}) {
  return (
    <div onDoubleClick={onDoubleClick} onContextMenu={onContextMenu} style={{ display: 'contents' }}>
      <ObjView o={o} font={font} selected={selected} onPointerDown={onPointerDown}>
        {editing && o.type === 'text' ? <TextBoxEditor o={o} font={font} onChange={onTextChange} onEditor={onEditor} onExit={onExit} /> : undefined}
      </ObjView>
    </div>
  )
}
