import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { EditorContent, useEditor, useEditorState } from '@tiptap/react'
import { StarterKit } from '@tiptap/starter-kit'
import { TextStyleKit } from '@tiptap/extension-text-style'
import Highlight from '@tiptap/extension-highlight'
import TextAlign from '@tiptap/extension-text-align'
import Image from '@tiptap/extension-image'
import Subscript from '@tiptap/extension-subscript'
import Superscript from '@tiptap/extension-superscript'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, ArrowDownUp, Bold, CaseSensitive, ChevronDown, ChevronUp, Columns3, Eye, FileDown, Highlighter, ImagePlus, Indent, Italic,
  LayoutList, Link2, List, ListChecks, ListOrdered, Minus, Outdent, PaintRoller, Plus, Printer, Redo2, RemoveFormatting, Replace, Ruler, Search, SeparatorHorizontal, Sigma, Smile,
  Strikethrough, Subscript as SubIcon, Superscript as SupIcon, Table2, TextQuote, Type, Underline as UnderlineIcon, Undo2, X,
} from 'lucide-react'
import { Button, ColorMenu, ConfirmModal, ContextMenu, Divider, IconButton, Menu, Modal, PromptModal, Ribbon, RibbonGroup, cx, type MenuEntry } from '../../design-system'
import { FontPicker, SizeBox } from '../../design-system/FontPicker'
import { cssFamily, preloadFonts } from '../../services/fonts'
import { EditorHeader, type MenuItem } from '../../components/EditorHeader'
import { useAutosave } from '../../hooks/useAutosave'
import { createFile, saveContent } from '../../file-system/fileService'
import { download, safeName } from '../../services/download'
import { docxToHtml, odtToHtml, textToHtml } from '../../services/importers'
import { useToast } from '../../contexts/ToastContext'
import { Find, findKey } from './find'
import { exportDocx, exportPdf } from './io'
import { DEFAULT_PAGE, MARGINS, PAPERS, pageMetrics, readDoc, type MarginId, type PageSetup, type PaperId } from './page'
import type { OFile } from '../../types'

const LINE_SPACINGS = ['1', '1.15', '1.5', '2', '2.5', '3']
const SYMBOLS = ['©', '®', '™', '°', '±', '×', '÷', '≠', '≤', '≥', '≈', '∞', 'µ', 'π', 'Σ', 'Ω', '√', '∫', '†', '‡', '§', '¶', '•', '–', '—', '…', '‹', '›', '«', '»', '“', '”', '‘', '’', '€', '£', '¥', '¢', '₹', '←', '→', '↑', '↓', '✓', '✗', '★', '☆']
const EMOJI = ['😀', '😄', '🙂', '😍', '🤔', '😎', '👍', '👎', '👏', '🙌', '🙏', '💪', '🎉', '🎯', '🚀', '🔥', '💡', '⚠️', '✅', '❌', '📌', '📎', '📊', '📈', '📉', '📅', '📝', '📁', '💼', '⭐', '❤️', '☕']
const STYLES: { id: string; label: string }[] = [
  { id: 'p', label: 'Normal' },
  { id: 'nospacing', label: 'No Spacing' },
  { id: 'h1', label: 'Heading 1' },
  { id: 'h2', label: 'Heading 2' },
  { id: 'h3', label: 'Heading 3' },
  { id: 'quote', label: 'Quote' },
]

const changeCase = (t: string, mode: 'upper' | 'lower' | 'title' | 'sentence') =>
  mode === 'upper' ? t.toUpperCase()
    : mode === 'lower' ? t.toLowerCase()
      : mode === 'title' ? t.replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase())
        : t.charAt(0).toUpperCase() + t.slice(1).toLowerCase()

export function WordEditor({ file }: { file: OFile }) {
  const nav = useNavigate()
  const toast = useToast()
  const initial = useMemo(() => readDoc(file.content), [file.content])
  const [page, setPage] = useState<PageSetup>(initial.page)
  const [version, setVersion] = useState(0)
  const [tab, setTab] = useState('Home')
  const [zoom, setZoom] = useState(100)
  const [findOpen, setFindOpen] = useState(false)
  const [showReplace, setShowReplace] = useState(false)
  const [query, setQuery] = useState('')
  const [replaceWith, setReplaceWith] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const [dialog, setDialog] = useState<null | 'table' | 'link' | 'wordcount' | 'clearAll'>(null)
  const [navOpen, setNavOpen] = useState(false)
  const [ruler, setRuler] = useState(false)
  const [reading, setReading] = useState(false)
  const [spellcheck, setSpellcheck] = useState(true)
  const [stats, setStats] = useState({ words: 0, chars: 0, charsNoSpace: 0, paragraphs: 0, pages: 1, page: 1 })
  const [ctx, setCtx] = useState<{ x: number; y: number; items: MenuEntry[] } | null>(null)
  const [painter, setPainter] = useState<Record<string, unknown> | null>(null)
  const imgInput = useRef<HTMLInputElement>(null)
  const importInput = useRef<HTMLInputElement>(null)
  const pageRef = useRef<HTMLDivElement>(null)
  const statTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const metrics = pageMetrics(page)

  const editorRef = useRef<ReturnType<typeof useEditor>>(null)
  const zoomRef = useRef(zoom); zoomRef.current = zoom
  const metricsRef = useRef(metrics); metricsRef.current = metrics

  /** Word, character and page counts, debounced so typing in a long document stays smooth. */
  const scheduleStats = useCallback(() => {
    clearTimeout(statTimer.current)
    statTimer.current = setTimeout(() => {
      const ed = editorRef.current, el = pageRef.current
      if (!ed || !el) return
      const text = ed.getText()
      const t = text.trim()
      const z = zoomRef.current / 100
      const rect = el.getBoundingClientRect()
      const pageH = metricsRef.current.height
      const pages = Math.max(1, Math.ceil(rect.height / z / pageH - 0.02))
      let cur = 1
      try {
        const c = ed.view.coordsAtPos(ed.state.selection.head)
        cur = Math.min(pages, Math.max(1, Math.floor((c.top - rect.top) / z / pageH) + 1))
      } catch { /* selection not laid out yet */ }
      const next = {
        words: t ? t.split(/\s+/).length : 0,
        chars: text.length,
        charsNoSpace: text.replace(/\s/g, '').length,
        paragraphs: ed.state.doc.content.childCount,
        pages, page: cur,
      }
      // only publish a genuinely different object, so a re-render cannot feed back into measuring
      setStats(prev => (prev.words === next.words && prev.chars === next.chars && prev.charsNoSpace === next.charsNoSpace
        && prev.paragraphs === next.paragraphs && prev.pages === next.pages && prev.page === next.page ? prev : next))
    }, 160)
  }, [])
  useEffect(() => () => clearTimeout(statTimer.current), [])

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ link: { openOnClick: false, autolink: true } }),
      TextStyleKit, Highlight.configure({ multicolor: true }), Find, Subscript, Superscript,
      TaskList, TaskItem.configure({ nested: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Image.configure({ allowBase64: true }),
      Table.configure({ resizable: true }), TableRow, TableHeader, TableCell,
    ],
    content: initial.html,
    editorProps: { attributes: { 'aria-label': 'Document body' } },
    onUpdate: () => { setVersion(v => v + 1); scheduleStats() },
    onCreate: scheduleStats,
    onSelectionUpdate: scheduleStats,
  })
  editorRef.current = editor

  useEffect(() => {
    const fonts = new Set<string>()
    const dom = new DOMParser().parseFromString(initial.html, 'text/html')
    dom.querySelectorAll('[style*="font-family"]').forEach(el => {
      const f = (el as HTMLElement).style.fontFamily.split(',')[0].replace(/["']/g, '').trim()
      if (f) fonts.add(f)
    })
    void preloadFonts(fonts)
  }, [initial.html])

  const save = useCallback(async () => {
    if (editor) await saveContent(file.id, { html: editor.getHTML(), page })
  }, [editor, file.id, page])
  const state = useAutosave(version, save)
  const setPageSetup = (patch: Partial<PageSetup>) => { setPage(p => ({ ...p, ...patch })); setVersion(v => v + 1); scheduleStats() }

  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => !e ? null : ({
      bold: e.isActive('bold'), italic: e.isActive('italic'), underline: e.isActive('underline'), strike: e.isActive('strike'),
      sub: e.isActive('subscript'), sup: e.isActive('superscript'),
      bullet: e.isActive('bulletList'), ordered: e.isActive('orderedList'), task: e.isActive('taskList'),
      quote: e.isActive('blockquote'), inTable: e.isActive('table'),
      alignL: e.isActive({ textAlign: 'left' }), alignC: e.isActive({ textAlign: 'center' }), alignR: e.isActive({ textAlign: 'right' }), alignJ: e.isActive({ textAlign: 'justify' }),
      block: e.isActive('heading', { level: 1 }) ? 'h1' : e.isActive('heading', { level: 2 }) ? 'h2' : e.isActive('heading', { level: 3 }) ? 'h3' : e.isActive('blockquote') ? 'quote' : 'p',
      font: (e.getAttributes('textStyle').fontFamily as string | undefined)?.replace(/["']/g, '').split(',')[0] ?? '',
      size: (e.getAttributes('textStyle').fontSize as string | undefined)?.replace('px', '') ?? '',
      lineHeight: String(e.getAttributes('textStyle').lineHeight ?? ''),
      canUndo: e.can().undo(), canRedo: e.can().redo(),
      findCount: findKey.getState(e.state)?.hits.length ?? 0, findCur: findKey.getState(e.state)?.current ?? 0,
      // serialised, because useEditorState compares selector results shallowly: a fresh array
      // on every call would re-render forever
      outline: JSON.stringify(outlineOf(e.state.doc)),
    }),
  })

  const headings = useMemo<Heading[]>(() => (s?.outline ? JSON.parse(s.outline) as Heading[] : []), [s?.outline])

  // Chains are built when a control is used, never cached across renders.
  const cmd = useCallback(() => editor!.chain().focus(), [editor])

  const openFind = (replace: boolean) => {
    setFindOpen(true); setShowReplace(replace)
    setTimeout(() => document.getElementById(replace ? 'replace-input' : 'find-input')?.focus(), 0)
  }
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey)) return
      const key = e.key.toLowerCase()
      if (key === 's') { e.preventDefault(); void save().then(() => toast('Saved')) }
      else if (key === 'f') { e.preventDefault(); openFind(false) }
      else if (key === 'h' && !e.shiftKey) { e.preventDefault(); openFind(true) }
      else if (key === 'p') { e.preventDefault(); window.print() }
    }
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [save, toast])

  /* ----------------------------- find and replace ----------------------------- */
  const pushFind = (q: string, opts?: { caseSensitive?: boolean; wholeWord?: boolean }) => {
    if (!editor) return
    const next = { caseSensitive: opts?.caseSensitive ?? caseSensitive, wholeWord: opts?.wholeWord ?? wholeWord }
    editor.view.dispatch(editor.state.tr.setMeta(findKey, { query: q, opts: next }))
  }
  const stepFind = (d: 1 | -1) => {
    if (!editor) return
    const st = findKey.getState(editor.state)
    if (!st?.hits.length) return
    const cur = (st.current + d + st.hits.length) % st.hits.length
    editor.view.dispatch(editor.state.tr.setMeta(findKey, { current: cur }))
    const hit = st.hits[cur]
    const dom = editor.view.domAtPos(hit.from).node as HTMLElement
    ;(dom.nodeType === 3 ? dom.parentElement : dom)?.scrollIntoView({ block: 'center' })
  }
  const replaceOne = () => {
    if (!editor) return
    const st = findKey.getState(editor.state)
    if (!st?.hits.length) { toast('Nothing to replace', 'error'); return }
    const hit = st.hits[Math.min(st.current, st.hits.length - 1)]
    editor.chain().focus().insertContentAt({ from: hit.from, to: hit.to }, replaceWith).run()
  }
  const replaceAll = () => {
    if (!editor) return
    const st = findKey.getState(editor.state)
    if (!st?.hits.length) { toast('Nothing to replace', 'error'); return }
    const n = st.hits.length
    const chain = editor.chain().focus()
    // replace from the end so earlier positions stay valid
    for (const hit of [...st.hits].sort((a, b) => b.from - a.from)) chain.insertContentAt({ from: hit.from, to: hit.to }, replaceWith)
    chain.run()
    toast(`Replaced ${n} ${n === 1 ? 'match' : 'matches'}`)
  }

  /* ----------------------------- commands ----------------------------- */
  function applyStyle(id: string) {
    const c = cmd()
    if (id === 'quote') { c.setParagraph().toggleBlockquote().run(); return }
    c.lift('blockquote')
    if (id === 'p') c.setParagraph().unsetFontSize().run()
    else if (id === 'nospacing') c.setParagraph().setLineHeight('1').run()
    else c.setHeading({ level: Number(id[1]) as 1 | 2 | 3 }).run()
  }
  function scaleFont(dir: 1 | -1) {
    if (!editor) return
    const cur = Number((editor.getAttributes('textStyle').fontSize as string | undefined)?.replace('px', '')) || 16
    const steps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72, 96, 120]
    const i = steps.findIndex(x => x >= cur)
    const next = dir > 0 ? steps[Math.min(steps.length - 1, (i < 0 ? steps.length - 1 : i) + (steps[i] === cur ? 1 : 0))] : steps[Math.max(0, (i < 0 ? steps.length - 1 : i) - 1)]
    cmd().setFontSize(`${next}px`).run()
  }
  function applyCase(mode: 'upper' | 'lower' | 'title' | 'sentence') {
    if (!editor) return
    const { from, to, empty } = editor.state.selection
    if (empty) { toast('Select some text first', 'error'); return }
    const text = editor.state.doc.textBetween(from, to, ' ')
    editor.chain().focus().insertContentAt({ from, to }, changeCase(text, mode)).setTextSelection({ from, to: from + text.length }).run()
  }
  function insertToc() {
    if (!editor) return
    if (!headings.length) { toast('Add some headings first', 'error'); return }
    const items = headings.map(h => `<p style="margin-left:${(h.level - 1) * 18}px">${escapeHtml(h.text)}</p>`).join('')
    cmd().insertContent(`<div class="toc"><div class="toc-title">Contents</div>${items}</div>`).run()
    toast('Table of contents inserted')
  }
  const copyFormat = () => {
    if (!editor) return
    setPainter({ ...editor.getAttributes('textStyle'), bold: editor.isActive('bold'), italic: editor.isActive('italic'), underline: editor.isActive('underline'), strike: editor.isActive('strike'), highlight: editor.getAttributes('highlight').color })
    toast('Formatting copied — select text to apply it')
  }
  const pasteFormat = () => {
    if (!editor || !painter) return
    const c = cmd()
    const p = painter as { fontFamily?: string; fontSize?: string; color?: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; highlight?: string }
    c.unsetAllMarks()
    if (p.fontFamily) c.setFontFamily(p.fontFamily)
    if (p.fontSize) c.setFontSize(p.fontSize)
    if (p.color) c.setColor(p.color)
    if (p.bold) c.setBold()
    if (p.italic) c.setItalic()
    if (p.underline) c.setUnderline()
    if (p.strike) c.setStrike()
    if (p.highlight) c.setHighlight({ color: p.highlight })
    c.run()
    setPainter(null)
  }

  async function doExport(kind: 'docx' | 'pdf') {
    if (!editor) return
    try {
      const html = editor.getHTML()
      toast(`Preparing ${kind.toUpperCase()}…`)
      if (kind === 'docx') download(await exportDocx(file.name, html, page), `${safeName(file.name)}.docx`)
      else download(await exportPdf(file.name, html, page), `${safeName(file.name)}.pdf`)
      toast(`Exported ${kind.toUpperCase()}`)
    } catch (e) { toast(`Export failed: ${(e as Error).message}`, 'error') }
  }
  async function importDocx(f: File | undefined) {
    if (!f || !editor) return
    try {
      const html = /\.docx$/i.test(f.name) ? await docxToHtml(await f.arrayBuffer()) : /\.odt$/i.test(f.name) ? await odtToHtml(await f.arrayBuffer()) : textToHtml(await f.text())
      editor.commands.setContent(html, { emitUpdate: true })
      toast(`Imported ${f.name}`)
    } catch (e) { toast(`Import failed: ${(e as Error).message}`, 'error') }
  }
  function insertImage(f: File | undefined) {
    if (!f || !editor) return
    const r = new FileReader()
    r.onload = () => {
      // collapse to the end of the selection first: inserting a picture must never
      // replace whatever happens to be selected
      editor.chain().focus().setTextSelection(editor.state.selection.to).setImage({ src: String(r.result), alt: f.name }).run()
    }
    r.onerror = () => toast('That image could not be read', 'error')
    r.readAsDataURL(f)
  }

  const fileMenu: MenuItem[] = [
    { label: 'New document', onClick: async () => { const n = await createFile('docx'); nav(`/word/${n.id}`) } },
    { label: 'Import DOCX / text…', onClick: () => importInput.current?.click() },
    'sep',
    { label: 'Save now', onClick: () => { void save().then(() => toast('Saved')) } },
    { label: 'Print (⌘P)', onClick: () => window.print() },
    { label: 'Export as PDF', onClick: () => doExport('pdf') },
    { label: 'Export as DOCX', onClick: () => doExport('docx') },
  ]

  if (!editor) return null
  const c = cmd

  return (
    <div className="h-full flex flex-col min-h-0">
      <EditorHeader file={file} color="var(--color-word)" appName="Word" save={state} fileMenu={fileMenu}>
        <IconButton label="Navigation pane" active={navOpen} onClick={() => setNavOpen(o => !o)}><LayoutList size={16} /></IconButton>
        <IconButton label="Find (⌘F)" active={findOpen} onClick={() => (findOpen ? setFindOpen(false) : openFind(false))}><Search size={16} /></IconButton>
      </EditorHeader>

      <Ribbon tabs={['Home', 'Insert', 'Layout', 'References', 'Review', 'View']} tab={tab} onTab={setTab}>
        {tab === 'Home' && (<>
          <RibbonGroup label="Undo">
            <IconButton label="Undo (⌘Z)" disabled={!s?.canUndo} onClick={() => c().undo().run()}><Undo2 size={16} /></IconButton>
            <IconButton label="Redo (⇧⌘Z)" disabled={!s?.canRedo} onClick={() => c().redo().run()}><Redo2 size={16} /></IconButton>
            <IconButton label="Format painter" active={!!painter} onClick={() => (painter ? pasteFormat() : copyFormat())}><PaintRoller size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Font">
            <FontPicker value={s?.font ?? ''} onPick={name => (name ? c().setFontFamily(cssFamily(name)).run() : c().unsetFontFamily().run())} width={150} />
            <SizeBox value={s?.size ? Number(s.size) : ''} onPick={n => c().setFontSize(`${n}px`).run()} />
            <IconButton label="Grow font" onClick={() => scaleFont(1)}><Plus size={15} /></IconButton>
            <IconButton label="Shrink font" onClick={() => scaleFont(-1)}><Minus size={15} /></IconButton>
            <IconButton label="Bold (⌘B)" active={s?.bold} onClick={() => c().toggleBold().run()}><Bold size={16} /></IconButton>
            <IconButton label="Italic (⌘I)" active={s?.italic} onClick={() => c().toggleItalic().run()}><Italic size={16} /></IconButton>
            <IconButton label="Underline (⌘U)" active={s?.underline} onClick={() => c().toggleUnderline().run()}><UnderlineIcon size={16} /></IconButton>
            <IconButton label="Strikethrough" active={s?.strike} onClick={() => c().toggleStrike().run()}><Strikethrough size={16} /></IconButton>
            <IconButton label="Subscript" active={s?.sub} onClick={() => c().toggleSubscript().run()}><SubIcon size={16} /></IconButton>
            <IconButton label="Superscript" active={s?.sup} onClick={() => c().toggleSuperscript().run()}><SupIcon size={16} /></IconButton>
            <Menu title="Change case" trigger={<CaseSensitive size={16} />} items={[
              { label: 'Sentence case', onClick: () => applyCase('sentence') },
              { label: 'lowercase', onClick: () => applyCase('lower') },
              { label: 'UPPERCASE', onClick: () => applyCase('upper') },
              { label: 'Title Case', onClick: () => applyCase('title') },
            ]} />
            <ColorMenu label="Text colour" icon={<Type size={16} />} allowNone onPick={col => (col ? c().setColor(col).run() : c().unsetColor().run())} />
            <ColorMenu label="Highlight colour" icon={<Highlighter size={16} />} allowNone onPick={col => (col ? c().setHighlight({ color: col }).run() : c().unsetHighlight().run())} />
            <IconButton label="Clear formatting" onClick={() => c().unsetAllMarks().clearNodes().run()}><RemoveFormatting size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Paragraph">
            <IconButton label="Bullets" active={s?.bullet} onClick={() => c().toggleBulletList().run()}><List size={16} /></IconButton>
            <IconButton label="Numbering" active={s?.ordered} onClick={() => c().toggleOrderedList().run()}><ListOrdered size={16} /></IconButton>
            <IconButton label="Checklist" active={s?.task} onClick={() => c().toggleTaskList().run()}><ListChecks size={16} /></IconButton>
            <IconButton label="Decrease indent" onClick={() => c().liftListItem('listItem').run()}><Outdent size={16} /></IconButton>
            <IconButton label="Increase indent" onClick={() => c().sinkListItem('listItem').run()}><Indent size={16} /></IconButton>
            <Divider />
            <IconButton label="Align left" active={s?.alignL} onClick={() => c().setTextAlign('left').run()}><AlignLeft size={16} /></IconButton>
            <IconButton label="Centre" active={s?.alignC} onClick={() => c().setTextAlign('center').run()}><AlignCenter size={16} /></IconButton>
            <IconButton label="Align right" active={s?.alignR} onClick={() => c().setTextAlign('right').run()}><AlignRight size={16} /></IconButton>
            <IconButton label="Justify" active={s?.alignJ} onClick={() => c().setTextAlign('justify').run()}><AlignJustify size={16} /></IconButton>
            <Menu title="Line spacing" trigger={<><ArrowDownUp size={15} />Spacing</>} items={LINE_SPACINGS.map(v => ({ label: `${v}${(s?.lineHeight || '') === v ? '  ✓' : ''}`, onClick: () => c().setLineHeight(v).run() }))} />
            <IconButton label="Quote" active={s?.quote} onClick={() => c().toggleBlockquote().run()}><TextQuote size={16} /></IconButton>
          </RibbonGroup>
          <RibbonGroup label="Styles">
            <Menu title="Paragraph style" trigger={<span className="w-[92px] text-left truncate">{STYLES.find(x => x.id === s?.block)?.label ?? 'Normal'}</span>}
              items={STYLES.map(st => ({ label: `${st.label}${s?.block === st.id ? '  ✓' : ''}`, onClick: () => applyStyle(st.id) }))} />
          </RibbonGroup>
          <RibbonGroup label="Editing">
            <Button variant="ghost" onClick={() => openFind(false)}><Search size={15} />Find</Button>
            <Button variant="ghost" onClick={() => openFind(true)}><Replace size={15} />Replace</Button>
          </RibbonGroup>
        </>)}

        {tab === 'Insert' && (<>
          <RibbonGroup label="Pages">
            <Button variant="ghost" onClick={() => c().setHorizontalRule().run()}><SeparatorHorizontal size={15} />Page break</Button>
          </RibbonGroup>
          <RibbonGroup label="Tables">
            <Button variant="ghost" onClick={() => setDialog('table')}><Table2 size={16} />Table</Button>
            {s?.inTable && (<>
              <Button variant="ghost" onClick={() => c().addRowAfter().run()}>+ Row</Button>
              <Button variant="ghost" onClick={() => c().addColumnAfter().run()}>+ Column</Button>
              <Button variant="ghost" onClick={() => c().deleteRow().run()}>− Row</Button>
              <Button variant="ghost" onClick={() => c().deleteColumn().run()}>− Column</Button>
              <Button variant="ghost" onClick={() => c().toggleHeaderRow().run()}>Header row</Button>
              <Button variant="ghost" onClick={() => c().mergeOrSplit().run()}>Merge / split</Button>
              <Button variant="ghost" onClick={() => c().deleteTable().run()}>Delete table</Button>
            </>)}
          </RibbonGroup>
          <RibbonGroup label="Illustrations">
            <Button variant="ghost" onClick={() => imgInput.current?.click()}><ImagePlus size={16} />Picture</Button>
          </RibbonGroup>
          <RibbonGroup label="Links"><Button variant="ghost" onClick={() => setDialog('link')}><Link2 size={16} />Link</Button></RibbonGroup>
          <RibbonGroup label="Symbols">
            <GridMenu label="Symbol" icon={<Sigma size={15} />} items={SYMBOLS} onPick={ch => c().insertContent(ch).run()} />
            <GridMenu label="Emoji" icon={<Smile size={15} />} items={EMOJI} onPick={ch => c().insertContent(ch).run()} />
          </RibbonGroup>
        </>)}

        {tab === 'Layout' && (<>
          <RibbonGroup label="Page setup">
            <Menu title="Margins" trigger={<>Margins<ChevronDown size={12} /></>} items={(Object.keys(MARGINS) as MarginId[]).map(id => ({ label: `${MARGINS[id].name}${page.margin === id ? '  ✓' : ''}`, onClick: () => setPageSetup({ margin: id }) }))} />
            <Menu title="Orientation" trigger={<>Orientation<ChevronDown size={12} /></>} items={[
              { label: `Portrait${!page.landscape ? '  ✓' : ''}`, onClick: () => setPageSetup({ landscape: false }) },
              { label: `Landscape${page.landscape ? '  ✓' : ''}`, onClick: () => setPageSetup({ landscape: true }) },
            ]} />
            <Menu title="Paper size" trigger={<>Size<ChevronDown size={12} /></>} items={(Object.keys(PAPERS) as PaperId[]).map(id => ({ label: `${PAPERS[id].name}${page.paper === id ? '  ✓' : ''}`, onClick: () => setPageSetup({ paper: id }) }))} />
            <Menu title="Columns" trigger={<><Columns3 size={15} />Columns<ChevronDown size={12} /></>} items={([1, 2, 3] as const).map(n => ({ label: `${n === 1 ? 'One' : n === 2 ? 'Two' : 'Three'}${page.columns === n ? '  ✓' : ''}`, onClick: () => setPageSetup({ columns: n }) }))} />
            <Button variant="ghost" onClick={() => setPageSetup(DEFAULT_PAGE)}>Reset</Button>
          </RibbonGroup>
          <RibbonGroup label="Current">
            <span className="text-[12px] text-muted px-1 leading-[34px]">{PAPERS[page.paper].name.split(' (')[0]} · {page.landscape ? 'Landscape' : 'Portrait'} · {MARGINS[page.margin].name.split(' (')[0]} margins · {page.columns} column{page.columns > 1 ? 's' : ''}</span>
          </RibbonGroup>
        </>)}

        {tab === 'References' && (
          <RibbonGroup label="Table of contents">
            <Button variant="ghost" onClick={insertToc}><LayoutList size={15} />Insert table of contents</Button>
            <span className="text-[12px] text-muted leading-[34px] px-2">{headings.length} heading{headings.length === 1 ? '' : 's'} found</span>
          </RibbonGroup>
        )}

        {tab === 'Review' && (<>
          <RibbonGroup label="Proofing">
            <Button variant="ghost" onClick={() => setDialog('wordcount')}>Word count</Button>
            <Button variant="ghost" aria-pressed={spellcheck} className={spellcheck ? 'bg-accent-soft text-accent' : ''}
              onClick={() => { setSpellcheck(v => { editor.view.dom.setAttribute('spellcheck', String(!v)); return !v }) }}>Spelling{spellcheck ? ' on' : ' off'}</Button>
          </RibbonGroup>
          <RibbonGroup label="Editing">
            <Button variant="ghost" onClick={() => openFind(true)}><Replace size={15} />Find & replace</Button>
            <Button variant="ghost" onClick={() => setDialog('clearAll')}><RemoveFormatting size={15} />Clear all formatting</Button>
          </RibbonGroup>
        </>)}

        {tab === 'View' && (<>
          <RibbonGroup label="Views">
            <Button variant="ghost" aria-pressed={reading} className={reading ? 'bg-accent-soft text-accent' : ''} onClick={() => setReading(r => !r)}><Eye size={15} />Reading view</Button>
            <Button variant="ghost" onClick={() => window.print()}><Printer size={15} />Print layout</Button>
          </RibbonGroup>
          <RibbonGroup label="Show">
            <Button variant="ghost" aria-pressed={ruler} className={ruler ? 'bg-accent-soft text-accent' : ''} onClick={() => setRuler(r => !r)}><Ruler size={15} />Ruler</Button>
            <Button variant="ghost" aria-pressed={navOpen} className={navOpen ? 'bg-accent-soft text-accent' : ''} onClick={() => setNavOpen(o => !o)}><LayoutList size={15} />Navigation pane</Button>
          </RibbonGroup>
          <RibbonGroup label="Zoom">
            <select aria-label="Zoom preset" value={zoom} onChange={e => setZoom(Number(e.target.value))} className="h-8 rounded-md border border-line bg-surface px-1.5 text-[13px]">
              {[50, 75, 100, 125, 150, 200].map(z => <option key={z} value={z}>{z}%</option>)}
            </select>
          </RibbonGroup>
          <RibbonGroup label="Export">
            <Button variant="ghost" onClick={() => doExport('pdf')}><FileDown size={15} />PDF</Button>
            <Button variant="ghost" onClick={() => doExport('docx')}><FileDown size={15} />DOCX</Button>
          </RibbonGroup>
        </>)}
      </Ribbon>

      {findOpen && (
        <div className="bg-panel border-b border-line px-4 py-1.5 flex flex-col gap-1.5" data-testid="find-bar">
          <div className="flex items-center gap-2">
            <Search size={14} className="text-muted shrink-0" />
            <input id="find-input" aria-label="Find text" value={query} onChange={e => { setQuery(e.target.value); pushFind(e.target.value) }}
              onKeyDown={e => { if (e.key === 'Enter') stepFind(e.shiftKey ? -1 : 1); if (e.key === 'Escape') { setFindOpen(false); pushFind('') } }}
              placeholder="Find in document" className="h-7 w-56 px-2 rounded border border-line bg-surface" />
            <span className="text-[12px] text-muted w-24" data-testid="find-count">{query ? (s?.findCount ? `${(s.findCur ?? 0) + 1} of ${s.findCount}` : 'No results') : ''}</span>
            <IconButton label="Previous match" onClick={() => stepFind(-1)}><ChevronUp size={15} /></IconButton>
            <IconButton label="Next match" onClick={() => stepFind(1)}><ChevronDown size={15} /></IconButton>
            <IconButton label="Match case" active={caseSensitive} onClick={() => { const v = !caseSensitive; setCaseSensitive(v); pushFind(query, { caseSensitive: v }) }}><span className="text-[12px] font-semibold px-0.5">Aa</span></IconButton>
            <IconButton label="Whole word" active={wholeWord} onClick={() => { const v = !wholeWord; setWholeWord(v); pushFind(query, { wholeWord: v }) }}><span className="text-[12px] font-semibold px-0.5">ab|</span></IconButton>
            <IconButton label={showReplace ? 'Hide replace' : 'Show replace'} active={showReplace} onClick={() => setShowReplace(v => !v)}><Replace size={15} /></IconButton>
            <div className="flex-1" />
            <IconButton label="Close find" onClick={() => { setFindOpen(false); pushFind('') }}><X size={15} /></IconButton>
          </div>
          {showReplace && (
            <div className="flex items-center gap-2 pl-6">
              <input id="replace-input" aria-label="Replace with" value={replaceWith} onChange={e => setReplaceWith(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') replaceOne(); if (e.key === 'Escape') setFindOpen(false) }}
                placeholder="Replace with" className="h-7 w-56 px-2 rounded border border-line bg-surface" />
              <Button onClick={replaceOne} disabled={!s?.findCount}>Replace</Button>
              <Button onClick={replaceAll} disabled={!s?.findCount}>Replace all</Button>
            </div>
          )}
        </div>
      )}

      <div className="flex-1 min-h-0 flex">
        {navOpen && (
          <aside aria-label="Navigation pane" className="w-56 shrink-0 border-r border-line bg-panel overflow-y-auto p-2" data-testid="nav-pane">
            <div className="text-[11px] uppercase tracking-wide text-muted px-2 py-1">Headings</div>
            {headings.length === 0 && <div className="px-2 py-2 text-[12.5px] text-muted">Apply Heading 1–3 to build an outline.</div>}
            {headings.map(h => (
              <button key={h.pos} onClick={() => { editor.chain().focus().setTextSelection(h.pos + 1).scrollIntoView().run() }}
                style={{ paddingLeft: 8 + (h.level - 1) * 12 }}
                className="w-full text-left py-1 pr-2 rounded text-[12.5px] hover:bg-panel2 truncate">{h.text || 'Untitled heading'}</button>
            ))}
          </aside>
        )}

        <div className="flex-1 overflow-auto px-4 py-8" style={{ background: 'var(--canvas)' }}
          onMouseDown={e => { if (e.target === e.currentTarget) editor.commands.focus('end') }}>
          {ruler && (
            <div aria-hidden className="mx-auto mb-1 relative h-5 bg-surface border border-line rounded" style={{ width: metrics.width * (zoom / 100), maxWidth: '100%' }}>
              <div className="absolute inset-y-0 bg-accent-soft" style={{ left: metrics.margin[3] * (zoom / 100), right: metrics.margin[1] * (zoom / 100) }} />
            </div>
          )}
          <div
            ref={pageRef}
            className={cx('doc-page', page.columns === 2 && 'doc-columns-2', page.columns === 3 && 'doc-columns-3', reading && 'doc-reading')}
            style={{
              zoom: zoom / 100,
              width: metrics.width,
              minHeight: metrics.height,
              padding: `${metrics.margin[0]}px ${metrics.margin[1]}px ${metrics.margin[2]}px ${metrics.margin[3]}px`,
              backgroundImage: `repeating-linear-gradient(to bottom, transparent 0, transparent ${metrics.height - 1}px, var(--border) ${metrics.height - 1}px, var(--border) ${metrics.height}px)`,
            }}
            onMouseDown={e => {
              // Clicking the page margin places the caret, the way Word does. Clicks inside the
              // text area are left entirely to ProseMirror.
              if (e.target !== e.currentTarget) return
              e.preventDefault()
              const first = editor.view.dom.firstElementChild?.getBoundingClientRect()
              editor.commands.focus(first && e.clientY < first.top ? 'start' : 'end')
            }}
            onContextMenu={e => {
              e.preventDefault()
              setCtx({ x: e.clientX, y: e.clientY, items: [
                { label: 'Cut', hint: '⌘X', onClick: () => document.execCommand('cut') },
                { label: 'Copy', hint: '⌘C', onClick: () => document.execCommand('copy') },
                { label: 'Paste', hint: '⌘V', onClick: () => toast('Use ⌘V to paste — the browser blocks reading the clipboard here') },
                'sep',
                { label: 'Bold', onClick: () => c().toggleBold().run() },
                { label: 'Italic', onClick: () => c().toggleItalic().run() },
                { label: 'Clear formatting', onClick: () => c().unsetAllMarks().run() },
                'sep',
                { label: 'Insert link…', onClick: () => setDialog('link') },
                { label: 'Insert table…', onClick: () => setDialog('table') },
              ] })
            }}>
            <EditorContent editor={editor} />
          </div>
        </div>
      </div>

      <div className="h-8 shrink-0 bg-panel border-t border-line flex items-center gap-4 px-4 text-[12px] text-muted">
        <span data-testid="page-indicator">Page {stats.page} of {stats.pages}</span>
        <button onClick={() => setDialog('wordcount')} className="hover:text-ink" data-testid="word-count">{stats.words} {stats.words === 1 ? 'word' : 'words'}</button>
        <span>{stats.chars} characters</span>
        {painter && <span className="text-accent font-medium">Format painter ready — select text</span>}
        <div className="flex-1" />
        <IconButton label="Zoom out" onClick={() => setZoom(z => Math.max(50, z - 10))}><Minus size={14} /></IconButton>
        <input aria-label="Zoom" type="range" min={50} max={200} step={10} value={zoom} onChange={e => setZoom(Number(e.target.value))} className="w-28" />
        <IconButton label="Zoom in" onClick={() => setZoom(z => Math.min(200, z + 10))}><Plus size={14} /></IconButton>
        <button onClick={() => setZoom(100)} title="Reset zoom to 100%" className="w-11 text-right hover:text-ink">{zoom}%</button>
      </div>

      <input ref={imgInput} type="file" accept="image/*" hidden data-testid="word-image-input" onChange={e => { insertImage(e.target.files?.[0]); e.target.value = '' }} />
      <input ref={importInput} type="file" accept=".docx,.odt,.txt,.md" hidden data-testid="word-import-input" onChange={e => { void importDocx(e.target.files?.[0]); e.target.value = '' }} />
      {ctx && <ContextMenu x={ctx.x} y={ctx.y} items={ctx.items} onClose={() => setCtx(null)} />}
      {dialog === 'table' && <TableDialog onClose={() => setDialog(null)} onInsert={(r, cc) => editor.chain().focus().setTextSelection(editor.state.selection.to).insertTable({ rows: r, cols: cc, withHeaderRow: true }).run()} />}
      {dialog === 'link' && (
        <PromptModal title="Insert link" label="Address (https://…)" initial={(editor.getAttributes('link').href as string) ?? 'https://'} confirm="Apply" onClose={() => setDialog(null)}
          onSubmit={v => {
            const u = v.trim()
            if (!u || u === 'https://') editor.chain().focus().unsetLink().run()
            else editor.chain().focus().extendMarkRange('link').setLink({ href: /^[a-z]+:/i.test(u) ? u : `https://${u}` }).run()
          }} />
      )}
      {dialog === 'clearAll' && (
        <ConfirmModal title="Clear all formatting?" danger confirm="Clear formatting"
          message="Every font, size, colour and style in this document is reset. The text itself is kept, and you can undo this."
          onConfirm={() => { editor.chain().focus().selectAll().unsetAllMarks().clearNodes().setTextSelection(1).run(); toast('Formatting cleared') }}
          onClose={() => setDialog(null)} />
      )}
      {dialog === 'wordcount' && (
        <Modal title="Word count" onClose={() => setDialog(null)} width={340} footer={<Button variant="primary" onClick={() => setDialog(null)}>Close</Button>}>
          <dl className="text-[13px] divide-y divide-line">
            {([['Pages', stats.pages], ['Words', stats.words], ['Characters (with spaces)', stats.chars], ['Characters (no spaces)', stats.charsNoSpace], ['Paragraphs', stats.paragraphs]] as const).map(([k, v]) => (
              <div key={k} className="flex justify-between py-1.5"><dt className="text-muted">{k}</dt><dd className="font-medium">{v.toLocaleString()}</dd></div>
            ))}
          </dl>
        </Modal>
      )}
    </div>
  )
}

/* ------------------------------------------------------------------ helpers */
const escapeHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

interface Heading { level: number; text: string; pos: number }
function outlineOf(doc: { descendants: (f: (n: { type: { name: string }; attrs: Record<string, unknown>; textContent: string }, pos: number) => void) => void }): Heading[] {
  const out: Heading[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'heading') out.push({ level: Number(node.attrs.level) || 1, text: node.textContent, pos })
  })
  return out
}

function GridMenu({ label, icon, items, onPick }: { label: string; icon: React.ReactNode; items: string[]; onPick: (v: string) => void }) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const h = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) setOpen(false) }
    document.addEventListener('mousedown', h)
    return () => document.removeEventListener('mousedown', h)
  }, [open])
  return (
    <div ref={ref} className="relative inline-block">
      <Button variant="ghost" onClick={() => setOpen(o => !o)} aria-expanded={open} aria-label={`Insert ${label.toLowerCase()}`}>{icon}{label}</Button>
      {open && (
        <div className="absolute z-[120] mt-1 p-2 bg-surface border border-line rounded-lg shadow-soft grid grid-cols-8 gap-1 w-[280px] max-h-[220px] overflow-y-auto">
          {items.map(ch => (
            <button key={ch} onMouseDown={e => e.preventDefault()} onClick={() => { onPick(ch); setOpen(false) }}
              className="h-7 rounded hover:bg-panel2 text-[15px]">{ch}</button>
          ))}
        </div>
      )}
    </div>
  )
}

function TableDialog({ onClose, onInsert }: { onClose: () => void; onInsert: (r: number, c: number) => void }) {
  const [r, setR] = useState(3)
  const [c, setC] = useState(3)
  const clamp = (n: number) => Math.max(1, Math.min(20, n || 1))
  return (
    <Modal title="Insert table" onClose={onClose} width={320} footer={<><Button onClick={onClose}>Cancel</Button><Button variant="primary" onClick={() => { onInsert(clamp(r), clamp(c)); onClose() }}>Insert</Button></>}>
      <div className="flex gap-4">
        <label className="flex-1 text-[13px]">Rows<input aria-label="Rows" type="number" min={1} max={20} value={r} onChange={e => setR(Number(e.target.value))} className="mt-1 w-full h-9 px-2 rounded-md border border-line bg-panel" /></label>
        <label className="flex-1 text-[13px]">Columns<input aria-label="Columns" type="number" min={1} max={20} value={c} onChange={e => setC(Number(e.target.value))} className="mt-1 w-full h-9 px-2 rounded-md border border-line bg-panel" /></label>
      </div>
    </Modal>
  )
}
