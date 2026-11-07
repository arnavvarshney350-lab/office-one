import { useEffect } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import type { Editor } from '@tiptap/core'
import { richExtensions } from './richText'
import { boxStyle } from './SlideView'
import { objHtml, type SlideObj } from './model'

/** In-place rich-text editing for one slide text box. The parent owns formatting commands via `onEditor`. */
export function TextBoxEditor({ o, font, onChange, onEditor, onExit }: {
  o: SlideObj; font: string; onChange: (html: string) => void; onEditor: (e: Editor | null) => void; onExit: () => void
}) {
  const editor = useEditor({
    extensions: richExtensions(),
    content: objHtml(o),
    autofocus: 'end',
    editorProps: { attributes: { 'aria-label': 'Edit slide text', spellcheck: 'true' }, handleKeyDown: (_v, e) => { if (e.key === 'Escape') { onExit(); return true } return false } },
    onUpdate: ({ editor: ed }) => onChange(ed.getHTML()),
  })
  useEffect(() => { onEditor(editor); return () => onEditor(null) }, [editor]) // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="slide-editor" style={{ ...boxStyle(o, font), width: '100%', height: '100%', overflow: 'hidden', outline: '2px solid var(--accent)', background: 'transparent' }} onPointerDown={e => e.stopPropagation()}>
      <EditorContent editor={editor} style={{ height: '100%' }} />
    </div>
  )
}
