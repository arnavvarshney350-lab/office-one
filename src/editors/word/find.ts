import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { Node as PMNode } from '@tiptap/pm/model'

export interface FindOptions { caseSensitive: boolean; wholeWord: boolean }
export interface FindState { query: string; opts: FindOptions; hits: { from: number; to: number }[]; current: number }
export const findKey = new PluginKey<FindState>('find')

function search(doc: PMNode, q: string, opts: FindOptions) {
  const hits: { from: number; to: number }[] = []
  if (!q) return hits
  const needle = opts.caseSensitive ? q : q.toLowerCase()
  const isWord = (ch: string | undefined) => !!ch && /[\wÀ-ɏ]/.test(ch)
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return
    const raw = node.text
    const hay = opts.caseSensitive ? raw : raw.toLowerCase()
    for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + Math.max(1, needle.length))) {
      if (opts.wholeWord && (isWord(raw[i - 1]) || isWord(raw[i + needle.length]))) continue
      hits.push({ from: pos + i, to: pos + i + needle.length })
    }
  })
  return hits
}

/** Find and replace: matches are decorations, so searching never mutates the document. */
export const Find = Extension.create({
  name: 'find',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: findKey,
      state: {
        init: (): FindState => ({ query: '', opts: { caseSensitive: false, wholeWord: false }, hits: [], current: 0 }),
        apply(tr, prev, _o, state): FindState {
          const meta = tr.getMeta(findKey) as Partial<FindState> | undefined
          if (meta?.query !== undefined || meta?.opts !== undefined) {
            const query = meta.query ?? prev.query
            const opts = meta.opts ?? prev.opts
            return { query, opts, hits: search(state.doc, query, opts), current: meta.current ?? 0 }
          }
          if (meta?.current !== undefined) return { ...prev, current: meta.current }
          if (tr.docChanged && prev.query) {
            const hits = search(state.doc, prev.query, prev.opts)
            return { ...prev, hits, current: Math.min(prev.current, Math.max(hits.length - 1, 0)) }
          }
          return prev
        },
      },
      props: {
        decorations(state) {
          const s = findKey.getState(state)
          if (!s?.hits.length) return DecorationSet.empty
          return DecorationSet.create(state.doc, s.hits.map((h, i) =>
            Decoration.inline(h.from, h.to, { class: i === s.current ? 'find-hit find-hit-current' : 'find-hit' })))
        },
      },
    })]
  },
})
