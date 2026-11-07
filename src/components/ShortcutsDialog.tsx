import { Modal, Button } from '../design-system'

const MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform)
const mod = MAC ? '⌘' : 'Ctrl'

const SECTIONS: { title: string; rows: [string, string][] }[] = [
  { title: 'Everywhere', rows: [['Show this help', '?'], ['Save now', `${mod} S`], ['Undo', `${mod} Z`], ['Redo', `⇧ ${mod} Z`], ['Find', `${mod} F`]] },
  { title: 'Word', rows: [['Bold / Italic / Underline', `${mod} B / I / U`], ['Find and replace', MAC ? '⌃ H' : 'Ctrl H'], ['Print', `${mod} P`], ['Select all', `${mod} A`]] },
  { title: 'Excel', rows: [['Edit cell', 'Enter / F2'], ['Commit and move', 'Enter / Tab'], ['Cancel edit', 'Esc'], ['Fill down / right', `${mod} D / R`], ['Replace', MAC ? '⌃ H' : 'Ctrl H'], ['Accept function suggestion', 'Tab'], ['Clear cells', 'Delete'], ['Select all', `${mod} A`]] },
  { title: 'PowerPoint', rows: [['Edit text', 'Enter / double-click'], ['Add to selection', '⇧ click'], ['Select all objects', `${mod} A`], ['Group / Ungroup', `${mod} G / ⇧ ${mod} G`], ['Duplicate', `${mod} D`], ['Nudge (10 px with ⇧)', 'Arrow keys'], ['Disable snapping while dragging', '⌥ drag'], ['Presenter view while presenting', 'V'], ['Next / previous slide', '→ ← Space']] },
]

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="Keyboard shortcuts" onClose={onClose} width={560} footer={<Button variant="primary" onClick={onClose}>Done</Button>}>
      <div className="grid sm:grid-cols-2 gap-x-8 gap-y-5 max-h-[60vh] overflow-y-auto pr-1">
        {SECTIONS.map(s => (
          <section key={s.title}>
            <h2 className="text-[12px] uppercase tracking-wide text-muted mb-1.5">{s.title}</h2>
            <dl className="text-[13px] divide-y divide-line">
              {s.rows.map(([k, v]) => <div key={k} className="flex justify-between gap-3 py-1.5"><dt>{k}</dt><dd><kbd className="px-1.5 py-0.5 rounded border border-line bg-panel text-[11.5px] whitespace-nowrap">{v}</kbd></dd></div>)}
            </dl>
          </section>
        ))}
      </div>
    </Modal>
  )
}
