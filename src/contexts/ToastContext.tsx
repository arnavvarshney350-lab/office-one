import { createContext, useCallback, useContext, useMemo, useState, type ReactNode } from 'react'

export interface ToastAction { label: string; onClick: () => void }
interface Toast { id: number; text: string; kind: 'info' | 'error'; action?: ToastAction }
type Notify = (text: string, kind?: Toast['kind'], action?: ToastAction) => void
const Ctx = createContext<{ toast: Notify } | null>(null)
let n = 0

export function ToastProvider({ children }: { children: ReactNode }) {
  const [items, setItems] = useState<Toast[]>([])
  const dismiss = useCallback((id: number) => setItems(i => i.filter(t => t.id !== id)), [])
  const toast = useCallback<Notify>((text, kind = 'info', action) => {
    const id = ++n
    setItems(i => [...i.slice(-3), { id, text, kind, action }])
    setTimeout(() => dismiss(id), action ? 7000 : kind === 'error' ? 6000 : 3500) // an undo window needs longer than a plain confirmation
  }, [dismiss])
  const value = useMemo(() => ({ toast }), [toast])
  return (
    <Ctx.Provider value={value}>
      {children}
      <div className="fixed bottom-12 left-1/2 -translate-x-1/2 z-[100] flex flex-col gap-2 items-center pointer-events-none" role="status" aria-live="polite">
        {items.map(t => (
          <div key={t.id} className={`pointer-events-auto flex items-center gap-3 pl-4 pr-2 h-11 rounded-xl shadow-pop text-[13px] menu-in ${t.kind === 'error' ? 'bg-danger text-white' : 'bg-ink text-bg'}`}>
            <span>{t.text}</span>
            {t.action && <button onClick={() => { t.action!.onClick(); dismiss(t.id) }} className="h-7 px-2.5 rounded-md font-semibold bg-white/15 hover:bg-white/25">{t.action.label}</button>}
          </div>
        ))}
      </div>
    </Ctx.Provider>
  )
}
export function useToast() {
  const c = useContext(Ctx)
  if (!c) throw new Error('useToast outside ToastProvider')
  return c.toast
}
