import { useEffect, useState, type ReactNode } from 'react'
import { DatabaseZap } from 'lucide-react'
import { db } from '../storage/db'

/** Blocks the app with a clear explanation when IndexedDB cannot be opened (e.g. some private-browsing modes). */
export function StorageGuard({ children }: { children: ReactNode }) {
  const [state, setState] = useState<'checking' | 'ok' | 'failed'>('checking')
  const [detail, setDetail] = useState('')
  useEffect(() => {
    let on = true
    db.open().then(() => on && setState('ok')).catch((e: unknown) => { if (on) { setDetail(e instanceof Error ? e.message : String(e)); setState('failed') } })
    return () => { on = false }
  }, [])
  if (state === 'ok') return <>{children}</>
  if (state === 'checking') return <div className="h-full" aria-busy="true" />
  return (
    <div role="alert" className="h-full flex items-center justify-center p-8">
      <div className="max-w-md text-center">
        <div className="mx-auto w-12 h-12 rounded-full bg-danger/10 text-danger flex items-center justify-center"><DatabaseZap size={24} /></div>
        <h1 className="mt-4 text-lg font-semibold">Local storage is unavailable</h1>
        <p className="mt-2 text-[13px] text-muted">Office One keeps your documents in your browser's IndexedDB, and this browser is not allowing it. This usually means private browsing or blocked site data. Allow site data for this page, or use a normal window, then reload.</p>
        {detail && <pre className="mt-3 text-left text-[11px] bg-panel border border-line rounded-md p-2 overflow-auto max-h-24 text-muted">{detail}</pre>}
        <button onClick={() => window.location.reload()} className="mt-5 h-9 px-4 rounded-md bg-accent text-accent-fg font-medium">Try again</button>
      </div>
    </div>
  )
}
