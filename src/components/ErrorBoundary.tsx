import { Component, type ErrorInfo, type ReactNode } from 'react'
import { AlertTriangle } from 'lucide-react'

interface Props { children: ReactNode; /** Changing this value clears a previous error (e.g. the route path). */ resetKey?: string }
interface State { error: Error | null }

/**
 * Catches render-time failures (including failed lazy chunk loads) so a bug in one editor never
 * blanks the whole app. Documents are saved continuously to IndexedDB, so nothing is lost.
 */
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null }
  static getDerivedStateFromError(error: Error): State { return { error } }
  componentDidCatch(error: Error, info: ErrorInfo) { console.error('Office One crashed:', error, info.componentStack) }
  componentDidUpdate(prev: Props) { if (this.state.error && prev.resetKey !== this.props.resetKey) this.setState({ error: null }) }

  render() {
    const { error } = this.state
    if (!error) return this.props.children
    const chunk = /Loading chunk|dynamically imported module|Importing a module script failed/i.test(error.message)
    return (
      <div role="alert" className="h-full flex items-center justify-center p-8">
        <div className="max-w-md text-center">
          <div className="mx-auto w-12 h-12 rounded-full bg-danger/10 text-danger flex items-center justify-center"><AlertTriangle size={24} /></div>
          <h1 className="mt-4 text-lg font-semibold">{chunk ? 'A new version is available' : 'Something went wrong'}</h1>
          <p className="mt-2 text-[13px] text-muted">
            {chunk ? 'Part of the app could not be loaded, usually because it was updated while this tab was open. Reload to continue.'
              : 'This screen hit an unexpected error. Your documents are saved automatically on this device, so your work is safe.'}
          </p>
          {!chunk && <pre className="mt-3 text-left text-[11px] bg-panel border border-line rounded-md p-2 overflow-auto max-h-24 text-muted">{error.message}</pre>}
          <div className="mt-5 flex gap-2 justify-center">
            <button onClick={() => window.location.reload()} className="h-9 px-4 rounded-md bg-accent text-accent-fg font-medium">Reload</button>
            <a href="/" className="h-9 px-4 rounded-md border border-line bg-surface inline-flex items-center hover:bg-panel2">Go to Home</a>
          </div>
        </div>
      </div>
    )
  }
}
