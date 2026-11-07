import { useState } from 'react'
import { Monitor, Moon, Sun } from 'lucide-react'
import { TopBar } from '../../components/TopBar'
import { Button, ConfirmModal, cx } from '../../design-system'
import { useTheme } from '../../contexts/ThemeContext'
import { db } from '../../storage/db'
import { useToast } from '../../contexts/ToastContext'

export function SettingsPage() {
  const { theme, setTheme } = useTheme()
  const toast = useToast()
  const [confirm, setConfirm] = useState(false)
  async function wipe() {
    await db.transaction('rw', db.files, db.meta, async () => { await Promise.all([db.files.clear(), db.meta.clear()]) })
    toast('All local data erased')
  }
  const Opt = ({ id, label, Icon }: { id: 'light' | 'dark'; label: string; Icon: typeof Sun }) => (
    <button onClick={() => setTheme(id)} aria-pressed={theme === id}
      className={cx('flex items-center gap-2 h-10 px-4 rounded-xl border text-[13.5px] transition-colors', theme === id ? 'border-accent bg-accent-soft text-accent font-medium' : 'border-line hover:bg-panel2')}><Icon size={16} />{label}</button>
  )
  return (
    <div className="h-full flex flex-col min-h-0">
      <TopBar contained />
      <div className="flex-1 overflow-auto">
        <div className="max-w-[720px] mx-auto px-10 pt-6 pb-16">
          <h1 className="font-display text-[40px] leading-none">Settings</h1>
          <section className="mt-10" aria-labelledby="appearance">
            <h2 id="appearance" className="text-[15px] font-semibold">Appearance</h2>
            <p className="text-muted text-[13.5px] mt-1">Currently using the {theme} theme. Your choice is remembered on this device.</p>
            <div className="mt-4 flex gap-2"><Opt id="light" label="Light" Icon={Sun} /><Opt id="dark" label="Dark" Icon={Moon} /></div>
          </section>
          <section className="mt-12" aria-labelledby="storage">
            <h2 id="storage" className="text-[15px] font-semibold">Your data</h2>
            <p className="text-muted text-[13.5px] mt-1 max-w-[56ch] leading-relaxed">Documents live in this browser’s storage, and nothing is uploaded. Clearing site data removes them, so download anything you want to keep from Drive.</p>
            <Button variant="danger" className="mt-4" onClick={() => setConfirm(true)}>Erase all local data</Button>
          </section>
          <section className="mt-12 flex items-start gap-3 text-muted text-[13px]"><Monitor size={16} className="mt-0.5 shrink-0" /><p className="max-w-[56ch] leading-relaxed">Office One is an independent project and isn’t affiliated with Microsoft.</p></section>
        </div>
      </div>
      {confirm && <ConfirmModal title="Erase everything?" confirm="Erase everything" message="Every file on this device will be permanently deleted. This can’t be undone." onConfirm={wipe} onClose={() => setConfirm(false)} />}
    </div>
  )
}
