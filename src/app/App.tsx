import { lazy, Suspense, useEffect } from 'react'
import { Navigate, Route, Routes, useLocation } from 'react-router-dom'
import { ErrorBoundary } from '../components/ErrorBoundary'
import { StorageGuard } from '../components/StorageGuard'
import { Shell } from './Shell'
import { HomePage } from '../applications/home/HomePage'
import { FilesPage } from '../applications/home/FilesPage'
import { SettingsPage } from '../applications/home/SettingsPage'
import { AppLanding } from '../applications/home/AppLanding'
import { Loader2 } from 'lucide-react'

const WordEditor = lazy(() => import('../applications/word/WordApp'))
const ExcelEditor = lazy(() => import('../applications/excel/ExcelApp'))
const PowerPointEditor = lazy(() => import('../applications/powerpoint/PowerPointApp'))

const Loading = () => <div className="h-full flex items-center justify-center text-muted"><Loader2 className="animate-spin" /></div>

export function App() {
  const { pathname } = useLocation()
  useEffect(() => { void navigator.storage?.persist?.() }, [])
  return (
    <Shell>
      <StorageGuard>
      <ErrorBoundary resetKey={pathname}>
      <Suspense fallback={<Loading />}>
        <Routes>
          <Route path="/" element={<HomePage />} />
          <Route path="/files" element={<FilesPage />} />
          <Route path="/settings" element={<SettingsPage />} />
          <Route path="/word" element={<AppLanding kind="docx" />} />
          <Route path="/excel" element={<AppLanding kind="xlsx" />} />
          <Route path="/powerpoint" element={<AppLanding kind="pptx" />} />
          <Route path="/word/:id" element={<WordEditor />} />
          <Route path="/excel/:id" element={<ExcelEditor />} />
          <Route path="/powerpoint/:id" element={<PowerPointEditor />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </Suspense>
      </ErrorBoundary>
      </StorageGuard>
    </Shell>
  )
}
