import { useParams } from 'react-router-dom'
import { Loader2 } from 'lucide-react'
import { useFileLoader } from '../../hooks/useFileLoader'
import { FileMissing } from '../../components/FileMissing'
import { ExcelEditor } from '../../editors/excel/ExcelEditor'

export default function ExcelApp() {
  const { id } = useParams()
  const file = useFileLoader(id)
  if (file === null) return <div className="h-full flex items-center justify-center"><Loader2 className="animate-spin text-muted" /></div>
  if (file === 'missing') return <FileMissing />
  return <ExcelEditor key={file.id} file={file} />
}
