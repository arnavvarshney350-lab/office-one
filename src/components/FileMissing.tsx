import { Link } from 'react-router-dom'
export function FileMissing() {
  return <div className="h-full flex flex-col items-center justify-center gap-3 text-muted"><div>This file doesn't exist or is in the trash.</div><Link className="text-accent underline" to="/files">Go to files</Link></div>
}
