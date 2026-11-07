import { saveAs } from 'file-saver'
export function download(data: Blob | string, filename: string, type = 'application/octet-stream') {
  saveAs(data instanceof Blob ? data : new Blob([data], { type }), filename)
}
export const safeName = (n: string) => (n.trim() || 'Untitled').replace(/[\\/:*?"<>|]/g, '_')
