// Office One for macOS: a thin Electron shell around the production web build.
// The app is served from a private app:// origin (not file://), so IndexedDB, fonts and the
// Content-Security-Policy behave exactly as they do on the web. Nothing is fetched from a server.
const { app, BrowserWindow, Menu, protocol, net, shell, nativeTheme } = require('electron')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const fs = require('node:fs')

const DIST = path.join(__dirname, 'dist')
const ORIGIN = 'app://office-one'

protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, corsEnabled: true } },
])

function serve(request) {
  const url = new URL(request.url)
  let rel = decodeURIComponent(url.pathname)
  if (rel === '/' || rel === '') rel = '/index.html'
  const file = path.normalize(path.join(DIST, rel))
  // never leave the bundled folder
  if (!file.startsWith(DIST)) return new Response('Forbidden', { status: 403 })
  // real files are served as-is; unknown paths are app routes (/word/123), so they get the SPA shell
  const target = fs.existsSync(file) && fs.statSync(file).isFile() ? file : path.join(DIST, 'index.html')
  return net.fetch(pathToFileURL(target).toString())
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1440, height: 900, minWidth: 900, minHeight: 600,
    title: 'Office One',
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#0d1019' : '#edeff5',
    show: false,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: true },
  })
  win.once('ready-to-show', () => win.show())
  void win.loadURL(`${ORIGIN}/`)

  // links to other sites open in the default browser, never inside the app window
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) void shell.openExternal(url); return { action: 'deny' } })
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith(ORIGIN)) { e.preventDefault(); if (/^https?:/.test(url)) void shell.openExternal(url) } })
  // the editors set the document name as the page title; keep the window title in step
  win.webContents.on('page-title-updated', (_e, title) => win.setTitle(title))
  return win
}

function buildMenu() {
  const isMac = process.platform === 'darwin'
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    ...(isMac ? [{ role: 'appMenu' }] : []),
    { label: 'File', submenu: [{ label: 'New Window', accelerator: 'CmdOrCtrl+Shift+N', click: () => createWindow() }, { type: 'separator' }, { role: isMac ? 'close' : 'quit' }] },
    { role: 'editMenu' },
    { label: 'View', submenu: [{ role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }] },
    { role: 'windowMenu' },
  ]))
}

app.whenReady().then(() => {
  protocol.handle('app', serve)
  buildMenu()
  createWindow()
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow() })
})
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit() })
