/* Office One service worker: makes the app shell available offline.
 * - Navigations: network first, falling back to the cached shell (so a reload offline still opens the app).
 * - /assets/* are content-hashed by the build, so they are cache-first and safe to keep forever.
 * Documents live in IndexedDB, not here, so this worker never touches user data. */
const CACHE = 'office-one-v1'
const SHELL = ['/', '/icon.svg', '/manifest.webmanifest', '/theme-init.js']

self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())) })
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()))
})
self.addEventListener('fetch', e => {
  const req = e.request
  if (req.method !== 'GET') return
  const url = new URL(req.url)
  if (url.origin !== location.origin) return // Google Fonts etc. are left to the browser's own cache
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('/', copy)); return res }).catch(() => caches.match('/')))
    return
  }
  if (url.pathname.startsWith('/assets/')) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => { if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)) } return res })))
  }
})
