/* Olisa Tools service worker — keeps the app shell and libraries cached so the installed app
   opens instantly. Google Drive/API traffic is NEVER cached: live data must stay live. */
const CACHE = 'olisa-tools-v22';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.hostname.includes('googleapis.com') || url.hostname.includes('accounts.google.com')) return; // live only
  if (url.hostname.includes('cdnjs.cloudflare.com') || url.hostname.includes('fonts.g')) {
    // libraries/fonts: cache-first (they're versioned URLs)
    e.respondWith(caches.open(CACHE).then(async c => (await c.match(e.request)) || fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; })));
    return;
  }
  if (url.origin === location.origin) {
    // app shell: network-first so a newly uploaded index.html wins; cached copy as offline fallback
    e.respondWith(fetch(e.request).then(r => { if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; }).catch(() => caches.match(e.request).then(m => m || caches.match('./index.html'))));
  }
});
