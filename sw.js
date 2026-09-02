/* Olisa Tools service worker — keeps the app shell and libraries cached so the installed app
   opens instantly. Google Drive/API traffic is NEVER cached: live data must stay live. */

/* Bump this on EVERY release. The name is the only thing that evicts the old cache: the activate
   handler deletes every cache whose key isn't this one. It sat on v24 through three releases,
   which left anyone opening the app offline on a stale build. */
const CACHE = 'olisa-tools-v51';
/* The libraries are now OURS, served from this repo, so they belong in the shell alongside the
   pages. Before this they were fetched from cdnjs on every cold start: the app only worked offline
   AFTER a successful online launch had populated the cache, and a blocked or slow cdnjs stopped it
   dead. Now everything the app needs to open ships with the app. */
const LIB = ['./lib/xlsx.full.min.js', './lib/exceljs.min.js', './lib/pdf.min.js',
             './lib/pdf-lib.min.js', './lib/mammoth.browser.min.js', './lib/pdf.worker.min.js'];
const SHELL = ['./', './index.html', './olisa.html', './calculator.html', './DC_Bypass_Bill.html', './manifest.json', './icon-192.png', './icon-512.png', './calc-icon-512.png', './apple-touch-icon.png'].concat(LIB);

self.addEventListener('install', e => {
  // cache.addAll() is ATOMIC: one 404 anywhere in the list rejects the whole install, the worker
  // never activates, and offline support silently stops existing — with nothing to indicate why.
  // Caching each entry independently means a single renamed or missing icon costs that one file
  // instead of the entire offline capability.
  e.waitUntil(
    caches.open(CACHE)
      .then(c => Promise.all(SHELL.map(u =>
        c.add(u).catch(err => { console.warn('[sw] skipped (not cached):', u, err && err.message); })
      )))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (url.hostname.includes('googleapis.com') || url.hostname.includes('accounts.google.com')) return; // live only
  if (url.hostname.includes('cdnjs.cloudflare.com') || url.hostname.includes('fonts.g')) {
    // libraries/fonts: cache-first (they're versioned URLs)
    e.respondWith(caches.open(CACHE).then(async c =>
      (await c.match(e.request)) || fetch(e.request).then(r => { if (r.status === 200) c.put(e.request, r.clone()); return r; })
    ));
    return;
  }
  // The libraries carry their version in the file itself and never change between releases, so
  // they are served cache-first. Leaving them on the shell's network-first path would mean asking
  // the network for 3 MB on every single launch — which would make self-hosting SLOWER on a weak
  // connection than the CDN was, and losing that was the whole point.
  if (url.origin === location.origin && /\/lib\/[^/]+\.js$/.test(url.pathname)) {
    e.respondWith(caches.open(CACHE).then(async c => {
      const hit = await c.match(e.request);
      if (hit) return hit;
      const r = await fetch(e.request);
      if (r && r.status === 200 && r.type === 'basic') c.put(e.request, r.clone());
      return r;
    }));
    return;
  }
  if (url.origin === location.origin) {
    // app shell: network-first so a newly uploaded index.html wins; cached copy as offline fallback.
    // Only basic (same-origin, non-opaque) 200s are stored — caching an opaque or partial response
    // here would serve an unusable body back later with no way to tell it was broken.
    e.respondWith(
      fetch(e.request)
        .then(r => {
          // r.ok spans 200-299, which includes 206 Partial Content. Caching a partial body under
          // the full URL hands back a truncated file on the next load, and a range request is the
          // one case where that can happen — so require a plain 200 and no Range header.
          if (r.status === 200 && r.type === 'basic' && !e.request.headers.has('range')) {
            const copy = r.clone();
            e.waitUntil(caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {}));
          }
          return r;
        })
        .catch(() => caches.match(e.request).then(m => m || caches.match('./index.html')))
    );
  }
});
