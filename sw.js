/* Olisa Tools service worker — keeps the app shell and libraries cached so the installed app
   opens instantly. Google Drive/API traffic is NEVER cached: live data must stay live. */

/* Bump this on EVERY release. The name is the only thing that evicts the old cache: the activate
   handler deletes every cache whose key isn't this one. It sat on v24 through three releases,
   which left anyone opening the app offline on a stale build. */
const CACHE = 'olisa-tools-v139';
/* The libraries are now OURS, served from this repo, so they belong in the shell alongside the
   pages. Before this they were fetched from cdnjs on every cold start: the app only worked offline
   AFTER a successful online launch had populated the cache, and a blocked or slow cdnjs stopped it
   dead. Now everything the app needs to open ships with the app. */
const LIB = ['./lib/xlsx.full.min.js', './lib/exceljs.min.js', './lib/pdf.min.js',
             './lib/pdf-lib.min.js', './lib/mammoth.browser.min.js', './lib/pdf.worker.min.js',
             // DC Bypass Bill Maker's own two: zip packaging and PDF output. It shares pdf.js with
             // the main app rather than carrying a second copy of it.
             './lib/jszip.min.js', './lib/jspdf.umd.min.js'];
const SHELL = ['./', './index.html', './olisa.html', './calculator.html', './DC_Bypass_Bill.html',
               './Pending_PI_Summary.html', './manifest.json', './icon-192.png', './icon-512.png', './calc-icon-512.png', './apple-touch-icon.png'].concat(LIB);

self.addEventListener('install', e => {
  // cache.addAll() is ATOMIC: one 404 anywhere in the list rejects the whole install, the worker
  // never activates, and offline support silently stops existing — with nothing to indicate why.
  // Caching each entry independently means a single renamed or missing icon costs that one file
  // instead of the entire offline capability.
  //
  // v139: every entry is fetched with cache:'reload'. A plain c.add('./olisa.html') is allowed to
  // be answered by the browser's own HTTP cache, and GitHub Pages lets a page be reused for ten
  // minutes. Installing a new release inside that window therefore filled the NEW cache with the
  // OLD page — and that stale copy is exactly what the app falls back to when it is opened
  // offline. 'reload' goes to the network every time, so the shell that gets stored is the one
  // that was just published.
  e.waitUntil(
    caches.open(CACHE)
      .then(c => Promise.all(SHELL.map(path => new Request(path, { cache: 'reload' })).map(u =>
        c.add(u).catch(err => { console.warn('[sw] skipped (not cached):', u.url, err && err.message); })
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
  // ==================== v139: THE FONTS WERE NEVER ACTUALLY CACHED ====================
  // Two separate things stopped it, and either one alone was enough:
  //   1. fonts.googleapis.com CONTAINS "googleapis.com", so the live-only rule just below returned
  //      before the font rule was ever reached. The stylesheet went to the network every time.
  //   2. A <link rel="stylesheet"> to another origin is a no-cors request. Its answer is opaque:
  //      status 0, unreadable. The old test was `r.status === 200`, which an opaque answer can
  //      never pass, so even with (1) fixed nothing would have been stored.
  // The result offline was the fallback typeface on every page, and font files sitting in the
  // cache that no stylesheet ever asked for. The font hosts are now matched by exact name BEFORE
  // the live-only rule, and a cache miss is fetched in CORS mode (both hosts allow it), which
  // gives a real 200 that can be checked before it is kept. Same treatment for the cdnjs fallback.
  const assetHost = url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com'
                 || url.hostname === 'cdnjs.cloudflare.com';
  if (assetHost) {
    // fonts/fallback libraries: cache-first (their URLs are versioned)
    e.respondWith(caches.open(CACHE).then(async c => {
      // ignoreVary: the font stylesheet is answered with a Vary header, and a stored copy must
      // still be found for the next page that asks for the same address.
      const hit = await c.match(e.request.url, { ignoreVary: true });
      if (hit) return hit;
      try {
        const r = await fetch(e.request.url, { mode: 'cors', credentials: 'omit' });
        if (r && r.status === 200) c.put(e.request.url, r.clone()).catch(() => {});
        return r;
      } catch (err) {
        return fetch(e.request);   // CORS refused for some reason: behave exactly as the browser would
      }
    }));
    return;
  }
  if (url.hostname.includes('googleapis.com') || url.hostname.includes('accounts.google.com')) return; // live only
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
        .catch(async () => {
          // Offline. The exact cached copy first, always.
          const exact = await caches.match(e.request);
          if (exact) return exact;
          // v139: the landing page used to be handed back for ANYTHING that was not cached — an
          // icon, the manifest, a script — so the browser received a page of HTML where it asked
          // for an image or JSON. That fallback only makes sense for a page being opened. A page
          // opened with something after a "?" is also the same page, so the cached copy is found
          // without the query before giving up and showing the tool menu.
          if (e.request.mode === 'navigate') {
            const samePage = await caches.match(e.request, { ignoreSearch: true });
            return samePage || (await caches.match('./index.html')) || Response.error();
          }
          return Response.error();
        })
    );
  }
});
