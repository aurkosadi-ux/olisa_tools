/* autoindextest.js — "can challans and work orders update themselves, like the Master File?"
 *
 * This does not read the source and hope. It stands a whole fake Google Drive up behind the real
 * page — folders, listings, downloads, uploads, and a real changes feed with page tokens — boots
 * olisa.html in a DOM, and then does what Sadi does: drops a new challan into the folder and walks
 * away. What it measures is behaviour:
 *
 *   1. it indexes by itself, with nothing pressed
 *   2. it reads ONLY the new file; everything already read is reused untouched
 *   3. nothing at all happens when nothing changed — no build, no upload, no downloads
 *   4. a change elsewhere in the Drive is not mistaken for ours
 *   5. searches keep answering from the complete index while an automatic build runs
 *   6. the phone carries on from the laptop's cursor instead of re-reading every work order
 *   7. a phone on mobile data with a real backlog declines it, says so, and goes when Wi-Fi arrives
 *   8. the tool's own OCR output is never mistaken for news
 *
 * Usage: node autoindextest.js olisa.html                                                       */
const fs = require('fs');
const path = require('path');
require('fake-indexeddb/auto');
const { IDBFactory } = require('fake-indexeddb');
const { JSDOM, VirtualConsole } = require('jsdom');
const XLSX = require('xlsx');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x !== undefined ? '  -> ' + x : ''))); };
const section = s => console.log('\n' + s);

const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
let html = fs.readFileSync(file, 'utf8');
const xlsxLib = fs.readFileSync(path.join(dir, 'lib', 'xlsx.full.min.js'), 'utf8');
html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${xlsxLib}</script>`);
html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');

// ---- test probe, injected into THIS COPY only, just inside the Ask IIFE ----
// Two parser stubs go in with it. This test is about WHICH files are read and WHEN, not about the
// insides of a PDF — those have tests of their own — so a "PDF" here is plain text the stubs read.
const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
if (!html.includes(anchor)) { console.log('FAIL probe anchor not found'); process.exit(1); }
html = html.replace(anchor, () => anchor + `
const __realBuild = buildStyleIndex;
buildStyleIndex = async function (o) { window.__t.builds++; return __realBuild(o); };
scanChallanDoc = async function (f, p, driveId) {
  const txt = new TextDecoder().decode(new Uint8Array(await f.arrayBuffer()));
  const m = txt.match(/CHALLAN:(\\S+) PI:(\\S+) DATE:(\\S+) STYLE:(\\S+) QTY:(\\d+)/);
  if (!m) return [];
  return [{ no: m[1], pages: [1], parts: 1, piRef: m[2], date: m[3], verified: true, hand: false,
            path: p, driveId, lines: [{ style: m[4], qty: Number(m[5]), desc: 'Master Carton' }] }];
};
extractPdfLines = async function (f) {
  const txt = new TextDecoder().decode(new Uint8Array(await f.arrayBuffer()));
  return txt.split('\\n');
};
// A person waits minutes for these; a test cannot. Applied here, inside the page, before the first
// schedule is made — the logic under test is identical, only the clock is impatient.
Object.assign(AUTOIDX, window.__testCfg || {});
window.__t = {
  builds: 0,
  scheduleAutoIndex,
  pollDriveChanges, runAutoIndex, autoIndexStart, checkForNewerSharedIndex, noteSelfWrite,
  cfg: AUTOIDX,
  get feed() { return { ...autoFeed }; },
  get pendingCount() { return autoPendingCount(); },
  get pendingRaw() { return autoPending ? Object.keys(autoPending.items).length : 0; },
  get blocked() { return autoBlocked; },
  get building() { return indexBuilding; },
  get autoBuilding() { return autoBuildRunning; },
  get styleIndex() { return styleIndex; },
  get piIndex() { return piIndex; },
  get challanIndex() { return challanIndex; },
  get styleCacheSize() { return Object.keys(styleFileCache).length; },
  get bell() { return bellLog.map(b => b.msg); },
  get connected() { return !!(savedRootHandle && piRootHandle); },
  get isPhone() { return isPhoneDevice(); },
  payload: () => currentIndexPayload(),
  navGuard: () => window.__olisaBuilding(),
  setPhone(v) { isPhoneDevice = () => !!v; }
};`);

// ============================ a fake Google Drive ============================
const FOLDER = 'application/vnd.google-apps.folder';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Drive's query language, enough of it: 'id' in parents / name contains / name= / mimeType(!)= /
// trashed= / modifiedTime > , joined by and/or with brackets.
function parseQuery(q) {
  let i = 0;
  const ws = () => { while (i < q.length && /\s/.test(q[i])) i++; };
  const str = () => { ws(); if (q[i] !== "'") throw new Error('quote expected: ' + q.slice(i)); i++; let o = ''; while (i < q.length && q[i] !== "'") o += q[i++]; i++; return o; };
  const word = () => { ws(); let o = ''; while (i < q.length && /[A-Za-z]/.test(q[i])) o += q[i++]; return o; };
  function term() {
    ws();
    if (q[i] === '(') { i++; const p = expr(); ws(); if (q[i] === ')') i++; return p; }
    if (q[i] === "'") { const id = str(); word(); word(); return f => (f.parents || []).includes(id); }
    const key = word(); ws();
    if (key === 'name' && q.startsWith('contains', i)) { i += 8; const v = str().toLowerCase(); return f => String(f.name).toLowerCase().includes(v); }
    if (q[i] === '!' && q[i + 1] === '=') { i += 2; const v = str(); return f => String(f[key]) !== v; }
    if (q[i] === '=') {
      i++;
      if (key === 'trashed') { const v = word(); return f => String(!!f.trashed) === v; }
      const v = str(); return f => String(f[key]) === v;
    }
    if (q[i] === '>' || q[i] === '<') { const op = q[i++]; const v = Date.parse(str()); return f => op === '>' ? Date.parse(f.modifiedTime) > v : Date.parse(f.modifiedTime) < v; }
    throw new Error('unsupported query near: ' + q.slice(Math.max(0, i - 12)));
  }
  function andE() { let p = term(); for (;;) { const s = i; if (word() === 'and') { const r = term(), l = p; p = f => l(f) && r(f); } else { i = s; return p; } } }
  function expr() { let p = andE(); for (;;) { const s = i; if (word() === 'or') { const r = andE(), l = p; p = f => l(f) || r(f); } else { i = s; return p; } } }
  return expr();
}

function makeDrive() {
  const files = new Map();
  const changes = [];
  const counters = { list: 0, meta: 0, changes: 0, upload: 0, startToken: 0 };
  const downloads = [];          // ids, in order, so a test can say exactly what came down the wire
  let seq = 1, win = null, clock = Date.parse('2026-09-20T08:00:00.000Z');
  const stalls = new Map();      // file id -> ms to hold that one download open
  const bytes = v => (typeof v === 'string' ? new TextEncoder().encode(v) : v);
  const stamp = () => { clock += 60000; return new Date(clock).toISOString(); };
  function touch(id) { changes.push({ seq: seq++, fileId: id, time: new Date(clock).toISOString() }); }
  function add(o) {
    const c = o.content === undefined || o.content === null ? null : bytes(o.content);
    const f = {
      id: o.id, name: o.name, mimeType: o.mimeType || 'application/octet-stream',
      parents: o.parents || [], trashed: false, content: c,
      size: String(c ? c.length : 0), modifiedTime: o.modifiedTime || stamp()
    };
    files.set(f.id, f);
    if (o.quiet !== true) touch(f.id);
    return f;
  }
  const meta = f => ({ id: f.id, name: f.name, mimeType: f.mimeType, parents: f.parents.slice(), trashed: !!f.trashed, size: f.size, modifiedTime: f.modifiedTime });
  const json = o => ({ ok: true, status: 200, json: async () => o, text: async () => JSON.stringify(o) });
  async function bodyText(b) {
    if (!b) return '';
    if (typeof b === 'string') return b;
    if (b.text) return await b.text();
    return String(b);
  }
  async function handle(url, opts) {
    opts = opts || {};
    const u = new URL(url);
    const p = u.pathname, qs = u.searchParams;
    const method = (opts.method || 'GET').toUpperCase();

    if (p === '/drive/v3/about') return json({ user: { emailAddress: 'sadi@example.com' } });
    if (p === '/drive/v3/changes/startPageToken') { counters.startToken++; return json({ startPageToken: String(seq) }); }
    if (p === '/drive/v3/changes') {
      counters.changes++;
      const from = parseInt(qs.get('pageToken') || '1', 10);
      const size = Math.min(parseInt(qs.get('pageSize') || '100', 10), 1000);
      const all = changes.filter(c => c.seq >= from);
      const page = all.slice(0, size);
      const out = page.map(c => {
        const f = files.get(c.fileId);
        if (!f) return { changeType: 'file', time: c.time, removed: true, fileId: c.fileId };
        return { changeType: 'file', time: c.time, removed: false, fileId: f.id, file: meta(f) };
      });
      const body = { changes: out };
      if (all.length > page.length) body.nextPageToken = String(page[page.length - 1].seq + 1);
      else body.newStartPageToken = String(seq);
      return json(body);
    }
    if (p.startsWith('/upload/drive/v3/files')) {
      counters.upload++;
      const id = p.split('/').pop();
      if (method === 'PATCH' && files.has(id)) {
        const f = files.get(id);
        f.content = bytes(await bodyText(opts.body));
        f.size = String(f.content.length);
        f.modifiedTime = stamp();
        touch(id);
        return json({ id, modifiedTime: f.modifiedTime });
      }
      // multipart create: FormData (index upload) or a hand-built multipart body (OCR)
      let name = 'upload', parents = [], content = '';
      const b = opts.body;
      if (b && typeof b.get === 'function') {
        const m = b.get('metadata');
        const j = JSON.parse(await bodyText(m));
        name = j.name; parents = j.parents || [];
        content = await bodyText(b.get('file'));
      } else {
        const raw = await bodyText(b);
        const m = raw.match(/\{[\s\S]*?\}/);
        if (m) { const j = JSON.parse(m[0]); name = j.name; parents = j.parents || []; }
        content = raw;
      }
      const made = add({ id: 'U' + (files.size + 1), name, parents, mimeType: 'application/json', content });
      return json({ id: made.id, name: made.name, size: made.size, modifiedTime: made.modifiedTime, parents: made.parents });
    }
    if (p === '/drive/v3/files' && method === 'POST') {
      const j = JSON.parse(await bodyText(opts.body));
      const made = add({ id: 'D' + (files.size + 1), name: j.name, mimeType: j.mimeType, parents: j.parents || [] });
      return json({ id: made.id });
    }
    if (p === '/drive/v3/files' && method === 'GET') {
      counters.list++;
      const pred = parseQuery(qs.get('q') || '');
      let hits = [...files.values()].filter(f => { try { return pred(f); } catch (e) { return false; } });
      if ((qs.get('orderBy') || '').startsWith('modifiedTime desc')) hits.sort((a, b) => Date.parse(b.modifiedTime) - Date.parse(a.modifiedTime));
      const size = Math.min(parseInt(qs.get('pageSize') || '100', 10), 1000);
      const start = parseInt(qs.get('pageToken') || '0', 10);
      const page = hits.slice(start, start + size);
      const body = { files: page.map(meta) };
      if (start + size < hits.length) body.nextPageToken = String(start + size);
      return json(body);
    }
    const one = p.match(/^\/drive\/v3\/files\/([^/]+)$/);
    if (one) {
      const id = one[1];
      if (id === 'root') return json({ id: 'MYDRIVE' });
      const f = files.get(id);
      if (!f) return { ok: false, status: 404, json: async () => ({ error: { message: 'not found' } }) };
      if (method === 'PATCH') {
        const j = JSON.parse(await bodyText(opts.body));
        if (j.name) f.name = j.name;
        f.modifiedTime = stamp(); touch(id);
        return json(meta(f));
      }
      if (qs.get('alt') === 'media') {
        downloads.push(id);
        const hold = stalls.get(id);
        if (hold) { stalls.delete(id); await new Promise(r => setTimeout(r, hold)); }
        const B = (win && win.Blob) || Blob;
        return { ok: true, status: 200, blob: async () => new B([f.content || new Uint8Array()]), json: async () => ({}), text: async () => new TextDecoder().decode(f.content || new Uint8Array()) };
      }
      counters.meta++;
      return json(meta(f));
    }
    return { ok: false, status: 404, json: async () => ({ error: { message: 'no route ' + p } }) };
  }
  return {
    files, counters, downloads, add, touch,
    attach(w) { win = w; },
    stall(id, ms) { stalls.set(id, ms); },
    rename(id, name) { const f = files.get(id); f.name = name; f.modifiedTime = stamp(); touch(id); },
    edit(id, content) { const f = files.get(id); f.content = bytes(content); f.size = String(f.content.length); f.modifiedTime = stamp(); touch(id); },
    trash(id) { const f = files.get(id); f.trashed = true; f.modifiedTime = stamp(); touch(id); },
    text(id) { const f = files.get(id); return f && f.content ? new TextDecoder().decode(f.content) : ''; },
    idByName(n) { for (const f of files.values()) if (f.name === n) return f.id; return null; },
    fetch: (url, opts) => handle(String(url), opts)
  };
}

// ============================ the folder, as it really looks ============================
const HEAD = ['PI Ref', 'Inventory Date', 'Style Ref', 'Item', 'Item Description', 'Delivery Qty',
  'Received Qty', 'Short&Exs', 'Challan No', 'Challan Date', 'Remarks'];
function masterBytes(rows) {
  const ws = XLSX.utils.aoa_to_sheet([HEAD, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '2026- Main');
  return new Uint8Array(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}
function workOrderBytes(style) {
  const ws = XLSX.utils.aoa_to_sheet([['Work Order'], ['Style', style], ['Item', 'Master Carton'], ['Qty', 500]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'WO');
  return new Uint8Array(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}
const challanDoc = (no, pi, style, qty) =>
  `%PDF-1.4\nDELIVERY CHALLAN\nCHALLAN:${no} PI:${pi} DATE:09-06-2026 STYLE:${style} QTY:${qty}\n`;
// A Proforma Invoice as the reader sees it once the text is off the page: a header row, numbered
// item rows ending in "Piece", and the invoice's own stated total.
const invoiceDoc = (pi, style, qty) =>
  ['%PDF-1.4', 'PROFORMA INVOICE', `PI NO : 121/700768/0${pi}062026`, 'PI Issue Date: 08/06/2026',
   'SL  Item Description  Quantity',
   `1  ${style} Master Carton L71.5 x W45 x H32 cm  ${qty}  Piece`,
   `Total Qty (Pcs): ${qty}`].join('\n') + '\n';

const PIS = [
  { ref: '124', style: 'M82039-1B1-5' },
  { ref: '130', style: 'M82040-2A1-1' },
  { ref: '131', style: 'M82041-3C1-2' }
];
function buildWorld() {
  const d = makeDrive();
  d.add({ id: 'MYDRIVE', name: 'My Drive', mimeType: FOLDER, parents: [], quiet: true });
  d.add({ id: 'ROOT', name: 'OLISA GROUP', mimeType: FOLDER, parents: ['MYDRIVE'], quiet: true });
  d.add({ id: 'WO', name: 'Work Orders', mimeType: FOLDER, parents: ['ROOT'], quiet: true });
  d.add({ id: 'CH', name: 'Challans', mimeType: FOLDER, parents: ['ROOT'], quiet: true });
  d.add({ id: 'CHJUN', name: 'June - 2026', mimeType: FOLDER, parents: ['CH'], quiet: true });
  d.add({ id: 'ELSE', name: 'Family photos', mimeType: FOLDER, parents: ['MYDRIVE'], quiet: true });
  d.add({
    id: 'MASTER', name: 'Olisa Master Inventory \u2014 09-08-2026, 04-28 pm \u2014 v9.xlsx',
    mimeType: XLSX_MIME, parents: ['ROOT'], quiet: true,
    content: masterBytes(PIS.map((p, i) => [Number(p.ref), '09/06/2026', p.style, 'Master Carton',
      'L71.5 x W45 x H32 cm', 100, 100 - i, -i, '3191079' + i, '09/06/2026', '']))
  });
  PIS.forEach(p => {
    d.add({ id: 'F' + p.ref, name: `PI-${p.ref}; ${p.style}`, mimeType: FOLDER, parents: ['WO'], quiet: true });
    d.add({ id: 'W' + p.ref, name: `121-700768-0${p.ref} Work Order.xlsx`, mimeType: XLSX_MIME, parents: ['F' + p.ref], quiet: true, content: workOrderBytes(p.style) });
    d.add({ id: 'I' + p.ref, name: `Proforma Invoice-Olisa Group -700768-0${p.ref}.pdf`, mimeType: 'application/pdf', parents: ['F' + p.ref], quiet: true, content: invoiceDoc(p.ref, p.style, 500) });
  });
  d.add({ id: 'C1', name: '09-6-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'], quiet: true, content: challanDoc('31910790', '124', 'M82039-1B1-5', 100) });
  d.add({ id: 'C2', name: '10-6-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'], quiet: true, content: challanDoc('31910791', '130', 'M82040-2A1-1', 200) });
  return d;
}

// ============================ booting the real page ============================
const allErrors = [];
async function boot(drive, idb, opts) {
  opts = opts || {};
  const errors = [];
  const vc = new VirtualConsole();
  vc.on('jsdomError', e => { const m = String(e && (e.detail || e.message || e)); errors.push(m); allErrors.push(m); });
  const conn = { type: opts.conn || 'wifi', saveData: !!opts.saveData, __on: [], addEventListener(ev, fn) { this.__on.push(fn); }, removeEventListener() {} };
  const cfg = Object.assign({ quietMs: 40, phoneExtraMs: 0, maxWaitMs: 200, minGapMs: 0, retryMs: 500, busyRetryMs: 250 }, opts.cfg || {});
  // A phone is a phone to the page itself — mobile user agent, a coarse pointer, no mouse, touch —
  // rather than something the test switches on afterwards, which is always a step too late.
  const PHONE_UA = 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Mobile Safari/537.36';
  const dom = new JSDOM(html, {
    runScripts: 'dangerously', pretendToBeVisual: true,
    url: 'https://aurkosadi-ux.github.io/olisa_tools/olisa.html', virtualConsole: vc,
    beforeParse(w) {
      w.indexedDB = idb; w.IDBKeyRange = IDBKeyRange;
      w.__testCfg = cfg;
      if (opts.phone) {
        Object.defineProperty(w.navigator, 'userAgent', { value: PHONE_UA, configurable: true });
        Object.defineProperty(w.navigator, 'maxTouchPoints', { value: 5, configurable: true });
      }
      w.matchMedia = q => ({ matches: !!(opts.phone && /coarse/.test(q)), media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
      w.scrollTo = () => {}; w.Element.prototype.scrollIntoView = function () {};
      w.requestAnimationFrame = cb => setTimeout(() => cb(Date.now()), 0);
      w.cancelAnimationFrame = id => clearTimeout(id);
      if (!w.Blob.prototype.text) w.Blob.prototype.text = function () { return new Promise((res, rej) => { const r = new w.FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(this); }); };
      if (!w.Blob.prototype.arrayBuffer) w.Blob.prototype.arrayBuffer = function () { return new Promise((res, rej) => { const r = new w.FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(this); }); };
      Object.defineProperty(w.navigator, 'connection', { value: conn, configurable: true });
      w.navigator.serviceWorker = { register: () => Promise.resolve({ addEventListener() {} }), addEventListener() {}, ready: Promise.resolve({}) };
      w.PDFLib = {}; w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {};
      w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
      w.CompressionStream = undefined; w.DecompressionStream = undefined;
      w.confirm = () => true; w.prompt = () => 'x';
      w.fetch = (url, o) => drive.fetch(url, o);
      drive.attach(w);
    }
  });
  const w = dom.window;
  await new Promise(r => setTimeout(r, 300));
  return { w, d: w.document, dom, errors, conn };
}
const tick = (ms = 50) => new Promise(r => setTimeout(r, ms));
async function until(fn, ms, label) {
  const stop = Date.now() + (ms || 15000);
  while (Date.now() < stop) { if (fn()) return true; await tick(25); }
  return false;
}
// A whole automatic cycle: it notices, it waits for the burst to settle, it builds, it finishes.
async function settle(w, ms) {
  await until(() => w.__t.autoBuilding, 4000);
  await until(() => !w.__t.autoBuilding && !w.__t.building, ms || 20000);
  await tick(150);
}
(async function run() {
  // The page's own IndexedDB names are read from the source so this cannot drift from them.
  const dbName = (html.match(/indexedDB\.open\('([^']+)'/) || [])[1];
  const storeName = (html.match(/createObjectStore\('([^']+)'\)/) || [])[1];
  if (!dbName || !storeName) { console.log('FAIL could not read the IndexedDB names out of the page'); process.exit(1); }
  const seed = (idb) => (key, val) => new Promise((res, rej) => {
    const req = idb.open(dbName, 1);
    req.onupgradeneeded = () => { const db = req.result; if (!db.objectStoreNames.contains(storeName)) db.createObjectStore(storeName); };
    req.onsuccess = () => {
      const db = req.result;
      const tx = db.transaction(storeName, 'readwrite');
      tx.objectStore(storeName).put(val, key);
      tx.oncomplete = () => { db.close(); res(); };
      tx.onerror = () => rej(tx.error);
    };
    req.onerror = () => rej(req.error);
  });

  const drive = buildWorld();
  const laptopDb = new IDBFactory();
  const put = seed(laptopDb);
  await put('driveSession', { token: 'TEST-TOKEN', exp: Date.now() + 3600e3, hint: 'sadi@example.com' });
  await put('driveRoot', { id: 'ROOT', name: 'OLISA GROUP' });

  section('1. First run: it connects, notices there is no index, and builds one unasked');
  const laptop = await boot(drive, laptopDb);
  const w = laptop.w;
  // On a build with no auto-index the probe cannot attach at all, and every check below would throw
  // an unreadable harness error instead of saying what is missing.
  if (!w.__t) {
    t('this build has an auto-index (the probe found its bindings)', false, 'no window.__t — older build?');
    console.log('\nFAILED — ' + pass + ' passed, ' + fail + ' failed');
    process.exit(1);
  }
  t('the page connected to the folder by itself', await until(() => w.__t.connected, 8000));
  t('an automatic build started with nothing pressed', await until(() => w.__t.autoBuilding, 12000));
  await settle(w, 30000);
  t('the style index knows every PI folder', PIS.every(p => (w.__t.styleIndex[p.style] || []).includes(p.ref)),
    JSON.stringify(Object.keys(w.__t.styleIndex)));
  t('both challans were read', !!w.__t.challanIndex['31910790'] && !!w.__t.challanIndex['31910791'],
    Object.keys(w.__t.challanIndex).join(','));
  t('the index was uploaded for the phone to pick up', !!drive.idByName('olisa-index.json'));
  t('and it carries a Drive change cursor', !!w.__t.payload().changesToken);
  t('and the per-file readings, so another device need not re-read them',
    Object.keys(w.__t.payload().styleFiles).length === PIS.length,
    Object.keys(w.__t.payload().styleFiles).length);

  section('2. Nothing changed: no build, no downloads, no upload');
  let dl = drive.downloads.length, ups = drive.counters.upload, builds = w.__t.builds;
  await w.__t.pollDriveChanges();
  await tick(400);
  t('the poll found nothing of ours', w.__t.pendingCount === 0);
  t('no build was started', w.__t.builds === builds, w.__t.builds - builds);
  t('not one byte was downloaded', drive.downloads.length === dl, drive.downloads.slice(dl).join(','));
  t('and nothing was uploaded', drive.counters.upload === ups);
  t('the index counts as checked against Drive just now', Date.now() - w.__t.feed.syncedAt < 5000);

  section('3. A change somewhere else in the Drive is not ours');
  drive.add({ id: 'P1', name: 'holiday.jpg', mimeType: 'image/jpeg', parents: ['ELSE'], content: 'x' });
  dl = drive.downloads.length; builds = w.__t.builds;
  await w.__t.pollDriveChanges();
  await tick(400);
  t('nothing is waiting', w.__t.pendingCount === 0);
  t('no build was started', w.__t.builds === builds);
  t('nothing was downloaded', drive.downloads.length === dl);

  section('4. A new challan lands in the folder — the whole point');
  drive.add({ id: 'C3', name: '14-9-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'],
              content: challanDoc('31910950', '131', 'M82041-3C1-2', 75) });
  dl = drive.downloads.length; builds = w.__t.builds; ups = drive.counters.upload;
  await w.__t.pollDriveChanges();
  t('it is noticed', w.__t.pendingCount === 1, w.__t.pendingCount);
  await settle(w);
  t('exactly one automatic build ran', w.__t.builds === builds + 1, w.__t.builds - builds);
  t('the challan is now in the index', !!w.__t.challanIndex['31910950']);
  t('its line was read too', (w.__t.challanIndex['31910950'].lines || []).length === 1);
  const after4 = drive.downloads.slice(dl);
  t('ONLY the new file was downloaded — nothing already read was fetched again',
    after4.length === 1 && after4[0] === 'C3', after4.join(','));
  t('the updated index went back to Drive', drive.counters.upload > ups);
  t('and it says in plain words what it added',
    w.__t.bell.some(m => /Indexed automatically/.test(m) && /31910950/.test(m)),
    w.__t.bell.slice(0, 3).join(' | '));

  section('5. A new PI folder with a work order in it');
  drive.add({ id: 'F145', name: 'PI-145; M82077', mimeType: FOLDER, parents: ['WO'] });
  drive.add({ id: 'W145', name: '121-700768-0145 Work Order.xlsx', mimeType: XLSX_MIME, parents: ['F145'], content: workOrderBytes('M82077-9A1-1') });
  dl = drive.downloads.length; builds = w.__t.builds;
  await w.__t.pollDriveChanges();
  t('the new folder and its work order are both noticed', w.__t.pendingCount >= 1, w.__t.pendingCount);
  await settle(w);
  t('one build again', w.__t.builds === builds + 1, w.__t.builds - builds);
  t('the new style is searchable', (w.__t.styleIndex['M82077-9A1-1'] || []).includes('145'),
    JSON.stringify(w.__t.styleIndex['M82077-9A1-1']));
  const after5 = drive.downloads.slice(dl);
  t('the new work order was read', after5.includes('W145'), after5.join(','));
  t('nothing already indexed was read again — no old work order, no old challan',
    !after5.some(id => ['W124', 'W130', 'W131', 'C1', 'C3', 'C4'].includes(id)), after5.join(','));
  t('the summary names the PI', w.__t.bell.some(m => /Indexed automatically/.test(m) && /145/.test(m)),
    w.__t.bell[0]);

  // PI 145 arrived with no invoice in its folder, so this build hunted for one across the whole
  // connected folder and opened the other PIs' invoices once to read the PI number printed inside.
  // That is a one-off: those files are remembered as "not what we were looking for", and the PI is
  // not hunted for again unless something in it changes. The next build must download nothing but
  // the new file — which is the difference between an index that updates itself and one that costs
  // a little more every time it does.
  section('5b. The next automatic build pays nothing for that hunt');
  drive.add({ id: 'C7', name: '18-9-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'],
              content: challanDoc('31910960', '124', 'M82039-1B1-5', 15) });
  dl = drive.downloads.length; builds = w.__t.builds;
  await w.__t.pollDriveChanges();
  await settle(w);
  const after5b = drive.downloads.slice(dl);
  t('one build', w.__t.builds === builds + 1);
  t('and exactly one file downloaded: the new challan', after5b.length === 1 && after5b[0] === 'C7', after5b.join(','));
  t('which is indexed', !!w.__t.challanIndex['31910960']);

  section('6. A deleted challan is noticed as well');
  drive.trash('C2');
  builds = w.__t.builds;
  await w.__t.pollDriveChanges();
  await settle(w);
  t('the challan that left the folder is out of the index', !w.__t.challanIndex['31910791'],
    Object.keys(w.__t.challanIndex).join(','));
  t('and the one still there is untouched', !!w.__t.challanIndex['31910790']);

  section('7. While an automatic build runs, searches keep the complete index');
  const styleBefore = w.__t.styleIndex;
  drive.add({ id: 'F150', name: 'PI-150; M82088', mimeType: FOLDER, parents: ['WO'] });
  drive.add({ id: 'W150', name: '121-700768-0150 Work Order.xlsx', mimeType: XLSX_MIME, parents: ['F150'], content: workOrderBytes('M82088-2B1-1') });
  drive.add({ id: 'C4', name: '15-9-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'],
              content: challanDoc('31910951', '124', 'M82039-1B1-5', 50) });
  // Held open for a moment, so "while the build is running" is a real, observable moment rather
  // than a race against a fake Drive that answers instantly.
  drive.stall('W150', 1200);
  await w.__t.pollDriveChanges();
  await until(() => w.__t.autoBuilding && drive.downloads.includes('W150'), 8000);
  t('the index answering searches is still the finished one, not a half-built copy',
    w.__t.styleIndex === styleBefore);
  t('an old style still resolves mid-build', !!w.__t.styleIndex['M82039-1B1-5']);
  t('and the app-switcher links are NOT hijacked for a build nobody asked for', w.__t.navGuard() === false);
  await settle(w, 30000);
  t('the swap happened at the end', w.__t.styleIndex !== styleBefore);
  t('the new work order is in', (w.__t.styleIndex['M82088-2B1-1'] || []).includes('150'),
    JSON.stringify(w.__t.styleIndex['M82088-2B1-1']));
  t('and so is the new challan', !!w.__t.challanIndex['31910951']);

  section('8. The tool\u2019s own OCR output is not news');
  builds = w.__t.builds;
  w.__t.noteSelfWrite('OCR1');
  drive.add({ id: 'OCR1', name: '15-9-26 Challan-Improved.pdf', mimeType: 'application/pdf', parents: ['CHJUN'], content: challanDoc('31910952', '124', 'M82039-1B1-5', 10) });
  await w.__t.pollDriveChanges();
  await tick(400);
  t('a file this tool just wrote does not start another build', w.__t.builds === builds && w.__t.pendingCount === 0);

  section('9. Pressing the button still works, and covers what was waiting');
  drive.add({ id: 'C5', name: '16-9-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'],
              content: challanDoc('31910953', '130', 'M82040-2A1-1', 20) });
  await w.__t.pollDriveChanges();
  t('something is waiting', w.__t.pendingCount === 1);
  laptop.d.getElementById('askBuildIndexBtn').click();
  await until(() => w.__t.building, 5000);
  t('a build a person started DOES protect itself from navigation', w.__t.navGuard() === true);
  await until(() => !w.__t.building, 30000);
  await tick(200);
  t('the waiting list was cleared by the manual build', w.__t.pendingCount === 0);
  t('and the challan it was waiting for is indexed', !!w.__t.challanIndex['31910953']);

  section('10. The phone continues from the laptop\u2019s cursor instead of re-reading everything');
  const phoneDb = new IDBFactory();
  const putP = seed(phoneDb);
  await putP('driveSession', { token: 'TEST-TOKEN', exp: Date.now() + 3600e3, hint: 'sadi@example.com' });
  await putP('driveRoot', { id: 'ROOT', name: 'OLISA GROUP' });
  const phone = await boot(drive, phoneDb, { conn: 'cellular', phone: true });
  const pw = phone.w;
  t('the page knows it is a phone', pw.__t.isPhone === true);
  t('the phone connected', await until(() => pw.__t.connected, 8000));
  t('it loaded the laptop\u2019s index instead of building one', await until(() => Object.keys(pw.__t.challanIndex).length > 0, 10000),
    Object.keys(pw.__t.challanIndex).length);
  t('it took the laptop\u2019s cursor with it', await until(() => !!pw.__t.feed.token && pw.__t.feed.syncedAt > 0, 6000), JSON.stringify(pw.__t.feed));
  t('and the work-order readings, so it never has to download them',
    await until(() => pw.__t.styleCacheSize >= PIS.length, 6000), pw.__t.styleCacheSize);
  const pdl = drive.downloads.length;
  drive.add({ id: 'C6', name: '17-9-26 Challan.pdf', mimeType: 'application/pdf', parents: ['CHJUN'],
              content: challanDoc('31910954', '131', 'M82041-3C1-2', 30) });
  await pw.__t.pollDriveChanges();
  await settle(pw, 30000);
  t('the phone indexed the new challan by itself', !!pw.__t.challanIndex['31910954']);
  const afterP = drive.downloads.slice(pdl).filter(id => id !== drive.idByName('olisa-index.json'));
  t('on mobile data it downloaded the one new challan and nothing else',
    afterP.filter(id => /^W/.test(id)).length === 0 && afterP.includes('C6'), afterP.join(','));

  section('11. The shared index file is not downloaded again when it has not changed');
  const idxFile = drive.idByName('olisa-index.json');
  const before12 = drive.downloads.filter(id => id === idxFile).length;
  await pw.__t.checkForNewerSharedIndex();
  await pw.__t.checkForNewerSharedIndex();
  await tick(200);
  t('two polls, no repeat download of the same index file',
    drive.downloads.filter(id => id === idxFile).length === before12,
    drive.downloads.filter(id => id === idxFile).length - before12);


  section('12. A phone with a real backlog on mobile data declines it, and says so');
  // Both live devices are shut down first. They are still watching this Drive, and an index upload
  // from either of them mid-scenario would hand the cold phone exactly the readings it is supposed
  // to be without.
  [laptop, phone].forEach(b => { try { b.dom.window.close(); } catch (e) {} });
  await tick(300);
  const coldDb = new IDBFactory();
  const putC = seed(coldDb);
  await putC('driveSession', { token: 'TEST-TOKEN', exp: Date.now() + 3600e3, hint: 'sadi@example.com' });
  await putC('driveRoot', { id: 'ROOT', name: 'OLISA GROUP' });
  // An index file from an older build: no cursor, no readings — exactly what this phone would find
  // the first time it runs this version.
  const idxId = drive.idByName('olisa-index.json');
  const old = JSON.parse(drive.text(idxId));
  delete old.changesToken; delete old.listedAt; delete old.styleFiles;
  drive.edit(idxId, JSON.stringify(old));
  t('the index file in Drive is the older kind, with no readings in it', !JSON.parse(drive.text(idxId)).styleFiles);
  const cold = await boot(drive, coldDb, { conn: 'cellular', phone: true, cfg: { phoneMaxReadsCell: 2 } });
  const cw = cold.w;
  t('the cold phone connected', await until(() => cw.__t.connected, 8000));
  t('it adopted the index but has no work-order readings',
    await until(() => Object.keys(cw.__t.challanIndex).length > 0, 10000) && cw.__t.styleCacheSize === 0,
    cw.__t.styleCacheSize);
  const cdl = drive.downloads.length;
  // The listing runs first; the refusal comes at the end of it, once the size of the job is known.
  await until(() => cw.__t.blocked, 30000);
  t('it refused to pull the backlog over mobile data', !!cw.__t.blocked && cw.__t.blocked.why === 'reads',
    JSON.stringify(cw.__t.blocked));
  t('no work orders were downloaded', drive.downloads.slice(cdl).filter(id => /^W/.test(id)).length === 0,
    drive.downloads.slice(cdl).join(','));
  t('and it says why, on screen', /mobile data/i.test(cold.d.getElementById('askIndexStamp').textContent),
    cold.d.getElementById('askIndexStamp').textContent.slice(0, 90));
  // Wi-Fi arrives.
  cold.conn.type = 'wifi';
  cold.conn.__on.forEach(fn => fn());
  await until(() => drive.downloads.slice(cdl).some(id => /^W/.test(id)), 30000);
  await settle(cw, 40000);
  t('on Wi-Fi it goes ahead and reads them', drive.downloads.slice(cdl).filter(id => /^W/.test(id)).length > 0,
    drive.downloads.slice(cdl).join(','));
  t('and the phone is no longer blocked', !cw.__t.blocked, JSON.stringify(cw.__t.blocked));
  t('and it now holds the readings, so the next update costs nothing', cw.__t.styleCacheSize >= 4, cw.__t.styleCacheSize);

  section('13. Nothing threw anywhere along the way');
  t('no script errors while driving the page', allErrors.length === 0, allErrors.slice(0, 2).join(' | '));

  [laptop, phone, cold].forEach(b => { try { b.dom.window.close(); } catch (e) {} });
  console.log('\n' + (fail ? 'FAILED' : 'PASSED') + ` — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL harness error: ' + (e && e.stack || e)); process.exit(1); });
