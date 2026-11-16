/* v139test.js — one guard per bug fixed in v139.
   Every check marked [v139] FAILS on v138 and PASSES on v139. Checks marked [same] pass on both:
   they pin behaviour that v139 was NOT allowed to change while fixing the thing next to it.

   Nothing here is a text match against the source. The real page is opened in a real DOM (real
   SheetJS, fake IndexedDB), the real service worker is run against a fake cache and network, the
   real calculator and Pending PI pages are driven the way a person drives them, and the DC Bypass
   workbook code is run against a real .xlsx built with the real JSZip.

   Usage: node v139test.js olisa.html                                                          */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
require('fake-indexeddb/auto');
const { JSDOM, VirtualConsole } = require('jsdom');
const XLSX = require('xlsx');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x !== undefined ? '  -> ' + x : ''))); };
const tick = (ms = 60) => new Promise(r => setTimeout(r, ms));
const until = async (fn, ms = 8000) => { const end = Date.now() + ms; while (Date.now() < end) { if (fn()) return true; await tick(40); } return false; };
const within = (p, ms) => Promise.race([p.then(v => ({ v }), e => ({ e })), new Promise(r => setTimeout(() => r({ timedOut: true }), ms))]);

// An error thrown inside the page that nothing catches must show up as a FAILED CHECK, not kill
// this file half-way through: v138's sign-in bug is exactly that, and without this line the run
// against v138 stopped at section 6 and never reached the other tools.
const stray = [];
process.on('unhandledRejection', e => { stray.push(String(e && e.message || e)); });

const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
const read = name => fs.readFileSync(path.join(dir, name), 'utf8');

/* ------------------------------------------------------------------ the main page, booted for real */
const net = { handler: () => Promise.reject(new Error('offline in test')) };
async function bootOlisa() {
  let html = read(path.basename(file));
  html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${read('lib/xlsx.full.min.js')}</script>`);
  html = html.replace(/<script src="\.\/lib\/exceljs\.min\.js"><\/script>/, () => `<script>${read('lib/exceljs.min.js')}</script>`);
  html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');
  const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
  if (!html.includes(anchor)) { console.log('FAIL probe anchor not found'); process.exit(1); }
  html = html.replace(anchor, () => anchor + `
window.__t = {
  parseDMY, masterVersionOf, buildChallanIndex, buildStyleIndex, driveFetch, requestDriveToken,
  get challanIndex() { return challanIndex; }, get styleIndex() { return styleIndex; },
  set savedRootHandle(v) { savedRootHandle = v; }, set piRootHandle(v) { piRootHandle = v; },
  set styleFileCache(v) { styleFileCache = v; },
  set driveToken(v) { driveToken = v; }, set driveTokenExp(v) { driveTokenExp = v; }
};`);
  const errors = [];
  const vc = new VirtualConsole();
  vc.on('jsdomError', e => errors.push(String(e && (e.detail || e.message || e))));
  const dom = new JSDOM(html, {
    runScripts: 'dangerously', pretendToBeVisual: true,
    url: 'https://aurkosadi-ux.github.io/olisa_tools/olisa.html', virtualConsole: vc,
    beforeParse(w) {
      w.indexedDB = indexedDB; w.IDBKeyRange = IDBKeyRange;
      w.matchMedia = w.matchMedia || (q => ({ matches: false, media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.scrollTo = () => {}; w.Element.prototype.scrollIntoView = function () {};
      w.requestAnimationFrame = cb => setTimeout(() => cb(Date.now()), 0);
      w.cancelAnimationFrame = id => clearTimeout(id);
      w.fetch = (...a) => net.handler(...a);
      w.navigator.serviceWorker = { register: () => Promise.resolve({ addEventListener() {} }), addEventListener() {}, ready: Promise.resolve({}) };
      w.PDFLib = {}; w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
      w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {};
      w.CompressionStream = undefined; w.DecompressionStream = undefined;
      w.prompt = () => 'x'; w.confirm = () => true;
      w.URL.createObjectURL = () => 'blob:test'; w.URL.revokeObjectURL = () => {};
    }
  });
  await tick(400);
  return { w: dom.window, d: dom.window.document, errors, T: dom.window.__t };
}

/* fake folder tree: the same shape the File System Access API and the Drive handles both present */
const dirH = (name, kids, ctl) => ({ kind: 'directory', name, entries: async function* () { if (ctl && ctl.fail) throw new Error('Google Drive error 503'); for (const k of kids) yield [k.name, k]; } });

/* ------------------------------------------------------------------ the service worker, run for real */
function loadSW(swPath, netFn) {
  const stores = new Map(); const listeners = {};
  const BASE = 'https://aurkosadi-ux.github.io/olisa_tools/';
  const abs = req => new URL(typeof req === 'string' ? req : req.url, BASE).href;
  class FakeCache {
    constructor() { this.m = new Map(); }
    async match(req, opt) {
      const k = abs(req);
      if (opt && opt.ignoreSearch) { for (const [u, r] of this.m) if (u.split('?')[0] === k.split('?')[0]) return r; return undefined; }
      return this.m.get(k);
    }
    async put(req, res) { this.m.set(abs(req), res); }
    async add(req) { const r = await sandbox.fetch(req); if (!r || r.status !== 200) throw new Error('bad status'); this.m.set(abs(req), r); }
  }
  const caches = {
    async open(n) { if (!stores.has(n)) stores.set(n, new FakeCache()); return stores.get(n); },
    async keys() { return [...stores.keys()]; }, async delete(n) { return stores.delete(n); },
    async match(req, opt) { for (const c of stores.values()) { const r = await c.match(req, opt); if (r) return r; } return undefined; }
  };
  class Req {
    constructor(u, init) {
      this.url = abs(u); this.method = (init && init.method) || 'GET';
      this.mode = (init && init.mode) || 'no-cors'; this.cache = (init && init.cache) || 'default';
      this.headers = { has: () => false };
    }
  }
  const sandbox = {
    self: { addEventListener: (n, f) => { listeners[n] = f; }, skipWaiting: async () => {}, clients: { claim: async () => {} } },
    caches, URL, Request: Req, Response: { error: () => ({ type: 'error', status: 0, isError: true }) },
    location: { origin: 'https://aurkosadi-ux.github.io' }, console: { warn() {}, log() {} }, Promise,
    fetch: async (req, init) => netFn(typeof req === 'string' ? new Req(req, init) : (init ? new Req(req.url, Object.assign({ mode: req.mode, cache: req.cache }, init)) : req))
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(swPath, 'utf8'), sandbox);
  const fire = async (name, req) => {
    let responded; const waits = [];
    listeners[name]({ request: req, respondWith: p => { responded = Promise.resolve(p); }, waitUntil: p => waits.push(Promise.resolve(p).catch(() => {})) });
    let res; if (responded) res = await responded.catch(e => ({ rejected: e }));
    await Promise.all(waits);
    return { handled: !!responded, res };
  };
  return { fire, Req };
}

/* ------------------------------------------------------------------ DC Bypass: a real workbook */
const TPL_SHEET = (() => {
  const c = (ref, inner, attrs) => '<c r="' + ref + '" s="3"' + (attrs || '') + (inner ? '>' + inner + '</c>' : '/>');
  let rows = '';
  const row = (n, cells) => { rows += '<row r="' + n + '" spans="1:9">' + cells + '</row>'; };
  row(5, c('H5', '<v>1</v>', ' t="s"')); row(6, c('H6', '<v>46000</v>')); row(7, c('H7', '<v>707001</v>'));
  row(9, c('H9', '<v>2</v>', ' t="s"')); row(10, c('H10', '<v>3</v>', ' t="s"')); row(19, c('A19', '<v>4</v>', ' t="s"'));
  row(20, 'ABCDEFGH'.split('').map(L => c(L + '20', '<v>1</v>')).join('') + c('I20', '<f>F20*H20</f><v>1</v>'));
  row(21, c('A21', '<v>5</v>', ' t="s"') + c('F21', '<f>SUM(F20:F20)</f><v>1</v>') + c('I21', '<f>SUM(I20:I20)</f><v>1</v>'));
  row(22, c('F22', '<v>6</v>', ' t="s"') + c('I22', '<f>I21</f><v>1</v>'));
  row(23, c('F23', '<v>7</v>', ' t="s"') + c('I23', '<f>ROUND(I22*122.5,0)</f><v>123</v>'));
  row(26, c('A26', '<v>8</v>', ' t="s"')); row(30, c('A30', '<v>9</v>', ' t="s"'));
  return '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:K49"/><sheetData>' + rows +
    '</sheetData><mergeCells count="3"><mergeCell ref="A21:E21"/><mergeCell ref="A26:I26"/><mergeCell ref="F23:H23"/></mergeCells>' +
    '<rowBreaks count="1" manualBreakCount="1"><brk id="49" max="16383" man="1"/></rowBreaks></worksheet>';
})();
const rowNums = xml => [...xml.matchAll(/<row r="(\d+)"/g)].map(m => +m[1]).join(',');
const cellInner = (xml, ref) => { const m = xml.match(new RegExp('<c r="' + ref + '"(?:[^>]*?/>|[^>]*?>([\\s\\S]*?)</c>)')); return m ? (m[1] || '') : null; };
const mkBill = (n, no) => ({
  billNo: no, sheetName: 'Bill-' + no, billNoText: '121/707001/' + no + '/2026', date: new Date(2026, 10, 15),
  woNumber: 'DCB/12/2026', rate: 122.5, statedBdt: null,
  items: Array.from({ length: n }, (_, i) => ({ type: 'Type-' + 'ABC'[i], L: 500 + i, W: 400, H: 300, ply: 5, qty: 1000 * (i + 1), price: 0.85 + i / 10, desc: 'Master Carton' }))
});

(async () => {
  /* ================================================================== 1. olisa.html */
  const { w, d, errors, T } = await bootOlisa();
  t('the page loads with the probe and no errors', !!T && !errors.length, errors.join(' | '));

  console.log('\n1. Dates typed as text — the forms that still read as "no date"');
  const D = (y, m, day) => new Date(y, m - 1, day).getTime();
  const want = D(2026, 11, 15);
  t('[v139] 2026/11/15 (year first, slashes) is a date', T.parseDMY('2026/11/15') === want, T.parseDMY('2026/11/15'));
  t('[v139] 2026.11.15 (year first, dots) is a date', T.parseDMY('2026.11.15') === want, T.parseDMY('2026.11.15'));
  t('[v139] 2026年11月15日 is a date', T.parseDMY('2026年11月15日') === want, T.parseDMY('2026年11月15日'));
  t('[v139] 20261115 (no separators) is a date', T.parseDMY('20261115') === want, T.parseDMY('20261115'));
  t('[v139] "15th November 2026" is a date', T.parseDMY('15th November 2026') === want, T.parseDMY('15th November 2026'));
  t('[v139] "15-Nov-\'26" is a date', T.parseDMY("15-Nov-'26") === want, T.parseDMY("15-Nov-'26"));
  t('[v139] "Nov 15, 2026" (month first) is a date', T.parseDMY('Nov 15, 2026') === want, T.parseDMY('Nov 15, 2026'));
  t('[v139] "November 15,2026" and "Nov 15th 2026" too', T.parseDMY('November 15,2026') === want && T.parseDMY('Nov 15th 2026') === want, T.parseDMY('November 15,2026') + ' / ' + T.parseDMY('Nov 15th 2026'));
  t('[same] 15/11/2026 still reads day-first', T.parseDMY('15/11/2026') === want);
  t('[same] 2026-11-15 still reads as ISO', T.parseDMY('2026-11-15') === want);
  t('[same] 15-Nov-26 still reads', T.parseDMY('15-Nov-26') === want);
  t('[same] 15/11/26 still reads', T.parseDMY('15/11/26') === want);
  t('[same] an Excel serial still reads', T.parseDMY('46341') === want, T.parseDMY('46341'));
  t('[same] a range keeps its FIRST date, not a year-first misreading', T.parseDMY('5/11/2026-12/11/2026') === D(2026, 11, 5), new Date(T.parseDMY('5/11/2026-12/11/2026')).toDateString());
  t('[same] text that is not a date is still not a date', T.parseDMY('pending') === 0 && T.parseDMY('') === 0 && T.parseDMY('12 pcs short') === 0);

  console.log('\n2. Master File version read from an underscored name');
  t('[v139] "Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx" is version 8', T.masterVersionOf('Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx') === 8, T.masterVersionOf('Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx'));
  t('[v139] "Master_v12 (1).xlsx" is version 12', T.masterVersionOf('Master_v12 (1).xlsx') === 12);
  t('[same] the tool’s own naming still reads', T.masterVersionOf('Olisa Master Inventory — 09-08-2026, 04-28 pm — v9.xlsx') === 9);
  t('[same] a name with no version is still 0', T.masterVersionOf('Olisa Master Inventory.xlsx') === 0 && T.masterVersionOf('Review8.xlsx') === 0);

  console.log('\n3. A challan folder that will not LIST is not an EMPTY folder');
  const File = w.File;
  const rtf = no => `{\\rtf1\\ansi Delivery Challan\\par Challan No: ${no}\\par Challan Date: 05-Nov-2026\\par PI : 700768/2026/000174260815\\par }`;
  const docH = (name, no, ctl) => ({ kind: 'file', name, getFile: async () => { if (ctl && ctl.fail) throw new Error('Lost connection to Google Drive'); return new File([rtf(no)], name, { lastModified: (ctl && ctl.lm) || 1700000000000 }); } });
  const ctlB = { fail: false }, ctlA2 = { fail: false, lm: 1700000000000 };
  const kidsA = [docH('a1.doc', 117000001), docH('a2.doc', 117000002, ctlA2)];
  const challans = dirH('Challans', [dirH('A', kidsA), dirH('B', [docH('b1.doc', 117000003), docH('b2.doc', 117000004)], ctlB)]);
  T.savedRootHandle = dirH('Olisa Group', [challans]);
  await T.buildChallanIndex(null, {});
  const keys = () => Object.keys(T.challanIndex).sort().join(',');
  t('[same] a clean build indexes all four challans', keys() === '117000001,117000002,117000003,117000004', keys());
  ctlB.fail = true;
  const r2 = await T.buildChallanIndex(null, {});
  t('[v139] the challans in the folder that failed to list are still in the index', keys() === '117000001,117000002,117000003,117000004', keys());
  t('[v139] the build says which folder it could not list', Array.isArray(r2.listFailed) && r2.listFailed.join() === 'B/', JSON.stringify(r2.listFailed));
  t('[v139] and how many challans it kept from the last scan', r2.keptFromLast === 2, r2.keptFromLast);
  ctlB.fail = false;
  ctlA2.fail = true; ctlA2.lm = 1700000999000;                    // a2.doc was re-saved, and the download drops
  const r3 = await T.buildChallanIndex(null, {});
  t('[v139] a changed challan whose download failed keeps its last known reading', keys() === '117000001,117000002,117000003,117000004' && r3.keptFromLast === 1, keys() + ' kept ' + r3.keptFromLast);
  ctlA2.fail = false;
  kidsA.shift();                                                  // a1.doc is really deleted from a folder that lists fine
  await T.buildChallanIndex(null, {});
  t('[same] a challan genuinely deleted from a folder that DID list still drops out', keys() === '117000002,117000003,117000004', keys());
  ctlB.fail = true;
  T.savedRootHandle = dirH('Olisa Group', [dirH('Challans', [dirH('B', [docH('b1.doc', 117000003)], ctlB)])]);
  const before = keys();
  const r5 = await T.buildChallanIndex(null, {}).then(() => null, e => e);
  t('[v139] when nothing could be listed the build fails loudly instead of reporting an empty folder', !!r5 && /could not be listed/.test(r5.message), r5 && r5.message);
  t('[same] and the index is left exactly as it was', keys() === before, keys());
  ctlB.fail = false;

  console.log('\n4. A PI folder that will not list, or a work order that will not download');
  const wo = styles => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['STYLE', 'SIZE', 'QTY'], ...styles.map(s => [s, '50x40x30 CM', 100])]), 'S'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); };
  const ctlP = { fail: false }, ctlF = { fail: false };
  const xf = (name, styles, ctl) => ({ kind: 'file', name, getFile: async () => { if (ctl && ctl.fail) throw new Error('Lost connection to Google Drive'); return new File([wo(styles)], name, { lastModified: 1700000000000 }); } });
  const piRoot = dirH('PI', [
    dirH('PI-171', [xf('wo171.xlsx', ['M11111', 'M11112'])]),
    dirH('PI-172', [xf('wo172.xlsx', ['M22221'])], ctlP),
    dirH('PI-173', [xf('wo173.xlsx', ['M33331'], ctlF)])
  ]);
  T.piRootHandle = piRoot;
  T.savedRootHandle = dirH('Olisa Group', [piRoot, challans]);
  await T.buildStyleIndex({});
  const styles = () => Object.keys(T.styleIndex || {}).filter(k => /^M\d{5}$/.test(k)).sort().map(k => k + '>' + T.styleIndex[k].join('/')).join(' ');
  t('[same] a clean build links every style to its PI', styles() === 'M11111>171 M11112>171 M22221>172 M33331>173', styles());
  ctlP.fail = true; ctlF.fail = true; T.styleFileCache = {};      // nothing cached on this device: the files must really be fetched
  const sres = await T.buildStyleIndex({});
  t('[v139] the PI whose folder would not list keeps its styles', /M22221>172/.test(styles()), styles());
  t('[v139] the PI whose work order would not download keeps its styles', /M33331>173/.test(styles()), styles());
  t('[v139] the build reports which PIs it could not reach', !!sres && Array.isArray(sres.unreached) && sres.unreached.slice().sort().join() === '172,173', sres && JSON.stringify(sres.unreached));
  t('[v139] and the progress line says so instead of a bare "Done"', /could not be listed or downloaded/.test(d.getElementById('askProgressText').textContent), d.getElementById('askProgressText').textContent.slice(0, 160));
  ctlP.fail = false; ctlF.fail = false;

  console.log('\n5. Google Drive answering 503 once');
  T.driveToken = 'tok'; T.driveTokenExp = Date.now() + 3600e3;
  let calls = 0;
  net.handler = async () => { calls++; return calls === 1 ? { ok: false, status: 503, clone() { return this; }, json: async () => ({}) } : { ok: true, status: 200, json: async () => ({ files: [{ id: 'x' }] }) }; };
  const df = await within(T.driveFetch('https://www.googleapis.com/drive/v3/files?q=test'), 6000);
  t('[v139] a single 503 is retried, not reported as a failed folder', !!df.v && df.v.files && df.v.files.length === 1 && calls === 2, df.e ? df.e.message : JSON.stringify(df));
  calls = 0;
  net.handler = async () => { calls++; return { ok: false, status: 404, clone() { return this; }, json: async () => ({}) }; };
  const df2 = await within(T.driveFetch('https://www.googleapis.com/drive/v3/files/none'), 6000);
  t('[same] a 404 is still an immediate, honest error', !!df2.e && /404/.test(df2.e.message) && calls === 1, df2.e ? df2.e.message : 'no error');
  net.handler = () => Promise.reject(new Error('offline in test'));

  console.log('\n6. Google sign-in that fails to start');
  w.google = { accounts: { oauth2: { initTokenClient() { throw new Error('gis half-loaded'); } } } };
  const tok = await within(T.requestDriveToken(true), 2500);
  t('[v139] Connect fails with a message instead of waiting forever', !!tok.e && /could not start/i.test(tok.e.message), tok.timedOut ? 'still waiting after 2.5 s' : (tok.e ? tok.e.message : 'resolved?'));
  delete w.google;

  console.log('\n7. PO Summary: sizes with nothing ordered');
  const poBook = run => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['STYLE', 'SIZE', 'QTY'], ['M12345', 'CARTON 50*40*30 CM (6-7#)', 150]]), 'Order');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['COLOR', 6, 6.5, 7, 'STYLE'], ['RED', ...run, 'M12345']]), 'Run');
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  };
  const poRun = async (name, run) => {
    const ev = new w.Event('drop', { bubbles: true, cancelable: true });
    ev.dataTransfer = { files: [new File([poBook(run)], name, { lastModified: Date.now() })] };
    d.getElementById('podrop').dispatchEvent(ev);
    d.getElementById('pogenBtn').click();
    const st = d.getElementById('postatus');
    await until(() => /^Done/.test(st.textContent) || /err/.test(st.className));
    return [...d.querySelectorAll('#popreviewTable tr')].slice(1).map(tr => { const c = [...tr.children].map(td => td.textContent); return c[1].replace(/^.*\(/, '(') + '=' + c[c.length - 1]; });
  };
  const po1 = await poRun('po-a.xlsx', [100, 0, 50]);
  t('[v139] a size with zero pieces is not a work-order line', po1.join(' ') === '(6#)=100 (7#)=50', po1.join(' '));
  const po2 = await poRun('po-b.xlsx', ['x', '', 'n/a']);
  t('[v139] an unreadable size run falls back to the row’s own quantity', po2.join(' ') === '(6-7#)=150', po2.join(' '));
  const po3 = await poRun('po-c.xlsx', [100, 20, 30]);
  t('[same] a full size run still expands into one line per size', po3.join(' ') === '(6#)=100 (6.5#)=20 (7#)=30', po3.join(' '));
  t('no page errors after all of that', !errors.length, errors.join(' | ').slice(0, 300));
  w.close();

  /* ================================================================== 2. sw.js */
  console.log('\n8. Service worker: what is stored, and what comes back offline');
  {
    const log = []; let online = true;
    const netFn = async req => {
      log.push({ url: req.url, mode: req.mode, cache: req.cache });
      if (!online) throw new Error('offline');
      const cross = !req.url.startsWith('https://aurkosadi-ux.github.io');
      if (cross && req.mode === 'no-cors') return { status: 0, type: 'opaque', clone() { return this; } };
      return { status: 200, type: cross ? 'cors' : 'basic', body: 'BODY ' + req.url, clone() { return this; } };
    };
    const sw = loadSW(path.join(dir, 'sw.js'), netFn);
    await sw.fire('install', {});
    const shell = log.filter(l => /\/olisa\.html$/.test(l.url))[0];
    t('[v139] the install fetches the shell past the browser cache (cache: reload)', !!shell && shell.cache === 'reload', shell && shell.cache);
    t('[same] the install still stores every page', ['olisa.html', 'calculator.html', 'DC_Bypass_Bill.html', 'Pending_PI_Summary.html', 'index.html'].every(p => log.some(l => l.url.endsWith('/' + p))));
    const css = 'https://fonts.googleapis.com/css2?family=Outfit:wght@400;600&display=swap';
    const f1 = await sw.fire('fetch', new sw.Req(css, { mode: 'no-cors' }));
    t('[v139] the font stylesheet is handled by the worker at all', f1.handled && f1.res && f1.res.status === 200, f1.handled ? 'status ' + (f1.res && f1.res.status) : 'passed straight to the network');
    const api = await sw.fire('fetch', new sw.Req('https://www.googleapis.com/drive/v3/files?q=x', { mode: 'cors' }));
    t('[same] Google Drive API calls are never touched by the worker', api.handled === false);
    const gis = await sw.fire('fetch', new sw.Req('https://accounts.google.com/gsi/client', { mode: 'no-cors' }));
    t('[same] nor is Google sign-in', gis.handled === false);
    online = false;
    const f2 = await sw.fire('fetch', new sw.Req(css, { mode: 'no-cors' }));
    t('[v139] offline, the font stylesheet comes back from the cache', f2.handled && f2.res && f2.res.status === 200, f2.handled ? JSON.stringify(f2.res && (f2.res.rejected ? String(f2.res.rejected.message) : f2.res.status)) : 'not handled');
    const icon = await sw.fire('fetch', new sw.Req('https://aurkosadi-ux.github.io/olisa_tools/not-cached.png', { mode: 'no-cors' }));
    t('[v139] offline, a missing image is an error — not the landing page’s HTML', !!icon.res && icon.res.isError === true, icon.res && String(icon.res.body).slice(0, 60));
    const nav = await sw.fire('fetch', new sw.Req('https://aurkosadi-ux.github.io/olisa_tools/olisa.html?source=pwa', { mode: 'navigate' }));
    t('[v139] offline, olisa.html?x=y opens olisa.html, not the tool menu', !!nav.res && /olisa\.html$/.test(String(nav.res.body)), nav.res && String(nav.res.body).slice(0, 80));
    const nav2 = await sw.fire('fetch', new sw.Req('https://aurkosadi-ux.github.io/olisa_tools/never-seen.html', { mode: 'navigate' }));
    t('[same] offline, an unknown page still falls back to the tool menu', !!nav2.res && /index\.html$/.test(String(nav2.res.body)), nav2.res && String(nav2.res.body).slice(0, 80));
    const page = await sw.fire('fetch', new sw.Req('https://aurkosadi-ux.github.io/olisa_tools/calculator.html', { mode: 'navigate' }));
    t('[same] offline, a cached page still opens', !!page.res && /calculator\.html$/.test(String(page.res.body)));
  }
  for (const p of ['index.html', 'calculator.html', 'DC_Bypass_Bill.html', 'Pending_PI_Summary.html']) {
    t(`[v139] ${p} registers the service worker itself`, /navigator\.serviceWorker\.register\('sw\.js'\)/.test(read(p)));
  }

  /* ================================================================== 3. calculator.html */
  console.log('\n9. Calculator: every copy says what happened, where it can be seen');
  {
    const dom = new JSDOM(read('calculator.html'), { runScripts: 'dangerously', pretendToBeVisual: true, url: 'https://aurkosadi-ux.github.io/olisa_tools/calculator.html' });
    const cw = dom.window, cd = cw.document; await tick(120);
    const App = cw.eval('App');
    let copyWorks = true; cd.execCommand = () => copyWorks;
    const hiddenBy = el => { for (let e = el; e && e !== cd.body; e = e.parentElement) if (e.classList.contains('hidden')) return e.id || e.tagName; return ''; };
    const shown = el => !!el && !hiddenBy(el) && el.style.opacity === '1';
    const msgEl = () => cd.getElementById(App.viewMode === 'multiple' ? 'multiCopyMessage' : 'copyMessage');

    const price = () => cd.getElementById('price_output').textContent;
    t('[v139] nothing typed: the price is $0.0000, not the wastage priced on its own', price() === '$0.0000', price());
    cd.getElementById('lengthInput').value = '50'; App.calculatePrice();
    t('[v139] only the length typed: still no price', price() === '$0.0000', price());
    cd.getElementById('widthInput').value = '40'; cd.getElementById('heightInput').value = '30'; App.calculatePrice();
    t('[same] fully measured: 2 x (50+40+6) x (40+30+4) cm\u00b2 at $0.80 is $1.1366', price() === '$1.1366', price());
    cd.getElementById('lengthInput').value = ''; cd.getElementById('widthInput').value = ''; cd.getElementById('heightInput').value = ''; App.calculatePrice();

    App.switchView();
    App.copyMultipleQuotes();                                       // the one default row is empty
    let m = cd.getElementById('multiCopyMessage');
    t('[v139] Multiple view: "Nothing copied" is on a line that is actually visible', shown(m) && /Nothing copied/.test(m.textContent), m ? `"${m.textContent}" hidden by ${hiddenBy(m) || 'nothing'}` : 'no message line exists in this view');
    App.addRow(); App.addRow(); App.removeRow(1);
    const heads = [...cd.querySelectorAll('#multipleRowsContainer h4')].map(h => h.textContent.trim()).join(', ');
    t('[v139] after removing the middle row the headings read 1, 2 — the numbers the quote uses', heads === 'Item #1, Item #2', heads);
    cd.getElementById('multiL-0').value = '50'; cd.getElementById('multiW-0').value = '40'; cd.getElementById('multiH-0').value = '30';
    App.copyMultipleQuotes();
    t('[v139] the refusal names the item by the number on its heading', /item 2 \(/.test(m ? m.textContent : '') && heads.includes('Item #2'), m && m.textContent);
    cd.getElementById('multiL-2').value = '60'; cd.getElementById('multiW-2').value = '45'; cd.getElementById('multiH-2').value = '35';
    copyWorks = false; App.copyMultipleQuotes(); await tick(80);
    t('[v139] Multiple view: a copy that failed says so', shown(m) && /Copy failed/.test(m.textContent), m && m.textContent);
    copyWorks = true; App.copyMultipleQuotes(); await tick(80);
    t('[same] Multiple view: a copy that worked is confirmed', /Multiple Quotes Copied!/.test(cd.getElementById('calculator-app').textContent));

    App.switchView();                                               // back to the single price
    const sm = cd.getElementById('copyMessage');
    App.copySingleQuote(); await tick(60);                          // nothing typed yet
    t('[same] Single view: an unmeasured carton is refused', /Nothing copied/.test(sm.textContent) && shown(sm), sm.textContent);
    cd.getElementById('lengthInput').value = '50'; cd.getElementById('widthInput').value = '40'; cd.getElementById('heightInput').value = '30'; App.calculatePrice();
    copyWorks = false; App.copySingleQuote(); await tick(80);
    t('[v139] Single view: "Quote Copied!" is not shown when nothing was copied', !/Quote Copied!/.test(sm.textContent) && /Copy failed/.test(sm.textContent), sm.textContent);
    await tick(4700);                                               // let the red message fade the way it does on the page
    copyWorks = true; App.copySingleQuote(); await tick(80);
    t('[v139] after a refusal has faded, the next "Quote Copied!" is still visible', /Quote Copied!/.test(sm.textContent) && sm.style.opacity !== '0', `text "${sm.textContent}", inline opacity "${sm.style.opacity}"`);
    cw.close();
  }

  /* ================================================================== 4. DC_Bypass_Bill.html */
  console.log('\n10. DC Bypass Bill: the taka total, the rate and the price');
  {
    const dc = read('DC_Bypass_Bill.html');
    const grab = id => (dc.match(new RegExp('<script id="' + id + '">([\\s\\S]*?)</script>')) || [])[1];
    const JSZip = require(path.join(dir, 'lib', 'jszip.min.js'));
    const { CORE, PKG } = new Function('JSZip', grab('core') + '\n' + grab('pkg') + '\nreturn { CORE: CORE, PKG: PKG };')(JSZip);
    const exact = (cents, rate4) => { const p = BigInt(cents) * BigInt(rate4), den = 1000000n; return Number(p / den + ((p % den) * 2n >= den ? 1n : 0n)); };
    let bad = 0, n = 0, first = '';
    for (let c = 10000; c < 3000000; c += 977) for (const rate of [117.5, 119.75, 121.25, 122.5, 122.75, 123.45, 121.87]) {
      n++; const got = CORE.roundBdt(c / 100, rate), ref = exact(c, Math.round(rate * 10000));
      if (got !== ref) { bad++; if (!first) first = `${c / 100} x ${rate}: tool ${got}, Excel ${ref}`; }
    }
    t(`[v139] BDT totals match Excel’s ROUND() on all ${n} sampled bills`, bad === 0, `${bad} differ, e.g. ${first}`);
    t('[v139] 1,135.62 x 119.75 is 135,990 (it is 135,990.495 — not a half)', CORE.roundBdt(1135.62, 119.75) === 135990, CORE.roundBdt(1135.62, 119.75));
    t('[same] a true half still rounds away from zero: 1,153.00 x 122.5 = 141,242.5 -> 141,243', CORE.roundBdt(1153, 122.5) === 141243, CORE.roundBdt(1153, 122.5));
    t('[v139] a rate of 117.125 is printed, and put in the formula, as 117.125', CORE.rateText(117.125) === '117.125', CORE.rateText(117.125));
    t('[same] ordinary rates print as before', CORE.rateText(122) === '122' && CORE.rateText(122.5) === '122.5' && CORE.rateText(122.85) === '122.85' && CORE.rateText(122.05) === '122.05');
    t('[v139] a price of 0.8525 is shown in full, not as 0.853', typeof CORE.priceText === 'function' && CORE.priceText(0.8525) === '0.8525', typeof CORE.priceText === 'function' ? CORE.priceText(0.8525) : 'always toFixed(3)');
    t('[same] a three-decimal price looks as it always did', typeof CORE.priceText !== 'function' || (CORE.priceText(0.85) === '0.850' && CORE.priceText(1.125) === '1.125'));

    console.log('\n11. DC Bypass Bill: copying a sheet that already has two carton lines');
    const clean1 = CORE.fillSheet(TPL_SHEET, mkBill(1, 30));
    const two = CORE.fillSheet(TPL_SHEET, mkBill(2, 31));
    t('[same] one line from the one-line original', rowNums(clean1) === '5,6,7,9,10,19,20,21,22,23,26,30');
    t('[same] two lines from the one-line original', rowNums(two) === '5,6,7,9,10,19,20,21,22,23,24,27,31' && cellInner(two, 'F22').includes('SUM(F20:F21)'));
    t('[v139] a manual page break moves down with the rows (id), and its column span (max) is left alone', /<brk id="50" max="16383"/.test(two), (two.match(/<brk [^>]*>/) || [''])[0]);
    let again = null, err = '';
    try { again = CORE.fillSheet(two, mkBill(1, 32)); } catch (e) { err = e.message; }
    t('[v139] a one-line bill copied FROM the two-line sheet is identical to a clean one', !!again && again === CORE.fillSheet(TPL_SHEET, mkBill(1, 32)), err || 'sheets differ: rows ' + (again && rowNums(again)));
    let again2 = null; err = '';
    try { again2 = CORE.fillSheet(two, mkBill(3, 33)); } catch (e) { err = e.message; }
    t('[v139] a three-line bill copied from the two-line sheet is identical to a clean one', !!again2 && again2 === CORE.fillSheet(TPL_SHEET, mkBill(3, 33)), err || 'sheets differ');

    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Bill-29" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'Bill-29\'!$A$1:$K$49</definedName></definedNames></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', TPL_SHEET);
    const book1 = await zip.generateAsync({ type: 'arraybuffer' });
    const out1 = await (await PKG.buildWorkbook(book1, [mkBill(2, 30)])).generateAsync({ type: 'arraybuffer' });   // this month: a two-line bill
    let out2 = null; err = '';
    try { out2 = await PKG.buildWorkbook(out1, [mkBill(1, 31)]); } catch (e) { err = e.message; }                  // next month: starts from that workbook
    t('[v139] next month’s workbook builds at all when last month’s newest bill had two lines', !!out2, err);
    if (out2) {
      const wbXml = await out2.file('xl/workbook.xml').async('string');
      const area = n => (wbXml.match(new RegExp("'Bill-" + n + "'!([^<]+)<")) || [])[1];
      t('[same] the two-line bill prints one row taller', area(30) === '$A$1:$K$50', area(30));
      t('[v139] the one-line bill copied from it goes back to the original print area', area(31) === '$A$1:$K$49', area(31));
      const b31 = await out2.file('xl/worksheets/sheet3.xml').async('string');
      t('[v139] and its TOTAL sits under its single carton line', rowNums(b31) === '5,6,7,9,10,19,20,21,22,23,26,30' && cellInner(b31, 'F21').includes('SUM(F20:F20)'), rowNums(b31));
    }
  }

  /* ================================================================== 5. Pending_PI_Summary.html */
  console.log('\n12. Pending PI Summary: two invoices that share a file name');
  {
    const starts = [];
    const mkPdf = (name, rel, piNo, value, delay) => ({
      name, type: 'application/pdf', webkitRelativePath: rel || '',
      arrayBuffer: async () => { starts.push(piNo); if (delay) await tick(delay); return new TextEncoder().encode(['PROFORMA INVOICE', 'PI No: ' + piNo + ' Date: 14-05-2026', 'Buyer: Gildan', 'TOTAL ' + value, 'Concern Person: Rifat'].join('\n')).buffer; }
    });
    const dom = new JSDOM(read('Pending_PI_Summary.html'), {
      runScripts: 'dangerously', pretendToBeVisual: true, url: 'https://example.com/Pending_PI_Summary.html',
      beforeParse(pw) {
        pw.fetch = () => Promise.resolve({ ok: true });
        pw.matchMedia = pw.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {} }));
        pw.pdfjsLib = {
          GlobalWorkerOptions: {},
          getDocument: ({ data }) => ({ promise: Promise.resolve({
            numPages: 1, destroy: async () => {},
            getPage: async () => ({
              cleanup() {}, getViewport: () => ({ width: 595 }),
              getTextContent: async () => ({ items: new TextDecoder().decode(data).split('\n').map((s, i) => ({ str: s, transform: [1, 0, 0, 10, 40, 800 - i * 14], width: s.length * 5, hasEOL: true })) })
            })
          }) })
        };
      }
    });
    const pw = dom.window; await new Promise(r => pw.addEventListener('load', r)); await tick(80);
    const rowsNow = () => pw.eval('state.rows.map(function(r){return r.pi+"="+r.value;}).join(" | ")');
    await pw.ingest([mkPdf('PI.pdf', 'Parties/Alpha/PI.pdf', 'PI-1001', '1,000.00'), mkPdf('PI.pdf', 'Parties/Beta/PI.pdf', 'PI-2002', '2,500.00')]);
    t('[v139] two different invoices both called PI.pdf, in two folders, are two rows', pw.eval('state.rows.length') === 2 && /PI-1001/.test(rowsNow()) && /PI-2002/.test(rowsNow()), rowsNow());
    const afterFirstPick = pw.eval('state.rows.length');
    await pw.ingest([mkPdf('PI.pdf', 'Parties/Alpha/PI.pdf', 'PI-1001', '1,250.00'), mkPdf('PI.pdf', 'Parties/Beta/PI.pdf', 'PI-2002', '2,500.00')]);
    t('[same] picking the same folder again re-reads its rows instead of adding more', pw.eval('state.rows.length') === afterFirstPick, rowsNow());
    t('[v139] and each of the two keeps its own re-read value', /PI-1001=1250/.test(rowsNow()) && /PI-2002=2500/.test(rowsNow()), rowsNow());
    await pw.ingest([mkPdf('other.pdf', '', 'PI-3003', '300.00')]);
    await pw.ingest([mkPdf('other.pdf', '', 'PI-3003', '333.00')]);
    t('[same] a loose file dropped twice still replaces itself', pw.eval('state.rows.length') === afterFirstPick + 1 && /PI-3003=333/.test(rowsNow()), rowsNow());

    console.log('\n13. Pending PI Summary: a second drop while the first is still being read');
    starts.length = 0;
    const rowsBefore = pw.eval('state.rows.length');
    const p1 = pw.ingest([mkPdf('s1.pdf', '', 'A1', '10.00', 120), mkPdf('s2.pdf', '', 'A2', '20.00', 120)]);
    const p2 = pw.ingest([mkPdf('s3.pdf', '', 'B1', '30.00', 10)]);
    await p1; const busyBetween = pw.eval('busy');
    await p2; await tick(30);
    t('[v139] the second batch waits for the first: files are read in the order they were dropped', starts.join(',') === 'A1,A2,B1', starts.join(','));
    t('[same] all three invoices reach the table and nothing is left "busy"', pw.eval('state.rows.length') === rowsBefore + 3 && pw.eval('busy') === false, pw.eval('state.rows.length') + ' rows, busy=' + pw.eval('busy') + ', between=' + busyBetween);
    pw.close();
  }

  /* ================================================================== 6. one version everywhere */
  console.log('\n14. One version everywhere');
  const v = (read(path.basename(file)).match(/const APP_VERSION = (\d+);/) || [])[1];
  const c = (read('sw.js').match(/olisa-tools-v(\d+)/) || [])[1];
  t('[v139] page and service worker carry the same version, v139 or later', v === c && Number(v) >= 139, `page ${v}, sw ${c}`);

  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL the test itself crashed  -> ' + (e && e.stack || e)); process.exit(1); });
