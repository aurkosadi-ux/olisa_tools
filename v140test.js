/* v140test.js — automatic indexing must never get stuck behind unrelated Drive files.
   The bug (from v135): once Drive's change feed held more than 12 PDF/Excel/Word files in folders
   outside Olisa ahead of a new work order, every poll re-asked the same 12 questions and put off
   the rest, so the cursor never moved and nothing new was ever indexed again.
   [v140] checks FAIL on v139 and earlier, PASS on v140. Usage: node v140test.js olisa.html */
const fs = require('fs'); const path = require('path');
require('fake-indexeddb/auto');
const { JSDOM, VirtualConsole } = require('jsdom');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x !== undefined ? '  -> ' + x : ''))); };
const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
async function boot(fetch) {
  let html = fs.readFileSync(file, 'utf8');
  html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${fs.readFileSync(path.join(dir, 'lib', 'xlsx.full.min.js'), 'utf8')}</script>`);
  html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');
  const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
  html = html.replace(anchor, () => anchor + `\nwindow.__t = { pollDriveChanges, get autoFeed() { return autoFeed; }, get autoPending() { return autoPending; } };\nwindow.__set = (k, v) => eval(k + ' = v');`);
  const errors = []; const vc = new VirtualConsole(); vc.on('jsdomError', e => errors.push(String(e && (e.detail || e.message || e))));
  const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, url: 'https://aurkosadi-ux.github.io/olisa_tools/olisa.html', virtualConsole: vc,
    beforeParse(w) {
      w.indexedDB = indexedDB; w.IDBKeyRange = IDBKeyRange;
      w.matchMedia = q => ({ matches: false, media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} });
      w.scrollTo = () => {}; w.Element.prototype.scrollIntoView = function () {};
      w.requestAnimationFrame = cb => setTimeout(() => cb(Date.now()), 0); w.cancelAnimationFrame = id => clearTimeout(id);
      w.fetch = fetch; w.navigator.serviceWorker = { register: () => Promise.resolve({}), addEventListener() {}, ready: Promise.resolve({}) };
      w.PDFLib = {}; w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
      w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {}; w.ExcelJS = {};
      w.prompt = () => 'x'; w.confirm = () => true; w.URL.createObjectURL = () => 'blob:t'; w.URL.revokeObjectURL = () => {};
    } });
  await new Promise(r => setTimeout(r, 400));
  return { w: dom.window, T: dom.window.__t, errors };
}
// Drive: MYDRIVE > Olisa Group(ROOT) > Work Orders(PIROOT) > PI-180 > wo180.xlsx ; MYDRIVE > Office > D0..Dn (unrelated)
async function scenario(unrelated, syncedAgoMs, polls, newFolderForWo) {
  const parents = { ROOT: ['MYDRIVE'], PIROOT: ['ROOT'], CHROOT: ['ROOT'], PI180: ['PIROOT'], OFF: ['MYDRIVE'] };
  const changes = []; let reqs = 0;
  for (let k = 0; k < unrelated; k++) { parents['D' + k] = ['OFF']; changes.push({ changeType: 'file', time: '2026-11-12T05:00:00Z', fileId: 'F' + k, file: { id: 'F' + k, name: 'report' + k + '.pdf', mimeType: 'application/pdf', parents: ['D' + k] } }); }
  changes.push({ changeType: 'file', time: '2026-11-12T06:00:00Z', fileId: 'WO180', file: { id: 'WO180', name: 'wo180.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', parents: ['PI180'] } });
  const resp = j => ({ ok: true, status: 200, json: async () => j, clone() { return this; } });
  const fetch = async url => { reqs++; const u = new URL(url);
    if (/\/files\/root$/.test(u.pathname)) return resp({ id: 'MYDRIVE' });
    if (/\/changes\/startPageToken/.test(u.pathname)) return resp({ startPageToken: 'T-NOW' });
    if (/\/changes$/.test(u.pathname)) return resp({ changes, newStartPageToken: 'T-END' });
    const m = u.pathname.match(/\/files\/([^/]+)$/); if (m) return resp({ parents: parents[m[1]] || [] });
    return { ok: false, status: 404, json: async () => ({}), clone() { return this; } }; };
  const { w, T, errors } = await boot(fetch);
  w.__set('driveToken', 'tok'); w.__set('driveTokenExp', Date.now() + 3600e3);
  w.__set('savedRootHandle', { isDrive: true, id: 'ROOT', name: 'Olisa Group', kind: 'directory' });
  w.__set('piRootHandle', { id: 'PIROOT', name: 'Work Orders', kind: 'directory' });
  w.__set('autoFeed', { rootId: 'ROOT', token: 'T-OLD', syncedAt: Date.now() - syncedAgoMs, sharedOk: false });
  w.__set('autoWatch', { rootId: 'ROOT', chRootId: 'CHROOT', folders: {}, files: {} });
  w.__set('runAutoIndex', async () => {});
  let foundAt = 0, how = '';
  for (let p = 1; p <= polls && !foundAt; p++) {
    await T.pollDriveChanges();
    const items = T.autoPending ? Object.values(T.autoPending.items) : [];
    if (items.some(i => i.name === 'wo180.xlsx')) { foundAt = p; how = 'seen directly'; }
    else if (items.some(i => i.kind === 'check')) { foundAt = p; how = 'full check queued'; }
  }
  w.close();
  return { foundAt, how, reqs, errors };
}
(async () => {
  let r = await scenario(5, 10 * 60000, 3);
  t('[same] a new work order behind a few unrelated files is seen on the first poll', r.foundAt === 1 && r.how === 'seen directly', JSON.stringify(r));
  r = await scenario(60, 10 * 60000, 8);
  t('[v140] behind 60 unrelated office files it is still seen, within two polls', r.foundAt >= 1 && r.foundAt <= 2 && r.how === 'seen directly', r.foundAt ? r.foundAt + ' polls, ' + r.how : 'never, after 8 polls');
  r = await scenario(400, 10 * 60000, 8);
  t('[v140] behind 400 it is caught by a full check within five polls (~10 minutes)', r.foundAt >= 1 && r.foundAt <= 5, r.foundAt ? r.foundAt + ' polls, ' + r.how : 'never, after 8 polls');
  r = await scenario(60, 8 * 86400000, 3);
  t('[v140] an index 8 days behind Drive starts a full check on the very first poll', r.foundAt === 1 && r.how === 'full check queued', r.foundAt ? r.foundAt + ' polls, ' + r.how : 'never');
  t('no page errors', !r.errors.length, r.errors.join(' | '));
  const v = (fs.readFileSync(file, 'utf8').match(/const APP_VERSION = (\d+);/) || [])[1];
  const c = (fs.readFileSync(path.join(dir, 'sw.js'), 'utf8').match(/olisa-tools-v(\d+)/) || [])[1];
  t('[v140] page and service worker carry the same version, v140 or later', v === c && +v >= 140, `page ${v}, sw ${c}`);
  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL crashed -> ' + e.stack); process.exit(1); });
