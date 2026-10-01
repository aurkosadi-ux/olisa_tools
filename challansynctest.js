/* challansynctest.js (v137) — the phone never got the laptop's challans; clicking a missing
   challan asked for a rebuild every time; text dates and stray spaces stopped a short settling;
   the index stamp showed two unlabelled clocks; '\u2014' appeared literally. Real page, real DOM.
   Usage: node challansynctest.js olisa.html */
const fs = require('fs');
const path = require('path');
require('fake-indexeddb/auto');
const { IDBFactory } = require('fake-indexeddb');
const { JSDOM, VirtualConsole } = require('jsdom');
const XLSX = require('xlsx');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
let html = fs.readFileSync(file, 'utf8');
const xlsxLib = fs.readFileSync(path.join(dir, 'lib', 'xlsx.full.min.js'), 'utf8');
html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${xlsxLib}</script>`);
const exceljsLib = fs.readFileSync(path.join(dir, 'lib', 'exceljs.min.js'), 'utf8');
html = html.replace(/<script src="\.\/lib\/exceljs\.min\.js"><\/script>/, () => `<script>${exceljsLib}</script>`);
html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');
// Test probe, injected into THIS COPY only, just inside the Ask IIFE so it can see its bindings.
const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
if (!html.includes(anchor)) { console.log('FAIL probe anchor not found'); process.exit(1); }
html = html.replace(anchor, () => anchor + `
window.__t = {
  loadFile, adoptIndexPayload, mergeIncomingChallans, currentIndexPayload, openChallanCopy, renderHealth,
  parseDMY, reconciliationInvariants, updateTextFor, combinedRemarks, normChallanKey, idbGet,
  get allRecords() { return allRecords; },
  get challanIndex() { return challanIndex; }, set challanIndex(v) { challanIndex = v; },
  get challanIndexBuiltAt() { return challanIndexBuiltAt; }, set challanIndexBuiltAt(v) { challanIndexBuiltAt = v; },
  get challanMisses() { return challanMisses; }, set challanMisses(v) { challanMisses = v; },
  set savedRootHandle(v) { savedRootHandle = v; }, set piRootHandle(v) { piRootHandle = v; },
  set autoFeed(v) { autoFeed = v; },
  get piIndex() { return piIndex; }, set piIndex(v) { piIndex = v; },
  get piIndexBuiltAt() { return piIndexBuiltAt; }, set piIndexBuiltAt(v) { piIndexBuiltAt = v; },
  stub(name, fn) { eval(name + ' = fn'); }
};`);

const HEAD = ['PI Ref', 'Inventory Date', 'Style Ref', 'Item', 'Item Description', 'Delivery Qty',
  'Received Qty', 'Short&Exs', 'Challan No', 'Challan Date', 'Remarks'];
function master(rows) {
  const ws = XLSX.utils.aoa_to_sheet([HEAD, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '2026- Main');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
// Two shorts on different PIs, one clean line, one excess on a different item.
const ROWS_V1 = [
  [124, '09/06/2026', 'M82039-1B1-5', 'Master Carton', 'L71.5 x W45 x H32 cm', 100, 99, -1, '31910793', '09/06/2026', ''],
  [130, '12/06/2026', 'M82040-2A1-1', 'Master Carton', 'L60 x W40 x H30 cm', 200, 195, -5, '31910800', '12/06/2026', ''],
  [131, '13/06/2026', 'M82041-3C1-2', 'Master Carton', 'L50 x W40 x H30 cm', 50, 50, 0, '31910801', '13/06/2026', ''],
  [132, '14/06/2026', 'M82042-4D1-3', 'Inner Carton', 'L30 x W20 x H10 cm', 80, 82, 2, '31910802', '14/06/2026', '']
];
// Olisa re-saves the file: the PI 130 line was recounted (195 -> 196 received), so its Short&Exs
// moved from -5 to -4. That changes the overlay key, which is exactly what the repair pass exists for.
const ROWS_V2 = ROWS_V1.map(r => r[0] === 130 ? [...r.slice(0, 6), 196, -4, ...r.slice(8)] : r);

// Each boot is a page reload on the SAME device (same IndexedDB) unless a separate factory is
// passed in — that is how a second device, with storage of its own, is simulated.
const allErrors = [];   // every boot's errors, so a problem in an early part cannot hide
async function boot(idb) {
  const errors = [];
  const vc = new VirtualConsole();
  vc.on('jsdomError', e => { const m = String(e && (e.detail || e.message || e)); errors.push(m); allErrors.push(m); });
  const dom = new JSDOM(html, {
    runScripts: 'dangerously', pretendToBeVisual: true,
    url: 'https://aurkosadi-ux.github.io/olisa_tools/olisa.html', virtualConsole: vc,
    beforeParse(w) {
      w.indexedDB = idb || indexedDB; w.IDBKeyRange = IDBKeyRange;   // one store shared across "reloads"
      w.matchMedia = w.matchMedia || (q => ({ matches: false, media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
      w.scrollTo = () => {}; w.Element.prototype.scrollIntoView = function () {};
      w.requestAnimationFrame = cb => setTimeout(() => cb(Date.now()), 0);
      w.cancelAnimationFrame = id => clearTimeout(id);
      w.fetch = () => Promise.reject(new Error('offline in test'));
      w.navigator.serviceWorker = { register: () => Promise.resolve({ addEventListener() {} }), addEventListener() {}, ready: Promise.resolve({}) };
      w.PDFLib = {};
      w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
      w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {};
      w.CompressionStream = undefined; w.DecompressionStream = undefined;
      w.__prompts = []; w.__confirms = [];
      w.prompt = (m) => { w.__prompts.push(String(m)); return 'test reason'; };
      w.confirm = (m) => { w.__confirms.push(String(m)); return true; };
    }
  });
  const w = dom.window;
  await new Promise(r => setTimeout(r, 400));
  return { w, d: w.document, errors };
}
const tick = (ms = 60) => new Promise(r => setTimeout(r, ms));
const chat = d => (d.getElementById('askChat') || d.body);
async function load(w, rows) {
  await w.__t.loadFile(new w.File([master(rows)], 'Olisa Master Inventory - v1.xlsx'));
  await tick();
}
const T0 = new Date(2026, 9, 1, 1, 21).getTime();
const entry = (no, fp) => ({ no, pages: [1], path: `Challans/${no}.pdf`, date: '27-SEP-26', piRef: '174', fp, lines: [] });
const payload = (o) => Object.assign({ kind: 'olisa-index', version: undefined, builtAt: 0, piIndex: {}, styleIndex: {} }, o);

(async () => {
  console.log('\n1. The laptop\'s challans reach the phone even when the phone\'s PI index is newer');
  let { w, d, errors } = await boot();
  t('the page loads with the probe', !!w.__t, errors.join(' | '));
  const P = w.__t;
  // phone: newer PI index (it auto-indexed at 01:21), 2 challans of its own
  P.piIndex = { '174': { rowCount: 3, items: [] } };
  P.piIndexBuiltAt = T0;
  P.challanIndex = { '117396543': entry('117396543', 'a|1|1'), '117327270': entry('117327270', 'b|1|1') };
  P.challanIndexBuiltAt = T0;
  // a stale "unreadable" note for the file the laptop has since read
  P.challanMisses = { 'm|1|1': { path: 'Challans/hand.doc', why: 'nochallan', at: 1 } };
  const laptop = payload({
    version: P.currentIndexPayload().version, builtAt: T0 - 3600000, challanBuiltAt: T0 - 3600000,
    piIndex: { '170': { rowCount: 1 } },
    challanIndex: { '117396543': entry('117396543', 'a|1|1'), '1173272701': { ...entry('1173272701', 'm|1|1'), pages: null, hand: true } },
    challanMisses: { 'm|1|1': { path: 'Challans/hand.doc', why: 'nochallan', at: 1 } }
  });
  const adopted = P.adoptIndexPayload(laptop, 'test');
  t('the PI part is still refused (the phone\'s own is newer)', adopted === false);
  t('THE BUG: the laptop-only challan is now on the phone', !!P.challanIndex['1173272701'], Object.keys(P.challanIndex).join(','));
  t('the phone\'s own challans are kept, not wiped', !!P.challanIndex['117327270']);
  t('the stale "unreadable" note for that file is dropped', !P.challanMisses['m|1|1'], JSON.stringify(P.challanMisses));
  t('the PI index itself was not replaced', !!P.piIndex['174'] && !P.piIndex['170']);
  await tick(150);
  const saved = await P.idbGet('challanIndex');
  t('the merged list is saved with its version stamp', saved && saved.version === 2 && !!saved.data['1173272701'], JSON.stringify(saved && Object.keys(saved.data)));

  console.log('\n2. Merge rules');
  const before = JSON.stringify(P.challanIndex);
  P.adoptIndexPayload(payload({ version: laptop.version, builtAt: T0 + 9e6, challanIndex: {} }), 'old file');
  t('a file with NO challans wipes nothing', Object.keys(P.challanIndex).length === 3);
  const olderCopy = { ...entry('117327270', 'b|1|1'), date: 'OLD' };
  P.mergeIncomingChallans(payload({ challanBuiltAt: 1, challanIndex: { '117327270': olderCopy } }));
  t('an OLDER read of the same challan does not overwrite the local one', P.challanIndex['117327270'].date === '27-SEP-26');
  const newerCopy = { ...entry('117327270', 'b|2|2'), date: 'NEW' };
  P.mergeIncomingChallans(payload({ challanBuiltAt: T0 + 99e6, challanIndex: { '117327270': newerCopy } }));
  t('a NEWER read of the same challan replaces it', P.challanIndex['117327270'].date === 'NEW');
  t('the payload now carries its own challan build time', typeof P.currentIndexPayload().challanBuiltAt === 'number' && P.currentIndexPayload().challanBuiltAt > 0);
  void before;

  console.log('\n3. Clicking a missing challan finds it by itself');
  ({ w, d, errors } = await boot(new IDBFactory()));
  let Q = w.__t;
  Q.challanIndex = { '117327270': entry('117327270', 'b|1|1') };
  Q.challanIndexBuiltAt = T0;
  Q.savedRootHandle = { isDrive: true, name: 'OLISA GROUP', id: 'root' };
  let forced = null;
  Q.stub('checkForNewerSharedIndex', async (o) => { forced = o; Q.mergeIncomingChallans(payload({ challanBuiltAt: T0 + 1, challanIndex: { '1173272701': { ...entry('1173272701', 'm|1|1'), pages: null } } })); });
  Q.stub('challanFileHandle', async () => { throw new Error('stub: no Drive in test'); });
  await Q.openChallanCopy('1173272701'); await tick(100);
  let txt = chat(d).textContent;
  t('it asked the shared index with force', forced && forced.force === true);
  t('it went straight on to open the challan', /Couldn't open challan|Opening Challan/.test(txt), txt.slice(-200));
  t('no "not in the challan index" message was shown', !/not in the challan index/.test(txt));

  ({ w, d, errors } = await boot(new IDBFactory()));
  Q = w.__t;
  Q.challanIndex = { '117327270': entry('117327270', 'b|1|1') };
  Q.challanIndexBuiltAt = T0;
  Q.savedRootHandle = { isDrive: true, name: 'OLISA GROUP', id: 'root' };
  Q.stub('checkForNewerSharedIndex', async () => {});
  Q.stub('challanTreeDelta', async () => ({ total: 109, unknown: [], retryable: [], unfingerprinted: 0, folder: 'Challans' }));
  await Q.openChallanCopy('1173272701'); await tick(100);
  txt = chat(d).textContent;
  t('when it truly is not anywhere, it says so plainly', /not printed on any challan copy in Drive/.test(txt), txt.slice(-300));
  t('it names the one-digit-away challan', /one digit away is 117327270/.test(txt));
  t('no pointless "Check for new challans" button after a full check', !d.querySelector('.chCheckBtn'));
  t('THE BUG: no literal \\u2014 in the message', !/\\u2014/.test(txt) && /\u2014/.test(txt), txt.slice(-300));

  ({ w, d, errors } = await boot(new IDBFactory()));
  Q = w.__t;
  Q.challanIndex = { '117327270': entry('117327270', 'b|1|1') };
  Q.challanIndexBuiltAt = T0;
  Q.savedRootHandle = { isDrive: true, name: 'OLISA GROUP', id: 'root' };
  let buildOpts = null;
  Q.stub('checkForNewerSharedIndex', async () => {});
  Q.stub('challanTreeDelta', async () => ({ total: 110, unknown: ['new hand.doc'], retryable: [], unfingerprinted: 0, folder: 'Challans' }));
  Q.stub('runChallanIndexBuild', async (r, o) => { buildOpts = o; Q.challanIndex = { ...Q.challanIndex, '1173272701': { ...entry('1173272701', 'n|1|1'), pages: null } }; });
  Q.stub('challanFileHandle', async () => { throw new Error('stub'); });
  await Q.openChallanCopy('1173272701'); await tick(100);
  txt = chat(d).textContent;
  t('a new file in the folder is read automatically', buildOpts && buildOpts.retryMisses === true);
  t('and the challan then opens', /Couldn't open challan|Opening Challan/.test(txt) && !/not printed on any/.test(txt), txt.slice(-200));

  ({ w, d, errors } = await boot(new IDBFactory()));
  Q = w.__t;
  Q.challanIndex = { '117327270': entry('117327270', 'b|1|1') };
  Q.savedRootHandle = { isDrive: false, name: 'OLISA GROUP' };
  Q.stub('challanTreeDelta', async () => ({ total: 109, unknown: [], retryable: [], unfingerprinted: 109, folder: 'Challans' }));
  await Q.openChallanCopy('1173272701'); await tick(100);
  txt = chat(d).textContent;
  t('a local-disk folder is never claimed as "checked"', !/not printed on any challan copy/.test(txt) && !!d.querySelector('.chCheckBtn'), txt.slice(-200));

  ({ w, d, errors } = await boot(new IDBFactory()));
  Q = w.__t;
  Q.challanIndex = { '117327270': entry('117327270', 'b|1|1') };
  await Q.openChallanCopy('1173272701'); await tick(100);
  txt = chat(d).textContent;
  t('with no folder connected it does not try to search, and still shows the reason', /not in the challan index/.test(txt) && !/\\u2014/.test(txt));

  console.log('\n4. The index stamp no longer shows two unlabelled clocks');
  Q.savedRootHandle = { isDrive: true, name: 'OLISA GROUP', id: 'root' };
  Q.piRootHandle = { id: 'wo' };
  Q.piIndex = { '174': { rowCount: 1 } };
  Q.piIndexBuiltAt = Date.now() - 12 * 60000;
  Q.autoFeed = { rootId: 'root', token: 'tok', syncedAt: Date.now() - 2000 };
  Q.renderHealth();
  const hb = d.getElementById('healthBar').textContent;
  t('it leads with "up to date with Drive"', /up to date with Drive \(checked just now; last file added \d\d\/\d\d\/\d{4} \d\d:\d\d\)/.test(hb), hb);
  t('THE BUG: no "(12 minutes ago)" beside it', !/minutes ago\)/.test(hb), hb);
  Q.autoFeed = { rootId: 'root', token: null, syncedAt: 0 };
  Q.renderHealth();
  t('without the Drive watcher, the age is still shown', /\(12 minutes ago\)/.test(d.getElementById('healthBar').textContent), d.getElementById('healthBar').textContent);

  console.log('\n5. Text dates are readable');
  const D = (y, m, dd) => new Date(y, m - 1, dd).getTime();
  t('27-SEP-26', Q.parseDMY('27-SEP-26') === D(2026, 9, 27));
  t('27 - SEP - 26 (as the challans print it)', Q.parseDMY('27 - SEP - 26') === D(2026, 9, 27));
  t('27 Sep 2026', Q.parseDMY('27 Sep 2026') === D(2026, 9, 27));
  t('27-Sept-2026', Q.parseDMY('27-Sept-2026') === D(2026, 9, 27));
  t('27/09/26', Q.parseDMY('27/09/26') === D(2026, 9, 27));
  t('an Excel serial (46292)', Q.parseDMY('46292') === D(2026, 9, 27), String(new Date(Q.parseDMY('46292'))));
  t('existing forms unchanged: 27/09/2026 and 2026-09-27', Q.parseDMY('27/09/2026') === D(2026, 9, 27) && Q.parseDMY('2026-09-27') === D(2026, 9, 27));
  t('rubbish is still 0', Q.parseDMY('L71.5 x W45') === 0 && Q.parseDMY('') === 0 && Q.parseDMY('99/99/26') === 0);

  console.log('\n6. PI 174: the 35-pc short settles by itself');
  ({ w, d, errors } = await boot(new IDBFactory()));
  Q = w.__t;
  const M = 'L60 x W40 x H30 cm';
  await load(w, [
    // original delivery 20/09, Olisa counted it 28/09 \u2014 35 short
    [174, '20/09/2026', 'M82100-1A1-1', 'Master Carton', M, 500, 465, -35, '117327200', '20/09/2026', ''],
    // manual re-delivery, date typed as TEXT the way the challan prints it, no inventory date yet
    [174, '', 'M82100-1A1-1', 'Master Carton', M, 35, 35, 0, '1173272701', '27-SEP-26', 'Manual challan'],
    // original went out 20/09 but Olisa counted it 29/09 — AFTER the manual challan of 27/09
    [174, '29/09/2026', 'M82105-1A1-1', 'Master Carton', M, 300, 288, -12, '117327205', '20/09/2026', ''],
    [174, '27/09/2026', 'M82105-1A1-1', 'Master Carton', M, 12, 12, 0, '117327295', '27/09/2026', 'Manual challan'],
    // style typed with a stray space, settled by its own manual challan
    [174, '20/09/2026', 'M82101-2B1-3', 'Master Carton', M, 100, 90, -10, '117327201', '20/09/2026', ''],
    [174, '29/09/2026', 'M82101-2B1 -3', 'Master Carton', M, 10, 10, 0, '117327299', '29/09/2026', 'Manual challan'],
    // a DIFFERENT style that only shares a prefix must never settle
    [174, '20/09/2026', 'M82102-1A1-5', 'Master Carton', M, 100, 95, -5, '117327202', '20/09/2026', ''],
    [174, '29/09/2026', 'M82102-1A1-50', 'Master Carton', M, 5, 5, 0, '117327298', '29/09/2026', 'Manual challan'],
    // re-delivered on an ordinary, unflagged challan: stays open, and says why
    [174, '20/09/2026', 'M82103-1A1-1', 'Master Carton', M, 100, 80, -20, '117327203', '20/09/2026', ''],
    [174, '30/09/2026', 'M82103-1A1-1', 'Master Carton', M, 20, 20, 0, '117327297', '30/09/2026', ''],
    // manual challan with no Received Qty yet
    [174, '20/09/2026', 'M82104-1A1-1', 'Master Carton', M, 100, 92, -8, '117327204', '20/09/2026', ''],
    [174, '30/09/2026', 'M82104-1A1-1', 'Master Carton', M, 8, null, null, '117327296', '30/09/2026', 'Manual challan']
  ]);
  const R = Q.allRecords;
  const row = (st, ch) => R.find(r => r.style === st && String(r.challanNo) === ch);
  const s35 = row('M82100-1A1-1', '117327200');
  t('THE BUG: the 35 short is settled by manual challan 1173272701', s35 && s35.resolved === true, JSON.stringify(s35 && { res: s35.resolved, notes: s35.autoNotes, why: s35.openWhy }));
  t('its note names the re-delivery date', s35 && /On 27-SEP-26, Full Short Qty is delivered/.test(Q.combinedRemarks(s35)), s35 && Q.combinedRemarks(s35));
  t('a short counted by Olisa after the manual challan\'s date still settles', row('M82105-1A1-1', '117327205').resolved === true);
  t('a style written with a stray space still settles', row('M82101-2B1-3', '117327201').resolved === true);
  t('a different style sharing a prefix does NOT settle', !row('M82102-1A1-5', '117327202').resolved);
  const unf = row('M82103-1A1-1', '117327203');
  t('an unflagged re-delivery is not settled silently', !unf.resolved);
  t('but the line says exactly why', /not marked Manual/.test(unf.openWhy || ''), unf.openWhy);
  const norec = row('M82104-1A1-1', '117327204');
  t('a manual challan with no Received Qty: says so', /no Received Qty filled in/.test(norec.openWhy || ''), norec.openWhy);
  t('the diagnostic never reaches the Olisa-facing Update column', !/Not settled/.test(Q.updateTextFor(unf)) && !/Not settled/.test(Q.combinedRemarks(unf)));
  t('settled lines carry no diagnostic', !s35.openWhy);
  t('the reconciliation ledger is still consistent', Q.reconciliationInvariants().length === 0, Q.reconciliationInvariants().join(' | '));
  t('no load errors in any boot', allErrors.length === 0, allErrors.slice(0, 3).join(' | '));

  console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL crashed: ' + (e && e.stack || e)); process.exit(1); });
