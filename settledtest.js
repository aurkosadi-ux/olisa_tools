/* settledtest.js — "I settled the last short by mistake and now I can't undo it."
   Runs the REAL page in a real DOM (with the real SheetJS and a real IndexedDB), loads a small
   Master File, and drives the shortage list the way a person does: Mark accepted, then Undo.

   It guards three separate ways an accidental settle could become permanent:
     1. The settled list vanished once nothing was left outstanding (the list returned early).
     2. Undo was reverted by the shared index file on the next connect (the merge only ever added).
     3. Undo was reverted by the overlay repair pass after Olisa re-saved the Master File
        (the repaired acceptance lived on under its old key and re-attached itself).
   Usage: node settledtest.js olisa.html                                                        */
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
  loadFile, showShortageList, mergeOverlays, currentIndexPayload, repairOverlayKeys, acceptKey,
  get acceptedShorts() { return acceptedShorts; }, get allRecords() { return allRecords; },
  effectiveShort, setRemark, effectiveRemark, buildUpdatedMasterBlob, idbSet, setAccepted,
  reconcile: reconcileManualChallans,
  get remarkOverrides() { return remarkOverrides; },
  setMasterBytes(b) { masterFileBytes = b; }
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
const chatText = d => (d.getElementById('askChat') || d.body).textContent;
const outstandingRows = d => [...d.querySelectorAll('.acc-btn')];
const undoBtns = d => [...d.querySelectorAll('.unacc-btn')];
const acceptedCount = w => w.__t.allRecords.filter(r => r.accepted).length;
async function load(w, rows, name) {
  await w.__t.loadFile(new w.File([master(rows)], name || 'Olisa Master Inventory - v1.xlsx'));
  await tick();
}
function setFilter(d, from, to) {
  d.getElementById('askPiFrom').value = from == null ? '' : String(from);
  d.getElementById('askPiTo').value = to == null ? '' : String(to);
}
async function openList(d) { d.getElementById('askShortageBtn').click(); await tick(); }

(async () => {
  console.log('\n1. Settling the LAST outstanding short must not hide the settled list');
  let { w, d, errors } = await boot();
  t('the page loads with the probe', !!w.__t, errors.join(' | '));
  await load(w, ROWS_V1);
  t('the Master File loads', w.__t.allRecords.length === 4, 'records: ' + (w.__t.allRecords || []).length);
  await openList(d);
  t('two outstanding shorts are listed', outstandingRows(d).length === 2, 'rows: ' + outstandingRows(d).length);
  outstandingRows(d)[0].click(); await tick();
  t('after settling one, one is still outstanding', outstandingRows(d).length === 1);
  t('and the settled section shows one Undo', undoBtns(d).length === 1);
  outstandingRows(d)[0].click(); await tick();
  t('after settling the last one, nothing is outstanding', outstandingRows(d).length === 0);
  t('THE BUG: the settled lines are still listed', undoBtns(d).length === 2,
    'undo buttons on screen: ' + undoBtns(d).length + ' — text: ' + chatText(d).slice(0, 160));
  t('the message does not claim every short was re-delivered',
    !/every short has been re-delivered\./i.test(chatText(d)) && /re-delivered or settled by agreement/.test(chatText(d)),
    chatText(d).slice(0, 160));
  if (undoBtns(d).length) {
    undoBtns(d)[0].click(); await tick();
    t('Undo puts the line back on the shortage list', outstandingRows(d).length === 1);
    t('and it is no longer settled', acceptedCount(w) === 1);
  }

  console.log('\n2. Same thing with a PI filter active');
  ({ w, d, errors } = await boot());   // storage persists: PI 130 or 124 is still settled from part 1
  await load(w, ROWS_V1);
  // Clear whatever part 1 left, so this part starts from a known state.
  Object.keys(w.__t.acceptedShorts).forEach(k => delete w.__t.acceptedShorts[k]);
  await load(w, ROWS_V1);
  setFilter(d, 130, 130);
  await openList(d);
  t('the filter shows only the PI 130 short', outstandingRows(d).length === 1);
  outstandingRows(d)[0].click(); await tick();
  t('after settling the only short in the filter, its Undo is still reachable', undoBtns(d).length === 1,
    'text: ' + chatText(d).slice(0, 200));
  setFilter(d, 124, 124);
  await openList(d);
  t('a filter that hides the settled line says it is there',
    /1 settled line\(s\) sit outside this filter/.test(chatText(d)), chatText(d).slice(0, 260));
  setFilter(d, 120, 135);
  await openList(d);
  t('the settled line is listed newest-first with its Undo when the filter includes it',
    undoBtns(d).length === 1 && undoBtns(d)[0].dataset.k === w.__t.acceptKey(w.__t.allRecords.find(r => r.piRef === 130)));
  setFilter(d, null, null);

  console.log('\n3. Undo must survive the shared index file on the next connect');
  ({ w, d, errors } = await boot());
  await load(w, ROWS_V1);
  Object.keys(w.__t.acceptedShorts).forEach(k => delete w.__t.acceptedShorts[k]);
  await load(w, ROWS_V1);
  await openList(d);
  const target = w.__t.allRecords.find(r => r.piRef === 130);
  outstandingRows(d).find(b => b.dataset.k === w.__t.acceptKey(target)).click(); await tick();
  t('the PI 130 short is settled', !!w.__t.allRecords.find(r => r.piRef === 130).accepted);
  // A build runs now and uploads the index — the file in Drive carries the acceptance.
  const driveCopy = JSON.parse(JSON.stringify(w.__t.currentIndexPayload()));
  await tick(5);
  undoBtns(d)[0].click(); await tick();
  t('Undo clears it', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  // Next connect: the stale index file in Drive is merged in.
  w.__t.mergeOverlays(driveCopy);
  t('THE BUG: the stale Drive copy does not re-settle the line', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  // And the other direction still works: a genuinely NEWER settle from another device wins.
  const newer = JSON.parse(JSON.stringify(driveCopy));
  Object.values(newer.acceptedShorts).forEach(v => { v.at = Date.now() + 60000; delete v.cleared; });
  w.__t.mergeOverlays(newer);
  t('a newer settle from another device is still picked up', !!w.__t.allRecords.find(r => r.piRef === 130).accepted);
  // An undo that travels in the index file reaches the other device too.
  await openList(d);
  undoBtns(d)[0].click(); await tick();
  const withUndo = JSON.parse(JSON.stringify(w.__t.currentIndexPayload()));
  t('the index file now carries the undo', !!(withUndo.overlayUndos && Object.keys(withUndo.overlayUndos.acc || {}).length));
  t('and not the undone settlement itself', !Object.keys(withUndo.acceptedShorts || {}).length);
  // The OTHER device: its own storage, so it has never seen this undo.
  const otherDeviceIdb = new IDBFactory();
  ({ w, d, errors } = await boot(otherDeviceIdb));
  await load(w, ROWS_V1);
  // this "other device" still believes the line is settled (older timestamp)
  const other = w.__t.acceptedShorts;
  Object.keys(other).forEach(k => delete other[k]);
  Object.entries(newer.acceptedShorts).forEach(([k, v]) => { other[k] = { ...v, at: v.at - 1000 }; });
  await load(w, ROWS_V1);
  t('(setup) the other device shows it settled', !!w.__t.allRecords.find(r => r.piRef === 130).accepted);
  w.__t.mergeOverlays(withUndo);
  t('the undo made on this device reaches the other one', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  t('merging the same file again reports no change (no false "update picked up")', w.__t.mergeOverlays(withUndo) === 0);
  t('nor does the old Drive copy that still carries the settlement', w.__t.mergeOverlays(driveCopy) === 0 &&
    !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  t('a settle made after the undo on this device is kept', (() => {
    const r = w.__t.allRecords.find(x => x.piRef === 130);
    w.__t.setAccepted(r, 'accepted again'); w.__t.repairOverlayKeys(); w.__t.reconcile();
    return !!w.__t.allRecords.find(x => x.piRef === 130).accepted;
  })());

  console.log('\n4. Undo must survive Olisa re-saving the Master File');
  ({ w, d, errors } = await boot());
  Object.keys(w.__t.acceptedShorts).forEach(k => delete w.__t.acceptedShorts[k]);
  await load(w, ROWS_V1, 'Olisa Master Inventory - v1.xlsx');
  await openList(d);
  const t130 = w.__t.allRecords.find(r => r.piRef === 130);
  outstandingRows(d).find(b => b.dataset.k === w.__t.acceptKey(t130)).click(); await tick();
  t('settled on v1 of the file', !!w.__t.allRecords.find(r => r.piRef === 130).accepted);
  await load(w, ROWS_V2, 'Olisa Master Inventory - v2.xlsx');
  t('still settled after Olisa recounted the line (repair re-linked it)', !!w.__t.allRecords.find(r => r.piRef === 130).accepted);
  await openList(d);
  undoBtns(d)[0].click(); await tick();
  t('Undo on v2 clears it', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  await load(w, ROWS_V2, 'Olisa Master Inventory - v2.xlsx');
  t('THE BUG: reloading the same file does not bring the settle back', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  await load(w, ROWS_V1, 'Olisa Master Inventory - v1.xlsx');
  t('nor does going back to the older file version', !w.__t.allRecords.find(r => r.piRef === 130).accepted);

  console.log('\n5. Undo survives a real reload (IndexedDB)');
  await tick(200);
  ({ w, d, errors } = await boot());
  await load(w, ROWS_V2, 'Olisa Master Inventory - v2.xlsx');
  await tick(200);
  t('after a reload the line is still on the shortage list', !w.__t.allRecords.find(r => r.piRef === 130).accepted);
  await openList(d);
  t('and the list shows it as outstanding', outstandingRows(d).some(b => b.dataset.k === w.__t.acceptKey(w.__t.allRecords.find(r => r.piRef === 130))));

  console.log('\n7. Removing a remark edit takes effect at once, and stays removed');
  ({ w, d, errors } = await boot());
  const ROWS_RMK = ROWS_V1.map(r => r[0] === 124 ? [...r.slice(0, 10), 'hole mark'] : r);
  await load(w, ROWS_RMK);
  let r124 = () => w.__t.allRecords.find(r => r.piRef === 124);
  t('(setup) Olisa\'s own remark is read', w.__t.effectiveRemark(r124()) === 'hole mark');
  w.__t.setRemark(r124(), 'ok done');
  t('an edit replaces it and colours the row green', r124().remarks === 'ok done' && r124().forceColor === 'green');
  const rmkDrive = JSON.parse(JSON.stringify(w.__t.currentIndexPayload()));   // index uploaded now
  await tick(5);
  w.__t.setRemark(r124(), '');
  t('clearing the edit shows Olisa\'s remark again immediately (no reload)', r124().remarks === 'hole mark',
    'remarks now: ' + JSON.stringify(r124().remarks));
  t('and the green colour from the removed edit is gone', r124().forceColor !== 'green');
  t('the edit marker is gone', !r124().remarkEdited);
  w.__t.mergeOverlays(rmkDrive);
  t('the stale Drive copy does not bring the removed edit back', w.__t.effectiveRemark(r124()) === 'hole mark');
  await load(w, ROWS_RMK);
  t('nor does reloading the Master File', w.__t.effectiveRemark(r124()) === 'hole mark' && r124().remarks === 'hole mark');
  w.__t.setRemark(r124(), 'recounted, fine');
  t('a new edit after that works normally', w.__t.effectiveRemark(r124()) === 'recounted, fine');
  t('and records Olisa\'s wording as the original, not the previous edit',
    w.__t.remarkOverrides[w.__t.acceptKey(r124())].original === 'hole mark');

  console.log('\n8. The saved Master File copy (real ExcelJS)');
  await load(w, ROWS_RMK);
  w.__t.setRemark(r124(), '');                      // PI 124: edit removed -> Olisa's "hole mark" must stay
  const r130b = w.__t.allRecords.find(r => r.piRef === 130);
  const k130b = w.__t.acceptKey(r130b);
  if (!w.__t.acceptedShorts[k130b] || w.__t.acceptedShorts[k130b].cleared) {
    await openList(d);
    outstandingRows(d).find(b => b.dataset.k === k130b).click(); await tick();
  }
  await openList(d);
  undoBtns(d).find(b => b.dataset.k === k130b).click(); await tick();   // PI 130: settle undone
  // Bytes must belong to the PAGE's realm — JSZip inside the page rejects a Node ArrayBuffer.
  const raw = master(ROWS_RMK);
  const pageBytes = new w.Uint8Array(raw.length); pageBytes.set(raw);
  w.__t.setMasterBytes(pageBytes.buffer);
  let copy = null, copyErr = '';
  try {
    const { blob } = await w.__t.buildUpdatedMasterBlob();
    const buf = Buffer.from(await blob.arrayBuffer());
    const wb2 = XLSX.read(buf, { type: 'buffer' });
    copy = XLSX.utils.sheet_to_json(wb2.Sheets['2026- Main'], { header: 1, defval: null });
  } catch (e) { copyErr = e.message; }
  t('the copy builds', !!copy, copyErr);
  if (copy) {
    const hdr = copy[0].map(v => String(v || ''));
    const rc = hdr.findIndex(h => /^remarks?$/i.test(h)), uc = hdr.findIndex(h => /^update$/i.test(h));
    const row124 = copy.find(r => r[0] === 124), row130 = copy.find(r => r[0] === 130);
    t('a removed remark edit leaves Olisa\'s own remark in the copy (not a blank cell)',
      row124 && row124[rc] === 'hole mark', 'Remarks cell: ' + JSON.stringify(row124 && row124[rc]));
    t('an undone settle writes no "Settled by agreement" note', row130 && !/Settled by agreement/i.test(String(row130[uc] || '')),
      'Update cell: ' + JSON.stringify(row130 && row130[uc]));
  }

  console.log('\n9. Identical lines on one challan, and a save that cannot complete');
  ({ w, d, errors } = await boot());
  Object.keys(w.__t.acceptedShorts).forEach(k => delete w.__t.acceptedShorts[k]);
  const twinRow = [140, '20/06/2026', 'M82050-5E1-1', 'Master Carton', 'L40 x W30 x H20 cm', 60, 58, -2, '31910900', '20/06/2026', ''];
  await load(w, [...ROWS_V1, twinRow, twinRow.slice()]);
  setFilter(d, 140, 140);
  await openList(d);
  w.__prompts.length = 0;
  outstandingRows(d)[0].click(); await tick();
  t('settling one of two identical lines says the other settles with it',
    /1 other line\(s\) on this challan are identical/.test(w.__prompts[0] || ''), (w.__prompts[0] || '').slice(0, 200));
  w.__confirms.length = 0;
  if (undoBtns(d)[0]) { undoBtns(d)[0].click(); await tick(); }
  t('and Undo says the identical line comes back with it', /1 identical line\(s\) on the same challan come back/.test(w.__confirms[0] || ''),
    undoBtns(d).length ? '' : 'no Undo button on screen');
  setFilter(d, null, null);
  // A transaction that aborts (what a full disk does) must reject rather than hang forever.
  const realOpen = w.indexedDB.open.bind(w.indexedDB);
  w.indexedDB.open = () => {
    const req = {};
    setTimeout(() => {
      req.result = { transaction() {
        const tx = { objectStore: () => ({ put() {} }) };
        setTimeout(() => { tx.error = null; tx.onabort && tx.onabort(); }, 5);
        return tx;
      } };
      req.onsuccess && req.onsuccess();
    }, 5);
    return req;
  };
  const settledOrTimeout = await Promise.race([
    w.__t.idbSet('probe', { x: 1 }).then(() => 'resolved', e => 'rejected: ' + (e && e.message)),
    new Promise(r => setTimeout(() => r('hung'), 800))
  ]);
  w.indexedDB.open = realOpen;
  t('a save whose transaction aborts reports the failure instead of hanging', /^rejected/.test(settledOrTimeout), settledOrTimeout);

  console.log('\n6. Nothing else broke');
  const real = allErrors.filter(e => !/offline in test|Not implemented/i.test(e));
  t('no script errors while driving the page (every reload, every part)', real.length === 0, real.slice(0, 3).join(' | '));
  t('shorts and excesses are still separate totals (the excess is not netted)',
    w.__t.effectiveShort(w.__t.allRecords.find(r => r.piRef === 132)) === 2);

  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL harness: ' + (e && e.stack || e)); process.exit(1); });
