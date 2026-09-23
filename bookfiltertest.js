/* bookfiltertest.js — the two v136 bugs, driven through the REAL page.
   1. "It is in the Master File, so why does it say not in Olisa's inventory?" The booking check read
      only the first challan number in a cell and compared style text exactly, so the search found
      Olisa's row and flagged the same delivery as unbooked underneath it.
   2. "Chip Box only shows me everything, and calls chip boxes Master Carton." PI coverage was worked
      out from the FILTERED rows, so every delivered line of another kind came back as "Not delivered
      yet (from PI)"; the filter never reached those lines; and the PI reader dropped "(Punch)".
   Loads a real Master workbook with the real SheetJS, sets the challan and PI indexes, types into the
   actual Ask box with the actual filter dropdown, and reads what the table shows.
   Usage: node bookfiltertest.js olisa.html                                                        */
const fs = require('fs');
const path = require('path');
require('fake-indexeddb/auto');
const { JSDOM, VirtualConsole } = require('jsdom');
const XLSX = require('xlsx');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
let html = fs.readFileSync(file, 'utf8');
const xlsxLib = fs.readFileSync(path.join(dir, 'lib', 'xlsx.full.min.js'), 'utf8');
html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${xlsxLib}</script>`);
html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');
const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
if (!html.includes(anchor)) { console.log('FAIL probe anchor not found'); process.exit(1); }
html = html.replace(anchor, () => anchor + `
// typeof guards: this probe must also load into OLDER builds, so the suite can prove it fails there.
const __opt = f => (typeof f === 'function' ? f : () => null);
window.__t = {
  loadFile, challanStatus, unbookedLinesForStyle, unbookedBlockHtml, parsePiLines, applyAskFilters, itemLabel,
  challanKeysOf: __opt(typeof challanKeysOf === 'function' ? challanKeysOf : null),
  piLineKind: __opt(typeof piLineKind === 'function' ? piLineKind : null),
  askFilterPassesPiLine: __opt(typeof askFilterPassesPiLine === 'function' ? askFilterPassesPiLine : null),
  get allRecords() { return allRecords; },
  setChallans(c) { challanIndex = c; challanIndexBuiltAt = Date.now(); },
  setPi(p) { piIndex = p; piIndexBuiltAt = Date.now(); },
  setMaster(m) { currentMasterMeta = m; }
};`);

const HEAD = ['PI Ref', 'Inventory Date', 'Style Ref', 'Item', 'Item Description', 'Delivery Qty',
  'Received Qty', 'Short&Exs', 'Challan No', 'Challan Date', 'Remarks'];
function master(rows) {
  const ws = XLSX.utils.aoa_to_sheet([HEAD, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '2026- Main');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
const MC = 'Master Carton, 5 Ply, L51.5 x W40 x H26.5cm;';
const PU = 'Master Carton (Punch), 3 Ply, L33 x W30.5 x H30.5cm;';
const CD = 'Cross Divider, 5 Ply, L50 x W39 x H0cm;';   // the real Master/challan wording (H0)
const ROWS = [
  // ---- part 1: bookings Olisa made, written the way people actually type them ----
  [140, '01/09/2026', 'M82036A6 - 1', 'Master Carton', MC, 50, 50, 0, 31910793, '01/09/2026', ''],            // spaces in the style
  [140, '01/09/2026', 'M82040-2A1', 'Master Carton', MC, 80, 80, 0, '31910799, 31910794', '01/09/2026', ''], // two challans, one cell
  [140, '01/09/2026', 'M82044-1A', 'Master Carton', MC, 30, 30, 0, '31910811 31910812', '01/09/2026', ''],   // two challans, space only
  [140, '01/09/2026', 'M82041-3C', 'Master Carton', MC, 20, 20, 0, '31910795', '01/09/2026', ''],            // _ vs -
  [140, '01/09/2026', 'M82050-5E1', 'Master Carton', MC, 60, 60, 0, '31910796', '01/09/2026', ''],           // parent code booked
  [140, '01/09/2026', 'M82039A-3', 'Master Carton', MC, 10, 10, 0, '31910797', '01/09/2026', ''],            // only ONE of two lines booked
  [140, '01/09/2026', 'M82035A1-31', 'Master Carton', MC, 15, 15, 0, '31910798', '01/09/2026', ''],          // a DIFFERENT style
  [140, '01/09/2026', 'M82060-1B', 'Master Carton', MC, 90, 90, 0, '31910999', '01/09/2026', ''],            // mistyped challan no.
  [140, '01/09/2026', 'M82061-1C', 'Master Carton', MC, 12, 12, 0, '31910793/94', '01/09/2026', ''],        // "/94" is not a challan
  // ---- part 2: one style carrying all three kinds on PI 150 ----
  [150, '05/09/2026', 'M82065-3B', 'Master Carton', MC, 100, 100, 0, '31920001', '05/09/2026', ''],
  [150, '05/09/2026', 'M82065-3B', 'Master Carton', PU, 200, 200, 0, '31920001', '05/09/2026', ''],
  [150, '05/09/2026', 'M82065-3B', 'Cross Divider', CD, 100, 100, 0, '31920001', '05/09/2026', ''],
  // a chip box Olisa's file describes only as "Master Carton" + dimensions
  [151, '06/09/2026', 'M82070-2A', 'Master Carton', 'L33 x W30.5 x H30.5cm', 40, 40, 0, '31920002', '06/09/2026', ''],
  [151, '06/09/2026', 'M82071-2A', 'Master Carton', 'Master Carton, 3 Ply, L28 x W25 x H22cm;', 40, 40, 0, '31920003', '06/09/2026', '']
];
const line = (style, qty, desc) => ({ style, qty, desc: desc || MC });
const CHALLANS = {
  a: { no: '31910793', date: '2026-09-01', verified: true, lines: [line('M82036A6-1', 50)] },
  b: { no: '31910794', date: '2026-09-01', verified: true, lines: [line('M82040-2A1', 80)] },
  c: { no: '31910812', date: '2026-09-01', verified: true, lines: [line('M82044-1A', 30)] },
  d: { no: '31910795', date: '2026-09-01', verified: true, lines: [line('M82041_3C', 20)] },
  e: { no: '31910796', date: '2026-09-01', verified: true, lines: [line('M82050-5E1-1', 60)] },
  f: { no: '31910797', date: '2026-09-01', verified: true, lines: [line('M82039A-3', 10), line('M82039A', 25)] },
  g: { no: '31910798', date: '2026-09-01', verified: true, lines: [line('M82035A1-3', 15)] },
  h: { no: '31910990', date: '2026-09-02', verified: true, lines: [line('M82060-1B', 90)] },
  i: { no: '94', hand: true, date: '2026-09-02', verified: true, lines: [line('M82061-1C', 12)] },
  j: { no: '31920001', date: '2026-09-05', verified: true, lines: [line('M82065-3B', 100, MC), line('M82065-3B', 200, PU), line('M82065-3B', 100, CD)] },
  k: { no: '31920002', date: '2026-09-06', verified: true, lines: [line('M82070-2A', 40, PU)] },
  // our challan for M82065-3B that Olisa has not booked: one chip box, one master carton
  n: { no: '31930000', date: '2026-09-10', verified: true, lines: [line('M82065-3B', 70, PU), line('M82065-3B', 30, MC)] }
};
const PI = {
  '150': { items: [
    { style: 'M82065-3B', dims: '51.5x40x26.5', meas: 'Ply: 5 (L51.5 x W40 x H26.5) CM', item: 'Master Carton', qty: 100 },
    { style: 'M82065-3B', dims: '33x30.5x30.5', meas: 'Ply: 3 (L33 x W30.5 x H30.5) CM', item: 'Master Carton', qty: 200 },   // an OLD cached read: "(Punch)" lost
    { style: 'M82065-3B', dims: '50x39x0', meas: 'Ply: 5 (50 x 39 x -) CM', item: 'Cross Divider', qty: 100 } ] },
  '152': { items: [
    // ordered, never delivered: one of each kind
    { style: 'M82065-3B', dims: '60x40x30', meas: 'Ply: 5 (L60 x W40 x H30) CM', item: 'Master Carton', qty: 11 },
    { style: 'M82065-3B', dims: '33x30.5x30.5', meas: 'Ply: 3 (L33 x W30.5 x H30.5) CM', item: 'Master Carton (Punch)', qty: 22 },
    { style: 'M82065-3B', dims: '59x39x0', meas: 'Ply: 5 (59 x 39 x -) CM', item: 'Cross Divider', qty: 33 } ] }
};

async function boot() {
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
      w.fetch = () => Promise.reject(new Error('offline in test'));
      w.navigator.serviceWorker = { register: () => Promise.resolve({ addEventListener() {} }), addEventListener() {}, ready: Promise.resolve({}) };
      w.PDFLib = {}; w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
      w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {};
      w.CompressionStream = undefined; w.DecompressionStream = undefined;
    }
  });
  await new Promise(r => setTimeout(r, 400));
  return { w: dom.window, d: dom.window.document, errors };
}
const tick = (ms = 80) => new Promise(r => setTimeout(r, ms));
// Every table row in the answer, as { item, remark, style, qty } — read off the screen.
function tableRows(d) {
  const out = [];
  d.querySelectorAll('#askChat table.preview').forEach(tb => {
    const head = [...tb.querySelectorAll('tr:first-child th')].map(th => th.textContent.trim());
    [...tb.querySelectorAll('tr')].slice(1).forEach(tr => {
      const c = [...tr.querySelectorAll('td')].map(td => td.textContent.trim());
      if (c.length < 3) return;
      const o = {}; head.forEach((h, i) => { o[h] = c[i]; });
      out.push(o);
    });
  });
  return out;
}
async function ask(w, d, q, mode) {
  d.getElementById('askFilterMode').value = mode || 'all';
  d.getElementById('askPiFrom').value = ''; d.getElementById('askPiTo').value = '';
  d.getElementById('askInput').value = q;
  d.getElementById('askBtn').click();
  await tick(250);
  return { rows: tableRows(d), text: d.getElementById('askChat').textContent };
}

(async () => {
  const { w, d, errors } = await boot();
  const T = w.__t;
  t('the page loads with the probe', !!T, errors.join(' | '));
  await T.loadFile(new w.File([master(ROWS)], 'Olisa Master Inventory - v1.xlsx'));
  await tick();
  t('the Master File loads', T.allRecords.length === ROWS.length, 'records: ' + T.allRecords.length);
  T.setChallans(CHALLANS); T.setPi(PI);
  T.setMaster({ id: 'x', name: 'Olisa Master Inventory - v1.xlsx', modifiedMs: Date.UTC(2026, 8, 22, 10, 5) });

  console.log('\n1. A delivery Olisa DID book is not reported as unbooked');
  const st = k => T.challanStatus(CHALLANS[k]).state;
  t('spaces in Olisa\'s style ("M82036A6 - 1")', st('a') === 'booked', st('a'));
  t('second challan in a comma cell ("31910799, 31910794")', st('b') === 'booked', st('b'));
  t('second challan in a space-only cell ("31910811 31910812")', st('c') === 'booked', st('c'));
  t('underscore vs hyphen (M82041_3C = M82041-3C)', st('d') === 'booked', st('d'));
  t('sub-code booked under its parent (M82050-5E1-1 under M82050-5E1)', st('e') === 'booked', st('e'));
  t('a challan carrying all three kinds of one style', st('j') === 'booked', st('j'));

  console.log('\n2. What really is unbooked still is — the looser match claims nothing it should not');
  const f = T.challanStatus(CHALLANS.f);
  t('one of two lines booked -> part-booked', f.state === 'part-booked', f.state);
  t('the unbooked neighbour M82039A is not hidden by the booked M82039A-3',
    f.missing.length === 1 && f.missing[0].style === 'M82039A', JSON.stringify(f.missing.map(x => x.style)));
  t('M82035A1-3 is NOT M82035A1-31 (a number is never split)', st('g') === 'not-booked', st('g'));
  t('a challan Olisa never entered is not-booked', st('h') === 'not-booked', st('h'));
  t('"31910793/94" does not invent challan 94', !(T.challanKeysOf('31910793/94') || []).includes('94'), JSON.stringify(T.challanKeysOf('31910793/94')));
  t('so hand challan 94 is still not booked', st('i') === 'not-booked', st('i'));
  t('a date in the cell is not a challan either', JSON.stringify(T.challanKeysOf('31910793 (12/06/2026)')) === '["31910793"]',
    JSON.stringify(T.challanKeysOf('31910793 (12/06/2026)')));
  t('a plain hand challan number still works', JSON.stringify(T.challanKeysOf('Manual 114')) === '["114"]');
  t('an Excel number still works', JSON.stringify(T.challanKeysOf(31910793)) === '["31910793"]');

  console.log('\n3. Every flagged line says why, and which Master File it was checked against');
  const h = T.unbookedLinesForStyle('M82060-1B');
  t('a missing challan says so', h.length === 1 && /is not in Olisa's file at all/.test(h[0].why), h[0] && h[0].why);
  t('and points at the same style + qty under another number (likely a typo)', h[0] && /31910999/.test(h[0].why), h[0] && h[0].why);
  const fl = T.unbookedLinesForStyle('M82039A').filter(r => r.style === 'M82039A');
  t('a challan that IS there, minus this style, says that instead', fl.length === 1 && /is in Olisa's file, but not this style/.test(fl[0].why), fl[0] && fl[0].why);
  const blk = T.unbookedBlockHtml('M82060-1B', { ok: false });
  t('the block has a "Why it is listed" column', /Why it is listed/.test(blk));
  t('and names the Master File it was checked against', /Checked against Olisa's file: Olisa Master Inventory - v1\.xlsx \(saved 22\/09\/2026/.test(blk), blk.slice(-220));

  console.log('\n4. Through the real search box: a booked delivery is not flagged under Olisa\'s own row');
  let r = await ask(w, d, 'M82040-2A1');
  t('Olisa\'s row is found', r.rows.some(x => x['Style'] === 'M82040-2A1' && x['Delivered Qty'] === '80'));
  t('THE BUG: it is not also listed as "not in Olisa\'s inventory"', !/not in Olisa's inventory yet/.test(r.text), r.text.slice(0, 300));

  console.log('\n5. Item type: what the tool already knows');
  t('"Master Carton (Punch)" on the row -> Chip Box', T.itemLabel({ styleNorm: 'X1', itemDesc: PU, item: 'Master Carton' }) === 'Chip Box (Punch)');
  t('a row saying only "Master Carton" + dims learns Punch from our challan for the same carton',
    T.itemLabel(T.allRecords.find(x => x.styleNorm === 'M82070-2A')) === 'Chip Box (Punch)', T.itemLabel(T.allRecords.find(x => x.styleNorm === 'M82070-2A')));
  t('"Master Carton, 3 Ply" -> Chip Box (the 3-ply carton)', T.itemLabel(T.allRecords.find(x => x.styleNorm === 'M82071-2A')) === 'Chip Box (Punch)');
  t('"Master Carton, 5 Ply" stays a Master Carton', T.itemLabel({ styleNorm: 'X2', itemDesc: MC, item: 'Master Carton' }) === 'Master Carton');
  t('an OLD cached PI line that lost "(Punch)" is recognised from the Master File', T.piLineKind(PI['150'].items[1]) === 'punch', T.piLineKind(PI['150'].items[1]));
  t('the PI reader now keeps "(Punch)"', (() => {
    const p = T.parsePiLines(['SL Description Quantity', '1 Master Carton (Punch) Ply: 3 (L33 x W30.5 x H30.5) CM M82099-1A 500 Piece', 'Total Qty (Pcs): 500']);
    return p.items.length === 1 && p.items[0].item === 'Master Carton (Punch)';
  })());

  console.log('\n6. Through the real search box: "Chip Box (Punch) only"');
  r = await ask(w, d, 'M82065-3B', 'punch');
  const items = [...new Set(r.rows.filter(x => x['Item']).map(x => x['Item']))];
  t('every Item on screen is a Chip Box', items.length === 1 && items[0] === 'Chip Box (Punch)', JSON.stringify(items));
  t('THE BUG: delivered Master Carton / Cross Divider lines do not come back as "Not delivered yet"',
    !r.rows.some(x => /Not delivered yet/.test(x['Remarks'] || '') && x['PI Ref'] === '150'), JSON.stringify(r.rows.filter(x => x['PI Ref'] === '150')));
  const nd = r.rows.filter(x => /Not delivered yet/.test(x['Remarks'] || ''));
  t('the one chip box still on order (PI 152, 22 pcs) is listed', nd.length === 1 && nd[0]['PI Ref'] === '152' && nd[0]['PI Qty'] === '22', JSON.stringify(nd));
  const ub = r.rows.filter(x => x['Why it is listed']);
  t('"not in Olisa\'s inventory" shows only the unbooked chip box (70 pcs), not the master carton',
    ub.length === 1 && ub[0]['Delivered Qty'] === '70' && ub[0]['Item'] === 'Chip Box (Punch)', JSON.stringify(ub));
  t('the answer says the filter is on', /Filter on: Chip Box \(Punch\) only/.test(r.text));

  console.log('\n7. The other two kinds, and no filter');
  r = await ask(w, d, 'M82065-3B', 'master');
  t('Master Carton only -> only Master Carton lines', r.rows.filter(x => x['Item']).every(x => x['Item'] === 'Master Carton'),
    JSON.stringify([...new Set(r.rows.map(x => x['Item']))]));
  t('and its undelivered line is the 11 pcs on PI 152', r.rows.filter(x => /Not delivered yet/.test(x['Remarks'] || '')).map(x => x['PI Qty']).join() === '11');
  r = await ask(w, d, 'M82065-3B', 'cross');
  t('Cross Divider only -> only Cross Divider lines', r.rows.filter(x => x['Item']).every(x => x['Item'] === 'Cross Divider'),
    JSON.stringify([...new Set(r.rows.map(x => x['Item']))]));
  t('and no "not in inventory" block (both unbooked lines are other kinds)', !/not in Olisa's inventory yet/.test(r.text));
  r = await ask(w, d, 'M82065-3B', 'all');
  const ndAll = r.rows.filter(x => /Not delivered yet/.test(x['Remarks'] || ''));
  t('no filter: exactly the three PI 152 lines are undelivered, none from PI 150', ndAll.length === 3 && ndAll.every(x => x['PI Ref'] === '152'),
    JSON.stringify(ndAll.map(x => x['PI Ref'] + ':' + x['Item'])));
  t('and they are labelled by kind', ndAll.map(x => x['Item']).sort().join('|') === 'Chip Box (Punch)|Cross Divider|Master Carton', ndAll.map(x => x['Item']).join('|'));
  t('both unbooked lines show when nothing is filtered', r.rows.filter(x => x['Why it is listed']).length === 2);

  console.log('\n8. A filter nothing delivered passes, and the Manual Challan filter');
  r = await ask(w, d, 'M82070-2A', 'cross');
  t('says nothing delivered matches, naming the filter', /none of the delivered lines match the filter \(Cross Divider only\)/.test(r.text), r.text.slice(0, 200));
  t('Manual Challan filter never lists a never-delivered PI line', T.askFilterPassesPiLine('152', PI['152'].items[0], { ok: true, manualOnly: true, kind: '', from: NaN, to: NaN }) === false);
  t('the PI range applies to undelivered lines too', T.askFilterPassesPiLine('152', PI['152'].items[0], { ok: true, manualOnly: false, kind: '', from: 140, to: 151 }) === false);

  t('no page errors', errors.length === 0, errors.join(' | '));
  console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL harness: ' + (e && e.stack || e)); process.exit(1); });
