/* v138test.js — one guard per bug fixed in v138. Every check here FAILS on v137 and PASSES on v138.
   It runs the REAL page in a real DOM (real SheetJS, real ExcelJS, fake IndexedDB) and drives the
   same functions the buttons drive, plus three small text lifts for the parts that live in other
   modules (Undelivered report dates, DC Bypass Bill).
   Usage: node v138test.js olisa.html                                                           */
const fs = require('fs');
const path = require('path');
require('fake-indexeddb/auto');
const { JSDOM, VirtualConsole } = require('jsdom');
const XLSX = require('xlsx');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x !== undefined ? '  -> ' + x : ''))); };

const file = process.argv[2] || 'olisa.html';
const dir = path.dirname(path.resolve(file));
let html = fs.readFileSync(file, 'utf8');
const rawHtml = html;
html = html.replace(/<script src="\.\/lib\/xlsx\.full\.min\.js"><\/script>/, () => `<script>${fs.readFileSync(path.join(dir, 'lib', 'xlsx.full.min.js'), 'utf8')}</script>`);
html = html.replace(/<script src="\.\/lib\/exceljs\.min\.js"><\/script>/, () => `<script>${fs.readFileSync(path.join(dir, 'lib', 'exceljs.min.js'), 'utf8')}</script>`);
html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');
const anchor = '// ==================== END GOOGLE DRIVE MODE ====================';
if (!html.includes(anchor)) { console.log('FAIL probe anchor not found'); process.exit(1); }
html = html.replace(anchor, () => anchor + `
window.__t = {
  loadFile, lookupStyle, piLineForRow, piQtyForRow, piFillerLines, sheetToPiLines, parsePiLines,
  remarkColorWord, applyAskFilters, unbookedChallans, piFromChallanText, uqRowsFromAoa,
  uqGenerateReport, normPiKey,
  get allRecords() { return allRecords; }, set allRecords(v) { allRecords = v; },
  get piIndex() { return piIndex; }, set piIndex(v) { piIndex = v; },
  set challanIndex(v) { challanIndex = v; }, set styleIndex(v) { styleIndex = v; },
  set uqCandidates(v) { uqCandidates = v; }, get uqOutputRows() { return uqOutputRows; },
  get uqOutputHeaders() { return uqOutputHeaders; }
};`);

const HEAD = ['PI Ref', 'Inventory Date', 'Style Ref', 'Item', 'Item Description', 'Delivery Qty',
  'Received Qty', 'Short&Exs', 'Challan No', 'Challan Date', 'Remarks'];
function book(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, rows, head] of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head || HEAD, ...rows]), name);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
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
      w.prompt = () => 'x'; w.confirm = () => true;
      w.URL.createObjectURL = () => 'blob:test'; w.URL.revokeObjectURL = () => {};
    }
  });
  const w = dom.window;
  await new Promise(r => setTimeout(r, 400));
  return { w, d: w.document, errors };
}
const tick = (ms = 80) => new Promise(r => setTimeout(r, ms));

(async () => {
  const { w, d, errors } = await boot();
  t('the page loads with the probe and no errors', !!w.__t && !errors.length, errors.join(' | '));
  const T = w.__t;

  console.log('\n1. Master File quantities typed as text with thousands separators');
  await T.loadFile(new w.File([book([['2026- Main', [
    [170, '01/10/2026', 'M81988-3A2', 'Master Carton', 'L56 x W38 x H30 cm', '1,200', '1,165', '-35', '117327270', '01/10/2026', '']
  ]]])], 'Olisa Master Inventory v9.xlsx'));
  await tick();
  const r0 = T.allRecords[0] || {};
  t('Delivery Qty "1,200" reads as 1200 (was 1)', r0.delQty === 1200, r0.delQty);
  t('Received Qty "1,165" reads as 1165 (was 1)', r0.recQty === 1165, r0.recQty);
  t('Short "-35" still reads as -35', r0.short === -35, r0.short);

  console.log('\n2. Master File sheets for a NEW year are read');
  await T.loadFile(new w.File([book([
    ['2027- Main', [[301, '05/01/2027', 'M90001-1A1', 'Master Carton', 'L50 x W40 x H30 cm', 100, 100, 0, '200000001', '05/01/2027', '']]],
    ['2026- Main', [[170, '01/10/2026', 'M81988-3A2', 'Master Carton', 'L56 x W38 x H30 cm', 100, 100, 0, '117327270', '01/10/2026', '']]],
    ['2026',       [[170, '01/10/2026', 'M81988-3A2', 'Master Carton', 'L56 x W38 x H30 cm', 100, 100, 0, '117327270', '01/10/2026', '']]],
    ['2025',       [[90,  '01/10/2025', 'M70000-1A1', 'Master Carton', 'L40 x W30 x H20 cm', 100, 100, 0, '100000001', '01/10/2025', '']]],
    ['Summary',    [['x', '', 'NOT-A-ROW']]]
  ])], 'Olisa Master Inventory v10.xlsx'));
  await tick();
  const sheetsRead = [...new Set(T.allRecords.map(r => r.sheet))].sort().join(',');
  t('"2027- Main" is read (was skipped from January)', T.allRecords.some(r => r.styleNorm === 'M90001-1A1'), sheetsRead);
  t('"2026- Main" and the "2025" archive are still read', /2026- Main/.test(sheetsRead) && /2025/.test(sheetsRead), sheetsRead);
  t('a bare "2026" copy beside "2026- Main" is NOT read twice', T.allRecords.filter(r => r.styleNorm === 'M81988-3A2').length === 1, sheetsRead);

  console.log('\n3. Excel-format PI: lines of 1,000 pcs or more');
  const ws = XLSX.utils.aoa_to_sheet([['SL', 'Description', 'Quantity', 'Unit'],
    [1, 'Master Carton (56 x 38 x 30) CM M81988-3A2', 1200, 'Piece'],
    [2, 'Cross Divider M81988-3A2', 450, 'Piece'], ['Total Qty (Pcs)', null, 1650]]);
  ws['C2'].z = '#,##0'; ws['C3'].z = '#,##0'; ws['C4'].z = '#,##0';
  const pwb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(pwb, ws, 'PI');
  const piBuf = XLSX.write(pwb, { type: 'array', bookType: 'xlsx' });
  const parsed = T.parsePiLines(T.sheetToPiLines(new w.Uint8Array(new Uint8Array(piBuf)).buffer)); // page-realm buffer, as a real upload gives
  t('the 1,200-pc line is read (was lost)', parsed.styles['M81988-3A2'] === 1650, JSON.stringify(parsed.styles));
  t('the stated total is read so a gap would be flagged', parsed.statedTotal === 1650 && parsed.totalMismatch === null, parsed.statedTotal);

  console.log('\n4. Remark colours understand "not"');
  t('"Not OK" is NOT green', T.remarkColorWord('Not OK') !== 'green', T.remarkColorWord('Not OK'));
  t('"not accepted by Olisa" is NOT green', T.remarkColorWord('not accepted by Olisa') !== 'green');
  t('"Not settled yet" is NOT green', T.remarkColorWord('Not settled yet') !== 'green');
  t('"no issue" is not red', T.remarkColorWord('no issue') !== 'red');
  t('"Settled" and "accepted, usable" stay green', T.remarkColorWord('Settled') === 'green' && T.remarkColorWord('accepted, usable') === 'green');

  console.log('\n5. A Master Carton row never takes the Cross Divider PI quantity');
  T.piIndex = { '170': { styles: {}, items: [
    { style: 'M81988-3A2', dims: '56.5x38x30', meas: '', item: 'Master Carton', qty: 1000 },
    { style: 'M81988-3A2', dims: null, meas: '', item: 'Cross Divider', qty: 2000 } ] } };
  const q5 = T.piQtyForRow(170, 'M81988-3A2', 'L56 x W38 x H30 cm');
  t('master row with slightly different dims is not given 2000 (the divider)', q5 !== 2000, q5);
  t('a divider row still finds the divider line', T.piQtyForRow(170, 'M81988-3A2', 'Cross Divider L56 x W38 cm') === 2000);

  console.log('\n6. A PI line matched by fallback is not ALSO listed as "not delivered yet"');
  T.piIndex = { '170': { styles: {}, items: [{ style: 'M81988-3A2', dims: '56.5x38x30', meas: '', item: 'Master Carton', qty: 1000 }] } };
  T.allRecords = [{ piRef: 170, styleNorm: 'M81988-3A2', style: 'M81988-3A2', itemDesc: 'Master Carton', item: 'Master Carton', delQty: 1000, recQty: 1000, short: 0 }];
  t('the row does reach the PI line (fallback tier)', T.piQtyForRow(170, 'M81988-3A2', 'Master Carton') === 1000);
  const fillers = T.piFillerLines(new w.Set(['M81988-3A2']));
  t('so that PI line is not listed again as undelivered (was double-counted)', fillers.length === 0, fillers.length);

  console.log('\n7. PI-range filter keeps "PI-174"');
  T.allRecords = [{ piRef: 'PI-174', styleNorm: 'A', style: 'A', itemDesc: '', short: -5 },
                  { piRef: 200, styleNorm: 'B', style: 'B', itemDesc: '', short: -5 }];
  d.getElementById('askPiFrom').value = '170'; d.getElementById('askPiTo').value = '180';
  const f7 = T.applyAskFilters(T.allRecords);
  t('"PI-174" sits inside 170–180 (was hidden)', f7.length === 1 && f7[0].piRef === 'PI-174', f7.map(r => r.piRef).join(','));
  d.getElementById('askPiFrom').value = ''; d.getElementById('askPiTo').value = '';

  console.log('\n8. Unbooked challans list newest first by real date');
  T.allRecords = [];
  T.challanIndex = {
    '1': { no: '111111111', date: '31-Oct-2026', lines: [{ style: 'M81988-3A2', qty: 10, desc: 'x' }] },
    '2': { no: '222222222', date: '01-Nov-2026', lines: [{ style: 'M81988-3A2', qty: 10, desc: 'x' }] } };
  const ub = T.unbookedChallans();
  t('01-Nov-2026 comes before 31-Oct-2026 (was day-number order)', ub[0] && ub[0].entry.date === '01-Nov-2026', ub.map(x => x.entry.date).join(' , '));
  T.challanIndex = {};

  console.log('\n9. Challan header PI without zero padding');
  t('"PI : 12/3/174011025" gives PI 174 (was none)', T.piFromChallanText('PI : 12/3/174011025') === '174', T.piFromChallanText('PI : 12/3/174011025'));
  t('the padded form still works', T.piFromChallanText('PI: 1/1/0174011025') === '174');

  console.log('\n10. A request file\'s Total row is not a request line');
  const rq = T.uqRowsFromAoa([['Style', 'Order Qty'], ['M81988-3A2', 500], [null, 300], ['Total', 800]], 0, 0, ['Style', 'Order Qty'], null);
  t('two request lines, not three', rq.length === 2, rq.length);

  console.log('\n11. Style lookup includes spelling variants, like settling does');
  T.piIndex = {};
  T.allRecords = [
    { piRef: 170, styleNorm: 'M81988-3A2', style: 'M81988-3A2', itemDesc: 'L56 x W38 x H30 cm', short: 35, delQty: 100, recQty: 135 },
    { piRef: 160, styleNorm: 'M81988 3A2', style: 'M81988 3A2', itemDesc: 'L56 x W38 x H30 cm', short: -35, delQty: 100, recQty: 65 } ];
  const lk = T.lookupStyle('M81988-3A2', { noFilters: true });
  t('both spellings are returned (was only the exact one)', lk.matches.length === 2, lk.matches.length);

  console.log('\n12. Request report counts PIs with nothing delivered yet');
  T.allRecords = [{ piRef: 170, styleNorm: 'M81988-3A2', style: 'M81988-3A2', item: 'Master Carton',
    itemDesc: 'L56 x W38 x H30 cm', delQty: 1000, recQty: 1000, short: 0, challanNo: '117327270', challanDate: '01/10/2026', invDate: '01/10/2026' }];
  T.piIndex = {
    '170': { styles: { 'M81988-3A2': 1000 }, items: [{ style: 'M81988-3A2', dims: '56x38x30', meas: '', item: 'Master Carton', qty: 1000 }] },
    '180': { styles: { 'M81988-3A2': 500 },  items: [{ style: 'M81988-3A2', dims: '56x38x30', meas: '', item: 'Master Carton', qty: 500 }] } };
  T.styleIndex = { 'M81988-3A2': ['170', '180'] };
  T.uqCandidates = [{ sheetName: 'S', headers: ['Style'], rows: [{ style: 'M81988-3A2', cells: ['M81988-3A2'] }], _checked: true }];
  await T.uqGenerateReport(); await tick();
  const H12 = T.uqOutputHeaders || [], row12 = (T.uqOutputRows || [])[0] || { cells: [] };
  const col = n => row12.cells[H12.indexOf(n)];
  t('PI Qty is 1500 (both PIs), was 1000', col('PI Qty') === 1500, col('PI Qty'));
  t('Undelivered Qty is 500, was 0', col('Undelivered Qty') === 500, col('Undelivered Qty'));
  t('the verdict no longer says "nothing pending"', !/nothing pending/i.test(String(col('Verdict'))), col('Verdict'));

  console.log('\n13. PO Summary: text quantities, merged style cells, and the Delivery Date box');
  function poFile(name, dateTxt, merged) {
    const aoa = [['ORDER SIZE', 'ORDER QTY', 'STYLE', 'DELIVERY DATE'],
      ['CARTON 56*38*30CM', '1,200', 'M81988-3A2', dateTxt],
      ['CARTON 60*40*30CM', 300, merged ? null : 'M81988-3A2', dateTxt]];
    const s = XLSX.utils.aoa_to_sheet(aoa);
    if (merged) s['!merges'] = [{ s: { r: 1, c: 2 }, e: { r: 2, c: 2 } }];
    const b = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(b, s, 'S');
    return new w.File([XLSX.write(b, { type: 'buffer', bookType: 'xlsx' })], name);
  }
  function dropFiles(files) {
    const ev = new w.Event('drop', { bubbles: true, cancelable: true });
    ev.dataTransfer = { files };
    d.getElementById('podrop').dispatchEvent(ev);
  }
  async function generatePo() { d.getElementById('pogenBtn').click(); for (let i = 0; i < 40 && /Reading|Classifying|Drawing/.test(d.getElementById('postatus').textContent); i++) await tick(50); await tick(150); }
  dropFiles([poFile('wo-a.xlsx', '01.11.26', true)]);
  await generatePo();
  const prev = d.getElementById('popreviewTable').textContent;
  t('the "1,200" text quantity reads as 1200 (was 1)', /1200/.test(prev) && !/\b1\s*$/m.test(prev), d.getElementById('postatus').textContent);
  t('the second size under a MERGED style cell is kept (was dropped)', (prev.match(/M81988-3A2/g) || []).length === 2, (prev.match(/M81988-3A2/g) || []).length);
  const box = d.getElementById('podeliveryDateInput');
  t('the Delivery Date box is filled from the first work order', box.value === '01.11.26', box.value);
  dropFiles([poFile('wo-b.xlsx', '15.11.26', false)]);
  await generatePo();
  t('a second work order refreshes it (was stuck on the first file\'s date)', box.value === '15.11.26', box.value);
  box.value = 'typed by hand';
  dropFiles([poFile('wo-c.xlsx', '20.11.26', false)]);
  await generatePo();
  t('a date typed by hand is never overwritten', box.value === 'typed by hand', box.value);

  console.log('\n14. Undelivered report dates');
  const lift = name => (rawHtml.match(new RegExp('function ' + name + '\\([\\s\\S]*?\\n\\}')) || [''])[0];
  const U = new Function('snapToLocalMidnight', lift('formatDateValue') + '\n' + lift('deliverySortKey') + '\nreturn {formatDateValue, deliverySortKey};')(x => x);
  t('"15-Nov-2026" has a delivery date (was "No delivery date given")', U.deliverySortKey('15-Nov-2026') === Date.UTC(2026, 10, 15), U.deliverySortKey('15-Nov-2026'));
  t('an unformatted serial 46341 reads as 15/11/2026 (was printed as 46341)', U.formatDateValue(46341) === '15/11/2026', U.formatDateValue(46341));
  t('the title is built from the date box at download time', /titleCell\.value = `OLISA's Undelivered Qty & Delivery Schedule\\n\(\$\{dateInput\.value\}\)`/.test(rawHtml));

  console.log('\n15. DC Bypass Bill');
  const dc = fs.readFileSync(path.join(dir, 'DC_Bypass_Bill.html'), 'utf8');
  t('the PDF header says "Total Value ($)" (was "Tota Value")', dc.includes('"Total Value ($)"') && !dc.includes('Tota Value'));
  const a = dc.indexOf('var CORE = (function(){'), b = dc.indexOf('\n})();', a);
  const CORE = new Function(dc.slice(a, b + 6) + '\nreturn CORE;')();
  t('amount in words at 125 crore has no "undefined"', !/undefined/i.test(CORE.amountInWords(1250000000)), CORE.amountInWords(1250000000));
  t('ordinary amounts are unchanged', CORE.amountInWords(1234567) === 'Twelve lakh thirty four thousand five hundred sixty seven', CORE.amountInWords(1234567));

  console.log('\n16. One version everywhere');
  const v = (rawHtml.match(/const APP_VERSION = (\d+);/) || [])[1];
  const c = (fs.readFileSync(path.join(dir, 'sw.js'), 'utf8').match(/olisa-tools-v(\d+)/) || [])[1];
  // Was pinned to exactly 138, which made this file fail on every release after its own. The
  // guard that matters is that the two numbers move together and never go backwards.
  t('page and service worker carry the same version, v138 or later', v === c && Number(v) >= 138, `page ${v}, sw ${c}`);

  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('FAIL crashed: ' + (e && e.stack || e)); process.exit(1); });
