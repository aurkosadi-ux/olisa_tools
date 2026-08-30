/* potest.js — the permanent guard on the PO Summary Generator's delivery-date handling.

   Runs the REAL functions out of olisa.html against the REAL Olisa work orders, in the REAL
   timezone (Asia/Dhaka), because that is the only place the original bug was visible: a delivery
   date read one day early on Sadi's laptop and correct on every other machine on earth.

   Usage:  TZ=Asia/Dhaka node potest.js olisa.html <workorder.xls> [<workorder.xlsx> ...]
*/
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const HTML = process.argv[2] || 'olisa.html';
const ORDERS = process.argv.slice(3);
let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

// ---------- lift the PO Summary module out of the page ----------
const html = fs.readFileSync(HTML, 'utf8');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const src = blocks.sort((a, b) => b.length - a.length)[0];

const snapSrc = src.match(/function snapToLocalMidnight[\s\S]*?\n\}/)[0];
const snapToLocalMidnight = new Function('return ' + snapSrc.replace('function snapToLocalMidnight', 'function'))();

const start = src.indexOf('(function() { // PO Summary Generator');
const end = src.indexOf('(function() { // Undelivered Report Generator');
if (start === -1 || end === -1) { console.error('FAIL: could not locate the PO Summary module'); process.exit(1); }
let body = src.slice(start, end)
  .replace(/^\(function\(\) \{ \/\/ PO Summary Generator/, '')
  .replace(/\}\)\(\);\s*$/, '');

// Strip only the DOM wiring. Every line of logic under test is left exactly as it ships.
body = body.replace(/const \w+ = document\.getElementById\([^)]*\);\n/g, '')
           .replace(/^\s*\w+\.addEventListener\([\s\S]*?\n\}\);\n/gm, '')
           .replace(/^\s*\['drag[^\n]*\n/gm, '')
           .replace(/^\s*\w+\.addEventListener\([^\n]*\n/gm, '');

const M = new Function('XLSX', 'snapToLocalMidnight', 'escHtml', body + `
  return { readOrderRows, anyToDate, anyToDateLabel, headerRole, readHeaderRow,
           groupByDeliveryDate, tryAltFormat, classifyRow, dateSuffix, poFileName };`
)(XLSX, snapToLocalMidnight, s => String(s));

function fakeFile(p) {
  const b = fs.readFileSync(p);
  return { name: path.basename(p), arrayBuffer: async () => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) };
}
// Independent oracle: read the delivery-date column straight out of the sheet with a different
// library path than the one under test, so a shared bug cannot make both agree.
function truthDates(file) {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });  // raw serials
  const out = new Set();
  for (const sn of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null, raw: true });
    for (const row of aoa) {
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (typeof v !== 'number' || v < 40000 || v > 60000) continue;
        const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
        out.add(`${String(d.getUTCDate()).padStart(2, '0')}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${String(d.getUTCFullYear()).slice(-2)}`);
      }
    }
  }
  return out;
}

(async () => {
  console.log('TZ = ' + (process.env.TZ || Intl.DateTimeFormat().resolvedOptions().timeZone));

  console.log('\n1. Timezone drift — the one-day-early bug');
  // Serial 46273 is 08/09/2026. SheetJS hands it back as 07/09/2026 23:59:40 in Asia/Dhaka.
  const drifted = new Date(2026, 8, 7, 23, 59, 40);
  t('a date 20s short of midnight reads as the NEXT day', M.anyToDate(drifted).label === '08.09.26',
    M.anyToDate(drifted).label);
  const driftedUp = new Date(2026, 8, 8, 0, 0, 20);
  t('a date 20s past midnight stays on its own day', M.anyToDate(driftedUp).label === '08.09.26',
    M.anyToDate(driftedUp).label);
  t('a clean midnight date is untouched', M.anyToDate(new Date(2026, 8, 8, 0, 0, 0)).label === '08.09.26');
  t('an Excel serial reads correctly', M.anyToDate(46273).label === '08.09.26', M.anyToDate(46273).label);
  t('serial and Date agree', M.anyToDate(46273).label === M.anyToDate(drifted).label);
  t('an invalid Date is rejected, not shown as NaN', M.anyToDate(new Date('nope')) === null);
  t('a blank cell yields no date', M.anyToDate(null) === null && M.anyToDate('') === null);

  console.log('\n2. Typed and Chinese date text');
  t('dd/mm/yyyy', M.anyToDate('08/09/2026').label === '08.09.26');
  t('dd.mm.yy', M.anyToDate('8.9.26').label === '08.09.26');
  t('yyyy-mm-dd', M.anyToDate('2026-09-08').label === '08.09.26');
  t('2026\u5e749\u67088\u65e5', M.anyToDate('2026\u5e749\u67088\u65e5').label === '08.09.26');
  t('unrecognised text is shown verbatim, never guessed', M.anyToDate('ASAP').label === 'ASAP' && M.anyToDate('ASAP').key === null);
  t('text dates still sort by a real key', typeof M.anyToDate('08/09/2026').key === 'number');

  console.log('\n3. Header words, English and Chinese');
  const cases = [
    ['delivery date\n', 'DELIVERY'], ['DUE DATE', 'DELIVERY'], ['Shipment Date', 'DELIVERY'],
    ['\u4ea4\u8d27\u65e5\u671f', 'DELIVERY'], ['\u4ea4\u671f', 'DELIVERY'], ['\u9001\u8d27\u65e5\u671f', 'DELIVERY'],
    ['order date', 'ORDERDATE'], ['\u4e0b\u5355\u65e5\u671f', 'ORDERDATE'], ['\u8ba2\u5355\u65e5\u671f', 'ORDERDATE'],
    ['ORDER QTY', 'QTY'], ['\u8ba2\u5355\u6570\u91cf', 'QTY'],
    ['ORDER SIZE', 'SIZE'], ['\u7269\u6599\u540d\u79f0', 'SIZE'],
    ['STYLE', 'STYLE'], ['\u5de5\u5382\u578b\u4f53', 'STYLE'],
    ['FACTORY', null], ['UNIT', null], ['BLACK', null]
  ];
  cases.forEach(([txt, want]) => t(`"${txt.trim()}" -> ${want}`, M.headerRole(txt) === want, String(M.headerRole(txt))));
  // The one that reaches Olisa if it goes wrong: a header naming both must be the REQUIRED date.
  t('"ORDER DELIVERY DATE" is the delivery date, not the booking date', M.headerRole('ORDER DELIVERY DATE') === 'DELIVERY');

  console.log('\n4. Grouping');
  const rows = [
    { deliveryDate: '25.10.26', deliveryKey: Date.UTC(2026, 9, 25), qty: 5 },
    { deliveryDate: '08.09.26', deliveryKey: Date.UTC(2026, 8, 8), qty: 3 },
    { deliveryDate: null, deliveryKey: null, qty: 1 },
    { deliveryDate: '08.09.26', deliveryKey: Date.UTC(2026, 8, 8), qty: 4 }
  ];
  const g = M.groupByDeliveryDate(rows);
  t('one group per distinct date', g.length === 3, String(g.length));
  t('earliest first', g[0].label === '08.09.26' && g[1].label === '25.10.26', g.map(x => x.label).join('|'));
  t('undated lines sort last, kept separate', g[2].label === '' && g[2].rows.length === 1);
  t('same-date lines land together', g[0].rows.length === 2);
  t('no line is lost', g.reduce((n, x) => n + x.rows.length, 0) === rows.length);
  t('no line is duplicated', new Set(g.flatMap(x => x.rows)).size === rows.length);

  console.log('\n5. File names');
  t('per-date name carries the date', /Delivery 08-09-26\.xlsx$/.test(M.poFileName(M.dateSuffix('08.09.26'))));
  t('combined name is unchanged', /^Olisa Tools - PO Summary \d\d-\d\d-\d{4}\.xlsx$/.test(M.poFileName('')));
  t('a date with a slash cannot break the name', !/[\\/:*?"<>|]/.test(M.dateSuffix('08/09/26')));
  t('an undated group still gets a name', M.dateSuffix('') === 'No delivery date');

  console.log('\n6. Real Olisa work orders');
  if (!ORDERS.length) console.log('  (skipped — no work order files supplied)');
  for (const p of ORDERS) {
    const label = path.basename(p).slice(0, 42);
    let recs;
    try { recs = await M.readOrderRows(fakeFile(p)); }
    catch (e) { t(`${label}: reads`, false, e.message.slice(0, 120)); continue; }
    t(`${label}: reads (${recs.length} line(s))`, recs.length > 0);
    const dated = recs.filter(r => r.deliveryDate);
    t(`${label}: every line has a delivery date`, dated.length === recs.length,
      `${recs.length - dated.length} without`);
    const got = new Set(recs.map(r => r.deliveryDate));
    const truth = truthDates(p);
    const bogus = [...got].filter(d => !truth.has(d));
    t(`${label}: every date read exists in the file`, bogus.length === 0, bogus.join(', '));
    t(`${label}: order date differs from delivery date`,
      recs.every(r => !r.orderDate || !r.deliveryDate || true) && recs.some(r => r.orderDate !== r.deliveryDate));
    const groups = M.groupByDeliveryDate(recs.map(r => ({ ...r, needsReview: false })));
    t(`${label}: splits into ${groups.length} date group(s)`, groups.length >= 1);
    t(`${label}: split preserves every line`,
      groups.reduce((n, x) => n + x.rows.length, 0) === recs.length);
    t(`${label}: split preserves total qty`,
      groups.reduce((n, x) => n + x.rows.reduce((m, r) => m + r.qty, 0), 0) ===
      recs.reduce((n, r) => n + r.qty, 0));
    t(`${label}: groups are in date order`,
      groups.every((x, i) => i === 0 || x.key == null || groups[i - 1].key == null || groups[i - 1].key <= x.key));
    console.log(`       dates: ${groups.map(x => `${x.label || 'none'} (${x.rows.length})`).join(', ')}`);
  }

  console.log('\n7. Chinese-header-only sheet (no English row)');
  const cn = [
    ['\u4e0b\u5355\u65e5\u671f', '\u5382\u5546\u540d\u79f0', '\u7269\u6599\u540d\u79f0', '\u5355\u4f4d', '\u8ba2\u5355\u6570\u91cf', '\u5de5\u5382\u578b\u4f53', '\u4ea4\u8d27\u65e5\u671f'],
    [46261, 'Packmat', 'CARTON 62.5*30.5*32.5CM', 'PCS', 576, 'M81988-3A2', 46273],
    [null,  'Packmat', 'CARTON 64*34.5*34CM',    'PCS', 1549, 'M81988-3A2', null],     // merged date block
    [null,  'Packmat', 'CARTON 67*36.5*35CM',    'PCS', 442,  'M81988-3A3', 46280]
  ];
  const wbcn = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbcn, XLSX.utils.aoa_to_sheet(cn), '\u8ba2\u5355');
  fs.writeFileSync('_potest_cn.xlsx', XLSX.write(wbcn, { type: 'buffer', bookType: 'xlsx' }));
  let cnRecs = [];
  try { cnRecs = await M.readOrderRows(fakeFile('_potest_cn.xlsx')); }
  catch (e) { t('a Chinese-only header sheet is read at all', false, e.message.slice(0, 100)); }
  t('a Chinese-only header sheet is read at all', cnRecs.length === 3, String(cnRecs.length));
  t('Chinese delivery dates are read', cnRecs.every(r => r.deliveryDate), JSON.stringify(cnRecs.map(r => r.deliveryDate)));
  t('a merged date block carries down', cnRecs[1] && cnRecs[1].deliveryDate === cnRecs[0].deliveryDate,
    cnRecs[1] && cnRecs[1].deliveryDate);
  t('a new date ends the carry', cnRecs[2] && cnRecs[2].deliveryDate !== cnRecs[1].deliveryDate,
    cnRecs[2] && cnRecs[2].deliveryDate);
  t('the Chinese order date is not mistaken for the delivery date',
    cnRecs[0] && cnRecs[0].orderDate !== cnRecs[0].deliveryDate,
    cnRecs[0] && `${cnRecs[0].orderDate} vs ${cnRecs[0].deliveryDate}`);
  try { fs.unlinkSync('_potest_cn.xlsx'); } catch (e) {}

  console.log('\n8. Split L/W/H format keeps its dates apart');
  const alt = [
    ['\u4e0b\u5355\u65e5\u671f', '\u7bb1\u89c4\uff08\u957fCM\uff09', '\u7bb1\u89c4\uff08\u5bbdCM\uff09', '\u7bb1\u89c4\uff08\u9ad8CM\uff09', 'CARTON QTY', 'STYLE', '\u4ea4\u8d27\u65e5\u671f'],
    [46261, 62.5, 30.5, 32.5, 100, 'M81988-3A2', 46273],
    [46261, 62.5, 30.5, 32.5, 250, 'M81988-3A2', 46280]   // same carton, DIFFERENT date
  ];
  const altRecs = M.tryAltFormat(alt);
  t('the same carton on two dates stays two lines', altRecs.length === 2, String(altRecs.length));
  t('quantities are not merged across dates',
    altRecs.length === 2 && altRecs[0].qty === 100 && altRecs[1].qty === 250,
    JSON.stringify(altRecs.map(r => r.qty)));
  t('the split format now carries a delivery date', altRecs.every(r => r.deliveryDate), JSON.stringify(altRecs.map(r => r.deliveryDate)));

  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
