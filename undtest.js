/* undtest.js — the permanent guard on the Undelivered Report.
   Runs the ACTUAL shipped functions, lifted out of olisa.html, against the REAL raw export, then
   builds the real workbooks with ExcelJS and reads every cell back.

   Usage:  TZ=Asia/Dhaka node undtest.js olisa.html <Undl_raw.xlsx>
*/
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const HTML = process.argv[2] || 'olisa.html';
const RAW = process.argv[3];
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

// ---------- lift the module ----------
const html = fs.readFileSync(HTML, 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const snap = new Function('return ' + src.match(/function snapToLocalMidnight[\s\S]*?\n\}/)[0].replace('function snapToLocalMidnight', 'function'))();
const st = src.indexOf('(function() { // Undelivered Report Generator');
const en = src.indexOf('(function() { // Ask — style lookup');
let body = src.slice(st, en)
  .replace(/^\(function\(\) \{ \/\/ Undelivered Report Generator/, '')
  .replace(/\}\)\(\);\s*$/, '');
body = body.replace(/const \w+ = document\.getElementById\([^)]*\);\n/g, '')
           .replace(/^\s*\w+\.addEventListener\([\s\S]*?\n\}\);\n/gm, '')
           .replace(/^\s*\['drag[^\n]*\n/gm, '')
           .replace(/^\s*\w+\.addEventListener\([^\n]*\n/gm, '')
           .replace(/^refreshDefaultDate\(\);\n/m, '');
const dateInput = { value: '31st August 2026' };
const statusEl = { textContent: '', className: '' };
const piwiseBtn = { disabled: false, classList: { add() {}, remove() {} } };
let captured = null;
function CapBlob(parts) { captured = parts[0]; }
const M = new Function('XLSX', 'ExcelJS', 'snapToLocalMidnight', 'escHtml', 'dateInput', 'saveBlobAs', 'Blob', 'document', 'statusEl', 'piwiseBtn',
  body + `
  return { readWorkbookRows, buildGroups, buildOutputRows, itemKind, KIND_LABEL, UND_HEADERS,
           defaultDateLabel, extractPiRef, formatDateValue, downloadXlsx, downloadPiWise,
           setState: (r, p, title) => { generatedRows = r; generatedPiRows = p; generatedTitle = title; } };`
)(XLSX, ExcelJS, snap, s => String(s), dateInput, () => 'saved.xlsx', CapBlob, { addEventListener() {} }, statusEl, piwiseBtn);

function fakeFile(p) {
  const b = fs.readFileSync(p);
  return { name: path.basename(p), arrayBuffer: async () => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) };
}

(async () => {
  console.log('1. Punch is its own kind');
  t('"Master Carton" is master', M.itemKind('Master Carton') === 'master');
  t('"Master Carton (Punch)" is punch', M.itemKind('Master Carton (Punch)') === 'punch', M.itemKind('Master Carton (Punch)'));
  t('lower-case "punch" still counts', M.itemKind('master carton punch') === 'punch');
  t('"Perforated Master Carton" is the same thing', M.itemKind('Perforated Master Carton') === 'punch');
  t('"Cross Divider" is cross', M.itemKind('Cross Divider') === 'cross');
  t('a Cross Divider variant does not leak into master', M.itemKind('Cross Divider (5 Ply)') === 'cross');
  t('punch is never also counted as master', M.itemKind('Master Carton (Punch)') !== 'master');
  t('every kind has a label', ['master','punch','cross'].every(k => M.KIND_LABEL[k]));

  console.log('\n2. Column layout is exactly what was asked for');
  const want = ['PI','Item','Style','Priority Delivery Date','Master Carton (qty)','Perforated Master Carton (qty)','Cross Divider (qty)'];
  t('seven columns, named and ordered', JSON.stringify(M.UND_HEADERS) === JSON.stringify(want), JSON.stringify(M.UND_HEADERS));
  ['Buyer','Priority','Previous PI'].forEach(c =>
    t(`"${c}" column is gone`, !M.UND_HEADERS.includes(c)));

  if (!RAW) { console.log('\n(skipped the real-file checks — no raw export supplied)'); }
  else {
    console.log('\n3. Real raw export');
    const recs = await M.readWorkbookRows(fakeFile(RAW));
    t(`reads (${recs.length} undelivered line(s))`, recs.length > 0);
    t('PI_DATE is carried through for the PI Wise sheet', recs.every(r => r.piDate), 'some blank');
    const groups = M.buildGroups(recs);
    const out = M.buildOutputRows(groups);

    // Independent oracle straight off the sheet — a different code path than the one under test.
    const wb = XLSX.read(fs.readFileSync(RAW), { type: 'buffer', cellDates: true });
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames.find(n => /rpt/i.test(n)) || wb.SheetNames[0]], { header: 1, defval: null, raw: true });
    const h = aoa[0].map(String);
    const ci = n => h.indexOf(n);
    let oMaster = 0, oPunch = 0, oCross = 0;
    const oPis = new Set();
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i]; if (!row || !row[ci('PI_NO')]) continue;
      const q = Number(row[ci('U_QTY')]) || 0; if (q <= 0) continue;
      const d = String(row[ci('ITEM_DESCRIPTION')] || '');
      oPis.add(String(row[ci('PI_NO')]).trim());
      if (/cross\s*divider/i.test(d)) oCross += q;
      else if (/punch|perforat/i.test(d)) oPunch += q;
      else oMaster += q;
    }
    t(`master total matches the sheet (${out.totalMaster})`, out.totalMaster === oMaster, `${out.totalMaster} vs ${oMaster}`);
    t(`punch total matches the sheet (${out.totalPunch})`, out.totalPunch === oPunch, `${out.totalPunch} vs ${oPunch}`);
    t(`cross total matches the sheet (${out.totalCross})`, out.totalCross === oCross, `${out.totalCross} vs ${oCross}`);
    t('grand total equals the sum of every undelivered piece',
      out.totalMaster + out.totalPunch + out.totalCross === oMaster + oPunch + oCross);
    t('punch is NOT silently added to master', out.totalMaster !== oMaster + oPunch || oPunch === 0);
    t(`one PI Wise line per PI (${out.piRows.length})`, out.piRows.length === oPis.size, `${out.piRows.length} vs ${oPis.size}`);
    t('every PI Wise line says Undelivered', out.piRows.every(p => p.status === 'Undelivered'));
    t('every PI Wise line has a DO date', out.piRows.every(p => p.doDate));
    t('every PI Wise line has a priority delivery date', out.piRows.every(p => p.deliveryDate));
    t('DO date is not the same field as the delivery date', out.piRows.some(p => p.doDate !== p.deliveryDate));
    t('dates stay in dd/mm/yyyy as the raw file wrote them',
      out.piRows.every(p => /^\d{2}\/\d{2}\/\d{4}$/.test(p.doDate) && /^\d{2}\/\d{2}\/\d{4}$/.test(p.deliveryDate)),
      JSON.stringify(out.piRows[0]));
    const punchRow = out.rows.find(r => r.punch);
    t('a punch line carries its qty in the punch column', !!punchRow && punchRow.punch > 0);
    t('a punch line leaves the master column empty', !!punchRow && !punchRow.master);
    const punchLabel = out.rows.find(r => r.item && /Punch/.test(r.item));
    t('the Item column names Master Carton (Punch)', !!punchLabel, out.rows.filter(r => r.item).map(r => r.item).join(' | '));

    console.log('\n4. The real workbook, read back cell by cell');
    M.setState(out.rows, out.piRows, `OLISA's Undelivered Qty & Delivery Schedule\n(${dateInput.value})`);
    // The shipped downloadXlsx hands its bytes to Blob; the harness's Blob keeps them.
    captured = null;
    await M.downloadXlsx();
    const buf = captured;
    t('a workbook was produced', !!buf);
    if (buf) {
      const back = new ExcelJS.Workbook();
      await back.xlsx.load(buf);
      const w = back.getWorksheet('Undelivered Report');
      const title = String(w.getCell('A1').value || '');
      t('title reads OLISA\'s Undelivered Qty & Delivery Schedule', title.startsWith("OLISA's Undelivered Qty & Delivery Schedule"), title.slice(0, 60));
      t('the date sits on a SECOND LINE of the same cell', title.includes('\n(') && title.trim().endsWith(')'), JSON.stringify(title));
      t('the title is one merged cell, not two rows', String(w.getCell('A2').value || '') === 'PI');
      t('title row is given height for two lines', (w.getRow(1).height || 0) >= 30, String(w.getRow(1).height));
      t('title cell wraps (without it Excel drops the break)', w.getCell('A1').alignment && w.getCell('A1').alignment.wrapText === true);
      const hrow = w.getRow(2);
      t('header row is taller than a body row', (hrow.height || 0) > (w.getRow(3).height || 15), `${hrow.height} vs ${w.getRow(3).height}`);
      let yellow = 0, centred = 0, cells = 0;
      hrow.eachCell({ includeEmpty: true }, c => {
        cells++;
        if (c.fill && c.fill.fgColor && c.fill.fgColor.argb === 'FFFFFF00') yellow++;
        if (c.alignment && c.alignment.horizontal === 'center' && c.alignment.vertical === 'middle') centred++;
      });
      t(`every header cell is yellow (${yellow}/${cells})`, yellow === cells && cells === 7);
      t('every header cell is centred and middle-aligned', centred === cells);
      t('headers are the seven asked for', want.every((v, i) => w.getCell(2, i + 1).value === v),
        want.map((v, i) => w.getCell(2, i + 1).value).join(' | '));
      let offCentre = 0, body = 0;
      for (let rr = 3; rr <= w.rowCount; rr++) {
        w.getRow(rr).eachCell({ includeEmpty: true }, c => {
          body++;
          if (!c.alignment || c.alignment.horizontal !== 'center' || c.alignment.vertical !== 'middle') offCentre++;
        });
      }
      t(`every body cell is centred and middle-aligned (${body} cells)`, offCentre === 0, String(offCentre) + ' off');
      // totals row
      const last = w.rowCount;
      t('total row still totals master', Number(w.getCell(last, 5).value) === out.totalMaster);
      t('total row totals the punch column', Number(w.getCell(last, 6).value) === out.totalPunch, String(w.getCell(last, 6).value));
      t('total row totals cross', Number(w.getCell(last, 7).value) === (out.totalCross || 0));
      t('no eighth column was left behind', w.columnCount <= 7, String(w.columnCount));
    }

    console.log('\n5. The PI Wise workbook');
    captured = null;
    await M.downloadPiWise();
    t('a PI Wise workbook was produced', !!captured);
    if (captured) {
      const back = new ExcelJS.Workbook();
      await back.xlsx.load(captured);
      const w = back.worksheets[0];
      t('title row', String(w.getCell('A1').value) === 'PI WISE PRIORITY DELIVERY DATE', String(w.getCell('A1').value));
      t('date row under the title', String(w.getCell('A2').value) === `(${dateInput.value})`, String(w.getCell('A2').value));
      t('headers PI / DO Date / Priority Delivery Date / Status',
        ['PI','DO Date','Priority Delivery Date','Status'].every((v, i) => w.getCell(3, i + 1).value === v));
      t(`one row per PI (${w.rowCount - 3})`, w.rowCount - 3 === out.piRows.length, `${w.rowCount - 3} vs ${out.piRows.length}`);
      t('Status is Undelivered on every row',
        Array.from({ length: w.rowCount - 3 }, (_, i) => w.getCell(i + 4, 4).value).every(v => v === 'Undelivered'));
      t('DO Date column holds PI_DATE', String(w.getCell(4, 2).value) === out.piRows[0].doDate,
        `${w.getCell(4, 2).value} vs ${out.piRows[0].doDate}`);
      t('Priority Delivery Date column holds PI_DLV_DATE', String(w.getCell(4, 3).value) === out.piRows[0].deliveryDate);
      let off = 0, n = 0;
      for (let rr = 1; rr <= w.rowCount; rr++) w.getRow(rr).eachCell({ includeEmpty: true }, c => {
        n++; if (!c.alignment || c.alignment.horizontal !== 'center') off++;
      });
      t(`every cell centred (${n} cells)`, off === 0, String(off));
    }
  }

  console.log('\n6. The after-midnight date bug');
  t('the default date is recomputed, not frozen at page load', /function refreshDefaultDate/.test(src));
  t('it refreshes when the app comes back into view', /visibilityState === 'visible'\) refreshDefaultDate/.test(src));
  t('it refreshes again the moment Generate is pressed', /refreshDefaultDate\(\);\s*\/\/ a report started after midnight/.test(src));
  t('a hand-typed date is left alone', /if \(dateTouched\) return;/.test(src));
  t('today reads as today', M.defaultDateLabel() === (() => {
    const d = new Date(), day = d.getDate();
    const sfx = (day % 10 === 1 && day !== 11) ? 'st' : (day % 10 === 2 && day !== 12) ? 'nd' : (day % 10 === 3 && day !== 13) ? 'rd' : 'th';
    const mo = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    return `${day}${sfx} ${mo[d.getMonth()]} ${d.getFullYear()}`;
  })());

  console.log('\n' + (fail ? 'FAILED' : 'PASSED') + ` — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
