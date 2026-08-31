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
  t('renaming the LABEL did not change the MATCH', M.itemKind('Master Carton (Punch)') === 'punch' && M.KIND_LABEL.punch === 'Chip Box (Punch)', M.KIND_LABEL.punch);
  t('lower-case "punch" still counts', M.itemKind('master carton punch') === 'punch');
  t('"Perforated Master Carton" is the same thing', M.itemKind('Perforated Master Carton') === 'punch');
  t('"Cross Divider" is cross', M.itemKind('Cross Divider') === 'cross');
  t('a Cross Divider variant does not leak into master', M.itemKind('Cross Divider (5 Ply)') === 'cross');
  t('punch is never also counted as master', M.itemKind('Master Carton (Punch)') !== 'master');
  t('every kind has a label', ['master','punch','cross'].every(k => M.KIND_LABEL[k]));

  console.log('\n2. Column layout is exactly what was asked for');
  const want = ['Priority Delivery Date','Item','Style','PI','Master Carton (qty)','Chip Box (Punch) (qty)','Cross Divider (qty)'];
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
    const punchLabel = out.rows.find(r => r.item && /Chip Box \(Punch\)/.test(r.item));
    t('the Item column names Chip Box (Punch)', !!punchLabel, out.rows.filter(r => r.item).map(r => r.item).join(' | '));
    t('no sheet still says "Perforated Master Carton"', !/Perforated Master Carton/.test(src));
    t('the Item column no longer says "Master Carton (Punch)"',
      !out.rows.some(r => r.item && /Master Carton \(Punch\)/.test(r.item)));

    console.log('\n3b. Rows run earliest delivery date first');
    const key = l => { const m = String(l||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? Date.UTC(+m[3],+m[2]-1,+m[1]) : null; };
    const seq = out.piRows.map(p => key(p.deliveryDate));
    t('delivery dates ascend, never descend',
      seq.every((v, i) => i === 0 || v === null || seq[i-1] === null || seq[i-1] <= v),
      out.piRows.map(p => p.deliveryDate).join(' -> '));
    t('undated PIs sort last, not into the middle',
      (() => { const f = seq.findIndex(v => v === null); return f === -1 || seq.slice(f).every(v => v === null); })());
    t('PI order is kept inside a shared date', (() => {
      for (let i = 1; i < out.piRows.length; i++)
        if (seq[i] === seq[i-1] && Number(out.piRows[i].ref) < Number(out.piRows[i-1].ref)) return false;
      return true;
    })(), out.piRows.map(p => `${p.deliveryDate}#${p.ref}`).join(' '));
    t('both sheets list the PIs in the SAME order',
      out.rows.filter(r => r.piRef).map(r => String(r.piRef)).join(',') === out.piRows.map(p => String(p.ref)).join(','));
    t('no PI was lost or duplicated by the sort',
      new Set(out.piRows.map(p => p.ref)).size === out.piRows.length);

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
      t('the title is one merged cell, not two rows', String(w.getCell('A2').value || '') === want[0], String(w.getCell('A2').value));
      t('title cell wraps (without it Excel drops the break)', w.getCell('A1').alignment && w.getCell('A1').alignment.wrapText === true);
      const hrow = w.getRow(2);
      t('header row is taller than a body row', (hrow.height || 0) > (w.getRow(3).height || 15), `${hrow.height} vs ${w.getRow(3).height}`);
      // The YELLOW belongs to the heading row, not the column-name row.
      let titleYellow = 0;
      for (let c = 1; c <= 7; c++) {
        const f = w.getCell(1, c).fill;
        if (f && f.fgColor && f.fgColor.argb === 'FFFFFF00') titleYellow++;
      }
      t(`the heading row is yellow all the way across (${titleYellow}/7)`, titleYellow === 7);
      let yellow = 0, centred = 0, cells = 0;
      hrow.eachCell({ includeEmpty: true }, c => {
        cells++;
        if (c.fill && c.fill.fgColor && c.fill.fgColor.argb === 'FFFFFF00') yellow++;
        if (c.alignment && c.alignment.horizontal === 'center' && c.alignment.vertical === 'middle') centred++;
      });
      t('the PI/Item/Style row is NOT yellow', yellow === 0, `${yellow} yellow header cell(s)`);
      t('every header cell is centred and middle-aligned', centred === cells && cells === 7);
      t('heading row height is 39', w.getRow(1).height === 39, String(w.getRow(1).height));
      t('column-name row height is 28', hrow.height === 28, String(hrow.height));
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
      const title = String(w.getCell('A1').value || '');
      t('title row', title.startsWith('PI WISE PRIORITY DELIVERY DATE'), title.slice(0, 40));
      t('the date sits on a second line of the same cell', title.includes(`\n(${dateInput.value})`), JSON.stringify(title));
      const PW = ['Priority Delivery Date','DO Date','PI','Master Carton','Chip Box (Punch)','Cross Divider','Status'];
      t('seven columns, quantities before Status',
        PW.every((v, i) => w.getCell(2, i + 1).value === v),
        PW.map((v, i) => w.getCell(2, i + 1).value).join(' | '));
      const nData = out.piRows.length;
      t(`one row per PI (${nData}) plus a Total row`, w.rowCount === 2 + nData + 1, `${w.rowCount} rows`);
      t('Status is Undelivered on every PI row',
        Array.from({ length: nData }, (_, i) => w.getCell(i + 3, 7).value).every(v => v === 'Undelivered'));
      t('column 1 is now the Priority Delivery Date', String(w.getCell(3, 1).value) === out.piRows[0].deliveryDate,
        `${w.getCell(3, 1).value} vs ${out.piRows[0].deliveryDate}`);
      t('DO Date column still holds PI_DATE', String(w.getCell(3, 2).value) === out.piRows[0].doDate);
      t('column 3 is now the PI', String(w.getCell(3, 3).value) === String(out.piRows[0].ref));
      t('the written sheet is in date order too', (() => {
        const k = l => { const m = String(l||'').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? Date.UTC(+m[3],+m[2]-1,+m[1]) : null; };
        const col = Array.from({length: out.piRows.length}, (_, i) => k(w.getCell(i + 3, 1).value));
        return col.every((v, i) => i === 0 || v === null || col[i-1] === null || col[i-1] <= v);
      })());

      // The quantities are the whole point of the new columns: they must equal the per-PI totals
      // in the main report, PI for PI, or the two sheets tell the business head different stories.
      let mism = 0, zeroBlank = 0;
      for (let i = 0; i < nData; i++) {
        const p = out.piRows[i], rr = i + 3;
        if (Number(w.getCell(rr, 4).value) !== p.master) mism++;
        if (Number(w.getCell(rr, 5).value) !== p.punch) mism++;
        if (Number(w.getCell(rr, 6).value) !== p.cross) mism++;
        [4, 5, 6].forEach(c => { if (w.getCell(rr, c).value === null || w.getCell(rr, c).value === undefined) zeroBlank++; });
      }
      t('every PI quantity matches its group total', mism === 0, mism + ' mismatched cell(s)');
      t('a zero quantity is written as 0, never left blank', zeroBlank === 0, zeroBlank + ' blank(s)');
      const lastR = w.rowCount;
      t('Total row sums Master Carton', Number(w.getCell(lastR, 4).value) === out.totalMaster,
        `${w.getCell(lastR, 4).value} vs ${out.totalMaster}`);
      t('Total row sums Master Carton (Punch)', Number(w.getCell(lastR, 5).value) === out.totalPunch);
      t('Total row sums Cross Divider', Number(w.getCell(lastR, 6).value) === out.totalCross);
      t('PI Wise totals equal the main report totals',
        Number(w.getCell(lastR, 4).value) + Number(w.getCell(lastR, 5).value) + Number(w.getCell(lastR, 6).value)
          === out.totalMaster + out.totalPunch + out.totalCross);
      t('heading row is yellow', (() => { let n = 0; for (let c = 1; c <= 7; c++) { const f = w.getCell(1, c).fill; if (f && f.fgColor && f.fgColor.argb === 'FFFFFF00') n++; } return n === 7; })());
      t('the column-name row is not yellow', (() => { let n = 0; for (let c = 1; c <= 7; c++) { const f = w.getCell(2, c).fill; if (f && f.fgColor && f.fgColor.argb === 'FFFFFF00') n++; } return n === 0; })());
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
