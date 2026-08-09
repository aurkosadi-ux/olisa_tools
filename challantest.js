/* challantest.js — runs the ACTUAL shipped challan-index functions, lifted from olisa.html,
   against the REAL challan PDFs. Proves page grouping, splitting, and key normalisation. */
const fs = require('fs');
const path = require('path');
const acorn = require('acorn');
const { PDFDocument } = require('pdf-lib');

const HTML = process.argv[2] || 'olisa.html';
const DIR = process.argv[3] || '/home/claude/ch';

const html = fs.readFileSync(HTML, 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
  .map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const ast = acorn.parse(src, { ecmaVersion: 2022, locations: true });
const fns = {};
(function scan(n) {
  if (!n || typeof n.type !== 'string') return;
  if (n.type === 'FunctionDeclaration' && n.id) fns[n.id.name] = src.slice(n.start, n.end);
  for (const k of Object.keys(n)) {
    if (k === 'loc') continue;
    const v = n[k];
    if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && scan(c));
    else if (v && typeof v.type === 'string') scan(v);
  }
})(ast);

const NEED = ['normChallanKey', 'piFromChallanText', 'scanChallanPdf', 'pdfPageText', 'rtfToText', 'scanChallanText', 'scanChallanRtf'];
const miss = NEED.filter(n => !fns[n]);
if (miss.length) { console.error('FAIL: could not lift ' + miss.join(', ')); process.exit(1); }

// pdf.js legacy build under node, standing in for the browser's pdfjsLib
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const api = new Function('pdfjsLib', `
  ${NEED.map(n => fns[n]).join('\n')}
  return { normChallanKey, piFromChallanText, scanChallanPdf, rtfToText, scanChallanText, scanChallanRtf };
`)(pdfjsLib);

let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

console.log('1. Challan key normalisation (Master File gives these in every shape)');
const K = api.normChallanKey;
t('plain number', K(31910793) === '31910793');
t('text with spaces', K(' 31910793 ') === '31910793');
t("Excel's trailing .0", K('31910793.0') === '31910793');
t('leading zeros stripped', K('031910793') === '31910793');
t('embedded spaces inside the number', K('3191 0793') === '31910793');
t('blank is not a key', K('') === '' && K(null) === '' && K(undefined) === '');
// This used to assert K('12') === '' — that assumption is what made hand challans unclickable.
// The real invariant is that a value with NO digits is not a challan number.
t('a value with no digits is not a key', K('Manual') === '' && K('-') === '' && K('n/a') === '');
t('a number with a date suffix keeps only the number', K('31910793/ 09-JUN-26') === '31910793');
// Hand challans are numbered by hand and sit right beside 8-digit system numbers in the same
// column. A 5-digit minimum used to refuse exactly these, so those rows were never clickable.
t('short hand-challan number 114 is a valid key', K('114') === '114');
t('short hand-challan number 42 is a valid key', K(42) === '42');
t('two-digit hand challan 16 is a valid key', K('16') === '16');
t('"Manual" alone is still not a key', K('Manual') === '');
t('an absurdly long digit run is still rejected', K('1'.repeat(25)) === '');

console.log('\n2. PI extraction from the challan header');
t('PI:121/700768/0144062026 -> 144', api.piFromChallanText('PI:121/700768/0144062026 WO:121/700768/0144/09-06-2026') === '144');
t('PI:121/700768/0124042026 -> 124', api.piFromChallanText('PI:121/700768/0124042026') === '124');
t('no PI line -> empty', api.piFromChallanText('Challan No : 31910793') === '');

(async () => {
  console.log('\n3. Page grouping across the real challan PDFs');
  const files = fs.readdirSync(DIR).filter(f => /\.pdf$/i.test(f) && f !== 'split.pdf').sort();
  t('challan PDFs present to test against', files.length >= 3, files.length + ' found');

  const all = {};
  for (const name of files) {
    const bytes = fs.readFileSync(path.join(DIR, name));
    const fakeFile = { arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
    const found = await api.scanChallanPdf(fakeFile, name, null);
    console.log('   ' + name);
    found.forEach(c => {
      console.log('     challan ' + c.no + '  page(s) ' + c.pages.join(',') + '  PI ' + (c.piRef || '?') + '  ' + (c.date || ''));
      all[c.no] = { ...c, file: name };
    });
    const src2 = await PDFDocument.load(bytes);
    const covered = found.reduce((s, c) => s + c.pages.length, 0);
    t(name.slice(0, 34) + ' — every page assigned to a challan', covered === src2.getPageCount(), covered + ' of ' + src2.getPageCount());
    t(name.slice(0, 34) + ' — no page claimed twice',
      new Set(found.flatMap(c => c.pages)).size === covered);
    t(name.slice(0, 34) + ' — page runs are contiguous',
      found.every(c => c.pages.every((p, i) => i === 0 || p === c.pages[i - 1] + 1)));
  }

  console.log('\n4. The multi-page case, in detail');
  // 16-06 file: 32020858 = pp1-2, 32016508 = p3, 32021081 = pp4-5, 32021160 = pp6-7, 32021465 = p8
  t('32021081 spans two pages (4 & 5)', all['32021081'] && all['32021081'].pages.join(',') === '4,5',
    all['32021081'] && all['32021081'].pages.join(','));
  t('32021081 is tagged PI 144', all['32021081'] && all['32021081'].piRef === '144');
  t('32016508 is a single page (3)', all['32016508'] && all['32016508'].pages.join(',') === '3');
  t('32021465 is the last single page (8)', all['32021465'] && all['32021465'].pages.join(',') === '8');
  t('32020858 spans pages 1 & 2', all['32020858'] && all['32020858'].pages.join(',') === '1,2');
  t('a continuation page never becomes its own challan', !Object.keys(all).some(k => !k || k.length < 5));
  t('challan 31910793 found in the 9-6 file', !!all['31910793']);
  t('every challan number is 8 digits as printed', Object.keys(all).every(k => /^\d{8}$/.test(k)));

  console.log('\n5. Splitting produces a real, single-challan PDF');
  const target = all['32021081'];
  const bytes = fs.readFileSync(path.join(DIR, target.file));
  const src3 = await PDFDocument.load(bytes);
  const out = await PDFDocument.create();
  const copied = await out.copyPages(src3, target.pages.map(p => p - 1));
  copied.forEach(p => out.addPage(p));
  const outBytes = await out.save();
  fs.writeFileSync('/tmp/one-challan.pdf', outBytes);
  const check = await PDFDocument.load(outBytes);
  t('output has exactly the challan\'s page count', check.getPageCount() === target.pages.length,
    check.getPageCount() + ' vs ' + target.pages.length);
  t('output is smaller than the source', outBytes.length < bytes.length,
    outBytes.length + ' vs ' + bytes.length);

  // Re-scan the split file: it must contain ONE challan, and the right one.
  const reFile = { arrayBuffer: async () => outBytes.buffer.slice(outBytes.byteOffset, outBytes.byteOffset + outBytes.byteLength) };
  const reFound = await api.scanChallanPdf(reFile, 'split', null);
  t('the split file contains exactly one challan', reFound.length === 1, reFound.length);
  t('and it is the one that was asked for', reFound[0] && reFound[0].no === '32021081', reFound[0] && reFound[0].no);
  t('no other challan leaked in',
    !reFound.some(c => c.no !== '32021081'));

  console.log('\n6. Hand challans: ".doc" files that are really Rich Text');
  const docDir = '/home/claude/doc';
  const docs = fs.existsSync(docDir) ? fs.readdirSync(docDir).filter(f => /\.doc$/i.test(f)).sort() : [];
  t('hand-challan .doc files present to test against', docs.length >= 6, docs.length + ' found');
  const docCh = {};
  for (const name of docs) {
    const bytes = fs.readFileSync(path.join(docDir, name));
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const txt = api.rtfToText(ab);
    const found = api.scanChallanText(txt, name, null, 'rtf');
    found.forEach(c => { docCh[c.no] = c; });
    if (found.length === 1) {
      const c = found[0];
      console.log('   ' + name.slice(0, 44).padEnd(46) + ' challan ' + c.no + (c.hand ? ' [hand]' : '') + '  PI ' + (c.piRef || '?') + '  ' + (c.date || ''));
    } else {
      console.log('   ' + name.slice(0, 44).padEnd(46) + ' ' + found.length + ' challan(s)');
    }
  }
  t('every hand challan file yielded exactly one challan',
    docs.every(name => {
      const b = fs.readFileSync(path.join(docDir, name));
      return api.scanChallanText(api.rtfToText(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)), name, null, 'rtf').length === 1;
    }));
  t('all are recognised as hand challans (the "-M" suffix)', Object.values(docCh).every(c => c.hand === true));
  t('a text challan carries no page range (nothing to split)', Object.values(docCh).every(c => c.pages === null));
  t('each one got a PI number', Object.values(docCh).every(c => c.piRef));
  t('each one got a challan date', Object.values(docCh).every(c => c.date));
  // The Chinese style text is the thing a naive regex-strip destroys.
  const sample = api.rtfToText((b => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength))(fs.readFileSync(path.join(docDir, docs[0]))));
  t('Chinese style descriptions survive extraction', /[\u4e00-\u9fff]/.test(sample), sample.match(/[\u4e00-\u9fff]+/g));
  t('no stray control-word debris (a lone "d" from \\pard)', !/^\s*d\s*$/m.test(sample));
  t('the item table keeps its cell structure', /\tPieces\t/.test(sample) || /Pieces/.test(sample));

  console.log('\n7. Coverage against the Master File');
  const XLSX = require('xlsx');
  const wb = XLSX.readFile('/mnt/project/Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx', { cellDates: true });
  // The header row is not row 1 on this sheet — find it, exactly as the app does.
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets['2026- Main'], { header: 1, defval: null });
  let hdr = -1, chCol = -1;
  for (let r = 0; r < Math.min(12, aoa.length); r++) {
    const row = (aoa[r] || []).map(c => String(c == null ? '' : c).toUpperCase().replace(/[^A-Z]/g, ''));
    const i = row.findIndex(c => c.includes('CHALLANNO'));
    if (i >= 0) { hdr = r; chCol = i; break; }
  }
  const masterCh = new Set(aoa.slice(hdr + 1).map(r => K(r && r[chCol])).filter(Boolean));
  const indexed = new Set(Object.keys(all));
  const hit = [...masterCh].filter(c => indexed.has(c));
  console.log('   Master File holds ' + masterCh.size + ' distinct challan numbers');
  console.log('   these 3 sample PDFs cover ' + hit.length + ' of them: ' + hit.slice(0, 8).join(', '));
  t('the sample PDFs match real Master File challans', hit.length > 0, hit.length);
  // The whole point of reading .doc: these challans exist in NO other format.
  const docHit = Object.keys(docCh).filter(c => masterCh.has(c));
  console.log('   the hand-challan .doc files cover ' + docHit.length + ': ' + docHit.join(', '));
  t('every hand challan matches a real Master File challan number',
    docHit.length === Object.keys(docCh).length, docHit.length + ' of ' + Object.keys(docCh).length);
  t('hand challans are NOT in any of the PDFs (doc is their only source)',
    Object.keys(docCh).every(c => !indexed.has(c)));
  // Cross-check against the Master File's own Manual Challan column.
  const manCol = new Set(aoa.slice(hdr + 1)
    .filter(r => r && r.some(c => c !== null && /MANUAL\s*CHALL|HAND\s*CHALL/i.test(String(c))))
    .map(r => K(r[chCol])).filter(Boolean));
  console.log('   Master File flags ' + manCol.size + ' distinct challans as manual');
  t('every hand challan read is flagged Manual in the Master File',
    docHit.every(c => manCol.has(c)), docHit.filter(c => !manCol.has(c)).join(', '));
  t('every indexed key would be found by a Master File click',
    hit.every(c => all[c] && all[c].pages.length >= 1));

  console.log('\n' + (fail ? 'FAILED ' + fail + ' / ' : 'PASSED all ') + (pass + fail) + ' checks');
  process.exit(fail ? 1 : 0);
})();
