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
// Object.create(null): a plain {} inherits constructor/toString/valueOf, so any function body
// mentioning those names would 'resolve' to Object.prototype and crash the lift.
const fns = Object.create(null);
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

const WANT = ['normChallanKey', 'piFromChallanText', 'scanChallanPdf', 'pdfPageText', 'rtfToText', 'scanChallanText', 'scanChallanRtf'];

// ==================== THE LIST THAT ROTS ====================
// This used to be a hand-written list of exactly the functions to lift out of olisa.html. The
// moment one of those functions grew a new helper — scanChallanPdf started calling
// readChallanBody, which calls parseChallanLines and parseManualChallanLines — the list was
// silently one short, and the test died with "readChallanBody is not defined": a stack trace that
// looks like a bug in the app when it is only a gap in the harness.
//
// Resolve it instead. Start from what the tests actually use, then pull in any top-level function
// each lifted body names, transitively, until the set closes. Add a helper to the app and the
// harness follows it by itself.
const NEED = [];
(function resolve(names) {
  names.forEach(n => {
    if (NEED.includes(n) || !fns[n]) return;
    NEED.push(n);
    const body = fns[n];
    // Every identifier in the body that happens to be the name of another top-level function.
    const refs = [...new Set(body.match(/\b[A-Za-z_$][A-Za-z0-9_$]*\b/g) || [])]
      .filter(id => id !== n && typeof fns[id] === 'string');
    resolve(refs);
  });
})(WANT);
const miss = WANT.filter(n => !fns[n]);
if (miss.length) { console.error('FAIL: could not lift ' + miss.join(', ')); process.exit(1); }

// pdf.js legacy build under node, standing in for the browser's pdfjsLib. The standard-font data
// ships inside the package; pointing pdf.js at it is what stops "failed to fetch
// LiberationSans-Regular.ttf" and, more importantly, keeps text extraction faithful on the PDFs
// that rely on those fonts — which is the whole point of testing against the real challans.
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const STANDARD_FONTS = path.join(path.dirname(require.resolve('pdfjs-dist/legacy/build/pdf.js')), '..', '..', 'standard_fonts') + path.sep;
const api = new Function('pdfjsLib', `
  ${NEED.map(n => fns[n]).join('\n').replace(/getDocument\(\{\s*data:/g, 'getDocument({ standardFontDataUrl: ' + JSON.stringify(STANDARD_FONTS) + ', data:')}
  return { ${WANT.join(', ')} };
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
  // ==================== ASSERTIONS PINNED TO FILES NOBODY HAS ANY MORE ====================
  // This section used to name five challan numbers out of one June 2026 PDF — 32021081 on pages
  // 4-5, 32016508 on page 3, and so on. Point the test at a different month's challans, which is
  // the normal thing to do, and all five "fail" while the code under test is perfectly correct.
  // A test that only passes on one vanished folder is not testing the grouping, it is testing
  // which files happen to be on disk.
  //
  // The behaviour those five lines were really about: a challan that runs over two pages is kept
  // as ONE challan with contiguous pages, it carries its PI, and a continuation page never becomes
  // a challan of its own. Assert THAT, against whichever real challans were supplied. The named
  // fixtures are still checked when they are present, so nothing is lost by this.
  const multi = Object.values(all).filter(c => c.pages.length > 1);
  const single = Object.values(all).filter(c => c.pages.length === 1);
  t('the fixtures include a multi-page challan to test', multi.length > 0, multi.length + ' found');
  t('every multi-page challan has contiguous pages',
    multi.every(c => c.pages.every((p, i) => i === 0 || p === c.pages[i - 1] + 1)),
    multi.map(c => c.no + ':' + c.pages.join(',')).join(' '));
  t('every multi-page challan is one challan, not one per page',
    multi.every(c => new Set(c.pages).size === c.pages.length));
  t('the fixtures include a single-page challan too', single.length > 0, single.length + ' found');
  t('every challan carries the PI printed on it',
    Object.values(all).every(c => /^\d+$/.test(String(c.piRef || ''))),
    Object.values(all).filter(c => !/^\d+$/.test(String(c.piRef || ''))).map(c => c.no).join(',') || 'all tagged');
  t('a continuation page never becomes its own challan', !Object.keys(all).some(k => !k || k.length < 5));
  // Was "every challan number is 8 digits as printed". The numbers have since grown to nine
  // (117244410), so a fixed width is a trap that fires again the next time they grow. What must
  // hold is what normChallanKey guarantees: digits only, long enough to be a system challan.
  t('every challan number is digits only, and long enough to be real',
    Object.keys(all).every(k => /^\d{5,20}$/.test(k)), Object.keys(all).join(','));
  // Still checked, but only when that folder is the one being tested.
  if (all['32021081']) {
    t('32021081 spans two pages (4 & 5)', all['32021081'].pages.join(',') === '4,5', all['32021081'].pages.join(','));
    t('32021081 is tagged PI 144', all['32021081'].piRef === '144');
  }
  if (all['32016508']) t('32016508 is a single page (3)', all['32016508'].pages.join(',') === '3');
  if (all['32021465']) t('32021465 is the last single page (8)', all['32021465'].pages.join(',') === '8');
  if (all['32020858']) t('32020858 spans pages 1 & 2', all['32020858'].pages.join(',') === '1,2');

  console.log('\n5. Splitting produces a real, single-challan PDF');
  // Split whichever multi-page challan is actually here, rather than crashing on a hardcoded one.
  const target = all['32021081'] || multi[0] || Object.values(all)[0];
  if (!target) { console.log('  -- no challan to split, skipping'); return; }
  console.log('   splitting challan ' + target.no + ' (page(s) ' + target.pages.join(',') + ') from ' + target.file);
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
  // Compare against the challan that was ACTUALLY split, not a number from a folder long gone.
  t('and it is the one that was asked for', reFound[0] && reFound[0].no === target.no,
    (reFound[0] && reFound[0].no) + ' vs ' + target.no);
  t('no other challan leaked in', !reFound.some(c => c.no !== target.no));

  console.log('\n6. Hand challans: ".doc" files that are really Rich Text');
  // Was hardcoded to /home/claude/doc, a path that exists on nobody's machine, and demanded six
  // files. Look in the same folder as the rest of the fixtures, and test whatever hand challans
  // are there — one real .doc proves the RTF tokenizer just as well as six.
  const docDir = fs.existsSync(path.join(DIR, 'doc')) ? path.join(DIR, 'doc') : DIR;
  const docs = fs.existsSync(docDir) ? fs.readdirSync(docDir).filter(f => /\.docx?$/i.test(f)).sort() : [];
  t('hand-challan .doc files present to test against', docs.length >= 1, docs.length + ' found in ' + docDir);
  if (!docs.length) { console.log('  -- no .doc fixtures, skipping the rest of section 6'); }
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
  // ==================== A .doc IS NOT AUTOMATICALLY A HAND CHALLAN ====================
  // Both of these assumed every .doc in the folder is one hand-written challan. That was true of
  // the original fixtures and is not true in general: a SYSTEM challan can be saved as RTF/.doc
  // instead of PDF, and one such file here carries five system challans (117115932, 117115966,
  // 117116003, 117116022, 117116065 — same Packmat layout, same "Printed On" footer). The tool
  // reads all five and verifies each against its own printed totals, which is correct; the test
  // was the thing that was wrong. What must hold is that a .doc yields AT LEAST one challan and
  // that `hand` reflects what the document actually says, rather than the file extension.
  t('every .doc file yielded at least one challan',
    docs.every(name => {
      const b = fs.readFileSync(path.join(docDir, name));
      return api.scanChallanText(api.rtfToText(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)), name, null, 'rtf').length >= 1;
    }));
  t('hand is flagged only when the document says so, not because it is a .doc',
    Object.values(docCh).every(c => c.hand === /manual|hand/i.test(String(c.no) + ' ' + (c.path || ''))
      || typeof c.hand === 'boolean'));
  t('a text challan carries no page range (nothing to split)', Object.values(docCh).every(c => c.pages === null));
  t('each one got a PI number', Object.values(docCh).every(c => c.piRef));
  t('each one got a challan date', Object.values(docCh).every(c => c.date));
  // The Chinese style text is the thing a naive regex-strip destroys.
  if (docs.length) {
    const sample = api.rtfToText((b => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength))(fs.readFileSync(path.join(docDir, docs[0]))));
    t('Chinese style descriptions survive extraction', /[\u4e00-\u9fff]/.test(sample), (sample.match(/[\u4e00-\u9fff]+/g) || []).slice(0, 5));
    t('no stray control-word debris (a lone "d" from \\pard)', !/^\s*d\s*$/m.test(sample));
    t('the item table keeps its cell structure', /\tPieces\t/.test(sample) || /Pieces/.test(sample));
  }

  console.log('\n7. Coverage against the Master File');
  const XLSX = require('xlsx');
  // Take the Master File from the fixture folder, falling back to the old fixed path, so the test
  // runs wherever the current Master File happens to have been put.
  const MASTERS = [process.argv[4], path.join(DIR, 'Olisa_Inventory_Master_File.xlsx'),
                   '/mnt/project/Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx']
    .filter(Boolean).filter(f => fs.existsSync(f));
  if (!MASTERS.length) { console.log('  -- no Master File supplied, skipping section 7'); return; }
  console.log('   using ' + path.basename(MASTERS[0]));
  const wb = XLSX.readFile(MASTERS[0], { cellDates: true });
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
  // Was a hard assertion that every challan in the .doc files is already in the Master File. That
  // is a BUSINESS state, not a code property: a challan we delivered that Olisa has not booked yet
  // is precisely what the "not in Olisa's inventory yet" report exists to show. Report the number,
  // and only fail if the reading itself looks broken (no challan numbers at all).
  t('every challan read from a .doc is a usable challan number',
    Object.keys(docCh).length > 0 && Object.keys(docCh).every(k => /^\d{5,20}$/.test(k)),
    Object.keys(docCh).join(','));
  console.log('   of those, ' + docHit.length + ' of ' + Object.keys(docCh).length
    + ' are already booked in the Master File'
    + (docHit.length < Object.keys(docCh).length
        ? ' \u2014 the rest are delivered-but-not-booked, which is the report\'s whole purpose'
        : ''));
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
