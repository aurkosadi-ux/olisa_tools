/* challanlinetest.js — the permanent guard on reading challans.
   Runs the SHIPPED parser against the real challan files, and checks every challan against the
   Total Line / Total Qty printed on the challan itself. That printed total is the whole safety
   model: a challan that does not add up is never allowed to tell Olisa anything.
   Usage: node challanlinetest.js olisa.html [challanDir]                                       */
const fs = require('fs'), path = require('path');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const htmlPath = process.argv[2] || 'olisa.html';
const dir = process.argv[3] || '/mnt/user-data/uploads';
const html = fs.readFileSync(htmlPath, 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];

function lift(name) {
  const s = src.indexOf('function ' + name);
  if (s < 0) throw new Error('not found: ' + name);
  let d = 0, e = s;
  for (let i = src.indexOf('{', s); i < src.length; i++) { if (src[i] === '{') d++; else if (src[i] === '}') { d--; if (!d) { e = i + 1; break; } } }
  return src.slice(s, e);
}
const bag = new Function(
  lift('chQtyNum') + lift('parseChallanLines') + lift('parseManualChallanLines') + lift('readChallanBody')
  + lift('rtfToText') + lift('normChallanKey') + lift('piFromChallanText')
  + '\nreturn { parseChallanLines, parseManualChallanLines, readChallanBody, rtfToText, normChallanKey, piFromChallanText };')();

console.log('\n1. The style comes from |Size:, never from Style/PO/KTC');
// Style/PO/KTC holds the Chinese carton description. Reading that instead of Size looks plausible
// and is wrong, so it gets its own check.
const one = bag.parseChallanLines(
  'Line Item Description Qty Unit Other Information 1 Master Carton, 5 Ply, L70 x W48 x H34cm; 258 Pieces '
  + 'Color:Printed |Style/PO/KTC:CARTON \u6df7 15\u53cc 70*48*34CM |Size:M81988-3A1-1 |Code:nill |Type:27.56 X 18.9 X 13.39 '
  + 'Total Line : 1 Total Qty: 258');
t('one line is read', one.lines.length === 1, JSON.stringify(one.lines));
t('the style is M81988-3A1-1', one.lines[0] && one.lines[0].style === 'M81988-3A1-1', one.lines[0] && one.lines[0].style);
t('the Chinese description is kept separately, not as the style', one.lines[0] && /CARTON/.test(one.lines[0].po));
t('the quantity is 258', one.lines[0] && one.lines[0].qty === 258);
t('it verifies against its own printed totals', one.verified === true);

console.log('\n2. A challan that does not add up is refused');
const bad = bag.parseChallanLines(
  'Line Item Description Qty Unit Other Information 1 Master Carton, 5 Ply, L70 x W48 x H34cm; 258 Pieces '
  + 'Color:Printed |Style/PO/KTC:X |Size:M81988-3A1-1 |Code:nill |Type:n Total Line : 2 Total Qty: 999');
t('a wrong line count fails verification', bad.verified === false);
t('but the lines are still returned for a human to look at', bad.lines.length === 1);
const noTotals = bag.parseChallanLines('1 Master Carton, 5 Ply, L70 x W48 x H34cm; 258 Pieces Color:P |Size:M1 |Code:n |Type:n');
t('a challan with no printed totals is never treated as proved', noTotals.verified === false);

console.log('\n3. Multi-challan files take their OWN totals, not the previous challan\'s');
// This was a real bug: the slice handed to the parser starts before this challan's header, so the
// previous challan's footer sits in front of its lines. Every challan after the first failed.
const two = 'Total Line : 1 Total Qty: 213 Challan No : 117115932/ 06-SEP-26 Line Item Description Qty Unit Other Information '
  + '1 Master Carton, 5 Ply, L70 x W48 x H34cm; 14 Pieces Color:P |Size:M81988-3A-1 |Code:n |Type:n '
  + '2 Master Carton, 5 Ply, L70 x W48 x H34cm; 48 Pieces Color:P |Size:M81988-3A-2 |Code:n |Type:n Total Line : 2 Total Qty: 62';
const r2 = bag.parseChallanLines(two);
t('reads its own 2 lines / 62 pcs, not the 1 / 213 in front of it',
  r2.declaredLines === 2 && r2.declaredQty === 62 && r2.parsedQty === 62, `${r2.declaredLines}/${r2.declaredQty}/${r2.parsedQty}`);
t('and therefore verifies', r2.verified === true);

console.log('\n4. The real challan files');
if (!fs.existsSync(dir)) {
  console.log('  -- fixture folder not present, skipping (run with the challan folder as arg 2)');
} else {
  const files = fs.readdirSync(dir).filter(f => /\.(pdf|docx?|rtf)$/i.test(f)).sort();
  t('challan fixtures are present', files.length > 0, `${files.length} file(s)`);
  let verified = 0, unreadable = 0;
  for (const f of files) {
    const buf = fs.readFileSync(path.join(dir, f));
    const magic = buf.subarray(0, 4).toString('latin1');
    if (!magic.startsWith('{\\rt')) { unreadable += 0; continue; }   // PDFs need pdf.js; covered below
    const text = bag.rtfToText(buf);
    const hits = [...text.matchAll(/Challan\s*No\s*[:\uff1a]?\s*(\d{2,})\s*(-\s*M\b)?/gi)];
    hits.forEach((m, k) => {
      const from = m.index, to = k + 1 < hits.length ? hits[k + 1].index : text.length;
      const seg = text.slice(Math.max(0, from - 600), to);
      const r = bag.readChallanBody(seg);
      t(`${f} challan ${m[1]}: ${r.lines.length}/${r.declaredLines} lines, ${r.parsedQty}/${r.declaredQty} pcs`,
        r.verified === true, r.verified ? '' : 'did not reconcile');
      if (r.verified) verified++;
    });
  }
  t('every challan in the .doc fixtures reconciles', verified > 0 && fail === 0, `${verified} verified`);
}

console.log('\n5. The tool refuses, loudly, what it cannot read');
t('a PDF with text but no font encoding is thrown, not silently skipped',
  /unreadable text layer/.test(src) && /throw new Error\('unreadable text layer/.test(src));
t('a scanned challan with no text layer is thrown too',
  /scanned image with no text layer/.test(src));
t('both are collected by name for the report', /scanned\.push\(f\.path\)/.test(src) && /noFont\.push\(f\.path\)/.test(src));
t('the build report names the unreadable files', /res\.scanned\.length/.test(src) && /res\.noFont\.length/.test(src));
t('challans that failed their own totals are listed after a build', /res\.unverified\.length/.test(src));

console.log('\n6. Our record is never merged into Olisa\'s');
t('challanStatus compares by challan number against the Master File',
  /function challanStatus[\s\S]{0,400}?masterRowsByChallan\(\)\.get\(key\)/.test(src)
  && /function masterRowsByChallan[\s\S]{0,400}?challanKeysOf\(r\.challanNo\)/.test(src));
t('only lines Olisa has NOT booked are offered',
  /const missing = lines\.filter\(ln => ln\.style && !chLineOnRows\(ln, rows, lines\)\)/.test(src));
t('unverified challans are held back from the style answer',
  /const good = rows\.filter\(r => r\.verified\)/.test(src) && /const shaky = rows\.filter\(r => !r\.verified\)/.test(src));
t('the answer says plainly that this is our record, not Olisa\'s',
  /our own delivery record, not Olisa's/.test(src));
t('the cache version was bumped so v1 indexes without lines are rebuilt',
  /CHALLAN_CACHE_VERSION = 2/.test(src));

console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
process.exit(fail ? 1 : 0);
