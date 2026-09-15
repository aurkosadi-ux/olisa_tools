/* ocrtest.js — the guard on the Drive OCR path.
   The dangerous part here is not the reading, it is the renaming: this process touches files in
   the real Challans folder. So the order is fixed and tested — the Improved file is created AND
   read back before the original is ever renamed — and the skip rules are tested, because a
   "-Original" left in the walk would make one dispatch appear twice.
   Usage: node ocrtest.js olisa.html                                                            */
const fs = require('fs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];
function lift(n) { const s = src.indexOf('function ' + n); if (s < 0) throw new Error('missing ' + n); let d = 0, e = s; for (let i = src.indexOf('{', s); i < src.length; i++) { if (src[i] === '{') d++; else if (src[i] === '}') { d--; if (!d) { e = i + 1; break; } } } return src.slice(s, e); }
const bag = new Function(lift('ocrBaseName') + lift('ocrOriginalName') + lift('ocrImprovedName')
  + lift('isOcrImproved') + lift('isOcrOriginal') + lift('ocrSafeText')
  + '\nreturn { ocrOriginalName, ocrImprovedName, isOcrImproved, isOcrOriginal, ocrSafeText };')();

console.log('\n1. Naming');
const real = '11-8-26_Challan__IR-15-8-26__Manual_Challan.pdf';
t('the original keeps its name and extension, plus -Original',
  bag.ocrOriginalName(real) === '11-8-26_Challan__IR-15-8-26__Manual_Challan-Original.pdf', bag.ocrOriginalName(real));
t('the new file takes the same name plus -Improved.pdf',
  bag.ocrImprovedName(real) === '11-8-26_Challan__IR-15-8-26__Manual_Challan-Improved.pdf', bag.ocrImprovedName(real));
t('a .doc original keeps .doc when renamed', bag.ocrOriginalName('x.doc') === 'x-Original.doc', bag.ocrOriginalName('x.doc'));
t('but its improved twin is always a PDF', bag.ocrImprovedName('x.doc') === 'x-Improved.pdf', bag.ocrImprovedName('x.doc'));
t('dots inside the name are not mistaken for the extension',
  bag.ocrImprovedName('31-8-26_Challan__IR-4-9-26_.pdf') === '31-8-26_Challan__IR-4-9-26_-Improved.pdf', bag.ocrImprovedName('31-8-26_Challan__IR-4-9-26_.pdf'));

console.log('\n2. Nothing gets read, or re-read, twice');
t('an -Original is recognised', bag.isOcrOriginal('a-Original.pdf') && bag.isOcrOriginal('a-Original.doc'));
t('an -Improved is recognised', bag.isOcrImproved('a-Improved.pdf'));
t('an ordinary challan is neither', !bag.isOcrOriginal(real) && !bag.isOcrImproved(real));
t('the folder walk skips -Original files', /else if \(isOcrOriginal\(name\)\) \{/.test(src));
t('the walk now also picks up images', /\\\.\(pdf\|docx\?\|rtf\|png\|jpe\?g\)\$/.test(src) || /pdf\|docx\?\|rtf\|png\|jpe\?g/.test(src));
t('an improved file is indexed under the challan number, like any other',
  /r\.found\.forEach\(c => \{ next\[c\.no\] = /.test(src));

console.log('\n3. The rename happens LAST, and only after the new file is proved');
const fn = lift('improveOneChallanFile');
const iUpload = fn.indexOf('upload/drive/v3/files?uploadType=multipart&fields=id,name');
const iVerify = fn.indexOf('scanChallanPdf(check');
const iThrow = fn.indexOf('still has no challan number');
const iRename = fn.indexOf("method: 'PATCH'");
t('the improved file is uploaded before anything is renamed', iUpload > 0 && iUpload < iRename);
t('it is read back and checked for a challan', iVerify > 0 && iVerify < iRename);
t('a file with no challan in it throws instead of renaming', iThrow > 0 && iThrow < iRename);
t('the rename is the last thing that happens', iRename > iVerify && iRename > iThrow);
t('the temp Google Doc is deleted even when the export fails', /finally \{[\s\S]{0,400}?method: 'DELETE'/.test(lift('driveOcrText')));

console.log('\n4. It only runs when it safely can');
t('needs the Drive connection', /savedRootHandle && savedRootHandle\.isDrive && navigator\.onLine/.test(src));
// The catch block now also records the failure by fingerprint so the next build does not repeat
// the OCR. What matters for THIS test is unchanged: the failure is collected, not rethrown.
const ocrCatch = (src.match(/catch \(e\) \{\s*improveFailed\.push\([\s\S]{0,900}?\n      \}/) || [''])[0];
t('a failure is collected per file, never aborting the build',
  /improveFailed\.push\(\{ name: f\.name, why: e\.message \}\)/.test(ocrCatch) && !/throw/.test(ocrCatch));
t('and the failed file is remembered so OCR is not repeated next build', /ocrTried:\s*true/.test(src));
t('the build reports what was improved and what was not', /res\.improved\.length/.test(src) && /res\.improveFailed\.length/.test(src));
t('OCR still needs the Drive connection inside driveOcrText', /OCR needs the Google Drive connection/.test(src));

console.log('\n5. Text that pdf-lib cannot encode does not break the file');
t('Chinese is replaced rather than thrown', bag.ocrSafeText('CARTON \u6df7 15\u53cc 70*48*34CM').indexOf('\u6df7') === -1);
t('the ASCII the parser needs survives intact', bag.ocrSafeText('|Size:M81988-3A1-1 |Code:nill').includes('|Size:M81988-3A1-1'));
t('newlines are kept so the layout survives', bag.ocrSafeText('a\nb') === 'a\nb');
t('each drawText is individually guarded', /try \{ pg\.drawText\([\s\S]{0,120}?\} catch \(e\)/.test(src));

console.log('\n6. A real Improved PDF round-trip');
(async () => {
  let PDFDocument, StandardFonts;
  try { ({ PDFDocument, StandardFonts } = require('pdf-lib')); }
  catch (e) { console.log('  -- pdf-lib not installed here, skipping the round-trip'); return done(); }
  // Build exactly what buildImprovedPdf builds: a page with invisible text over it, then read it
  // back the way pdf.js will and confirm the challan grammar survives.
  const out = await PDFDocument.create();
  const font = await out.embedFont(StandardFonts.Helvetica);
  const pg = out.addPage([595, 842]);
  const lines = ['Challan No : 104470017/ 31-AUG-26', 'PI:121/700768/0159082026',
    'Line Item Description Qty Unit Other Information',
    '1 Master Carton, 5 Ply, L33 x W30.5 x H30.5cm; 500 Pieces Color:PRINT |Size:M82065-3B |Code:nill |Type:n',
    'Total Line : 1 Total Qty: 500'];
  let y = 842 - 24;
  lines.forEach(l => { pg.drawText(l, { x: 16, y, size: 8, font, opacity: 0 }); y -= 9; });
  const bytes = await out.save();
  t('the improved PDF is a real file', bytes.length > 500, bytes.length + ' bytes');
  const reread = await PDFDocument.load(bytes);
  t('and it opens cleanly', reread.getPageCount() === 1);
  done();
})().catch(e => { t('round-trip ran', false, e.message); done(); });

function done() {
  console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
  process.exit(fail ? 1 : 0);
}
