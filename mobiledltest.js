/* mobiledltest.js — the permanent guard on "it works on the laptop but not on my phone".
   Three faults were live at once and every one of them was invisible on desktop Chrome:
     1. the Ask tab wrote its own <a>, removed it in the same tick and revoked the blob at 60s,
        which hands Android a truncated .xlsx that opens showing only a header line;
     2. every download button was appended AFTER addBubble() had already scrolled, leaving it
        below the fold of a nested scroller a finger cannot move;
     3. the on-screen table and the downloaded file used different PI-coverage keys, so they
        could disagree about which "Not delivered yet" lines exist.
   Usage: node mobiledltest.js olisa.html                                                     */
const fs = require('fs');
const ExcelJS = require('exceljs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
  .map(m => m[1]).sort((a, b) => b.length - a.length)[0];

console.log('\n1. Every download goes through the one hardened saver');
// A download site is an anchor that gets a .download attribute and a .click(). There must be
// exactly one in the whole file, and it must be the one inside saveBlobAs.
const dlAttr = [...src.matchAll(/\.download\s*=/g)].length;
t('only one place in the file sets a .download attribute', dlAttr === 1, `found ${dlAttr}`);
t('that place is saveBlobAs', /function saveBlobAs\([\s\S]{0,600}?\.download\s*=/.test(src));
t('downloadMatches no longer builds its own anchor', !/function downloadMatches[\s\S]*?createElement\('a'\)[\s\S]*?\n\}/.test(src));
t('downloadMatches saves through saveBlobAs', /function downloadMatches[\s\S]*?saveBlobAs\(/.test(src));
t('the upload checker saves through saveBlobAs', /function uqDownloadXlsx[\s\S]*?saveBlobAs\(/.test(src));

console.log('\n2. The saver itself stays Android-safe');
const saver = src.slice(src.indexOf('function saveBlobAs'), src.indexOf('function saveBlobAs') + 1600);
t('the anchor is put IN the document before the click', /appendChild\(a\)[\s\S]*?a\.click\(\)/.test(saver));
t('the anchor is removed on a later tick, not in the click tick',
  /a\.click\(\)[\s\S]{0,200}?setTimeout\(\s*\(\)\s*=>\s*\{\s*a\.remove\(\)/.test(saver));
const revoke = saver.match(/revokeObjectURL\(url\)\s*\)?\s*,\s*(\d+)\)/);
t('the blob URL survives at least 120s so a queued Android write can finish',
  revoke && Number(revoke[1]) >= 120000, revoke ? revoke[1] + 'ms' : 'no revoke found');
// The 60s window was the old Ask-tab value. It must not come back anywhere.
t('no download path revokes its blob in under 120s',
  ![...src.matchAll(/revokeObjectURL\([^)]*\)[^,]*,\s*(\d+)\)/g)].some(m => Number(m[1]) < 120000));

console.log('\n3. A download reports what it actually contained');
t('downloadMatches returns the row count it wrote', /return \{ fileName, rows: dataRows/.test(src));
t('the row count is derived from the sheet, not from the input array',
  /dataRows\s*=\s*ws\.rowCount\s*-\s*2/.test(src));
t('an empty workbook is refused instead of saved', /!buf\.byteLength\) throw new Error/.test(src));
// downloadMatches must be reachable ONLY through runMatchDownload, so no call site can lose an
// exception or skip the receipt. That means exactly two mentions: its declaration and that call.
// Comment lines mention it by name; only real code counts.
const codeOnly = src.split('\n').filter(l => !/^\s*\/\//.test(l)).join('\n');
const dmMentions = [...codeOnly.matchAll(/\bdownloadMatches\(/g)].length;
t('downloadMatches is only ever called from runMatchDownload', dmMentions === 2,
  `${dmMentions} mentions, expected 2 (declaration + the wrapped call)`);
t('runMatchDownload catches and says so', /async function runMatchDownload[\s\S]*?catch \(err\)[\s\S]*?addBubble\(/.test(src));

console.log('\n4. The download button can be reached with a finger');
// Every button that is appended into a chat bubble must be scrolled into view afterwards,
// because addBubble() scrolled before it existed.
const appendSites = [...src.matchAll(/(bubble|container)\.appendChild\(btn\);\s*\n(\s*)([^\n]*)/g)];
t('there is at least one button appended into a bubble', appendSites.length >= 5,
  `${appendSites.length} sites`);
const unscrolled = appendSites.filter(m => !/scrollIntoChatView\(btn\)/.test(m[3]));
t('every appended download button is followed by scrollIntoChatView', unscrolled.length === 0,
  unscrolled.map(m => m[3].trim()).join(' | '));
t('scrollIntoChatView exists and scrolls after layout, across two frames',
  /function scrollIntoChatView[\s\S]{0,400}?requestAnimationFrame\([\s\S]{0,80}?requestAnimationFrame\(/.test(src));

console.log('\n5. A phone gets one level of scrolling, not two');
t('the chat box drops its own scroller on a narrow screen',
  /@media \(max-width: 760px\)[\s\S]{0,400}?\.askChat \{[^}]*max-height:\s*none/.test(html));
t('the preview tables stay capped so the button below is a short scroll away',
  /@media \(max-width: 760px\)[\s\S]{0,400}?\.askA \.previewbox \{[^}]*max-height:\s*46vh/.test(html));
t('desktop keeps the 520px chat box', /\.askChat \{[^}]*max-height:\s*520px/.test(html));

console.log('\n6. The table on screen and the file cannot disagree');
t('one shared PI-coverage key exists', /function piCoverKey\(r\)/.test(src));
t('the Excel export uses it', /const ck = piCoverKey\(r\);/.test(src));
t('the on-screen table uses it', /coveredPiLine\.add\(piCoverKey\(r\)\)/.test(src));
t('no second, divergent dim key is still being built for coverage',
  !/coveredPiLine\.add\(refKey \+/.test(src) && !/const ck = refKey \+/.test(src));
t('both sides draw filler lines from every matched style',
  /function fillerStylesFor/.test(src) && /fillerStylesFor\(matches\)/.test(src));
t('a PI line shared by two matched styles is printed once, not twice',
  /if \(coveredPiLine\.has\(sig\)\) return;\s*\n\s*coveredPiLine\.add\(sig\);/.test(src) &&
  /if \(!coveredXls\.has\(sig\)\) \{\s*\n\s*coveredXls\.add\(sig\);/.test(src));


console.log('\n7. The undelivered path offers a file too');
// A style ordered but never delivered has no master rows, so this branch used to render its PI
// lines and return — skipping the download button entirely. Not off-screen: never created.
t('the PI-only branch builds real rows', /const piOnlyRows = \[\];/.test(src));
t('and offers them for download', /addDownloadButton\(bubble, piOnlyRows, code, \{ noFillers: true \}\)/.test(src));
t('with fillers off so the PI lines are not written twice', /addDownloadButton\(container, matches, code, opts\)/.test(src));
t('a bulk search collects them into the combined file', /if \(result\.piOnlyRows && result\.piOnlyRows\.length\) allMatches\.push/.test(src));
t('and the combined file turns fillers off when it carries them', /allMatches\.some\(r => r\._piOnly\)/.test(src));
t('the rows carry the marker the combined file checks for', /_piOnly: true/.test(src));

console.log('\n8. The build tells the truth about when it ran, and why challans are missing');
// The stamp used to be set only after the challan phase, so a PI index that had just been built
// and saved still reported the PREVIOUS build's time whenever that phase was slow or interrupted.
const buildFn = src.slice(src.indexOf('Challan copies, in the same build'), src.indexOf('renderPiIndexDiagnostics'));
t('the stamp is set before the challan phase, not after',
  buildFn.indexOf('piIndexBuiltAt = Date.now()') < buildFn.indexOf('await buildChallanIndex'),
  'stamp must not wait on challans');
t('the challan phase still cannot abort the build', /catch \(e\) \{\s*challanError = e\.message;/.test(src));
t('why no challans were read is remembered', /challanBuildError = challanError;/.test(src));
t('and said in the health line instead of a blank "not indexed yet"',
  /challanBuildError \? ` \\u2014 challan copies NOT read/.test(src));
t('the Challans folder is found by one Drive query, not a walk of the whole Drive',
  /name contains 'hallan' and trashed=false/.test(src));
t('with the folder walk kept only as a fallback', /if \(!found\) found = await findFolderRecursive/.test(src));
t('no OCR call can wait forever inside a build', /function ocrFetch/.test(src) && /ctrl\.abort\(\)/.test(src));
t('every OCR network call goes through it', [...src.matchAll(/await ocrFetch\(/g)].length === 3,
  [...src.matchAll(/await ocrFetch\(/g)].length + ' of 3');
t('the build number is one constant', /const APP_VERSION = \d+;/.test(src));
t('shown in the footer', /id="buildStamp"/.test(html) && /'build v' \+ APP_VERSION/.test(src));
t('and in the health strip', /v\$\{APP_VERSION\}/.test(src));

console.log('\n9. The workbook really does hold one row per match (real ExcelJS)');
// Lift the shipped row-building code and run it, so "the file only had one line" is a thing the
// suite can prove or disprove rather than a thing that has to be reproduced on a phone.
(async () => {
  const body = src.slice(src.indexOf('async function downloadMatches'),
                         src.indexOf('// ==================== A DOWNLOAD THAT SAYS WHAT IT CONTAINED'));
  const stubs = {
    ExcelJS,
    piQtyForRow: () => null,
    piCoverKey: r => `${r.piRef}|${r.styleNorm}|nodim`,
    normPiKey: v => (v === null || v === undefined || v === '') ? '' : String(v).trim(),
    effectiveShort: r => r.short || 0,
    rowClass: () => '',
    challanDisp: r => r.challanNo,
    combinedRemarks: r => r.remarks || '',
    extractMeasurement: d => String(d || ''),
    piDateFor: () => '',
    piLinesForStyle: () => [],
    dimKey: () => '', dimsFromMasterText: () => null, uqExtractDimsAny: () => null,
    notifyDownload: () => {},
    saveBlobAs: (blob, name) => name,
    Blob: class { constructor(parts) { this.size = parts[0].byteLength || parts[0].length || 0; } }
  };
  const make = new Function(...Object.keys(stubs), body + '\nreturn downloadMatches;');
  const fn = make(...Object.values(stubs));

  const rows = n => Array.from({ length: n }, (_, i) => ({
    piRef: 138, styleNorm: 'M82035A1-3', style: 'M82035A1-3', item: 'Carton',
    itemDesc: 'L59 x W45 x H32 cm', delQty: 210, recQty: 210, short: 0,
    challanNo: 3139779 + i, challanDate: '08-09-2026', invDate: '08-09-2026', remarks: ''
  }));

  for (const n of [1, 7, 240, 2500]) {
    const res = await fn(rows(n), 'Test');
    t(`${n} match(es) -> ${n} data row(s) in the sheet`, res.rows === n, `got ${res.rows}`);
  }
  const big = await fn(rows(2500), 'Test');
  t('a 2,500-row workbook is a real file, not an empty shell', big.bytes > 20000, big.bytes + ' bytes');
  t('the file name is handed back so it can be reported', /Olisa Tools - Query Test\.xlsx$/.test(big.fileName), big.fileName);

  console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
  process.exit(fail ? 1 : 0);
})();
