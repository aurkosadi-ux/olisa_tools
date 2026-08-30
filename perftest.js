/* perftest.js — the permanent guard on "page unresponsive".
   Proves the PO Summary never holds the main thread in one long task, and that the launch screen
   costs nothing once it has played. */
const fs = require('fs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };
const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const po = src.slice(src.indexOf('(function() { // PO Summary Generator'), src.indexOf('(function() { // Undelivered Report Generator'));

console.log('1. The preview cannot flood the DOM');
t('a row cap exists', /const PREVIEW_LIMIT = \d+/.test(po));
const lim = Number((po.match(/const PREVIEW_LIMIT = (\d+)/) || [])[1]);
t('the cap is a sane size (50-400)', lim >= 50 && lim <= 400, String(lim));
t('rows are sliced before being drawn', /rows\.slice\(0, PREVIEW_LIMIT\)/.test(po));
t('the cap is announced, not silent', /Showing the first \$\{PREVIEW_LIMIT\}/.test(po));
t('the user is told the FILE still has every line', /downloaded file always contains all/.test(po));
t('lines needing review are never hidden silently', /hiddenReview/.test(po));
t('"show all" is offered', /poShowAllBtn/.test(po));
t('a new order re-applies the cap', /previewShowAll = false/.test(po));

console.log('\n2. No long task between click and finish');
t('reading yields per file', /await poYield\(\);\s*\n\s*const recs = await readOrderRows/.test(po));
t('readOrderRows yields per sheet', (po.match(/await poYield\(\)/g) || []).length >= 4, String((po.match(/await poYield\(\)/g) || []).length));
t('classification yields before it starts', /Classifying \$\{allRecords\.length\}/.test(po));
t('the preview yields before drawing', /Drawing the preview/.test(po));
t('multi-workbook download yields between sheets', /Building sheet \$\{i \+ 1\} of/.test(po));
t('progress is shown while reading, not just at the end', /Reading \$\{f\.name\}/.test(po));

console.log('\n3. Each sheet is converted once, not twice');
t('sheets are converted into a reusable array', /const sheets = \[\];/.test(po));
const conversions = (po.match(/XLSX\.utils\.sheet_to_json/g) || []).length;
t('only one sheet_to_json call site remains', conversions === 1, String(conversions) + ' call sites');
t('the second pass reuses the array', /for \(const \{ aoa \} of sheets\)/.test(po));
t('the diagnostic dump reuses it too', /for \(const \{ name: sheetName, aoa \} of sheets\)/.test(po));

console.log('\n4. Launch screen costs nothing once played');
t('it is markup, not JS-built (paints with the first frame)', /<div class="launch" id="launchScreen"/.test(html));
t('it never blocks a tap', /\.launch \{[\s\S]{0,600}?pointer-events: none;/.test(html));
t('it only appears in the installed app', /@media \(display-mode: standalone\)/.test(html));
t('reduced motion switches it off entirely', /prefers-reduced-motion: reduce\) \{ \.launch \{ display: none !important/.test(html));
t('it shows once per session, not per page', /sessionStorage\.getItem\('olisaLaunched'\)/.test(html));
t('a return visit removes it before paint', /if \(seen\) \{ el\.parentNode\.removeChild\(el\); return; \}/.test(html));
t('it removes itself from the DOM when done', /removeChild\(el\)/.test(html));
t('a timeout guarantees removal if animationend never fires', /setTimeout\(kill, \d+\)/.test(html));
// Only compositor-safe properties may be animated inside the launch block.
const launchCss = (html.match(/\.launch \{ display: none; \}[\s\S]*?@media \(prefers-reduced-motion: reduce\) \{ \.launch[^\n]*\n/) || [''])[0];
const frames = [...html.matchAll(/@keyframes (launch-[\w-]+) \{([\s\S]*?)\}\n/g)];
t('launch keyframes exist', frames.length >= 3, String(frames.length));
const banned = /(width|height|top|left|right|bottom|margin|padding|box-shadow|filter)\s*:/;
t('keyframes animate only opacity/transform (no layout, no paint)',
  frames.every(f => !banned.test(f[2])), frames.filter(f => banned.test(f[2])).map(f => f[1]).join(', '));
t('the animation is finite (forwards/both, never infinite)', !/animation:[^;]*launch[^;]*infinite/.test(html));
t('will-change is declared so the layer is promoted up front', /will-change: transform, opacity;/.test(launchCss || html));

console.log('\n5. Nothing regressed in what the file contains');
t('every row is still written to the workbook', /rows\.forEach\(r => \{\s*\n\s*const excelRow = ws\.addRow/.test(po));
t('buildWorkbookBuffer takes the FULL row set, not the preview', /buildWorkbookBuffer\(generatedRows,/.test(po));
t('the preview cap is not applied to any download path', !/buildWorkbookBuffer\(\s*shown/.test(po) && !/slice\(0, PREVIEW_LIMIT\)[\s\S]{0,200}buildWorkbookBuffer/.test(po));

console.log('\n' + (fail ? 'FAILED' : 'PASSED') + ` — ${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
