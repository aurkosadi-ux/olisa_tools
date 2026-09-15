/* cachetest.js — guards the "it rescans everything, every time" family of bugs.
 *
 * Every check here is a bug that actually shipped. The version-stamp one in particular cost a full
 * re-read of every challan on every single app launch, and nothing in the suite could see it.
 */
const fs = require('fs');
const path = require('path');
const SRC = fs.readFileSync(path.join(__dirname, 'olisa.html'), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (detail ? '  -> ' + detail : '')); }
};
const section = t => console.log('\n' + t);

/* Pull one function's source out by brace matching, so a check can be scoped to it. */
function fnBody(name) {
  const m = SRC.indexOf('function ' + name + '(');
  if (m === -1) return '';
  let i = SRC.indexOf('{', m), depth = 0, start = i;
  for (; i < SRC.length; i++) {
    if (SRC[i] === '{') depth++;
    else if (SRC[i] === '}') { depth--; if (depth === 0) return SRC.slice(start, i + 1); }
  }
  return '';
}

section('1. Every cached index is written WITH its version stamp');
// initChallanIndex refuses `cached.version !== CHALLAN_CACHE_VERSION`. Any write that omits the
// version is therefore a write that will be silently discarded on the next launch.
const challanWrites = SRC.match(/idbSet\(\s*'challanIndex'\s*,\s*\{[^}]*\}/g) || [];
ok('challanIndex is written somewhere', challanWrites.length > 0, challanWrites.length + ' write(s)');
challanWrites.forEach((w, i) => {
  ok(`challanIndex write #${i + 1} carries a version`, /version\s*:/.test(w), w.replace(/\s+/g, ' ').slice(0, 90));
});
const piWrites = SRC.match(/idbSet\(\s*'piIndex'\s*,\s*\{[^}]*\}/g) || [];
piWrites.forEach((w, i) => {
  ok(`piIndex write #${i + 1} carries a version`, /version\s*:/.test(w), w.replace(/\s+/g, ' ').slice(0, 90));
});
// The reader's guard must still be the thing that makes the stamp matter.
ok('the challan cache reader still checks the version',
  /cached\.version\s*!==\s*CHALLAN_CACHE_VERSION/.test(SRC));
// Adopting a shared index must also set the in-memory stamp, or the UI reports "never read".
const adopt = SRC.slice(SRC.indexOf('data.challanIndex && Object.keys(data.challanIndex).length'));
ok('adopting a shared index sets challanIndexBuiltAt',
  /challanIndexBuiltAt\s*=\s*incoming/.test(adopt.slice(0, 1200)));

section('2. Unchanged files are recognised WITHOUT being downloaded');
const bci = fnBody('buildChallanIndex');
ok('buildChallanIndex exists', bci.length > 500);
// The original bug: getFile() ran first and the fingerprint was computed from the downloaded blob,
// so the reuse check saved nothing at all. metaFp must be consulted before any download.
const metaAt = bci.indexOf('metaFp(');
const getFileAt = bci.indexOf('.getFile()');
ok('buildChallanIndex fingerprints from metadata', metaAt !== -1);
ok('and does so BEFORE downloading the file', metaAt !== -1 && getFileAt !== -1 && metaAt < getFileAt,
  `metaFp at ${metaAt}, getFile at ${getFileAt}`);
ok('the style scan has a per-file token cache', /styleFileCache\s*\[/.test(SRC));
const bsi = fnBody('buildStyleIndex');
ok('the style scan checks that cache before reading', bsi.indexOf('styleFileCache[') < bsi.indexOf('XLSX.read'));
ok('Pass B remembers PDFs it has already rejected', /piPdfDuds\s*\[/.test(SRC));

section('3. metaFp and the downloaded fingerprint are the SAME string');
// If these two ever disagree, nothing is ever reused and the cache is dead weight that also lies.
// Drive returns `size` as a string and `modifiedTime` as RFC-3339; File carries a number and a ms
// epoch. The two must still render identically inside a template literal.
function metaFp(p, handle) {                     // mirrors the implementation under test
  const m = handle && handle.meta;
  if (!m) return '';
  const size = (m.size === undefined || m.size === null) ? '' : m.size;
  const mt = m.modifiedTime ? Date.parse(m.modifiedTime) : 0;
  if (size === '' && !mt) return '';
  return `${p}|${size}|${mt}`;
}
const meta = { size: '482913', modifiedTime: '2026-08-31T09:14:22.000Z' };
const fromMeta = metaFp('August - 2026/31-8-26 Challan.pdf', { meta });
const fromFile = `August - 2026/31-8-26 Challan.pdf|${482913}|${Date.parse(meta.modifiedTime)}`;
ok('metadata and downloaded fingerprints match exactly', fromMeta === fromFile, `${fromMeta} vs ${fromFile}`);
ok('a handle with no metadata yields no fingerprint', metaFp('x.pdf', {}) === '');
ok('and that empty result is falsy, so callers fall back to downloading', !metaFp('x.pdf', {}));

section('4. The synchronous parsers cannot hold the page');
ok('a deadline-based yield helper exists', /function yieldSoon\(/.test(SRC));
ok('the work-order read loop yields', /await yieldSoon\(\)/.test(bsi));
ok('it yields between sheets of one workbook too',
  bsi.indexOf('sheet_to_csv') < bsi.indexOf('await yieldSoon()', bsi.indexOf('sheet_to_csv')));
ok('the challan read loop yields', /await yieldSoon\(\)/.test(bci));
ok('the PI-quantity loop yields', /await yieldSoon\(\)/.test(fnBody('buildPiQtyIndex')));
ok('progress repaints are throttled', /function setProgress\(/.test(SRC));
// XLSX.read defaults parse formulas, styles and HTML — all discarded by this scan, all main-thread.
ok('XLSX.read skips work the scan throws away',
  /XLSX\.read\([^)]*cellFormula:\s*false[^)]*cellStyles:\s*false/.test(bsi));

section('5. Folder listing is batched, not one request per folder');
ok('a bulk listing function exists', /async function driveListChildrenBulk\(/.test(SRC));
const bulk = fnBody('driveListChildrenBulk');
ok('it ORs several parents into one query', /in parents.*join\(' or '\)|join\(' or '\)/.test(bulk));
ok('it paginates', /nextPageToken/.test(bulk));
ok('it files results by their actual parent', /f\.parents\s*\|\|\s*\[\]/.test(bulk));
ok('it only caches a batch that completed cleanly',
  bulk.indexOf('driveListCache.set') > bulk.indexOf('} while (pageToken)'));
ok('a breadth-first tree prefetch exists', /async function prefetchDriveTree\(/.test(SRC));
ok('the root PDF sweep no longer recurses one folder at a time',
  !/for \(const s of subs\) await walk\(/.test(fnBody('sweepRootForPdfs')));
// Batch arithmetic: the win only exists if the batch size is real.
const mBatch = SRC.match(/const DRIVE_PARENT_BATCH\s*=\s*(\d+)/);
ok('the batch size is set', !!mBatch);
if (mBatch) {
  const b = parseInt(mBatch[1], 10);
  ok('batch size is large enough to matter', b >= 20, String(b));
  // Drive rejects very long queries; each term is ~30 chars of file id plus " in parents or ".
  ok('and small enough to stay inside Drive query limits', b <= 60, String(b));
  ok(`1,000 folders becomes ${Math.ceil(1000 / b)} requests, not 1,000`, Math.ceil(1000 / b) <= 50);
}

section('6. The expensive recovery paths do not repeat themselves');
ok('challan misses are remembered by fingerprint', /challanMisses\s*\[/.test(SRC));
ok('a file that OCR already failed on is marked', /ocrTried:\s*true/.test(SRC));
ok('misses are rebuilt each run, so a fixed file clears itself',
  /challanMisses\s*=\s*freshMisses/.test(SRC));
ok('OCR is skipped when offline without marking files as tried',
  /needOcr\.forEach\(\(\{ f, fp, why \}\)/.test(SRC));

section('7. Files that are not challans are not treated as challans');
ok('Office lock files are skipped in the challan walk', /\^~\\\$/.test(bci) || /\/\^~\\\$\//.test(bci));
ok('the style scan uses the spreadsheet-name guard, not a bare extension test',
  /isSpreadsheetName\(fname\)/.test(bsi) && !/\/\\\.\(xlsx\|xls\)\$\/i\.test\(fname\)/.test(bsi));
// Was pinned to an `improvedBases` set that only ever looked one level deep, which is why a
// four-deep -Improved chain sailed straight past it. The rule is now family-based: group by the
// name everything descends from, keep the original and the deepest copy, skip what is between.
ok('OCR copies are grouped into families, not matched one level deep',
  /ocrFamilyRoot\(/.test(bci) && /families/.test(bci));
ok('the intermediate copies are collected for reporting', /ocrJunk/.test(bci));
ok('and the name builder cannot append a second -Improved',
  /\(\?:-\(\?:Improved\|Original\)\)\+\$/.test(SRC));

section('8. A challan still opens when its file has been renamed or moved');
const cfh = fnBody('challanFileHandle');
ok('the Drive file id is tried first', cfh.indexOf('entry.driveId') !== -1 &&
  cfh.indexOf('entry.driveId') < cfh.indexOf('findChallanRoot'));
ok('a trashed file is not accepted', /!meta\.trashed/.test(cfh));
ok('there is still a path fallback', /prefix \+ name === entry\.path/.test(cfh));
ok('and a name fallback below that', /byName/.test(cfh));

section('9. Checking for changes does not read any file');
const delta = fnBody('challanTreeDelta');
ok('a listing-only change check exists', delta.length > 200);
ok('it downloads nothing', !/getFile\(\)/.test(delta));
ok('it compares against both the index and the known misses',
  /challanIndex/.test(delta) && /challanMisses/.test(delta));
ok('it reports files it cannot fingerprint rather than guessing', /unfingerprinted/.test(delta));

section('10. Version discipline');
const appV = (SRC.match(/const APP_VERSION\s*=\s*(\d+)/) || [])[1];
const swV = (fs.readFileSync(path.join(__dirname, 'sw.js'), 'utf8').match(/olisa-tools-v(\d+)/) || [])[1];
ok('olisa.html carries a version', !!appV);
ok('sw.js cache name matches it', appV === swV, `page v${appV}, sw v${swV}`);

section('11. The yield actually yields, and near-misses are only suggested');
// The 45ms shared gate made seven of eight lanes get a microtask instead of a task. A microtask
// does not let the browser paint, so the page locked harder than before the fix went in.
const ys = fnBody('yieldSoon');
// Checked against code, not prose: the comment above the fix names the old gate, and a grep of
// raw text cannot tell an explanation of a bug from the bug itself.
const codeOnly = t => t.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
ok('yieldSoon has no shared time gate', !/lastYieldAt|YIELD_EVERY_MS/.test(codeOnly(SRC)));
ok('and always crosses a real task boundary', /setTimeout\(r, 0\)/.test(ys) && !/Promise\.resolve\(\)/.test(ys));
ok('a frame is given BEFORE the workbook parse too',
  bsi.indexOf('await yieldSoon()') < bsi.indexOf('XLSX.read'));
ok('the challan scan plans from metadata before reading', /const plan = \{ reuse: \[\]/.test(bci));
ok('progress counts what needs reading, not every file', /plan\.read\.length, 1\)/.test(bci));
ok('a full re-read despite a full index is reported', /challanFpMismatch/.test(SRC));

// One digit apart, exactly: the rule behind suggesting 117116022 for 1171160222.
function oneDigitApart(a, b) {
  if (a === b) return false;
  if (Math.abs(a.length - b.length) > 1) return false;
  if (a.length === b.length) { let d = 0; for (let i = 0; i < a.length; i++) if (a[i] !== b[i] && ++d > 1) return false; return d === 1; }
  const long = a.length > b.length ? a : b, short = a.length > b.length ? b : a;
  let j = 0, skipped = 0;
  for (let i = 0; i < long.length; i++) { if (long[i] === short[j]) j++; else if (++skipped > 1) return false; }
  return j === short.length;
}
ok('1171160222 is one digit from 117116022', oneDigitApart('1171160222', '117116022'));
ok('a substitution counts', oneDigitApart('117116022', '117116023'));
ok('an identical number is not a near miss', !oneDigitApart('117116022', '117116022'));
ok('two digits apart does not count', !oneDigitApart('117116022', '117116099'));
ok('two digits longer does not count', !oneDigitApart('11711602299', '117116022'));
ok('a completely different number does not count', !oneDigitApart('117116022', '320210810'));
ok('nothing is linked automatically \u2014 the suggestion is text only',
  /nothing is linked automatically/.test(SRC));

console.log('\n' + (fail ? `FAILED — ${pass} passed, ${fail} failed` : `PASSED — ${pass} passed, 0 failed`));
process.exit(fail ? 1 : 0);
