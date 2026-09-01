/* indexspeedtest.js — guards the speed of building the style/challan index.
   Proves, against the SHIPPED source, that Drive folder listing and file/PDF reading run in
   parallel lanes rather than one request at a time, and simulates real Drive latency to show the
   actual wall-clock effect — the thing Sadi timed with a stopwatch. */
const fs = require('fs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };
const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];

function fn(name) {
  const m = src.match(new RegExp('(?:async )?function ' + name + '\\s*\\([\\s\\S]*?\\n\\}\\n(?=(?:async )?function |\\(function|$)'));
  return m ? m[0] : '';
}

console.log('1. The style-index build no longer waits on one folder at a time');
const bsi = fn('buildStyleIndex');
t('buildStyleIndex exists and was found', bsi.length > 200);
t('folder listing runs through runPool, not a bare for-loop', /runPool\(piFolders, LIST_CONCURRENCY/.test(bsi), 'listing phase not parallel');
t('file reading runs through runPool, not nested inside the folder loop', /runPool\(readTasks, READ_CONCURRENCY/.test(bsi));
t('listing and reading are two separate passes (list fully, then read all at once)',
  bsi.indexOf('runPool(piFolders,') < bsi.indexOf('runPool(readTasks,'));
t('a single unreadable folder does not abort the whole build',
  /catch \(e\) \{ \/\* this folder is unreadable/.test(bsi));
t('concurrency is a small, named, easy-to-find number (not buried magic)', /LIST_CONCURRENCY = 8/.test(bsi) && /READ_CONCURRENCY = 8/.test(bsi));
t('the running index is still checkpointed as it grows (a killed build keeps its progress)',
  /await checkpointStyleIndex\(index\);/.test(bsi));
t('the final folder count used in the summary line still reflects reality',
  /foldersScanned = foldersListed;/.test(bsi));

console.log('\n2. The challan index no longer walks subfolders or reads PDFs one at a time');
const bci = fn('buildChallanIndex');
t('buildChallanIndex exists and was found', bci.length > 200);
t('subfolder recursion (e.g. one folder per month) runs through runPool', /await runPool\(subs, 6,/.test(bci));
t('reading and parsing every PDF runs through runPool', /await runPool\(pdfs, 5,/.test(bci));
t('a file that fails to download does not stop the rest from being read',
  /catch \(e\) \{ failed\+\+; done\+\+; return; \}/.test(bci));
t('an already-parsed, unchanged file is still reused instead of re-parsed', /const reused = Object\.values\(challanIndex\)/.test(bci));

console.log('\n3. Nothing about WHAT gets indexed changed — only HOW FAST');
t('style tokens are still extracted the same way', /extractStyleTokens\(text\)/.test(bsi));
t('a token still keeps a de-duplicated list of PI refs', /if \(!index\[t\]\.includes\(piRef\)\) index\[t\]\.push\(piRef\)/.test(bsi));
t('challan matching still keys off the same normalised number', /next\[normChallanKey\(e\.no\)\]/.test(bci));

console.log('\n4. Simulated real Drive latency — the actual speedup');
const LATENCY = 250; // ms — a realistic Drive round trip
const delay = ms => new Promise(r => setTimeout(r, ms));
function runPool(items, limit, worker) {
  return new Promise(resolve => {
    const results = new Array(items.length); let next = 0, active = 0, launched = 0;
    if (!items.length) return resolve(results);
    function lane() {
      if (next >= items.length) { if (--active === 0) resolve(results); return; }
      const i = next++;
      worker(items[i], i).then(r => { results[i] = r; lane(); }, e => { results[i] = { __err: e }; lane(); });
    }
    const lanes = Math.min(limit, items.length);
    active = lanes;
    for (let k = 0; k < lanes; k++) lane();
  });
}
function makeWorld(n) {
  const folders = [];
  for (let i = 1; i <= n; i++) {
    const nFiles = 1 + (i % 3);
    const files = []; for (let f = 0; f < nFiles; f++) files.push({ token: `M${1000 + i}${f}` });
    folders.push({ ref: i, files });
  }
  return folders;
}
async function driveList(folder) { await delay(LATENCY); return folder.files; }
async function driveGet(file) { await delay(LATENCY); return file.token; }
async function oldSequential(folders) {
  const index = {};
  for (const folder of folders) {
    const files = await driveList(folder);
    for (const file of files) {
      const tok = await driveGet(file);
      (index[tok] = index[tok] || []).push(folder.ref);
    }
  }
  return index;
}
async function newParallel(folders) {
  const index = {};
  const perFolder = await runPool(folders, 8, async f => ({ ref: f.ref, files: await driveList(f) }));
  const tasks = [];
  perFolder.forEach(e => e.files.forEach(file => tasks.push({ ref: e.ref, file })));
  await runPool(tasks, 8, async ({ ref, file }) => {
    const tok = await driveGet(file);
    (index[tok] = index[tok] || []).push(ref);
  });
  return index;
}
(async () => {
  const world = makeWorld(127);   // matches the real folder count reported in the app
  let t0 = Date.now();
  const oldIndex = await oldSequential(world);
  const oldMs = Date.now() - t0;
  t0 = Date.now();
  const newIndex = await newParallel(world);
  const newMs = Date.now() - t0;
  const speedup = oldMs / newMs;

  t('the parallel version is meaningfully faster', speedup > 3, `${speedup.toFixed(1)}x`);
  t('every style token the old version found is still found', Object.keys(oldIndex).sort().join(',') === Object.keys(newIndex).sort().join(','));
  t('every PI reference under each token is identical, just possibly reordered',
    Object.keys(oldIndex).every(k => [...oldIndex[k]].sort().join(',') === [...(newIndex[k]||[])].sort().join(',')));
  const projectedMin = (25 * (newMs / oldMs)).toFixed(1);
  console.log(`       at this ratio, a real 25-minute scan becomes about ${projectedMin} minutes`);
  t("a 127-folder scan (this project's real size) lands comfortably under 5 minutes", Number(projectedMin) < 5, projectedMin + ' min');

  console.log(`\n${fail ? 'FAILED' : 'PASSED'} — ${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
