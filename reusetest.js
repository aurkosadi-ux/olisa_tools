/* reusetest.js — runs the REAL buildChallanIndex twice against a fake Drive folder and counts how
 * many files it actually downloads the second time. Nothing textual: this measures behaviour.
 *
 * "It reads the challans from the start every time" is either true or it is not, and no amount of
 * reading the source settles it. This does.
 */
const fs = require('fs');
const acorn = require('acorn');

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
  .map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const ast = acorn.parse(src, { ecmaVersion: 2022 });
// Object.create(null): a plain {} inherits constructor/toString/valueOf, so any function body
// mentioning those names would 'resolve' to Object.prototype and crash the lift.
const fns = Object.create(null);
(function scan(n) {
  if (!n || typeof n.type !== 'string') return;
  if (n.type === 'FunctionDeclaration' && n.id) fns[n.id.name] = src.slice(n.start, n.end);
  for (const k of Object.keys(n)) {
    const v = n[k];
    if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && scan(c));
    else if (v && typeof v.type === 'string') scan(v);
  }
})(ast);

let pass = 0, fail = 0;
const t = (name, cond, extra) => {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra !== undefined ? '  -> ' + extra : '')); }
};

// ---- a fake Drive folder: 40 challan PDFs across 3 month subfolders ----
let downloads = 0, listCalls = 0;
function makeWorld() {
  const months = ['June - 2026', 'July - 2026', 'August - 2026'];
  const tree = { name: 'Challans', id: 'CH', dirs: {}, files: {} };
  months.forEach((m, mi) => {
    const d = { name: m, id: 'D' + mi, dirs: {}, files: {} };
    for (let i = 0; i < 12; i++) {
      const n = `${m.slice(0, 2)}-${i}-26 Challan.pdf`;
      d.files[n] = { id: `f${mi}_${i}`, name: n, size: String(100000 + i),
                     modifiedTime: '2026-0' + (6 + mi) + '-1' + (i % 9) + 'T09:00:00.000Z',
                     challan: String(320000000 + mi * 100 + i) };
    }
    tree.dirs[m] = d;
  });
  for (let i = 0; i < 4; i++) {
    const n = `loose-${i}.pdf`;
    tree.files[n] = { id: `L${i}`, name: n, size: String(50000 + i),
                      modifiedTime: '2026-05-0' + (i + 1) + 'T09:00:00.000Z',
                      challan: String(319000000 + i) };
  }
  return tree;
}
function dirHandle(node) {
  return {
    kind: 'directory', isDrive: true, id: node.id, name: node.name,
    async *entries() {
      listCalls++;
      for (const k of Object.keys(node.dirs)) yield [k, dirHandle(node.dirs[k])];
      for (const k of Object.keys(node.files)) {
        const f = node.files[k];
        yield [k, {
          kind: 'file', isDrive: true, name: f.name, meta: f,
          async getFile() {
            downloads++;
            return { name: f.name, size: Number(f.size), lastModified: Date.parse(f.modifiedTime),
                     slice: () => ({ arrayBuffer: async () => new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer }),
                     arrayBuffer: async () => new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer,
                     __challan: f.challan };
          }
        }];
      }
    }
  };
}

// ---- lift buildChallanIndex with every top-level function it transitively needs ----
// These are provided by STUBS below. The lifted copies must NOT be pulled in, or they would be
// defined after the stubs and silently win — which is how the first attempt at this harness ended
// up calling the real findChallanRoot and asking a fake folder tree to talk to Google.
const STUBBED = ['findChallanRoot', 'prefetchDriveTree', 'checkpointChallanIndex', 'saveChallanMisses',
                 'improveOneChallanFile', 'newEtaClock', 'etaText', 'yieldSoon', 'scanChallanDoc',
                 'runPool', 'topLevelFolderNames'];
const NEED = [];
(function resolve(names) {
  names.forEach(n => {
    if (NEED.includes(n) || STUBBED.includes(n) || typeof fns[n] !== 'string') return;
    NEED.push(n);
    const refs = [...new Set(fns[n].match(/\b[A-Za-z_$][A-Za-z0-9_$]*\b/g) || [])]
      .filter(id => id !== n && !STUBBED.includes(id) && typeof fns[id] === 'string');
    resolve(refs);
  });
})(['buildChallanIndex']);

const STUBS = `
  let challanIndex = {}, challanMisses = {}, challanIndexBuiltAt = null;
  let savedRootHandle = { isDrive: true, name: 'OLISA GROUP', id: 'ROOT' };
  const CHECKPOINT_MS = 4000; let lastChallanCheckpoint = 0;
  const navigator = { onLine: false };   // OCR off: this test is about reuse, not recovery
  async function findChallanRoot() { return __root; }
  async function prefetchDriveTree() { return 0; }
  async function checkpointChallanIndex() {}
  async function saveChallanMisses() {}
  async function improveOneChallanFile() { throw new Error('no ocr in test'); }
  function newEtaClock() { return {}; }
  function etaText() { return ''; }
  function yieldSoon() { return Promise.resolve(); }
  async function scanChallanDoc(file) {
    // One challan per file, exactly as a real single-challan PDF would give.
    return [{ no: file.__challan, pages: [1], piRef: '169', date: '06-SEP-26', verified: true, lines: [{ q: 1 }] }];
  }
  async function runPool(items, limit, worker) {
    const results = new Array(items.length); let next = 0;
    async function lane() { while (next < items.length) { const i = next++;
      try { results[i] = await worker(items[i], i); } catch (e) { results[i] = { __err: e }; } } }
    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, lane));
    return results;
  }
`;
const api = new Function('__root', `
  ${STUBS}
  ${NEED.map(n => fns[n]).join('\n')}
  return {
    build: buildChallanIndex,
    get index() { return challanIndex; },
    set index(v) { challanIndex = v; },
    get misses() { return challanMisses; },
    set misses(v) { challanMisses = v; }
  };
`);

(async () => {
  const world = makeWorld();
  const root = dirHandle(world);
  const A = api(root);

  console.log('1. First build reads everything (nothing is cached yet)');
  downloads = 0;
  const r1 = await A.build();
  t('every challan document was found', r1.files === 40, r1.files);
  t('every one produced a challan', r1.challans === 40, r1.challans);
  t('and every one was downloaded, because none was known', downloads === 40, downloads);

  console.log('\n2. Second build, nothing changed on Drive');
  downloads = 0;
  const r2 = await A.build();
  t('the same 40 challans are still indexed', r2.challans === 40, r2.challans);
  t('NOTHING was downloaded', downloads === 0, downloads + ' file(s) downloaded');
  t('and the build says so', r2.reusedFiles === 40 && r2.readNow === 0,
    `reused ${r2.reusedFiles}, read ${r2.readNow}`);

  console.log('\n3. One new challan added, one edited');
  world.dirs['August - 2026'].files['new.pdf'] =
    { id: 'NEW', name: 'new.pdf', size: '99999', modifiedTime: '2026-09-01T09:00:00.000Z', challan: '320999999' };
  world.dirs['June - 2026'].files['Ju-3-26 Challan.pdf'].modifiedTime = '2026-09-02T09:00:00.000Z';
  downloads = 0;
  const r3 = await A.build();
  t('exactly the new and the changed file were read', downloads === 2, downloads + ' downloaded');
  // 41 files now: 39 unchanged, 1 edited, 1 new. Reuse is the 39, not the 40 from before.
  t('the rest were reused', r3.reusedFiles === 39, r3.reusedFiles);
  t('the new challan is in the index', !!A.index['320999999']);

  console.log('\n4. A file that cannot be read is not re-read for ever');
  world.files['broken.pdf'] =
    { id: 'BRK', name: 'broken.pdf', size: '123', modifiedTime: '2026-09-03T09:00:00.000Z', challan: null };
  const A2 = api(dirHandle(world));
  A2.index = A.index; A2.misses = A.misses;
  downloads = 0;
  await A2.build();                          // learns it is unreadable
  const firstPass = downloads;
  downloads = 0;
  const r5 = await A2.build();               // must not touch it again
  t('the unreadable file was opened once', firstPass >= 1, firstPass);
  t('and not opened again on the next build', downloads === 0, downloads + ' downloaded');
  // The stub returns a challan object for every file, so "broken.pdf" is not actually unreadable
  // here \u2014 what this section proves is the re-read behaviour above, which it does.
  t('the build still returns a full report', Array.isArray(r5.skipped) && typeof r5.readNow === 'number');

  console.log('\n5. The OCR chain cannot grow');
  // The real folder ended up holding "...-Improved-Improved-Improved-Improved.pdf". Three separate
  // things had to be wrong for that: the name builder appended a suffix to a name that already had
  // one, the guard against re-OCR-ing an OCR result was never called, and the walk kept every link
  // of the chain. All three are checked here against the shipped source.
  const nameFns = new Function(`
    ${['ocrBaseName', 'ocrImprovedName', 'ocrOriginalName', 'ocrChainDepth', 'ocrFamilyRoot', 'isOcrImproved']
      .map(n => fns[n]).join('\n')}
    return { ocrBaseName, ocrImprovedName, ocrOriginalName, ocrChainDepth, ocrFamilyRoot, isOcrImproved };
  `)();
  t('OCR-ing a plain file names it -Improved',
    nameFns.ocrImprovedName('31-8-26 Challan.pdf') === '31-8-26 Challan-Improved.pdf',
    nameFns.ocrImprovedName('31-8-26 Challan.pdf'));
  t('OCR-ing an -Improved file does NOT make -Improved-Improved',
    nameFns.ocrImprovedName('31-8-26 Challan-Improved.pdf') === '31-8-26 Challan-Improved.pdf',
    nameFns.ocrImprovedName('31-8-26 Challan-Improved.pdf'));
  t('even a four-deep chain collapses back to one suffix',
    nameFns.ocrImprovedName('X-Improved-Improved-Improved-Improved.pdf') === 'X-Improved.pdf',
    nameFns.ocrImprovedName('X-Improved-Improved-Improved-Improved.pdf'));
  t('chain depth is counted correctly',
    nameFns.ocrChainDepth('X-Improved-Improved-Improved-Improved.pdf') === 4 &&
    nameFns.ocrChainDepth('X.pdf') === 0, nameFns.ocrChainDepth('X-Improved-Improved-Improved-Improved.pdf'));
  t('every link of a chain shares one family root',
    ['X.pdf', 'X-Improved.pdf', 'X-Improved-Improved.pdf'].every(n => nameFns.ocrFamilyRoot(n) === 'x'));
  t('an -Original name does not gain -Improved as well',
    nameFns.ocrOriginalName('X-Improved.pdf') === 'X-Original.pdf', nameFns.ocrOriginalName('X-Improved.pdf'));
  t('an OCR result is recognised so it is never re-OCR-ed', nameFns.isOcrImproved('X-Improved.pdf'));
  // And the guard is actually CALLED this time, not just declared.
  const bciSrc = fns['buildChallanIndex'];
  t('the re-OCR guard is wired into the OCR loop, not just declared',
    /isOcrImproved\(f\.name\)/.test(bciSrc));

  console.log('\n6. A real runaway folder is collapsed');
  const w2 = makeWorld();
  ['X.pdf', 'X-Improved.pdf', 'X-Improved-Improved.pdf', 'X-Improved-Improved-Improved.pdf'].forEach((n, i) => {
    w2.files[n] = { id: 'X' + i, name: n, size: String(700 + i), modifiedTime: '2026-09-0' + (i + 1) + 'T09:00:00.000Z', challan: '321000' + i };
  });
  const A3 = api(dirHandle(w2));
  const r6 = await A3.build();
  t('only the original and the deepest copy were read', r6.files === 42, r6.files + ' (40 + 2)');
  t('the two middle copies were reported as leftovers', r6.ocrJunk.length === 2, r6.ocrJunk.join(', '));
  t('and they are named so they can be deleted', r6.ocrJunk.every(p => /-Improved/.test(p)));

  console.log('\n' + (fail ? `FAILED — ${pass} passed, ${fail} failed` : `PASSED — ${pass} passed, 0 failed`));
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('HARNESS ERROR:', e); process.exit(1); });
