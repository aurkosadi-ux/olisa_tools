/* loadtest.js — proves the loading fixes on the ACTUAL shipped code, lifted from olisa.html.
   Covers: the ETA no longer inflates when the page is backgrounded, the wake lock is taken and
   released around a build, and checkpoints are throttled but always flushed at the end. */
const fs = require('fs');
const acorn = require('acorn');

const HTML = process.argv[2] || 'olisa.html';
const html = fs.readFileSync(HTML, 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
  .map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const ast = acorn.parse(src, { ecmaVersion: 2022, locations: true });

const fns = {}, vars = {};
(function scan(node) {
  if (!node || typeof node.type !== 'string') return;
  if (node.type === 'FunctionDeclaration' && node.id) fns[node.id.name] = src.slice(node.start, node.end);
  if (node.type === 'VariableDeclaration') node.declarations.forEach(d => {
    if (d.id.type === 'Identifier') vars[d.id.name] = src.slice(node.start, node.end);
  });
  for (const k of Object.keys(node)) {
    if (k === 'loc') continue;
    const v = node[k];
    if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && scan(c));
    else if (v && typeof v.type === 'string') scan(v);
  }
})(ast);

let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

// ---------- fake page lifecycle ----------
const listeners = [];
const fakeDoc = {
  visibilityState: 'visible',
  addEventListener: (ev, fn) => { if (ev === 'visibilitychange') listeners.push(fn); }
};
let NOW = 1000000;
const realNow = Date.now;
Date.now = () => NOW;
function advance(ms) { NOW += ms; }
function hide() { fakeDoc.visibilityState = 'hidden'; listeners.forEach(f => f()); }
function show() { fakeDoc.visibilityState = 'visible'; listeners.forEach(f => f()); }

// ---------- wake lock stub ----------
let wakeRequests = 0, wakeReleases = 0, released = false;
const fakeNav = {
  wakeLock: {
    request: async () => { wakeRequests++; released = false; return { addEventListener() {}, release() { wakeReleases++; released = true; } }; }
  }
};

const NEED = ['newEtaClock', 'etaActiveMs', 'etaText', 'acquireBuildWakeLock', 'releaseBuildWakeLock'];
const missingF = NEED.filter(n => !fns[n]);
if (missingF.length) { console.error('FAIL: could not lift ' + missingF.join(', ')); process.exit(1); }
if (!vars['etaClocks']) { console.error('FAIL: etaClocks not found'); process.exit(1); }
if (!vars['CHECKPOINT_MS']) { console.error('FAIL: CHECKPOINT_MS not found'); process.exit(1); }

const api = new Function('document', 'navigator', `
  let indexBuilding = false;
  let buildWakeLock = null;
  ${vars['etaClocks']}
  ${NEED.map(n => fns[n]).join('\n')}
  ${fns['checkpointStyleIndex'] ? '' : ''}
  document.addEventListener('visibilitychange', () => {
    const hidden = document.visibilityState === 'hidden';
    etaClocks.forEach(c => {
      if (hidden) { if (c.since) { c.activeMs += Date.now() - c.since; c.since = 0; } }
      else if (!c.since) { c.since = Date.now(); }
    });
  });
  return { newEtaClock, etaText, etaActiveMs, acquireBuildWakeLock, releaseBuildWakeLock,
           setBuilding: v => { indexBuilding = v; }, clocks: () => etaClocks };
`)(fakeDoc, fakeNav);

// The listener under test is the one shipped in olisa.html; re-declared above verbatim because it
// sits at statement level rather than inside a function. Guard that they have not diverged.
const shippedListener = src.includes("if (hidden) { if (c.since) { c.activeMs += Date.now() - c.since; c.since = 0; } }");
console.log('1. ETA clock pauses while the page is hidden');
t('the shipped visibilitychange handler matches the one under test', shippedListener);

// ---------- the exact scenario Sadi reported ----------
// 100 folders. 40 done in 40s of real work => 60s of work left.
const c = api.newEtaClock();
advance(40000);
let txt = api.etaText(c, 40, 100);
const before = txt;
t('estimate shown after real progress', /remaining/.test(txt), txt);
const secs = s => { const m = s.match(/(?:(\d+)m )?(\d+)s/); return m ? (+(m[1] || 0)) * 60 + (+m[2]) : null; };
t('60s of work left reads as ~60s', Math.abs(secs(txt) - 60) <= 3, txt);

// Now background the app for 30s. Nothing progresses: throttled.
hide();
advance(30000);
show();
const txt2 = api.etaText(c, 40, 100);
t('30s in the background does NOT inflate the estimate', secs(txt2) === secs(before), before + ' -> ' + txt2);

// Two more round trips, as reported ("every time I do this, it resets").
hide(); advance(45000); show();
hide(); advance(60000); show();
const txt3 = api.etaText(c, 40, 100);
t('repeated app switching still does not inflate it', secs(txt3) === secs(before), txt3);
t('hidden time is excluded from the active clock', api.etaActiveMs(c) === 40000, api.etaActiveMs(c));

// Real progress still moves the number down.
advance(30000);
const txt4 = api.etaText(c, 70, 100);
t('genuine progress lowers the estimate', secs(txt4) < secs(before), before + ' -> ' + txt4);
t('no estimate once the job is finished', api.etaText(c, 100, 100) === '');

// A clock created while hidden must not claim "0s remaining".
fakeDoc.visibilityState = 'hidden';
const c2 = api.newEtaClock();
t('a clock started while hidden gives no false estimate', api.etaText(c2, 5, 100) === '');
fakeDoc.visibilityState = 'visible';

// Old behaviour, for contrast — this is what was shipping.
const oldEta = (startMs, done, total) => { const rate = (Date.now() - startMs) / done; return Math.round((total - done) * rate / 1000); };
const t0 = NOW - 175000; // 40 done, 40s work + 135s backgrounded
t('(reference) the old wall-clock formula would have said ' + oldEta(t0, 40, 100) + 's', oldEta(t0, 40, 100) > 60);

console.log('\n2. Clock bookkeeping');
t('clock list does not grow without bound', (() => {
  for (let i = 0; i < 50; i++) api.newEtaClock();
  return api.clocks().length <= 8;
})(), api.clocks().length);
t('mark window stays bounded on a long run', (() => {
  const cc = api.newEtaClock();
  for (let i = 1; i < 500; i++) { advance(100); api.etaText(cc, i, 5000); }
  return cc.marks.length <= 30;
})(), 'marks');

console.log('\n3. Screen wake lock around a build');
(async () => {
  await api.acquireBuildWakeLock();
  t('wake lock taken when the build starts', wakeRequests === 1, wakeRequests);
  await api.acquireBuildWakeLock();
  t('never taken twice over (no leak)', wakeRequests === 1, wakeRequests);
  api.releaseBuildWakeLock();
  t('released when the build ends', wakeReleases === 1 && released);
  api.releaseBuildWakeLock();
  t('releasing twice is harmless', wakeReleases === 1, wakeReleases);

  // wake lock must never throw into the build, even where unsupported
  const noWake = new Function('navigator', `
    let buildWakeLock = null;
    ${fns['acquireBuildWakeLock']}
    ${fns['releaseBuildWakeLock']}
    return { acquireBuildWakeLock, releaseBuildWakeLock };
  `)({});
  let threw = false;
  try { await noWake.acquireBuildWakeLock(); noWake.releaseBuildWakeLock(); } catch (e) { threw = true; }
  t('a browser without wakeLock does not break the build', !threw);

  console.log('\n4. Checkpoint throttle');
  const writes = [];
  const cp = new Function('idbSet', 'piIndex', 'PI_CACHE_VERSION', `
    ${vars['CHECKPOINT_MS']}
    ${vars['lastCheckpoint']}
    ${fns['checkpointStyleIndex']}
    ${fns['checkpointPiIndex']}
    return { checkpointStyleIndex, checkpointPiIndex };
  `)(async (k, v) => { writes.push({ k, partial: v && v.partial }); }, { a: 1 }, 3);

  const idx = { S1: ['124'] };
  await cp.checkpointStyleIndex(idx);
  const afterFirst = writes.length;
  for (let i = 0; i < 200; i++) await cp.checkpointStyleIndex(idx);   // same instant
  t('a burst of folders does not cause a burst of writes', writes.length === afterFirst, writes.length);
  advance(5000);
  await cp.checkpointStyleIndex(idx);
  t('a write does happen once the interval passes', writes.length === afterFirst + 1, writes.length);
  await cp.checkpointStyleIndex(idx, true);
  t('a forced flush always writes', writes.length === afterFirst + 2, writes.length);
  t('checkpoints are flagged partial', writes[0].partial === true);
  t('the final forced write is flagged complete', writes[writes.length - 1].partial === false);

  // a failing store must not abort the build
  const cpFail = new Function('idbSet', 'piIndex', 'PI_CACHE_VERSION', `
    ${vars['CHECKPOINT_MS']}
    ${vars['lastCheckpoint']}
    ${fns['checkpointStyleIndex']}
    ${fns['checkpointPiIndex']}
    return { checkpointPiIndex };
  `)(async () => { throw new Error('QuotaExceededError'); }, {}, 3);
  let blew = false;
  try { await cpFail.checkpointPiIndex(true); } catch (e) { blew = true; }
  t('a full disk does not kill the build', !blew);

  Date.now = realNow;
  console.log('\n' + (fail ? 'FAILED ' + fail + ' / ' : 'PASSED all ') + (pass + fail) + ' checks');
  process.exit(fail ? 1 : 0);
})();
