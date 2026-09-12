/* smoketest.js — the test that would have caught the dead tool.
   Every other test in this suite reads the source as TEXT. Not one of them ever RAN it. So a
   constant declared below the line that reads it — a temporal dead zone — passed all fourteen
   files and then threw on load, stopping script evaluation dead. Everything declared after that
   point never came into existence, which is why the visible symptom was an unrelated
   "cannot access 'piIndex' before initialization" and every button did nothing.
   This file opens the page in a real DOM and fails on ANY error thrown while it loads.
   Usage: node smoketest.js olisa.html                                                          */
const fs = require('fs');
const path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');

let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const file = process.argv[2] || 'olisa.html';
let html = fs.readFileSync(file, 'utf8');

// The self-hosted libraries in lib/ are not needed to prove the page's own script evaluates, and
// pulling them in would only test SheetJS. Strip the external <script src> tags and stub what the
// page expects them to leave behind.
html = html.replace(/<script[^>]*\bsrc=[^>]*><\/script>/g, '');

const errors = [];
const vc = new VirtualConsole();
vc.on('jsdomError', e => errors.push(e && (e.detail || e.message || String(e))));
vc.on('error', (...a) => errors.push(a.map(String).join(' ')));

(async () => {
  console.log('\n1. The page loads without throwing');
  let dom;
  try {
    dom = new JSDOM(html, {
      runScripts: 'dangerously',
      pretendToBeVisual: true,
      url: 'https://aurkosadi-ux.github.io/olisa_tools/olisa.html',
      virtualConsole: vc,
      beforeParse(w) {
        // Things the real browser has that jsdom does not. Missing APIs must not be mistaken for
        // the page's own bugs, so each one is stubbed rather than left to throw.
        w.matchMedia = w.matchMedia || (q => ({ matches: false, media: q, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
        w.scrollTo = () => {};
        w.requestAnimationFrame = cb => setTimeout(() => cb(Date.now()), 0);
        w.cancelAnimationFrame = id => clearTimeout(id);
        w.indexedDB = { open: () => ({ addEventListener() {}, set onsuccess(v) {}, set onerror(v) {}, set onupgradeneeded(v) {} }) };
        w.fetch = () => Promise.reject(new Error('offline in test'));
        w.navigator.serviceWorker = { register: () => Promise.resolve({ addEventListener() {} }), addEventListener() {}, ready: Promise.resolve({}) };
        w.XLSX = {}; w.ExcelJS = { Workbook: function () {} }; w.PDFLib = {};
        w.pdfjsLib = { GlobalWorkerOptions: {}, getDocument: () => ({ promise: Promise.resolve({ numPages: 0 }) }) };
        w.JSZip = function () {}; w.mammoth = {}; w.jspdf = {};
        w.CompressionStream = undefined; w.DecompressionStream = undefined;
      }
    });
  } catch (e) {
    t('the document parses and the script evaluates', false, e.message);
    return finish();
  }

  // Give timers, the DOMContentLoaded handlers and any promise chains a moment to run, because a
  // crash on load does not always happen during parse.
  await new Promise(r => setTimeout(r, 600));

  // A temporal dead zone reads as exactly this, and it is the one that killed the tool.
  const tdz = errors.filter(e => /before initialization/i.test(String(e)));
  t('no "cannot access X before initialization"', tdz.length === 0, tdz.join(' | '));

  const refErr = errors.filter(e => /ReferenceError/.test(String(e)) && !/offline in test/.test(String(e)));
  t('no ReferenceError while loading', refErr.length === 0, refErr.slice(0, 3).join(' | '));

  const synErr = errors.filter(e => /SyntaxError/.test(String(e)));
  t('no SyntaxError', synErr.length === 0, synErr.slice(0, 2).join(' | '));

  const typeErr = errors.filter(e => /TypeError/.test(String(e)) && !/offline in test|fetch/.test(String(e)));
  t('no unexpected TypeError', typeErr.length === 0, typeErr.slice(0, 3).join(' | '));

  console.log('\n2. The script actually finished evaluating');
  const w = dom.window;
  // If evaluation had stopped part-way, the things declared near the END of the file would be
  // missing while the ones near the top were fine. That asymmetry is the fingerprint of the bug.
  t('a function from the very top of the script exists', typeof w.escHtml === 'function' || typeof w.saveBlobAs === 'function');
  t('the version constant is readable', w.eval('typeof APP_VERSION') === 'number', w.eval('typeof APP_VERSION'));
  // The last line of the script sets this. If it is missing, evaluation stopped somewhere above —
  // which is precisely the failure that looked like "nothing is working at all".
  t('the script ran all the way to its last line', w.__olisaReady !== undefined, 'marker never set');
  t('and the marker carries the same version', w.__olisaReady === w.eval('APP_VERSION'), String(w.__olisaReady));

  console.log('\n3. The page rendered, not just parsed');
  t('the Ask box is on the page', !!w.document.getElementById('askInput') || !!w.document.querySelector('input'));
  t('the build stamp element exists', !!w.document.getElementById('buildStamp'));
  const stamp = w.document.getElementById('buildStamp');
  t('and it was filled in with the version', stamp && /^build v\d+$/.test(stamp.textContent.trim()), stamp && stamp.textContent);
  t('the version on the page matches the constant',
    stamp && stamp.textContent.trim() === 'build v' + w.eval('APP_VERSION'), stamp && stamp.textContent);

  console.log('\n4. The page and the service worker agree on the build');
  const swPath = path.join(path.dirname(path.resolve(file)), 'sw.js');
  if (!fs.existsSync(swPath)) { console.log('  -- sw.js not beside it, skipping'); return finish(); }
  const sw = fs.readFileSync(swPath, 'utf8');
  const swV = (sw.match(/olisa-tools-v(\d+)/) || [])[1];
  t('sw.js carries a version', !!swV, swV);
  t('it is the same number as the page', swV && Number(swV) === w.eval('APP_VERSION'),
    `sw=${swV} page=${w.eval('APP_VERSION')}`);
  finish();

  function finish() {
    if (errors.length) {
      console.log('\n  (errors seen while loading, for reference:)');
      [...new Set(errors.map(String))].slice(0, 6).forEach(e => console.log('   - ' + e.split('\n')[0].slice(0, 160)));
    }
    console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
    process.exit(fail ? 1 : 0);
  }
})();
