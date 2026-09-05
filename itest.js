/* itest.js — whole-file integrity. Catches collateral damage from an edit: broken syntax anywhere,
   a DOM id referenced but never defined, a duplicated function, an artifact left behind by a patch. */
const fs = require('fs');
const acorn = require('acorn');

const HTML = process.argv[2] || 'olisa.html';
const html = fs.readFileSync(HTML, 'utf8');
let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

console.log('1. Syntax');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
let parsed = [];
blocks.forEach((b, i) => {
  try { parsed.push(acorn.parse(b, { ecmaVersion: 2022, locations: true })); t('inline script block ' + i + ' parses (' + b.length + ' chars)', true); }
  catch (e) { t('inline script block ' + i + ' parses', false, e.message); }
});

console.log('\n2. No duplicate function declarations in the SAME scope');
// Each tab is its own IIFE and legitimately has its own handleFiles/generate/renderPreview —
// that is not a duplicate. Only two declarations of the same name in one scope are a bug.
const dupes = [];
parsed.forEach(ast => {
  (function scope(node, names) {
    const here = names || new Map();
    const body = node.body && node.body.body ? node.body.body : (node.body || []);
    (Array.isArray(body) ? body : []).forEach(st => {
      if (st.type === 'FunctionDeclaration' && st.id) {
        if (here.has(st.id.name)) dupes.push(st.id.name + ' (line ' + st.loc.start.line + ')');
        here.set(st.id.name, true);
      }
    });
    (function descend(n) {
      if (!n || typeof n.type !== 'string') return;
      if (n !== node && (n.type === 'FunctionDeclaration' || n.type === 'FunctionExpression' || n.type === 'ArrowFunctionExpression')) { scope(n, new Map()); return; }
      for (const k of Object.keys(n)) {
        if (k === 'loc') continue;
        const v = n[k];
        if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && descend(c));
        else if (v && typeof v.type === 'string') descend(v);
      }
    })(node);
  })(ast, new Map());
});
t('no function is declared twice in one scope', dupes.length === 0, dupes.slice(0, 5).join(', '));

console.log('\n3. DOM ids referenced actually exist in the markup');
// Static markup ids, PLUS ids the code injects at runtime (innerHTML template literals and
// element.id = '...'), which are just as real as ones written in the page.
const definedIds = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
const runtimeIds = new Set([...html.matchAll(/\.id\s*=\s*['"]([^'"]+)['"]/g)].map(m => m[1]));
const used = new Set([...html.matchAll(/getElementById\(\s*['"]([^'"]+)['"]\s*\)/g)].map(m => m[1]));
const ghosts = [...used].filter(id => !definedIds.has(id) && !runtimeIds.has(id));
t('no getElementById targets a non-existent id', ghosts.length === 0, ghosts.join(', '));

console.log('\n4. Patch hygiene');
t('no merge-conflict markers', !/^(<{7}|={7}|>{7})/m.test(html));
t('no TODO/FIXME left in the shipped file', !/\b(TODO|FIXME|XXX):/.test(html));
t('no stray console.log in the app code', (html.match(/console\.log\(/g) || []).length === 0,
  (html.match(/console\.log\(/g) || []).length + ' found');
t('no debugger statements', !/\bdebugger\b/.test(html));

console.log('\n5. Business rules that must never regress');
t('delivering-party language: no "received N out of"', !/received\s+\$?\{?[\w.]+\}?\s+(pcs\s+)?out of/i.test(html));
t('manual challans still tagged "(Manual)"', /\(Manual\)/.test(html));
t('Master File is never written to Drive', !/files\/[^'"`]*\?uploadType/.test(html) || /appProperties|Olisa Inventory Master File/.test(html));
t('short and excess still summed as two separate totals',
  /Math\.min\(effectiveShort\(r\), 0\)/.test(html) && /Math\.max\(effectiveShort\(r\), 0\)/.test(html));
t('excess coverage pass is wired into reconciliation', /reconcileExcessAgainstShorts\(\);/.test(html));
t('reconciliation resets _consumed each run', /delete r\._consumed/.test(html));

console.log('\n6. Loading fixes present');
t('ETA runs on an active-time clock', /function newEtaClock/.test(html) && !/const rate = \(Date\.now\(\) - startMs\) \/ done/.test(html));
t('wake lock taken and released around a build', /acquireBuildWakeLock\(\)/.test(html) && /releaseBuildWakeLock\(\)/.test(html));
t('build state is unwound when a build throws', /buildStyleIndex\(\)\.catch/.test(html));
t('indexes are checkpointed during the build', /checkpointPiIndex\(\)/.test(html) && /checkpointStyleIndex\(index\)/.test(html));
t('an interrupted build is flagged to the user', /piIndexPartial/.test(html) && /Last build was interrupted/.test(html));
t('away-links open in a new tab while building', /__olisaBuilding/.test(html));

console.log('\n7. Challan copy feature');
t('challan numbers render as links', /class="ch-link"/.test(html));
t('links are wired by delegation, not per-render', /closest\('a\.ch-link'\)/.test(html));
t('pdf-lib is loaded for splitting', /pdf-lib/.test(html));
t('page grouping treats headerless pages as continuations', /current\.pages\.push\(p\)/.test(html));
t('challan index is fingerprint-cached', /e\.fp === fp/.test(html));
t('plain-text challan uses are untouched', /challanDisp\(r\)/.test(html));
t('challan index survives a reload', /idbGet\('challanIndex'\)/.test(html));
t('".doc" hand challans are indexed too', /pdf\|docx\?\|rtf/.test(html));
t('RTF is read with a real tokenizer, not a regex strip', /function rtfToText/.test(html) && /ucSkip/.test(html));
t('the skip-depth double-push bug stays fixed', /markSkip/.test(html));
t('format is decided by magic bytes, not extension', /startsWith\('%PDF'\)/.test(html) && /startsWith\('\{\\\\rt'\)/.test(html));
t('RTF renders in the preview instead of "no preview"', /head4\.startsWith\('\{\\\\rt'\)/.test(html));
t('unreadable binary Word files are named, not hidden', /binaryDoc/.test(html));
t('a text challan is served whole (nothing to split)', /if \(!entry\.pages\)/.test(html));
// The preview must outrank the PERSISTENT chrome — the sticky nav, the bell, the offline banner,
// the toast layer, the Drive log-out link. Transient overlays that deliberately sit on top of
// everything (the folder picker, the fatal-error box) are not in scope, and the particle burst is
// pointer-events:none so it can never swallow a click.
t('preview outranks every persistent fixed layer', (() => {
  const modal = Number((html.match(/\.pdfmodal \{[^}]*z-index:(\d+)/) || [])[1]);
  if (!modal) return false;
  const chrome = ['.appswitch', '.bellwrap', '.offbanner', '.toast', '.anote'].map(sel => {
    const m = html.match(new RegExp('\\' + sel + ' \\{[^}]*z-index:\\s*(\\d+)'));
    return m ? Number(m[1]) : 0;
  });
  const logout = Number((html.match(/driveLogoutLink[^>]*z-index:(\d+)/) || [])[1] || 0);
  return [...chrome, logout].every(z => z < modal);
})(), 'modal must outrank .appswitch/.bellwrap/.offbanner/.toast/.anote/log-out link');
t('preview closes on Escape', /ev\.key === 'Escape'/.test(html));
t('preview closes on a backdrop click', /ev\.target === modal/.test(html));
t('the Escape listener is removed with the modal', /removeEventListener\('keydown', onKey, true\)/.test(html));
t('preview still offers Download', /id="pvDl"/.test(html));
t('short hand-challan numbers are linkable', /digits\.length >= 1/.test(html));
t('challan copies are built by the same button as the style index', /await buildChallanIndex\(\(d, total, found, eta\)/.test(html));
t('a challan-folder failure does not lose the style index', /challanError = e\.message/.test(html));
t('one merged index line, not two pills', / and \$\{chCount\} challan copies/.test(html) && !/Challan copies: not read/.test(html));

console.log('\n8. Service worker');
const sw = fs.readFileSync(process.argv[3] || 'sw.js', 'utf8');
const ver = (sw.match(/olisa-tools-v(\d+)/) || [])[1];
t('cache version bumped past v30', ver && Number(ver) > 30, 'v' + ver);
const swCode = sw.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
t('shell is cached per-entry, not atomically', !/\.addAll\(/.test(swCode));
t('shell entries fail independently', /\.add\(u\)\.catch/.test(swCode));

console.log('\n' + (fail ? 'FAILED ' + fail + ' / ' : 'PASSED all ') + (pass + fail) + ' checks');
process.exit(fail ? 1 : 0);
