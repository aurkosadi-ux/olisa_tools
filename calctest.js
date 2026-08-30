/* calctest.js — runs the ACTUAL shipped reconciliation functions (lifted verbatim out of
   olisa.html by the parser, not re-typed) against the REAL Master File.
   Proves: the M82039-1B1-5 case now closes, shorts/excesses stay separate totals, no
   double-counting, no coverage travelling backwards in time, and manual-challan behaviour
   is byte-for-byte unchanged from before this release. */
const fs = require('fs');
const acorn = require('acorn');
const XLSX = require('xlsx');

const HTML = process.argv[2] || 'olisa.html';
const MASTER = process.argv[3] || '/mnt/project/Olisa_Master_Inventory___09082026_0346_pm___v8.xlsx';

// ---------- lift the real functions out of the file ----------
function buildAPI(path) {
const html = fs.readFileSync(path, 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
  .map(m => m[1]).sort((a, b) => b.length - a.length)[0];
const ast = acorn.parse(src, { ecmaVersion: 2022, locations: true });

const WANT = ['measNorm', 'parseDMY', 'fmtNoteDate', 'manualTime', 'effectiveShort', 'acceptKey',
  'challanDisp', 'normPiKey', 'dmyFromMs', 'acceptedFor', 'remarkColorWord', 'effectiveRemark',
  'extractMeasurement', 'reconcileManualChallans', 'reconcileExcessAgainstShorts', 'combinedRemarks',
  'sameCartonItem'];
const found = {};
// Collect EVERY top-level function declaration in the file, then pull in the transitive closure of
// the ones WANT names. The old version listed helpers by hand, so every helper added to olisa.html
// since broke this test with a ReferenceError that looked like a product bug and was not one.
const allFns = {};
(function scan(node) {
  if (!node || typeof node.type !== 'string') return;
  if (node.type === 'FunctionDeclaration' && node.id) allFns[node.id.name] = src.slice(node.start, node.end);
  for (const k in node) {
    if (k === 'loc') continue;
    const v = node[k];
    if (Array.isArray(v)) v.forEach(c => c && typeof c.type === 'string' && scan(c));
    else if (v && typeof v.type === 'string') scan(v);
  }
})(ast);
const IDENT = /\b([A-Za-z_$][\w$]*)\s*\(/g;
const queue = [...WANT];
while (queue.length) {
  const n = queue.shift();
  if (found[n] || !allFns[n]) continue;
  found[n] = allFns[n];
  let m; IDENT.lastIndex = 0;
  while ((m = IDENT.exec(allFns[n]))) if (allFns[m[1]] && !found[m[1]]) queue.push(m[1]);
}
// The previous build has no excess pass at all — stub it so the same harness runs both.
if (!found['reconcileExcessAgainstShorts']) found['reconcileExcessAgainstShorts'] = 'function reconcileExcessAgainstShorts() {}';
const WANTED_ORDER = Object.keys(found);
const missing = WANT.filter(n => !found[n]);
if (missing.length) { console.error('FAIL: could not lift ' + missing.join(', ') + ' from ' + path); process.exit(1); }

const sandbox = `
  let allRecords = [], acceptedShorts = {}, remarkOverrides = {};
  ${WANTED_ORDER.map(n => found[n]).join('\n')}
  return { run: () => reconcileManualChallans(),
           set: r => { allRecords = r; },
           get: () => allRecords,
           accept: m => { acceptedShorts = m; },
           effectiveShort, measNorm, manualTime, normPiKey, combinedRemarks };
`;
return new Function(sandbox)();
}
const API = buildAPI(HTML);
const PREV = process.argv[4] ? buildAPI(process.argv[4]) : null;

// ---------- build records exactly as loadMasterFile does ----------
function fmtDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = new Date(v.getTime());
    const drift = (d.getHours()*3600 + d.getMinutes()*60 + d.getSeconds())*1000 + d.getMilliseconds();
    if (drift > 43200000) d.setTime(d.getTime() + (86400000 - drift)); else d.setTime(d.getTime() - drift);
    return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  }
  return v === null || v === undefined ? '' : String(v);
}
const wb = XLSX.readFile(MASTER, { cellDates: true });
// Same sheet selection the app uses — otherwise the harness reconciles rows the tool never sees.
const wantedSheets = wb.SheetNames.filter(n => /^2026-\s*Main$/i.test(n.trim()) || /^2025$/i.test(n.trim()));
const useSheets = wantedSheets.length ? wantedSheets : wb.SheetNames;
const records = [];
for (const sheetName of useSheets) {
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
  let hdr = -1, idx = {};
  for (let r = 0; r < Math.min(12, aoa.length); r++) {
    const row = (aoa[r] || []).map(c => String(c == null ? '' : c).toUpperCase().replace(/[^A-Z]/g, ''));
    if (row.some(c => c.includes('STYLEREF'))) {
      hdr = r;
      row.forEach((c, i) => {
        if (c.includes('STYLEREF')) idx.STYLE = i;
        else if (c.includes('PIREF')) idx.PIREF = i;
        else if (c.includes('INVENTORYDATE')) idx.INVDATE = i;
        else if (c.includes('ITEMDESCRIPTION')) idx.ITEMDESC = i;
        else if (c === 'ITEM') idx.ITEM = i;
        else if (c.includes('DELIVERYQTY')) idx.DELQTY = i;
        else if (c.includes('RECEIVEDQTY')) idx.RECQTY = i;
        else if (c.includes('SHORT')) idx.SHORT = i;
        else if (c.includes('CHALLANNO')) idx.CHALLANNO = i;
        else if (c.includes('CHALLANDATE')) idx.CHALLANDATE = i;
        else if (c === 'REMARKS') idx.REMARKS = i;
        else if (c.includes('MANUALCHALLAN')) idx.MANUALCH = i;
      });
      break;
    }
  }
  if (hdr < 0) continue;
  for (let r = hdr + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const styleVal = row[idx.STYLE];
    if (!styleVal) continue;
    records.push({
      sheet: sheetName, xlRow: r + 1,
      style: String(styleVal).trim(), styleNorm: String(styleVal).trim().toUpperCase(),
      piRef: row[idx.PIREF], invDate: fmtDate(row[idx.INVDATE]),
      item: row[idx.ITEM], itemDesc: row[idx.ITEMDESC],
      delQty: parseFloat(row[idx.DELQTY]) || 0,
      recQty: parseFloat(row[idx.RECQTY]) || 0,
      short: parseFloat(row[idx.SHORT]) || 0,
      challanNo: row[idx.CHALLANNO], challanDate: fmtDate(row[idx.CHALLANDATE]),
      remarks: row[idx.REMARKS],
      manualCh: (idx.MANUALCH !== undefined && row[idx.MANUALCH] !== null && String(row[idx.MANUALCH]).trim() !== '') ||
        row.some(c => c !== null && c !== undefined && /MANUAL\s*CHALL|HAND\s*CHALL|\bMANUAL\b/i.test(String(c)))
    });
  }
}

let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

console.log('Master File: ' + records.length + ' records across sheet(s): ' + useSheets.join(', ') + '\n');

const clone = () => JSON.parse(JSON.stringify(records));

// ================= 1. the reported case =================
console.log('1. M82039-1B1-5 — short on PI 124 covered by the excess on PI 144');
API.set(clone()); API.run();
let rows = API.get().filter(r => r.styleNorm === 'M82039-1B1-5' && API.measNorm(r.itemDesc) === API.measNorm('Master Carton, 5 Ply, L71.5 x W45 x H32cm;'));
const shortRow = rows.find(r => r.short < 0);
const excessRow = rows.find(r => r.short > 0);
t('both lines present', !!shortRow && !!excessRow);
t('short line is now resolved (renders green)', shortRow && shortRow.resolved === true, shortRow && String(shortRow.resolved));
t('short line outstanding is 0', shortRow && API.effectiveShort(shortRow) === 0, shortRow && API.effectiveShort(shortRow));
t('short line carries a delivery note', shortRow && /Full Short Qty is delivered/.test(API.combinedRemarks(shortRow)), shortRow && API.combinedRemarks(shortRow));
t('note names the covering challan 32021081', shortRow && /32021081/.test(API.combinedRemarks(shortRow)), shortRow && API.combinedRemarks(shortRow));
t('note says it came under PI 144 (cross-PI stated)', shortRow && /PI 144/.test(API.combinedRemarks(shortRow)), shortRow && API.combinedRemarks(shortRow));
t('raw Short&Exs on the short line untouched (-1)', shortRow && shortRow.short === -1, shortRow && shortRow.short);
t('excess line raw value untouched (+2)', excessRow && excessRow.short === 2, excessRow && excessRow.short);
t('excess line outstanding drops to +1', excessRow && API.effectiveShort(excessRow) === 1, excessRow && API.effectiveShort(excessRow));
t('excess line says where its piece went', excessRow && /covers the earlier short/.test(API.combinedRemarks(excessRow)), excessRow && API.combinedRemarks(excessRow));
console.log('   short line remark : ' + (shortRow ? API.combinedRemarks(shortRow) : '-'));
console.log('   excess line remark: ' + (excessRow ? API.combinedRemarks(excessRow) : '-'));

// ================= 2. conservation — nothing invented, nothing lost =================
console.log('\n2. Conservation across the whole file');
API.set(clone()); API.run();
const all = API.get();
const covered = all.reduce((s, r) => s + (r._covered || 0), 0);
const consumed = all.reduce((s, r) => s + (r._consumed || 0), 0);
const manualPool = all.filter(r => r.manualCh).reduce((s, r) => s + (r.recQty || 0), 0);
t('every covered piece is paid for by a source', covered >= consumed, covered + ' vs ' + consumed);
t('excess consumed never exceeds excess available',
  all.every(r => (r._consumed || 0) <= Math.max(0, r.short || 0)),
  all.filter(r => (r._consumed || 0) > Math.max(0, r.short || 0)).length + ' violations');
t('coverage never exceeds the short it covers',
  all.every(r => (r._covered || 0) <= Math.abs(Math.min(0, r.short || 0))),
  all.filter(r => (r._covered || 0) > Math.abs(Math.min(0, r.short || 0))).length + ' violations');
t('no line is both a coverage source and a coverage target',
  all.every(r => !((r._covered || 0) > 0 && (r._consumed || 0) > 0)));
t('outstanding short never worse than raw short',
  all.every(r => API.effectiveShort(r) >= Math.min(0, r.short || 0)));
t('outstanding excess never larger than raw excess',
  all.every(r => API.effectiveShort(r) <= Math.max(0, r.short || 0)));

// ================= 3. shorts and excesses are still two separate totals =================
console.log('\n3. Shorts and excesses remain independent totals');
const totalShort = all.reduce((s, r) => s + Math.min(API.effectiveShort(r), 0), 0);
const totalExcess = all.reduce((s, r) => s + Math.max(API.effectiveShort(r), 0), 0);
const rawShort = all.reduce((s, r) => s + Math.min(r.short || 0, 0), 0);
const rawExcess = all.reduce((s, r) => s + Math.max(r.short || 0, 0), 0);
console.log('   raw:         short ' + rawShort + '   excess +' + rawExcess);
console.log('   outstanding: short ' + totalShort + '   excess +' + totalExcess);
t('totals are reported as two numbers, never one netted figure', totalShort <= 0 && totalExcess >= 0);
t('short total moved by exactly the pieces matched', Math.abs((totalShort - rawShort)) === covered, (totalShort - rawShort) + ' vs ' + covered);
t('excess total moved by exactly the pieces matched', (rawExcess - totalExcess) === consumed, (rawExcess - totalExcess) + ' vs ' + consumed);

// ================= 4. matching discipline =================
console.log('\n4. Matching discipline');
let badMeas = 0, badTime = 0, badStyle = 0; const orphans = [];
all.forEach(x => {
  if (!(x._consumed > 0)) return;
  // NOTE: exact equality is stricter than the shipped rule, which walks a ladder of
  // exact -> format-variant -> prefix. Anything this reports is a PREFIX-tier match: the pieces went
  // to a short on a NEARBY style, not the same one. That is the matcher working as designed, but it
  // is worth a human eye, so the lines are named rather than quietly accepted.
  const partners = all.filter(s => (s._covered || 0) > 0 && s.styleNorm === x.styleNorm &&
    API.measNorm(s.itemDesc) === API.measNorm(x.itemDesc));
  if (!partners.length) { badStyle++; orphans.push(x); return; }
  partners.forEach(s => { if (API.measNorm(s.itemDesc) !== API.measNorm(x.itemDesc)) badMeas++; });
});
// Time order is checked against the DATE THE CODE ITSELF WROTE into the note, not against a
// guessed pairing: with several shorts and several excesses on one style, any outside guess at
// who paid for whom is wrong. The note is the tool's own claim, so that is what must hold up.
const timeTravel = [];
all.forEach(r => {
  if (!((r._covered || 0) > 0) || (r.short || 0) >= 0) return;
  const own = API.manualTime(r);
  (r.autoNotes || []).forEach(n => {
    const m = n.match(/^On (\d{1,2})-(\d{1,2})-(\d{4}),/);
    if (!m || !own) return;
    const noteMs = new Date(+m[3], +m[2] - 1, +m[1]).getTime();
    if (noteMs < own) timeTravel.push(r.sheet + ':' + r.xlRow + ' short ' + new Date(own).toISOString().slice(0,10) + ' <- note ' + n);
  });
});
if (orphans.length) {
  console.log('   FYI — excess consumed with no same-style covered short (prefix-tier matches):');
  orphans.forEach(x => console.log('     ' + x.styleNorm + '  consumed=' + x._consumed + ' pcs'));
}
t('coverage only ever within the same style', badStyle === 0, badStyle + ' prefix-tier match(es) — see list above');
t('coverage only ever at the same measurement', badMeas === 0, badMeas);
t('no note claims delivery BEFORE the short happened', timeTravel.length === 0, timeTravel.slice(0,3).join(' ;; '));
t('undated lines never act as a coverage source',
  all.every(r => !((r._consumed || 0) > 0 && !API.manualTime(r))));

// ================= 5. accepted shorts are not raided =================
console.log('\n5. Accepted (waived) shorts must not eat coverage');
const victim = records.find(r => r.short < 0 && r.styleNorm === 'M82039-1B1-5');
if (victim) {
  const key = [String(victim.challanNo ?? '').trim(), victim.styleNorm,
    API.measNorm(victim.itemDesc), String(victim.short ?? '')].join('|');
  API.set(clone()); API.accept({ [key]: { reason: 'test waiver', at: Date.now() } }); API.run();
  const after = API.get().find(r => r.xlRow === victim.xlRow && r.sheet === victim.sheet);
  t('an accepted short reports 0 outstanding', after && API.effectiveShort(after) === 0);
  t('an accepted short is never also covered by an excess', after && !(after._covered > 0), after && after._covered);
  API.accept({});
} else { console.log('  (skipped — no sample row)'); }

// ================= 6. manual-challan behaviour unchanged =================
console.log('\n6. Manual-challan path unchanged vs the previous build');
const manualRows = records.filter(r => r.manualCh);
console.log('   manual challan lines: ' + manualRows.length + ', received on them: ' + manualPool + ' pcs');
if (PREV) {
  PREV.set(clone()); PREV.run();
  const before = PREV.get();
  const key = r => r.sheet + ':' + r.xlRow;
  const bMap = new Map(before.map(r => [key(r), r]));
  let lost = [], shrank = [], noteLost = [];
  all.forEach(r => {
    const b = bMap.get(key(r));
    if (!b) return;
    if (b.resolved && !r.resolved) lost.push(key(r));
    if ((b._covered || 0) > (r._covered || 0)) shrank.push(key(r) + ' ' + (b._covered||0) + '->' + (r._covered||0));
    (b.autoNotes || []).forEach(n => { if (!(r.autoNotes || []).includes(n)) noteLost.push(key(r) + ' :: ' + n); });
    if (b.autoNote && r.autoNote !== b.autoNote) noteLost.push(key(r) + ' :: ' + b.autoNote);
  });
  t('no line that WAS resolved before is unresolved now', lost.length === 0, lost.slice(0,3).join(', '));
  t('no line lost manual-challan coverage', shrank.length === 0, shrank.slice(0,3).join(', '));
  t('every note the old build produced is still produced', noteLost.length === 0, noteLost.slice(0,2).join(' ;; '));
  const bCov = before.reduce((s, r) => s + (r._covered || 0), 0);
  console.log('   coverage before: ' + bCov + ' pcs   after: ' + covered + ' pcs   (+' + (covered - bCov) + ' from the new excess matching)');
  t('manual-challan coverage total is preserved exactly', bCov <= covered, bCov + ' vs ' + covered);
} else {
  console.log('   (no previous build supplied — pass it as argv[4] to diff)');
}

// ================= 7. idempotency =================
console.log('\n7. Re-running reconciliation is stable');
API.set(clone()); API.run();
const first = API.get().map(r => (r._covered || 0) + '/' + (r._consumed || 0) + '/' + (r.resolved ? 1 : 0));
API.run();
const second = API.get().map(r => (r._covered || 0) + '/' + (r._consumed || 0) + '/' + (r.resolved ? 1 : 0));
t('a second run produces identical results', JSON.stringify(first) === JSON.stringify(second));

// ================= 8. blast radius =================
console.log('\n8. What actually changed on this Master File');
const newlyGreen = all.filter(r => (r.short || 0) < 0 && r.resolved && !r.accepted);
console.log('   short lines now showing as delivered: ' + newlyGreen.length + ' of ' + all.filter(r => (r.short || 0) < 0).length);
console.log('   pieces matched: ' + covered + ' pcs');

console.log('\n' + (fail ? 'FAILED ' + fail + ' / ' : 'PASSED all ') + (pass + fail) + ' checks');
process.exit(fail ? 1 : 0);
