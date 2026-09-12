/* kindtest.js — the guard on "it says every line is a Master Carton".
   itemKind() lived inside the Undelivered Report's IIFE. The Ask tab could not reach it, so it
   printed the Master File's Item column instead — and that column says "Master Carton" on punch
   lines too. One tab split them correctly, the other showed them all as the same thing. Same data,
   two answers, in the two places that must never disagree.
   Usage: node kindtest.js olisa.html                                                            */
const fs = require('fs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];
function lift(n) { const s = src.indexOf('function ' + n); if (s < 0) throw new Error('missing ' + n); let d = 0, e = s; for (let i = src.indexOf('{', s); i < src.length; i++) { if (src[i] === '{') d++; else if (src[i] === '}') { d--; if (!d) { e = i + 1; break; } } } return src.slice(s, e); }
const bag = new Function(lift('itemKind') + src.match(/const KIND_LABEL = \{[^}]*\};/)[0] + lift('itemLabel')
  + '\nreturn { itemKind, KIND_LABEL, itemLabel };')();

console.log('\n1. Real descriptions are classified correctly');
// Lifted from the actual challans and the Master File.
[['Master Carton, 5 Ply, L70 x W48 x H34cm;', 'master'],
 ['Cross Divider, 5 Ply, L69 x W31.5 x H0cm;', 'cross'],
 ['Master Carton (Punch), 3 Ply, L33 x W30.5 x H30.5cm;', 'punch'],
 ['Master Carton Punch 5 Ply', 'punch'],
 ['PERFORATED CARTON L42 x W38', 'punch'],
 ['Cross Divider (5 Ply)', 'cross'],
 ['MASTER CARTON, 5 PLY', 'master'],
 ['master carton (punch)', 'punch']
].forEach(([d, want]) => t(`${d.slice(0, 42)} -> ${want}`, bag.itemKind(d) === want, bag.itemKind(d)));

console.log('\n2. Punch beats the word "Master Carton" sitting next to it');
// This is the whole bug: the raw Item column says "Master Carton" on a punch line. Reading the
// description must win, or every punch line is reported as an ordinary carton.
t('a punch line whose Item column says Master Carton is still a punch',
  bag.itemLabel({ item: 'Master Carton', itemDesc: 'Master Carton (Punch), 3 Ply' }) === 'Chip Box (Punch)',
  bag.itemLabel({ item: 'Master Carton', itemDesc: 'Master Carton (Punch), 3 Ply' }));
t('a cross divider whose Item column says Master Carton is still a cross divider',
  bag.itemLabel({ item: 'Master Carton', itemDesc: 'Cross Divider, 5 Ply' }) === 'Cross Divider');
t('an ordinary carton is left alone',
  bag.itemLabel({ item: 'Master Carton', itemDesc: 'Master Carton, 5 Ply, L70 x W48' }) === 'Master Carton');
t('the word punch in the Item column alone is enough',
  bag.itemLabel({ item: 'Master Carton Punch', itemDesc: '' }) === 'Chip Box (Punch)');
t('a row with nothing at all does not crash', typeof bag.itemLabel({}) === 'string');
t('and neither does a null row', typeof bag.itemLabel(null) === 'string');

console.log('\n3. One definition, reachable by both tabs');
t('itemKind is at true top level, not inside an IIFE',
  /^function itemKind\(desc\) \{/m.test(src), 'must not be indented inside a scope');
t('KIND_LABEL sits beside it', /^const KIND_LABEL = \{/m.test(src));
t('there is only ONE itemKind in the file', [...src.matchAll(/function itemKind\(/g)].length === 1);
t('there is only ONE KIND_LABEL in the file', [...src.matchAll(/const KIND_LABEL = /g)].length === 1);
t('the Ask tab no longer prints the raw Item column',
  !/<td>\$\{escHtml\(String\(r\.item \?\? ''\)\)\}<\/td>/.test(src), 'raw r.item cell still present');
t('every Item cell goes through itemLabel', [...src.matchAll(/<td>\$\{escHtml\(itemLabel\(r\)\)\}<\/td>/g)].length >= 5,
  [...src.matchAll(/itemLabel\(r\)/g)].length + ' uses');
t('the Excel export agrees with the screen', /r\.style, itemLabel\(r\), extractMeasurement/.test(src));

console.log('\n4. The filter');
t('Master Carton only is offered', /<option value="master">Master Carton only<\/option>/.test(html));
t('Chip Box (Punch) only is offered', /<option value="punch">Chip Box \(Punch\) only<\/option>/.test(html));
t('Cross Divider only is offered', /<option value="cross">Cross Divider only<\/option>/.test(html));
t('Manual Challan only is still there', /<option value="manual">/.test(html));
t('the chosen kind is read out of the dropdown', /const kind = \(mode === 'master' \|\| mode === 'punch' \|\| mode === 'cross'\)/.test(src));
t('and filters on the description, the same way the Undelivered Report does',
  /if \(f\.kind && itemKind\(String\(r\.itemDesc \|\| ''\) \+ ' ' \+ String\(r\.item \|\| ''\)\) !== f\.kind\) return false;/.test(src));
t('a kind filter counts as "something is active"', /if \(!f\.manualOnly && !f\.kind && !hasFrom && !hasTo\) return matches;/.test(src));
t('and is named in the active-filter text', /bits\.push\(KIND_LABEL\[f\.kind\] \+ ' only'\)/.test(src));

console.log('\n5. An active filter is impossible to miss');
t('the bar turns red when a filter is on', /\.filterbar\.filter-on \.filterlabel \{ color: var\(--bad\); \}/.test(html));
t('the dropdown turns red too', /\.filterbar\.filter-on select/.test(html));
t('so does the hint', /\.filterbar\.filter-on \.filterhint/.test(html));
t('the class is toggled from the live filter state', /bar\.classList\.toggle\('filter-on', !!desc\)/.test(src));
t('and painted on load, not only on change', /\[mode, pf, pt\]\.forEach[\s\S]{0,140}?\n\s*upd\(\);/.test(src));

console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
process.exit(fail ? 1 : 0);
