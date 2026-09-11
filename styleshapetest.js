/* styleshapetest.js — the guard on "understand it is a style even though nothing says so".
   Every Olisa style is M + five digits + an optional tail. The value of this is entirely in what
   it REFUSES: vehicle numbers, phone numbers, PI and DO numbers, size ranges and carton dimensions
   all sit in the same documents and all look like codes at token level.
   Usage: node styleshapetest.js olisa.html                                                     */
const fs = require('fs');
let pass = 0, fail = 0;
const t = (n, c, x) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (x ? '  -> ' + x : ''))); };

const html = fs.readFileSync(process.argv[2] || 'olisa.html', 'utf8');
const src = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).sort((a, b) => b.length - a.length)[0];
function lift(n) { const s = src.indexOf('function ' + n); let d = 0, e = s; for (let i = src.indexOf('{', s); i < src.length; i++) { if (src[i] === '{') d++; else if (src[i] === '}') { d--; if (!d) { e = i + 1; break; } } } return src.slice(s, e); }
const bag = new Function(src.match(/const STYLE_SHAPE_RE = .*/)[0] + lift('styleShaped') + lift('findStyleShapes')
  + '\nreturn { styleShaped, findStyleShapes };')();

console.log('\n1. Real styles are recognised');
[['M81988-3A1-1', 'five digits then a hyphenated tail'], ['M82065C', 'bare letter tail'],
 ['M82065-1A', 'short hyphen tail'], ['M82065-3A10-1', 'two hyphen groups'],
 ['M82050-1A1-3', 'digits inside the tail'], ['M82035A1-3', 'no hyphen before the tail'],
 ['M82062B5', 'letter then digit'], ['M81988-5A', 'digit then letter'], ['M82073A1', 'plain tail'],
 ['M81988-5C1-2', 'long tail']].forEach(([s, why]) => t(`${s} (${why})`, bag.styleShaped(s) === true));

console.log('\n2. Everything that sits beside a style is refused');
// Each of these is real text lifted from the challans, and each would have been read as a code by
// a "has a digit and a letter" rule.
[['METRO-U-12-3590', 'vehicle number'], ['M-103', 'hand challan number'],
 ['5-8#', 'carton size range'], ['8.5-12#', 'the other size range'],
 ['27.56', 'a dimension'], ['01841213454', 'a phone number'], ['700768', 'the customer code'],
 ['036260826380815-1293003210', 'a DO number'], ['121/700768/0159082026', 'a PI reference'],
 ['M8206', 'only four digits'], ['MASTER', 'a word'], ['M123456789', 'runs past five digits into more digits']
].forEach(([s, why]) => t(`${s} (${why}) is refused`, bag.styleShaped(s) === false));

console.log('\n3. Pulled out of a pasted Olisa email, with the noise left behind');
const email = `Dear Sadi, Please hold the production of below styles. We will revise later.
M82065-3A10-1, M82065-3A10-2 and M81988-3A1-1. Also M82062A4 / M82062B4.
Driver Md. Shamim V# DHAKA METRO-U-12-3590 delivered on 31-AUG-26.
PI:121/700768/0159082026 WO:121/700768/0159/15-08-2026 Do No : 036260826380815-1293003210
PHONE: 01841213454, Mr. Suman - 01704132997. Size 5-8# and 8.5-12#. Type:12.99 X 12.01 X 12.01`;
const got = bag.findStyleShapes(email);
t('finds all five styles', got.length === 5, got.join(','));
t('in the order they appear', got.join(',') === 'M82065-3A10-1,M82065-3A10-2,M81988-3A1-1,M82062A4,M82062B4', got.join(','));
t('and nothing else at all', got.every(g => /^M\d{5}/.test(g)));

console.log('\n4. Survives OCR noise');
// Styles recovered from a genuinely bad OCR pass. The point is not that OCR is good — it is that
// the shape rule still picks the styles out of the wreckage without inventing any.
const ocr = 'L63 x W49 x H28.8cm 61 144 M82035A1-3 | I reteladhticeh ee a 144 __] '
  + 'L62 x W33.5 x HOcm 14 144 M82050-1A1-3 L67 x W34.5 x H28cm M82050A1 '
  + 'L71.5 x W39 x H27em M82063A4 L35 x W30.5 x H32.Sem an06se2 mis twas';
const o = bag.findStyleShapes(ocr);
t('picks the four legible styles out', o.length === 4, o.join(','));
t('does not invent one from the garbled token', !o.some(x => /AN06/.test(x)));

console.log('\n5. Wired in where it matters');
t('a blank Size: on a challan line falls back to the shape', /styleFromShape/.test(src));
t('the shape pass feeds the Ask box', /const shapes = findStyleShapes\(text\);/.test(src));
t('a long paste drops candidates that are neither known nor style-shaped', /const isPaste = String\(text\)\.length > 60/.test(src));
t('a short typed query is left exactly as it was', /if \(isPaste && shapes\.length\)/.test(src));
t('classifyStyle can say "looks like a style but I have never seen it"', /how: styleShaped\(u\) \? 'shape' : 'unknown'/.test(src));

console.log(fail ? `\nFAILED — ${pass} passed, ${fail} failed` : `\nPASSED — ${pass} passed, 0 failed`);
process.exit(fail ? 1 : 0);
