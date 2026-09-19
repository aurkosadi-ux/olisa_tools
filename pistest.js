/* Pending PI Summary — end-to-end checks on the two files the page now produces.
   The page's own code is lifted out of the HTML and run here, so what is tested is
   exactly what ships. */
const fs = require('fs');
const zlib = require('zlib');

const html = fs.readFileSync('Pending_PI_Summary.html', 'utf8');
const scripts = [];
const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;
let m;
while ((m = re.exec(html))) scripts.push(m[1]);
const code = scripts[scripts.length - 1];

/* everything before "4. STATE + UI" is pure logic with no DOM in it */
const cut = code.indexOf('/* ============================================================\n   3. PDF READING');
const logic = code.slice(0, cut);
/* the export helpers are further down and also DOM-free; take the two builders we need */
const pdfStart = code.indexOf('/* ============================================================\n   4b. PDF WRITER');
const pdfEnd = code.indexOf('/* ---- exports ---- */');
const pdfCode = code.slice(pdfStart, pdfEnd);
/* the file-name helpers, without the download plumbing */
const nameStart = code.indexOf('function safeFileName(');
const nameEnd = code.indexOf('var savedCap,capAsked=false;');
const nameCode = code.slice(nameStart, nameEnd);

const fn = new Function('TextEncoder', 'exportGroupsRef',
  logic + '\n' + pdfCode + '\nfunction exportGroups(){return exportGroupsRef();}\n' + nameCode +
  '\nreturn {buildSummaryWorkbook,buildSummaryPdf,exportBaseName,sheetTitle,safeFileName,' +
  'textWidth,pdfAscii,ageing,band,BANDS,fmtDate,fmtISO,fmtMoney,upperBuyer};');

let GROUPS = [];
const api = fn(TextEncoder, () => GROUPS);

/* ---------- a tiny zip reader, enough to pull the stored XML back out ---------- */
function unzip(buf) {
  const out = {};
  let i = 0;
  while (i < buf.length - 4) {
    if (buf.readUInt32LE(i) !== 0x04034b50) { i++; continue; }
    const size = buf.readUInt32LE(i + 18);
    const nlen = buf.readUInt16LE(i + 26);
    const elen = buf.readUInt16LE(i + 28);
    const name = buf.slice(i + 30, i + 30 + nlen).toString('utf8');
    const data = buf.slice(i + 30 + nlen + elen, i + 30 + nlen + elen + size);
    const method = buf.readUInt16LE(i + 8);
    out[name] = method === 8 ? zlib.inflateRawSync(data) : data;
    i = i + 30 + nlen + elen + size;
  }
  return out;
}

const D = (y, mo, d) => new Date(Date.UTC(y, mo - 1, d));
const ONE = [{
  party: 'A One Polar',
  rows: [
    { buyer: 'Gildan', pi: 'PI-2201/07/2026', value: 12500.5, date: D(2026, 5, 14), concern: 'Rifat', pendingDays: 128 },
    { buyer: 'Hanes', pi: 'PI-2202/07/2026', value: 8400, date: D(2026, 6, 2), concern: 'Rifat', pendingDays: 109 },
    { buyer: 'Primark', pi: 'PI-2203/08/2026', value: null, date: null, concern: '', pendingDays: null }
  ]
}];
const TWO = ONE.concat([{
  party: 'Beximco Textiles Ltd.',
  rows: [{ buyer: 'Zara', pi: 'PI-99/2026', value: 3000.25, date: D(2026, 7, 1), concern: 'Shuvo', pendingDays: 80 }]
}]);

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
}

console.log('\n== Excel ==');
GROUPS = ONE;
const xlsx = Buffer.from(api.buildSummaryWorkbook(ONE, { asOf: '2026-09-20' }));
const files = unzip(xlsx);
const sheet = files['xl/worksheets/sheet1.xml'].toString('utf8');
const styles = files['xl/styles.xml'].toString('utf8');
const wb = files['xl/workbook.xml'].toString('utf8');

check('the workbook holds all six parts', Object.keys(files).length === 6, Object.keys(files).join(', '));
check('the party name is in cell A1', /<c r="A1"[^>]*t="inlineStr"[^>]*><is><t[^>]*>A One Polar</.test(sheet));
check('A1 uses the title style (s=10)', /<c r="A1" t="inlineStr" s="10"/.test(sheet));
check('the title is merged right across A1:F1', sheet.includes('<mergeCell ref="A1:F1"/>'));
check('the as-of line sits under the party name, in day-month-year',
  /<c r="A2"[^>]*><is><t[^>]*>Pending as of 20-09-2026/.test(sheet));
check('the column headings sit on row 3', /<c r="A3"[^>]*><is><t[^>]*>Buyer<\/t>/.test(sheet) && /<c r="F3"[^>]*><is><t[^>]*>Concern Person</.test(sheet));
check('the workbook writes dates day-month-year', styles.includes('formatCode="dd\\-mm\\-yyyy"'));
check('buyers are set in capitals', /<is><t[^>]*>GILDAN<\/t>/.test(sheet) && !/>Gildan</.test(sheet));
check('nothing says "Generated ..." any more', !/Generated /.test(sheet));
check('no live-formula footnote survives', !/recalculates every time/.test(sheet));
check('Pending Days is still =TODAY()-<PI date>', sheet.includes('<f>TODAY()-D4</f>'));
check('the per-party total sums the value column', sheet.includes('<f>SUM(C4:C6)</f>'));
check('TOTAL - is merged across A:B', sheet.includes('<mergeCell ref="A7:B7"/>'));
check('the sheet is exactly six columns wide', (sheet.match(/<col /g) || []).length === 6);
check('no seventh column is written', !/<c r="G\d/.test(sheet));
check('the tab is named after the party', wb.includes('name="A One Polar"'));
check('a fill is declared for every ageing band', /<fills count="13">/.test(styles),
  (styles.match(/<fills count="\d+"/) || [])[0]);
check('each band colour reaches the workbook',
  api.BANDS.every(b => styles.includes(b.xl)), api.BANDS.map(b => b.xl).join(','));
/* white text on amber is not readable; each chip must carry ink it can hold */
const lin = c => { c /= 255; return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
const lum = h => 0.2126 * lin(parseInt(h.slice(0, 2), 16)) + 0.7152 * lin(parseInt(h.slice(2, 4), 16)) + 0.0722 * lin(parseInt(h.slice(4, 6), 16));
const contrast = (a, b) => { const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05); };
const weak = api.BANDS.map(b => ({ label: b.label, hex: b.xl.slice(2), dark: b.dark }))
  .concat([{ label: 'No PI date', hex: 'FF8B98A9'.slice(2), dark: 1 }])
  .filter(b => contrast(b.hex, b.dark ? '16261C' : 'FFFFFF') < 4.5);
check('every band chip carries readable text', weak.length === 0,
  weak.map(b => b.label + ' ' + contrast(b.hex, b.dark ? '16261C' : 'FFFFFF').toFixed(2)).join(', '));
check('the ageing block is written under the party', /AGEING OF PENDING VALUE/.test(sheet));
check('its four headings are there',
  ['Ageing','Invoices','Value','Share of Value'].every(h => sheet.includes('>' + h + '<')));
check('a share is written as a percentage', /numFmtId="166" formatCode="0.0%"/.test(styles));
check('the title style is bold 14 navy', /<font><b\/><sz val="14"\/><color rgb="FF102A6E"\/>/.test(styles));

/* well-formedness: every tag closed, in order */
function xmlOk(s) {
  const stack = [];
  const t = /<(\/?)([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  let mm;
  while ((mm = t.exec(s))) {
    if (mm[3].endsWith('/') || mm[4] === '/') continue;
    if (mm[1] === '/') { if (stack.pop() !== mm[2]) return 'mismatch at ' + mm[2]; }
    else stack.push(mm[2]);
  }
  return stack.length ? 'unclosed ' + stack.join(',') : true;
}
for (const f of Object.keys(files)) {
  const body = files[f].toString('utf8').replace(/<\?xml[^>]*\?>/, '');
  check('well-formed XML: ' + f, xmlOk(body) === true, xmlOk(body));
}

GROUPS = TWO;
const xlsx2 = Buffer.from(api.buildSummaryWorkbook(TWO, { asOf: '2026-09-20' }));
const sheet2 = unzip(xlsx2)['xl/worksheets/sheet1.xml'].toString('utf8');
const wb2 = unzip(xlsx2)['xl/workbook.xml'].toString('utf8');
check('a second party gets its own title band', (sheet2.match(/s="10"/g) || []).length >= 12);
check('two parties earn a GRAND TOTAL', sheet2.includes('GRAND TOTAL - '));
check('mixed parties fall back to a plain tab name', wb2.includes('name="Summary"'));

console.log('\n== File names ==');
GROUPS = ONE;
check('one party names the file after it', api.exportBaseName() === 'A One Polar - PI Pending Summary', api.exportBaseName());
GROUPS = TWO;
check('several parties get the plain name', api.exportBaseName() === 'PI Pending Summary', api.exportBaseName());
GROUPS = [{ party: 'A/B: Textiles *Ltd*', rows: ONE[0].rows }];
check('characters Windows refuses are stripped',
  api.exportBaseName() === 'A B Textiles Ltd - PI Pending Summary', api.exportBaseName());
GROUPS = [{ party: '(no party name)', rows: ONE[0].rows }];
check('an unnamed party still gets a sane name', api.exportBaseName() === 'PI Pending Summary', api.exportBaseName());

console.log('\n== PDF ==');
GROUPS = ONE;
const pdf = Buffer.from(api.buildSummaryPdf(ONE, { asOf: '2026-09-19' }));
const txt = pdf.toString('latin1');
check('starts with a PDF header', txt.startsWith('%PDF-1.4'));
check('ends with %%EOF', txt.trim().endsWith('%%EOF'));
check('declares a catalogue, a page tree and two fonts',
  txt.includes('/Type /Catalog') && txt.includes('/Type /Pages') && txt.includes('/BaseFont /Helvetica-Bold'));
check('the party name is printed on it', txt.includes('(A One Polar) Tj'));
check('the six headings are printed', ['BUYER', 'PI', 'VALUE', 'PI DATE', 'PENDING DAYS', 'CONCERN PERSON']
  .every(h => txt.includes('(' + h + ') Tj')));
check('a row value is printed', txt.includes('($12,500.50) Tj'));
check('the total is printed', txt.includes('($20,900.50) Tj'));

/* the cross-reference table must point at real object headers, or readers reject the file */
const startxref = +txt.slice(txt.lastIndexOf('startxref') + 9).trim().split('\n')[0];
const xrefBlock = txt.slice(startxref);
check('startxref lands on the xref table', xrefBlock.startsWith('xref'));
const offsets = [...xrefBlock.matchAll(/^(\d{10}) 00000 n $/gm)].map(x => +x[1]);
check('every object offset is right',
  offsets.every((off, idx) => txt.slice(off).startsWith((idx + 1) + ' 0 obj')),
  offsets.map((off, idx) => (idx + 1) + '->' + JSON.stringify(txt.substr(off, 9))).join(' '));
const lengths = [...txt.matchAll(/<< \/Length (\d+) >>\nstream\n/g)];
check('every stream /Length matches the bytes that follow', lengths.every(l => {
  const start = l.index + l[0].length;
  return txt.slice(start + +l[1], start + +l[1] + 10) === '\nendstream';
}));
check('no stray non-ASCII byte crept in', !/[^\x00-\x7f]/.test(txt));

GROUPS = TWO;
const pdf2 = Buffer.from(api.buildSummaryPdf(TWO, { asOf: '2026-09-19' })).toString('latin1');
check('a second party is printed too', pdf2.includes('(Beximco Textiles Ltd.) Tj'));
check('two parties earn a GRAND TOTAL', pdf2.includes('(GRAND TOTAL -) Tj'));
check('the grand total adds up', pdf2.includes('($23,900.75) Tj'));

/* a long run has to break onto more pages, repeat the party band and number the pages */
const many = [{ party: 'A One Polar', rows: [] }];
for (let i = 0; i < 90; i++) many[0].rows.push({
  buyer: 'Buyer ' + i, pi: 'PI-' + i + '/2026', value: 100 + i, date: D(2026, 3, 1), concern: 'Rifat', pendingDays: 200
});
GROUPS = many;
const big = Buffer.from(api.buildSummaryPdf(many, { asOf: '2026-09-19' })).toString('latin1');
const pageCount = (big.match(/\/Type \/Page[^s]/g) || []).length;
check('90 rows spill onto more than one page', pageCount > 1, 'pages=' + pageCount);
check('the page count in /Count agrees', big.includes('/Count ' + pageCount));
check('continued pages repeat the party name', big.includes('(A One Polar  \\(continued\\)) Tj'));
check('the pages are numbered', big.includes('(Page 1 of ' + pageCount + ') Tj'));

console.log('\n== Ageing ==');
const ag = api.ageing([
  { value: 100, pendingDays: 3 },   { value: 200, pendingDays: 15 },
  { value: 400, pendingDays: 16 },  { value: 800, pendingDays: 45 },
  { value: 1600, pendingDays: 91 }, { value: 50, pendingDays: null }
]);
check('the bands are fifteen days wide', api.BANDS.map(b => b.label).join('|') ===
  '0-15 days|16-30 days|31-45 days|46-60 days|61-75 days|76-90 days|Over 90 days',
  api.BANDS.map(b => b.label).join('|'));
check('15 days falls in the first band', api.band(15) === 0);
check('16 days falls in the second', api.band(16) === 1);
check('45 days falls in the third', api.band(45) === 2);
check('91 days lands in Over 90', api.band(91) === 6);
check('the first band collects both of its rows', ag.all[0].count === 2 && ag.all[0].value === 300);
check('an undated row is counted apart', ag.undated.count === 1 && ag.undated.value === 50);
check('empty bands are left out of the printed lines',
  ag.lines.map(l => l.label).join('|') === '0-15 days|16-30 days|31-45 days|Over 90 days|No PI date',
  ag.lines.map(l => l.label).join('|'));
check('the total is every row, dated or not', ag.total === 3150 && ag.count === 6);
check('the oldest ignores undated rows', ag.oldest === 91);

console.log('\n== Ageing reaches both files ==');
GROUPS = ONE;
const agPdf = Buffer.from(api.buildSummaryPdf(ONE, { asOf: '2026-09-20' })).toString('latin1');
check('the PDF prints an ageing block', agPdf.includes('(AGEING OF PENDING VALUE) Tj'));
check('it prints the band labels that apply',
  api.BANDS.some(b => agPdf.includes('(' + b.label + ') Tj')) && agPdf.includes('(No PI date) Tj'));
check('it prints the headline figures',
  agPdf.includes('(TOTAL PENDING) Tj') && agPdf.includes('(OLDEST PENDING) Tj') && agPdf.includes('(OVER 60 DAYS) Tj'));
check('it prints shares as percentages', /\(\d+\.\d%\) Tj/.test(agPdf));
check('the as-of date is day-month-year', agPdf.includes('(Pending as of 20-09-2026) Tj'));
check('the PDF says what currency it is in', agPdf.includes('(All values in USD) Tj'));
check('it carries document properties', agPdf.includes('/Title (') && agPdf.includes('/Info 5 0 R'));
check('buyers are printed in capitals', agPdf.includes('(GILDAN) Tj') || /\(GILDAN[^)]*\) Tj/.test(agPdf));

GROUPS = TWO;
const agPdf2 = Buffer.from(api.buildSummaryPdf(TWO, { asOf: '2026-09-20' })).toString('latin1');
check('two parties get an overall ageing block', agPdf2.includes('(AGEING OF PENDING VALUE - ALL PARTIES) Tj'));
check('and one panel each, side by side',
  agPdf2.includes('(A ONE POLAR) Tj') && agPdf2.includes('(BEXIMCO TEXTILES LTD.) Tj'));
check('the detail is introduced as its own section', agPdf2.includes('(INVOICE DETAIL BY PARTY) Tj'));
const sheetAge = unzip(Buffer.from(api.buildSummaryWorkbook(TWO, { asOf: '2026-09-20' })))['xl/worksheets/sheet1.xml'].toString('utf8');
check('the workbook gives each party its own ageing',
  (sheetAge.match(/AGEING OF PENDING VALUE/g) || []).length === 3,
  String((sheetAge.match(/AGEING OF PENDING VALUE/g) || []).length));
check('and an all-parties block at the end', sheetAge.includes('ALL PARTIES'));

console.log('\n== Dates and capitals ==');
check('a date is written day-month-year', api.fmtDate(new Date(Date.UTC(2026, 4, 3))) === '03-05-2026', api.fmtDate(new Date(Date.UTC(2026, 4, 3))));
check('an ISO as-of is turned round', api.fmtISO('2026-09-20') === '20-09-2026', api.fmtISO('2026-09-20'));
check('a buyer is capitalised', api.upperBuyer('bestseller a/s') === 'BESTSELLER A/S');

console.log('\n== Text measuring ==');
check('a space is narrower than a W', api.textWidth(' ', 10, false) < api.textWidth('W', 10, false));
check('bold is wider than regular', api.textWidth('Hello', 10, true) > api.textWidth('Hello', 10, false));
check('non-Latin text is folded, never dropped', api.pdfAscii('Café — Ltd').length > 0 && !/[^\x20-\x7e]/.test(api.pdfAscii('Café — Ltd')));
check('accents fold to their plain letter', api.pdfAscii('Café') === 'Cafe', api.pdfAscii('Café'));

console.log('\n' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
