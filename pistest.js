/* Pending PI Summary — end-to-end checks on the two files the page now produces.
   The page's own code is lifted out of the HTML and run here, so what is tested is
   exactly what ships. */
const fs = require('fs');
const zlib = require('zlib');

const html = fs.readFileSync('PPS_new.html', 'utf8');
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

const sandbox = {
  TextEncoder,
  fmtDate(d) {
    if (!d) return '';
    const p = n => String(n).padStart(2, '0');
    return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  },
  fmtMoney(n) {
    if (n == null) return '';
    return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  },
  exportGroups: null
};
const fn = new Function('TextEncoder', 'fmtDate', 'fmtMoney', 'exportGroupsRef',
  logic + '\n' + pdfCode + '\nfunction exportGroups(){return exportGroupsRef();}\n' + nameCode +
  '\nreturn {buildSummaryWorkbook,buildSummaryPdf,exportBaseName,sheetTitle,safeFileName,textWidth,pdfAscii};');

let GROUPS = [];
const api = fn(TextEncoder, sandbox.fmtDate, sandbox.fmtMoney, () => GROUPS);

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
const xlsx = Buffer.from(api.buildSummaryWorkbook(ONE));
const files = unzip(xlsx);
const sheet = files['xl/worksheets/sheet1.xml'].toString('utf8');
const styles = files['xl/styles.xml'].toString('utf8');
const wb = files['xl/workbook.xml'].toString('utf8');

check('the workbook holds all six parts', Object.keys(files).length === 6, Object.keys(files).join(', '));
check('the party name is in cell A1', /<c r="A1"[^>]*t="inlineStr"[^>]*><is><t[^>]*>A One Polar</.test(sheet));
check('A1 uses the title style (s=10)', /<c r="A1" t="inlineStr" s="10"/.test(sheet));
check('the title is merged right across A1:F1', sheet.includes('<mergeCell ref="A1:F1"/>'));
check('the column headings sit on row 2', /<c r="A2"[^>]*><is><t[^>]*>Buyer<\/t>/.test(sheet) && /<c r="F2"[^>]*><is><t[^>]*>Concern Person</.test(sheet));
check('nothing says "Generated ..." any more', !/Generated /.test(sheet));
check('no live-formula footnote survives', !/recalculates every time/.test(sheet));
check('Pending Days is still =TODAY()-<PI date>', sheet.includes('<f>TODAY()-D3</f>'));
check('the per-party total sums the value column', sheet.includes('<f>SUM(C3:C5)</f>'));
check('TOTAL - is merged across A:B', sheet.includes('<mergeCell ref="A6:B6"/>'));
check('the sheet is exactly six columns wide', (sheet.match(/<col /g) || []).length === 6);
check('no seventh column is written', !/<c r="G\d/.test(sheet));
check('the tab is named after the party', wb.includes('name="A One Polar"'));
check('the styles declare four fills', /<fills count="4">/.test(styles));
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
const xlsx2 = Buffer.from(api.buildSummaryWorkbook(TWO));
const sheet2 = unzip(xlsx2)['xl/worksheets/sheet1.xml'].toString('utf8');
const wb2 = unzip(xlsx2)['xl/workbook.xml'].toString('utf8');
check('a second party gets its own title band', (sheet2.match(/s="10"/g) || []).length === 12);
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
check('the six headings are printed', ['Buyer', 'PI', 'Value', 'PI Date', 'Pending Days', 'Concern Person']
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

console.log('\n== Text measuring ==');
check('a space is narrower than a W', api.textWidth(' ', 10, false) < api.textWidth('W', 10, false));
check('bold is wider than regular', api.textWidth('Hello', 10, true) > api.textWidth('Hello', 10, false));
check('non-Latin text is folded, never dropped', api.pdfAscii('Café — Ltd').length > 0 && !/[^\x20-\x7e]/.test(api.pdfAscii('Café — Ltd')));
check('accents fold to their plain letter', api.pdfAscii('Café') === 'Cafe', api.pdfAscii('Café'));

console.log('\n' + pass + ' passed, ' + fail + ' failed\n');
process.exit(fail ? 1 : 0);
