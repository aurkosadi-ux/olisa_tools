/* Awkward data, driven through the real page: many parties, a party with nothing
   but undated rows, huge and tiny values, long names, and a run long enough to
   spill over several pages. Nothing here should throw or come out crooked. */
const fs = require('fs');
const { JSDOM } = require('jsdom');
let pass = 0, fail = 0;
const check = (n, c, e) => { c ? (pass++, console.log('  ok   ' + n)) : (fail++, console.log('  FAIL ' + n + (e ? '  -> ' + e : ''))); };

const html = fs.readFileSync('Pending_PI_Summary.html', 'utf8');
const errors = [];
const dom = new JSDOM(html, {
  runScripts: 'dangerously', pretendToBeVisual: true, url: 'https://example.com/Pending_PI_Summary.html',
  beforeParse(w) {
    w.fetch = () => Promise.resolve({ ok: true });
    w.onerror = m => errors.push(String(m));
    w.__saved = [];
    w.URL.createObjectURL = b => { w.__lastBlob = b; return 'blob:x'; };
    w.URL.revokeObjectURL = () => {};
  }
});
const w = dom.window, d = w.document;
w.HTMLAnchorElement.prototype.click = function () { if (this.download) w.__saved.push({ name: this.download, blob: w.__lastBlob }); };

(async () => {
  await new Promise(r => w.addEventListener('load', r));
  await new Promise(r => setTimeout(r, 50));

  console.log('\n== Five parties, awkward data ==');
  w.eval(`
    var parties = ['A One Polar','Beximco Textiles Ltd.',
      'A Very Long Party Name That Will Not Fit In A Narrow Column Limited',
      'Square Fashions Ltd.','Undated Only Traders'];
    var id = 0; state.rows = [];
    parties.forEach(function (p, pi) {
      var n = pi === 4 ? 2 : 9;
      for (var i = 0; i < n; i++) {
        var undated = (pi === 4) || (i === 8);
        state.rows.push({
          id: ++id, party: p,
          buyer: 'buyer name number ' + i + ' of ' + p.slice(0, 6),
          pi: 'PI-' + (1000 + id) + '/09/2026',
          value: undated ? null : (i === 0 ? 1234567.89 : (i === 1 ? 0.01 : 1000 * (i + 1) + 0.5)),
          date: undated ? null : new Date(Date.UTC(2026, 8 - i, 1 + i)),
          concern: 'Concern Person ' + i, file: 'f' + id + '.pdf',
          issues: [], confidence: 'high', touched: {}, text: 'x'
        });
      }
    });
    state.asOf = '2026-09-20';
    render();
  `);
  check('no script errors with five parties', errors.length === 0, errors.join(' | '));
  check('every row is drawn', d.querySelectorAll('#tbody tr').length === 38, String(d.querySelectorAll('#tbody tr').length));
  check('five split panels appear', d.querySelectorAll('#partyAge .pa').length === 5);
  check('a party with no dates still gets a panel',
    [...d.querySelectorAll('#partyAge .pah b')].some(e => e.textContent === 'Undated Only Traders'));

  /* within each party, pending days must run high to low with undated rows last */
  const perParty = w.eval(`JSON.stringify(grouped(sortRows(state.rows)).map(function (g) {
    return g.rows.map(function (r) { return pendingDays(r); });
  }))`);
  const orders = JSON.parse(perParty);
  const descending = orders.every(list => {
    const dated = list.filter(v => v != null), undated = list.filter(v => v == null);
    const tailOk = list.slice(list.length - undated.length).every(v => v == null);
    return tailOk && dated.every((v, i) => i === 0 || dated[i - 1] >= v);
  });
  check('each party runs longest wait first, undated last', descending, perParty);

  console.log('\n== Both files still come out ==');
  d.getElementById('btnXlsx').click(); await new Promise(r => setTimeout(r, 200));
  d.getElementById('btnPdf').click(); await new Promise(r => setTimeout(r, 300));
  check('two files were produced', w.__saved.length === 2, JSON.stringify(w.__saved.map(s => s.name)));
  check('several parties get the plain name',
    w.__saved[0] && w.__saved[0].name === 'PI Pending Summary.xlsx', w.__saved[0] && w.__saved[0].name);
  check('neither file is empty', w.__saved.every(s => s.blob.size > 4000), w.__saved.map(s => s.blob.size).join(','));
  check('no errors while building', errors.length === 0, errors.join(' | '));

  console.log('\n== A single undated row ==');
  w.__saved.length = 0;
  w.eval(`state.rows = [{id:1,party:'Solo Ltd.',buyer:'x',pi:'PI-1',value:null,date:null,
    concern:'',file:'a.pdf',issues:[],confidence:'low',touched:{},text:''}]; render();`);
  check('the page survives one dateless row', errors.length === 0, errors.join(' | '));
  check('the oldest figure shows a dash', d.getElementById('figOldest').textContent === '\u2014',
    d.getElementById('figOldest').textContent);
  check('the total is zero', d.getElementById('figValue').textContent === '$0.00', d.getElementById('figValue').textContent);
  d.getElementById('btnPdf').click(); await new Promise(r => setTimeout(r, 200));
  check('a PDF is still produced', w.__saved.length === 1 && w.__saved[0].blob.size > 1000);
  check('still no errors', errors.length === 0, errors.join(' | '));

  console.log('\n' + pass + ' passed, ' + fail + ' failed\n');
  process.exit(fail ? 1 : 0);
})();
