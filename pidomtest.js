/* Boots Pending_PI_Summary.html in a real DOM, drives it the way a person would, and
   checks that both download buttons produce a correctly named file. */
const fs = require('fs');
const { JSDOM } = require('jsdom');

let pass = 0, fail = 0;
const check = (name, cond, extra) => {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (extra ? '  -> ' + extra : '')); }
};

const html = fs.readFileSync('Pending_PI_Summary.html', 'utf8');
const errors = [];

const dom = new JSDOM(html, {
  runScripts: 'dangerously',
  pretendToBeVisual: true,
  url: 'https://example.com/Pending_PI_Summary.html',
  resources: undefined,                 // never fetch the CDN copy of pdf.js
  beforeParse(w) {
    w.fetch = () => Promise.resolve({ ok: true });
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {} }));
    w.onerror = (msg) => errors.push(String(msg));
    const cl = w.console;
    w.console = Object.assign({}, cl, { error: (...a) => { errors.push(a.join(' ')); cl.error(...a); } });
    /* capture downloads instead of performing them */
    w.__saved = [];
    w.URL.createObjectURL = (blob) => { w.__lastBlob = blob; return 'blob:fake'; };
    w.URL.revokeObjectURL = () => {};
  }
});

const w = dom.window, d = w.document;
/* HTMLAnchorElement.click() is a navigation in jsdom; record it instead */
w.HTMLAnchorElement.prototype.click = function () {
  if (this.download) w.__saved.push({ name: this.download, blob: w.__lastBlob });
};

(async () => {
  await new Promise(r => w.addEventListener('load', r));
  await new Promise(r => setTimeout(r, 60));

  console.log('\n== The page boots ==');
  check('no script errors on load', errors.length === 0, errors.join(' | '));
  check('the app switcher carries all five tabs', d.querySelectorAll('.appswitch a').length === 5);
  check('this page is the one marked current',
    d.querySelector('.appswitch a.on').getAttribute('href') === 'Pending_PI_Summary.html');
  check('the empty state hides the ledger', d.getElementById('ledgerCard').hidden === true);
  check('the empty state hides the download bar', d.getElementById('foot').hidden === true);
  check('the drop zone is showing', !!d.getElementById('drop'));
  check('exactly two download buttons exist',
    !!d.getElementById('btnXlsx') && !!d.getElementById('btnPdf') &&
    !d.getElementById('btnCsv') && !d.getElementById('btnCopy'));
  check('the as-of date is filled in', /^\d{4}-\d{2}-\d{2}$/.test(d.getElementById('asof').value));
  check('the table head is set to stick', /position: sticky; top: 0;/.test(d.querySelector('style').textContent));

  console.log('\n== Rows appear ==');
  /* feed the page rows the way a read PDF would */
  const D = (y, m, day) => new w.Date(w.Date.UTC(y, m - 1, day));
  w.eval(`
    state.rows = [
      {id:1,party:'A One Polar',buyer:'Gildan',pi:'PI-2201/07/2026',value:12500.5,
       date:new Date(Date.UTC(2026,4,14)),concern:'Rifat',file:'a.pdf',issues:[],confidence:'high',touched:{},text:'x'},
      {id:2,party:'A One Polar',buyer:'Hanes',pi:'PI-2202/07/2026',value:8400,
       date:new Date(Date.UTC(2026,5,2)),concern:'Rifat',file:'b.pdf',issues:[],confidence:'high',touched:{},text:'x'}
    ];
    state.asOf='2026-09-19';
    render();
  `);
  check('still no script errors', errors.length === 0, errors.join(' | '));
  check('the ledger card is revealed', d.getElementById('ledgerCard').hidden === false);
  check('the download bar is revealed', d.getElementById('foot').hidden === false);
  check('the summary card is revealed', d.getElementById('summary').hidden === false);
  check('two rows are drawn', d.querySelectorAll('#tbody tr').length === 2);
  check('the party cell spans both rows',
    d.querySelector('#tbody td.party').getAttribute('rowspan') === '2');
  check('rows are ordered by Pending Days, longest wait first',
    [...d.querySelectorAll('#tbody td.days .daycell')].map(e => e.textContent.trim()).join(',') === '128,109',
    [...d.querySelectorAll('#tbody td.days .daycell')].map(e => e.textContent.trim()).join(','));
  check('buyers are shown in capitals',
    d.querySelectorAll('#tbody tr')[0].querySelector('td.buyer').textContent === 'GILDAN',
    d.querySelectorAll('#tbody tr')[0].querySelector('td.buyer').textContent);
  check('the as-of date is echoed day-month-year',
    d.getElementById('asofShown').textContent === '19-09-2026', d.getElementById('asofShown').textContent);
  check('the ageing legend uses fifteen-day bands',
    /Over 90 days/.test(d.getElementById('legend').textContent), d.getElementById('legend').textContent);
  check('the legend gives a share as well as a value', /%/.test(d.getElementById('legend').textContent));
  check('one party needs no split panel', d.getElementById('partyAge').hidden === true);
  check('the total reads $20,900.50',
    d.getElementById('figValue').textContent === '$20,900.50', d.getElementById('figValue').textContent);
  check('pending days were counted', d.querySelector('#tbody td.days .daycell').textContent.trim() === '128',
    d.querySelector('#tbody td.days .daycell').textContent);
  check('the tray shrinks out of the way', d.getElementById('tray').classList.contains('slim'));
  check('the heading counts what was read',
    d.getElementById('trayCount').textContent === '2 invoices read', d.getElementById('trayCount').textContent);
  check('the bar names the file it will save',
    /A One Polar - PI Pending Summary/.test(d.getElementById('barStatus').textContent),
    d.getElementById('barStatus').textContent);

  console.log('\n== The two downloads ==');
  d.getElementById('btnXlsx').click();
  await new Promise(r => setTimeout(r, 120));
  d.getElementById('btnPdf').click();
  await new Promise(r => setTimeout(r, 120));

  const saved = w.__saved;
  check('exactly two files were produced', saved.length === 2, JSON.stringify(saved.map(s => s.name)));
  check('the Excel file is named after the party',
    saved[0] && saved[0].name === 'A One Polar - PI Pending Summary.xlsx', saved[0] && saved[0].name);
  check('the PDF is named after the party',
    saved[1] && saved[1].name === 'A One Polar - PI Pending Summary.pdf', saved[1] && saved[1].name);
  check('the Excel blob carries the spreadsheet type',
    saved[0] && /spreadsheetml/.test(saved[0].blob.type), saved[0] && saved[0].blob.type);
  check('the PDF blob carries the pdf type',
    saved[1] && saved[1].blob.type === 'application/pdf', saved[1] && saved[1].blob.type);
  check('both files have real bytes in them',
    saved.every(s => s.blob.size > 1000), saved.map(s => s.blob.size).join(','));
  check('no errors while downloading', errors.length === 0, errors.join(' | '));

  console.log('\n== Two parties split the panel ==');
  w.eval(`
    state.rows.push({id:3,party:'Beximco Textiles Ltd.',buyer:'zara',pi:'PI-99/2026',value:3000,
      date:new Date(Date.UTC(2026,8,1)),concern:'Shuvo',file:'c.pdf',issues:[],confidence:'high',touched:{},text:'x'});
    render();
  `);
  check('the split panel appears', d.getElementById('partyAge').hidden === false);
  check('there is one panel per party', d.querySelectorAll('#partyAge .pa').length === 2);
  check('each panel is named', [...d.querySelectorAll('#partyAge .pah b')].map(e => e.textContent).sort().join('|') ===
    'A One Polar|Beximco Textiles Ltd.', [...d.querySelectorAll('#partyAge .pah b')].map(e => e.textContent).join('|'));
  check('each panel draws its own bar', d.querySelectorAll('#partyAge .pa .agebar i').length >= 2);
  check('the panels sit two across',
    /\.partyage \{ display: grid; grid-template-columns: 1fr 1fr;/.test(d.querySelector('style').textContent));
  w.eval("state.rows = state.rows.filter(function(r){return r.id!==3;}); render();");
  check('removing the party folds the panel away', d.getElementById('partyAge').hidden === true);

  console.log('\n== Editing still works ==');
  const buyerCell = d.querySelectorAll('#tbody tr')[0].querySelector('td.buyer');
  buyerCell.dispatchEvent(new w.MouseEvent('click', { bubbles: true }));
  const input = buyerCell.querySelector('input.cellin');
  check('clicking a cell opens an editor', !!input);
  if (input) {
    input.value = 'gildan activewear';
    input.dispatchEvent(new w.FocusEvent('blur'));
    await new Promise(r => setTimeout(r, 40));
    check('a typed buyer is capitalised too',
      d.querySelectorAll('#tbody tr')[0].querySelector('td.buyer').textContent === 'GILDAN ACTIVEWEAR',
      d.querySelectorAll('#tbody tr')[0].querySelector('td.buyer').textContent);
  }

  console.log('\n== Renaming the party renames the files ==');
  w.eval("state.rows.forEach(function(r){r.party='Beximco Textiles Ltd.';}); render();");
  w.__saved.length = 0;
  d.getElementById('btnXlsx').click();
  await new Promise(r => setTimeout(r, 120));
  check('the new party name reaches the file name',
    w.__saved[0] && w.__saved[0].name === 'Beximco Textiles Ltd. - PI Pending Summary.xlsx',
    w.__saved[0] && w.__saved[0].name);

  console.log('\n== Dates are never second-guessed ==');
  check('no "could be read either way" warning survives', !/could be read either way/.test(html));
  check('the page still lets you state the format', !!d.getElementById('dfmt'));
  check('day-month-year is what it starts on', d.getElementById('dfmt').value === 'DMY');

  console.log('\n== Sorting and filtering ==');
  d.querySelectorAll('#headRow th')[3].dispatchEvent(new w.MouseEvent('click', { bubbles: true }));
  check('a heading click re-sorts without error', errors.length === 0, errors.join(' | '));
  const search = d.getElementById('search');
  search.value = 'nothing-matches-this';
  search.dispatchEvent(new w.Event('input'));
  check('an empty filter shows the empty note',
    /No rows match/.test(d.getElementById('tbody').textContent));
  check('the heading says what is hidden',
    /showing 0 of 2/.test(d.getElementById('filterNote').textContent),
    d.getElementById('filterNote').textContent);
  search.value = '';
  search.dispatchEvent(new w.Event('input'));
  check('clearing the filter brings the rows back', d.querySelectorAll('#tbody tr').length === 2);

  console.log('\n' + pass + ' passed, ' + fail + ' failed\n');
  process.exit(fail ? 1 : 0);
})();
