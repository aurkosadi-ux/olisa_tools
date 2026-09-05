/* libtest.js — guards the self-hosted libraries.
   Usage: node libtest.js olisa.html sw.js [libDir]  */
const fs=require('fs'), path=require('path');
let pass=0,fail=0,warn=0;
const t=(n,c,x)=>{c?(pass++,console.log('  ok   '+n)):(fail++,console.log('  FAIL '+n+(x?'  -> '+x:'')))};
const w=(n,x)=>{warn++;console.log('  WARN '+n+(x?'  -> '+x:''))};
const html=fs.readFileSync(process.argv[2]||'olisa.html','utf8');
const sw=fs.readFileSync(process.argv[3]||'sw.js','utf8');
const dir=process.argv[4]||'lib';

const EXPECT={'xlsx.full.min.js':['SheetJS'],'exceljs.min.js':['ExcelJS'],
  'pdf.min.js':['pdfjs','PDFJS','pdfjsLib'],'pdf-lib.min.js':['PDFDocument','PDFLib'],
  'mammoth.browser.min.js':['mammoth']};

console.log('1. Nothing executable is fetched from a third party any more');
const srcs=[...html.matchAll(/<script[^>]*\bsrc="([^"]+)"/g)].map(m=>m[1]);
const remote=srcs.filter(s=>/^https?:/.test(s));
t('every <script src> is local', remote.length===0, remote.join(', '));
t('all five libraries are pointed at ./lib/', ['xlsx.full','exceljs','pdf.min','pdf-lib','mammoth']
  .every(k=>srcs.some(s=>s.startsWith('./lib/')&&s.includes(k))), srcs.join(' '));
// the ONE allowed remaining reference is the worker fallback, and only as a fallback
const cdnHits=[...html.matchAll(/https:\/\/cdnjs\.cloudflare\.com[^'"\s]*/g)].map(m=>m[0]);
t('the only cdnjs reference left is the worker fallback',
  cdnHits.length===1 && /pdf\.worker\.min\.js$/.test(cdnHits[0]), cdnHits.join(', '));
t('the worker is set to the LOCAL copy first', /workerSrc = LOCAL;/.test(html));
t('a missing worker falls back instead of breaking PDF reading', /workerSrc = CDN;/.test(html));
t('the fallback probe is cheap (HEAD, not a full download)', /method: 'HEAD'/.test(html));

console.log('\n2. The files are really there and really the libraries');
EXPECT['jszip.min.js']=['JSZip']; EXPECT['jspdf.umd.min.js']=['jsPDF','jspdf'];
for(const [f,marks] of Object.entries(EXPECT)){
  const p=path.join(dir,f);
  if(!fs.existsSync(p)){ t(`${f} present`,false,'not found in '+dir); continue; }
  const body=fs.readFileSync(p,'utf8');
  t(`${f} present (${(body.length/1024).toFixed(0)} KB)`, body.length>50000, body.length+' bytes');
  t(`${f} is the real library, not a 404 page`,
    marks.some(m=>body.includes(m)) && !/^\s*<!DOCTYPE html/i.test(body), body.slice(0,60));
}
const worker=path.join(dir,'pdf.worker.min.js');
if(fs.existsSync(worker)) t('pdf.worker.min.js present — self-hosting is complete', fs.readFileSync(worker,'utf8').length>50000);
else w('pdf.worker.min.js NOT present — PDF reading still uses the CDN worker','upload it to finish the job');

console.log('\n3. The calculator no longer compiles CSS in the browser');
const calc=fs.existsSync('calculator.html')?fs.readFileSync('calculator.html','utf8'):'';
if(calc){
  const calcLoads=[...calc.matchAll(/<(?:script|link)[^>]*(?:src|href)="(https?:\/\/[^"]+)"/g)].map(m=>m[1]);
  t('calculator.html loads nothing from the network', calcLoads.length===0, calcLoads.join(', '));
  t('the Tailwind compiler script is gone', !/<script[^>]*src="https:\/\/cdn\.tailwindcss\.com/.test(calc));
  t('compiled Tailwind CSS is inlined instead', /--tw-border-spacing-x/.test(calc));
  t('the compiled CSS is small, not a shipped compiler', calc.length<200000, (calc.length/1024).toFixed(0)+' KB total page');
  // every class the page uses must survive the switch
  const toks=new Set();
  for(const m of calc.matchAll(/class(?:Name)?\s*=\s*["'`]([^"'`]+)["'`]/g)) m[1].split(/\s+/).forEach(x=>x&&toks.add(x));
  for(const m of calc.matchAll(/classList\.(?:add|remove|toggle)\(\s*["']([^"']+)["']/g)) m[1].split(/\s+/).forEach(x=>x&&toks.add(x));
  const style=(calc.match(/<style>[\s\S]*?<\/style>/g)||[]).join('\n');
  const esc=c=>[...c].map(ch=>':/[].#()%,'.includes(ch)?'\\'+ch:ch).join('');
  const miss=[...toks].filter(x=>!x.startsWith('${')&&!style.includes('.'+esc(x))&&!style.includes(x));
  t(`all ${toks.size} classes still have styling`, miss.length===0, miss.join(', '));
}

console.log('\n3b. DC Bypass Bill Maker is no longer the odd one out');
const dc = fs.existsSync('DC_Bypass_Bill.html') ? fs.readFileSync('DC_Bypass_Bill.html','utf8') : '';
if (dc) {
  t('it uses the SAME self-hosted pdf.js as the main app', /src="\.\/lib\/pdf\.min\.js"/.test(dc));
  t('its worker points at the local copy first', /workerSrc = "\.\/lib\/pdf\.worker\.min\.js"/.test(dc));
  t('with a CDN fallback so it cannot break outright', /pdf\.worker\.min\.js";\s*\n\s*\}\);/.test(dc) || /cdnjs[^"]*pdf\.worker/.test(dc));
  t('each PDF page is released after its text is read', /_page\.cleanup\(\)/.test(dc));
  t('the document is destroyed when the file is done', /pdf\.destroy\(\)/.test(dc));
  t('and destroyed on failure too, so a bad PDF cannot leak', /\}, function\(err\)\{[\s\S]{0,120}pdf\.destroy/.test(dc));
  const left = [...dc.matchAll(/<script[^>]*src="(https?:\/\/[^"]+)"/g)].map(m => m[1]);
  t('DC Bypass loads NOTHING from the network', left.length === 0, left.join(', '));
  t('jszip is local', /src="\.\/lib\/jszip\.min\.js"/.test(dc));
  t('jspdf is local', /src="\.\/lib\/jspdf\.umd\.min\.js"/.test(dc));
  t('both are cached by the service worker', /jszip\.min\.js/.test(sw) && /jspdf\.umd\.min\.js/.test(sw));
}

console.log('\n4. Fonts cannot block the first paint');
for(const f of ['olisa.html','index.html']){
  if(!fs.existsSync(f)) continue;
  const p=fs.readFileSync(f,'utf8');
  const blocking=[...p.matchAll(/<link(?![^>]*media=)[^>]*fonts\.googleapis\.com\/css2[^>]*rel="stylesheet"[^>]*>/g)]
    .concat([...p.matchAll(/<link[^>]*rel="stylesheet"(?![^>]*media=)[^>]*fonts\.googleapis\.com\/css2[^>]*>/g)])
    .filter(m=>!/<noscript>/.test(p.slice(Math.max(0,m.index-12),m.index)));
  t(`${f}: the font stylesheet does not block render`, blocking.length===0, blocking.map(m=>m[0].slice(0,70)).join(' | '));
  t(`${f}: it still loads (media flips on load)`, /media="print" onload="this\.media='all'/.test(p));
  t(`${f}: a no-JavaScript fallback is kept`, /<noscript><link rel="stylesheet"/.test(p));
}

console.log('\n5. The service worker ships them with the app');
t('cache version was bumped', /olisa-tools-v(\d+)/.test(sw) && Number(sw.match(/olisa-tools-v(\d+)/)[1])>=48);
t('libraries are in the install shell', /const LIB = \[/.test(sw));
['xlsx.full','exceljs','pdf.min','pdf-lib','mammoth','pdf.worker'].forEach(k=>
  t(`  ${k} cached at install`, new RegExp('\\./lib/[^\']*'+k.replace('.','\\.')).test(sw)));
t('libraries are served cache-first, not network-first',
  /url\.origin === location\.origin && \/\\\/lib\\\/\[\^\/\]\+\\\.js\$\/\.test\(url\.pathname\)/.test(sw)
  || /\/lib\\\//.test(sw) && /hit\) return hit;/.test(sw));
t('install stays non-atomic so one missing file cannot kill offline support', /c\.add\(u\)\.catch/.test(sw));
t('Google API traffic is still never cached', /googleapis\.com[\s\S]{0,80}return;/.test(sw));

console.log('\n'+(fail?'FAILED':'PASSED')+` — ${pass} passed, ${fail} failed, ${warn} warning(s)`);
process.exit(fail?1:0);
