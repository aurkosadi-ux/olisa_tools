/* xsstest.js — no value from an uploaded file may reach the DOM as live markup.
   Proves it by rendering a crafted spreadsheet through the REAL shipped code. */
const fs=require('fs'), XLSX=require('xlsx');
let pass=0,fail=0;
const t=(n,c,x)=>{c?(pass++,console.log('  ok   '+n)):(fail++,console.log('  FAIL '+n+(x?'  -> '+x:'')))};
const html=fs.readFileSync(process.argv[2]||'olisa.html','utf8');
const src=[...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]).sort((a,b)=>b.length-a.length)[0];
const esc=new Function('return '+src.match(/function escHtml[\s\S]*?\n\}/)[0].replace('function escHtml','function'))();
const snap=new Function('return '+src.match(/function snapToLocalMidnight[\s\S]*?\n\}/)[0].replace('function snapToLocalMidnight','function'))();

console.log('1. escHtml itself');
[['<','&lt;'],['>','&gt;'],['&','&amp;'],['"','&quot;']].forEach(([raw,want])=>
  t(`escapes ${raw}`, esc(raw).includes(want), esc(raw)));

console.log('\n2. A crafted spreadsheet rendered through the real PO Summary preview');
const PAYLOADS=[
  'CARTON 30*20*7CM <img src=x onerror=alert(1)>',
  'CARTON 30*20*7CM <script>alert(2)</'+'script>',
  'CARTON 30*20*7CM "><svg onload=alert(3)>',
];
const aoa=[['order date','FACTORY','ORDER SIZE','UNIT','ORDER QTY','STYLE','delivery date']];
PAYLOADS.forEach((p,i)=>aoa.push(['27/08/2026','P',p,'PCS',100+i,'M81988-3A'+i,'08/09/2026']));
const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),'S');
const buf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});

const st=src.indexOf('(function() { // PO Summary Generator'), en=src.indexOf('(function() { // Undelivered Report Generator');
let b=src.slice(st,en).replace(/^\(function\(\) \{ \/\/ PO Summary Generator/,'').replace(/\}\)\(\);\s*$/,'')
 .replace(/const \w+ = document\.getElementById\([^)]*\);\n/g,'').replace(/^\s*\w+\.addEventListener\([\s\S]*?\n\}\);\n/gm,'')
 .replace(/^\s*\['drag[^\n]*\n/gm,'').replace(/^\s*\w+\.addEventListener\([^\n]*\n/gm,'');
let captured='';
const M=new Function('XLSX','snapToLocalMidnight','escHtml','previewTable','previewNote',
  b+'\nreturn {readOrderRows,classifyRow,buildLine,renderPreview};')
  (XLSX,snap,esc,{set innerHTML(v){captured=v},get innerHTML(){return captured}},{style:{},innerHTML:'',addEventListener(){}});

(async()=>{
  const recs=await M.readOrderRows({name:'x.xlsx',arrayBuffer:async()=>buf.buffer.slice(buf.byteOffset,buf.byteOffset+buf.byteLength)});
  const rows=recs.map(r=>{const c=M.classifyRow(r.sizeText);const l=c.classification?M.buildLine(c.classification,r.sizeText):null;
    return {classification:c.classification,sizeText:r.sizeText,style:r.style,color:l&&l.color,productCode:l&&l.productCode,
      ply:l&&l.ply,wastageW:l&&l.wastageW,wastageH:l&&l.wastageH,l:c.l,w:c.w,h:c.h,qty:r.qty,needsReview:c.needsReview};});
  M.renderPreview(rows);
  t('no live <img onerror> reaches the DOM', !/<img[^>]*onerror/i.test(captured));
  t('no live <script> reaches the DOM', !/<script/i.test(captured));
  t('no live <svg onload> reaches the DOM', !/<svg[^>]*onload/i.test(captured));
  t('no stray attribute-breaking quote survives', !/"><svg/i.test(captured));
  t('the payloads ARE shown, escaped, not silently dropped', /&lt;img src=x onerror=alert\(1\)&gt;/.test(captured));
  t('the NEEDS REVIEW flag is still real markup, built by the app not the data',
    /<span class="flag">NEEDS REVIEW<\/span>/.test(captured) || !rows.some(r=>r.needsReview));

  console.log('\n3. The source rule holds across the file');
  t('the PO preview escapes every cell', /escHtml\(String\(v\)\)/.test(src));
  t('shortage-table PI dates are escaped', /escHtml\(String\(piDateFor/.test(src));
  t('PI-derived review rows are escaped', /<td>\$\{escHtml\(String\(ln\.style/.test(src));
  t('quantities from the file are escaped too', /escHtml\(String\(r\.short/.test(src));
  t('style suggestions escape the code from the Master File', /escHtml\(disp\.slice\(0, pos\)\)/.test(src));
  t('the upload-checker escapes its cells', /return `<td>\$\{escHtml\(String\(disp\)\)\}<\/td>`/.test(src));
  t('the upload-checker escapes its HEADERS too', /headers\.map\(h => `<th>\$\{escHtml\(String\(h \?\? ''\)\)\}<\/th>`\)/.test(src));
  // A ternary like `ch ? escHtml(ch) : '<i>none</i>'` is safe but does not START with escHtml, and
  // escHtml(String(f(x) ?? '')) nests three deep, so regex cannot do this. Count parens instead:
  // blank out every complete escHtml(...) call and every string literal, then whatever identifier
  // survives is a value reaching the DOM raw.
  const stripCalls = e => {
    let out = e, guard = 0;
    while (guard++ < 20) {
      const at = out.indexOf('escHtml(');
      if (at < 0) break;
      let d = 0, k = at + 'escHtml'.length;
      for (; k < out.length; k++) {
        if (out[k] === '(') d++;
        else if (out[k] === ')') { d--; if (d === 0) { k++; break; } }
      }
      out = out.slice(0, at) + '@' + out.slice(k);
    }
    return out;
  };
  const bare = e => {
    const x = stripCalls(e)
      .replace(/'(?:[^'\\]|\\.)*'|"(?:[^"\\]|\\.)*"|`(?:[^`\\]|\\.)*`/g, '@')
      .replace(/\b(null|undefined|true|false)\b/g, '@')
      // An identifier immediately before `?` is a ternary CONDITION, not a value being written out.
      // `ch ? escHtml(ch) : '<i>none</i>'` writes nothing raw. `??` is left alone: that IS a value.
      .replace(/[A-Za-z_$][\w$.]*\s*\?(?!\?)/g, '@?');
    return /[A-Za-z_$][\w$]*/.test(x);
  };
  const raw=[...src.matchAll(/<td>\$\{([^}]*)\}/g)].map(m=>m[1].trim()).filter(bare)
    // Allowed: values the APP computes, and markup the app itself builds. Anything that could
    // carry file text must go through escHtml.
    .filter(e=>!/^(piCell|challanCell|remarkCell|piQ !== null|outstanding|act|g\.styles\.size|g\.lines|g\.del|total(Master|Punch|Cross)|v === null \|\| v === undefined \? '' : escHtml)/.test(e));
  t('no <td> interpolates an unescaped file value', raw.length===0, raw.join(' | '));

  console.log('\n'+(fail?'FAILED':'PASSED')+` — ${pass} passed, ${fail} failed`);
  process.exit(fail?1:0);
})();
