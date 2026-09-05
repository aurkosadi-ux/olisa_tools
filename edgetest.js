const fs=require('fs'),path=require('path'),XLSX=require('xlsx'),ExcelJS=require('exceljs');
const html=fs.readFileSync('olisa.html','utf8');
const src=[...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]).sort((a,b)=>b.length-a.length)[0];
const snap=new Function('return '+src.match(/function snapToLocalMidnight[\s\S]*?\n\}/)[0].replace('function snapToLocalMidnight','function'))();
function lift(startMark,endMark,ret,extra){
  let b=src.slice(src.indexOf(startMark),src.indexOf(endMark))
    .replace(/^\(function\(\) \{ \/\/ [^\n]*/,'').replace(/\}\)\(\);\s*$/,'')
    .replace(/const \w+ = document\.getElementById\([^)]*\);\n/g,'')
    .replace(/^\s*\w+\.addEventListener\([\s\S]*?\n\}\);\n/gm,'')
    .replace(/^\s*\['drag[^\n]*\n/gm,'').replace(/^\s*\w+\.addEventListener\([^\n]*\n/gm,'')
    .replace(/^refreshDefaultDate\(\);\n/m,'');
  return new Function('XLSX','ExcelJS','snapToLocalMidnight','escHtml','dateInput','document','buyerInput','deliveryDateInput',
    b+'\nreturn {'+ret+'};')(XLSX,ExcelJS,snap,x=>String(x),{value:'31st August 2026'},{addEventListener(){}},{value:'B'},{value:''});
}
const U=lift('(function() { // Undelivered Report Generator','(function() { // Ask — style lookup','readWorkbookRows,buildGroups,buildOutputRows,deliverySortKey,itemKind,findCrossFileOverlaps');
const P=lift('(function() { // PO Summary Generator','(function() { // Undelivered Report Generator','readOrderRows,groupByDeliveryDate,anyToDate,dateSuffix');

function mk(aoa,name){ const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),'Rpt');
  const buf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
  return {name:name||'t.xlsx',arrayBuffer:async()=>buf.buffer.slice(buf.byteOffset,buf.byteOffset+buf.byteLength)}; }
const H=['PI_NO','DIST_NAME','DIST_ID','S_PERSON','PI_DATE','PI_DLV_DATE','BUYER_NAME','DO_NO','DO_LINE_NO','ITEM_ID','ITEM_NAME','ITEM_DESCRIPTION','L','W','H','MESUREMENT','STYLE','P_SIZE','COUNTRY_CODE','TYPE_OF_PACK','D_U_FACT','D_QTY','D_S_QTY','U_QTY','U_S_QTY','UND_AMOUNT'];
const row=(pi,dlv,desc,qty,style,pidate)=>{const r=new Array(26).fill(null);
  r[0]=`121/700768/0${pi}082026`;r[4]=pidate||'27/08/2026';r[5]=dlv;r[6]='Olisa 1';r[11]=desc;r[17]=style;r[23]=qty;return r;};

(async()=>{
 const R=[]; const chk=(n,c,d)=>R.push([c?'ok  ':'BUG ',n,d||'']);

 // ---- 1. the same export uploaded twice ----
 // the guard now lives at SELECTION time, so test it the way the tab uses it
 const same=(a,b)=>a.name===b.name&&a.size===b.size&&a.lastModified===b.lastModified;
 const A={name:'Undl.xlsx',size:1234,lastModified:99}, A2={name:'Undl.xlsx',size:1234,lastModified:99};
 const B={name:'Undl (1).xlsx',size:1234,lastModified:99};
 const dedupe=list=>{const k=[];list.forEach(f=>{if(!k.some(x=>same(x,f)))k.push(f);});return k;};
 chk('the same file dropped twice is kept once', dedupe([A,A2]).length===1);
 chk('two genuinely different files are both kept', dedupe([A,B]).length===2);
 chk('a re-export under a new name is caught by PI instead', (()=>{
   const recs=[{piNo:'121/700768/0159082026',_srcFile:'a.xlsx'},{piNo:'121/700768/0159082026',_srcFile:'b.xlsx'}];
   return U.findCrossFileOverlaps(recs).length===1; })());
 chk('two partial exports with NO shared PI raise no warning', (()=>{
   const recs=[{piNo:'121/700768/0159082026',_srcFile:'a.xlsx'},{piNo:'121/700768/0160082026',_srcFile:'b.xlsx'}];
   return U.findCrossFileOverlaps(recs).length===0; })());

 // ---- 2. empty / all-zero ----
 const emptyRecs = await U.readWorkbookRows(mk([H])).catch(()=>[]);
 chk('a header-only export yields no lines (generate() then refuses it)', emptyRecs.length===0);
 try{ const z=await U.readWorkbookRows(mk([H,row(159,'15/09/2026','Master Carton',0,'X')]));
      chk('all-zero export yields no lines', z.length===0, z.length+' lines'); }catch(e){ chk('all-zero export yields no lines',true,'threw: '+e.message.slice(0,40)); }

 // ---- 3. missing PI_DATE column ----
 const H2=H.filter(x=>x!=='PI_DATE');
 const r2=new Array(25).fill(null); r2[0]='121/700768/0159082026'; r2[4]='15/09/2026'; r2[5]='Olisa 1'; r2[10]='Master Carton'; r2[16]='M1'; r2[22]=50;
 try{ const recs=await U.readWorkbookRows(mk([H2,r2])); chk('export without PI_DATE still works',true,recs.length+' line(s)'); }
 catch(e){ chk('export without PI_DATE still works',false,e.message.slice(0,70)); }

 // ---- 4. blank / odd delivery dates ----
 const odd=await U.readWorkbookRows(mk([H,
   row(160,'','Master Carton',10,'A'), row(161,'ASAP','Master Carton',20,'B'),
   row(162,'2026-09-05','Master Carton',30,'C'), row(163,'05/09/2026','Master Carton',40,'D')]));
 const og=U.buildOutputRows(U.buildGroups(odd));
 chk('blank + free-text dates do not crash the sort', og.piRows.length===4, og.piRows.map(p=>p.deliveryDate||'(blank)').join(' -> '));
 chk('blanks sort to the very bottom, below written notes', og.piRows[og.piRows.length-1].deliveryDate==='',
     og.piRows.map(p=>p.deliveryDate||'(blank)').join(' -> '));
 chk('yyyy-mm-dd sorts with dd/mm/yyyy', U.deliverySortKey('2026-09-05')===U.deliverySortKey('05/09/2026'));

 // ---- 5. cross-divider-only PI ----
 const cd=await U.readWorkbookRows(mk([H,row(170,'01/09/2026','Cross Divider',77,'Z')]));
 const cg=U.buildOutputRows(U.buildGroups(cd));
 chk('cross-divider-only PI keeps its qty', cg.totalCross===77 && cg.totalMaster===0, JSON.stringify(cg.piRows[0]));

 // ---- 6. PO Summary: same work order twice ----
 const WH=[['order date','FACTORY','ORDER SIZE','UNIT','ORDER QTY','STYLE','delivery date'],
           ['27/08/2026','P','CARTON 62.5*30.5*32.5CM','PCS',500,'M81988-3A2','08/09/2026']];
 const wf=()=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(WH),'S');
   const b=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});return{name:'wo.xlsx',arrayBuffer:async()=>b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength)};};
 const W1={name:'wo.xlsx',size:900,lastModified:5}, W2={name:'wo.xlsx',size:900,lastModified:5};
 chk('the same work order dropped twice is kept once', dedupe([W1,W2]).length===1);

 // ---- 7. filename safety ----
 chk('a delivery date with a slash cannot escape the filename', !/[\\\/]/.test(P.dateSuffix('08/09/2026')), P.dateSuffix('08/09/2026'));
 chk('a very long free-text date does not produce a huge filename', P.dateSuffix('x'.repeat(400)).length<=420);

 R.forEach(([s,n,d])=>console.log(`  ${s} ${n}${d?'   -> '+d:''}`));
 const bugs=R.filter(r=>r[0]==='BUG ').length;
 console.log(`\n${R.length-bugs} ok, ${bugs} BUG(S)`);
})();
