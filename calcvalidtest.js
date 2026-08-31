/* calcvalidtest.js — guards the calculator against sending a price nobody measured. */
const fs=require('fs');
let pass=0,fail=0; const t=(n,c,x)=>{c?(pass++,console.log('  ok   '+n)):(fail++,console.log('  FAIL '+n+(x?'  -> '+x:'')))};
const html=fs.readFileSync('calculator.html','utf8');
const js=[...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]).sort((a,b)=>b.length-a.length)[0];
const badNumbers=new Function(js.match(/function badNumbers[\s\S]*?\n    \}/)[0]+'\nreturn badNumbers;')();

console.log('1. What counts as a real number');
t('a normal measurement passes', badNumbers({Length:71.5,Width:45,Height:32,Rate:0.8}).length===0);
t('a negative length is rejected', badNumbers({Length:-50,Width:45,Rate:0.8}).includes('Length'));
t('a blank box is rejected', badNumbers({Length:'',Width:45,Rate:0.8}).includes('Length'));
t('zero is rejected', badNumbers({Length:0,Width:45,Rate:0.8}).includes('Length'));
t('a zero rate is rejected', badNumbers({Length:71.5,Width:45,Rate:0}).includes('Rate'));
t('text is rejected', badNumbers({Length:'abc',Width:45,Rate:0.8}).includes('Length'));
t('Infinity is rejected', badNumbers({Length:Infinity,Width:45,Rate:0.8}).includes('Length'));
t('every bad field is named, not just the first', badNumbers({Length:0,Width:-1,Rate:0}).length===3);
t('a numeric string still passes', badNumbers({Length:'71.5',Width:'45',Rate:'0.80'}).length===0);

console.log('\n2. The guards are actually wired into both copy paths');
t('the single quote refuses before building the string', /const _bad = badNumbers\(_need\);\s*\n\s*if \(_bad\.length\) return refuseQuote\(_bad\);/.test(js));
t('Top\\/Bottom is not asked for a height it does not have', /if \(itemName !== 'Top\/Bottom'\) _need\.Height/.test(js));
t('the multi-row copy collects bad rows', /badRows\.push\(`item \$\{index \+ 1\}/.test(js));
t('the multi-row copy refuses if any row is bad', /if \(badRows\.length\) return refuseQuote\(badRows\);/.test(js));
t('the refusal names which item number was wrong', /item \$\{index \+ 1\} \(\$\{_bad\.join/.test(js));

console.log('\n3. Browser-level guards');
const nums=[...html.matchAll(/type="number"[^>]*>/g)].map(m=>m[0]);
t(`every number input carries min="0" (${nums.length} inputs)`, nums.every(n=>/\bmin="0"/.test(n)), nums.filter(n=>!/\bmin="0"/.test(n)).length+' without');
t('the modern Clipboard API is used first', /navigator\.clipboard && window\.isSecureContext/.test(js));
t('execCommand is kept as the fallback, not deleted', /document\.execCommand\('copy'\)/.test(js));

console.log('\n'+(fail?'FAILED':'PASSED')+` — ${pass} passed, ${fail} failed`);
process.exit(fail?1:0);
