/* synctest.js — guards cross-device index sync (laptop builds, phone catches up). */
const fs=require('fs');
let pass=0,fail=0;
const t=(n,c,x)=>{c?(pass++,console.log('  ok   '+n)):(fail++,console.log('  FAIL '+n+(x?'  -> '+x:'')))};
const html=fs.readFileSync(process.argv[2]||'olisa.html','utf8');
const src=[...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]).sort((a,b)=>b.length-a.length)[0];

console.log('1. The laptop publishes its index');
t('a build uploads the index to Drive automatically', /if \(typeof uploadIndexToDrive === 'function'\) uploadIndexToDrive\(\);/.test(src));
t('the upload replaces the existing file rather than piling up copies', /method: 'PATCH'/.test(src));
t('the index is gzipped before upload', /const gz = await gzipBytes\(json\)/.test(src));

console.log('\n2. The phone picks it up — WARM resume (the bug)');
// The warm path used to `return` without ever looking for the shared index.
// slice from the warm-path marker to where the FULL connect begins (the second assignment,
// not the top-level declaration of the same name)
const warmStart = src.indexOf('// ---- Warm path');
const warmEnd = src.indexOf('lastConnectWasWarm = false;', warmStart);
const warm = warmStart >= 0 && warmEnd > warmStart ? src.slice(warmStart, warmEnd) : '';
t('the warm-resume block was found', warm.length > 500);
t('warm resume looks for a newer shared index before returning',
  /tryLoadIndexFromFolder\(dirHandle\)/.test(warm), 'warm path still returns without checking Drive');
t('it does not block the resume (stays fast)', /void \(async \(\) => \{[\s\S]{0,400}tryLoadIndexFromFolder/.test(warm));
t('a failure there cannot break the resume', /advisory only — must never break a warm resume/.test(warm));
t('the stamp is redrawn once a newer index is adopted', /adopted && typeof renderIndexStamp === 'function'/.test(warm));

console.log('\n3. The phone picks it up — while ALREADY OPEN');
t('a shared-index poller exists', /async function checkForNewerSharedIndex/.test(src));
t('it runs on the existing poll loop', /lastSharedIdxCheck >= SHAREDIDX_POLL_MS\) checkForNewerSharedIndex\(\)/.test(src));
t('it also runs every time the app returns to the foreground',
  /checkForNewerSharedIndex\(\);\s*\n\s*if \(Date\.now\(\) - lastNewPiCheck/.test(src));
t('and once shortly after connecting', /setTimeout\(checkForNewerSharedIndex, \d+\)/.test(src));

console.log('\n4. The poll is cheap — this runs on mobile data');
const poll = src.slice(src.indexOf('async function checkForNewerSharedIndex'), src.indexOf('async function checkForNewerSharedIndex') + 2200);
t('it asks only for metadata, not the file', /fields=files\(id,name,modifiedTime\)/.test(poll));
t('the index is downloaded ONLY when the stamp is newer',
  poll.indexOf('if (!stamp || stamp <=') < poll.indexOf('readIndexFile('), 'download happens before the freshness check');
t('a clock-skew margin prevents a download loop', /\+ 1000\) return;/.test(poll));
t('it never fights a running build', /indexBuilding\) return;/.test(poll));
t('it is skipped while offline', /navigator\.onLine === false\) return;/.test(poll));
t('overlapping runs are prevented', /if \(sharedIdxBusy\) return;/.test(poll));
t('a failed poll cannot disturb the loaded index', /must never disturb what is already loaded/.test(poll));

t('the file handle is built from the raw Drive listing object, not (id, name)',
  /new DriveFileHandle\(f\)/.test(poll) && !/new DriveFileHandle\(f\.id/.test(poll),
  'DriveFileHandle takes meta — (id,name) throws inside getFile()');

console.log('\n4b. The challan index travels too (it did not before)');
t('challanIndex is packed into the shared payload', /challanIndex: challanIndex \|\| \{\}/.test(src),
  'phone adopts an index with zero challans and challan lookups find nothing');
t('a challan count rides along for the status line', /challanCount: Object\.keys\(challanIndex/.test(src));
t('the receiving device adopts it', /challanIndex = data\.challanIndex;/.test(src));
t('it is persisted locally so it survives a reload', /idbSet\('challanIndex', \{ data: challanIndex/.test(src));
t('an OLD index file without challans cannot wipe a good local copy',
  /if \(data\.challanIndex && Object\.keys\(data\.challanIndex\)\.length\) \{/.test(src),
  'adopting an empty challanIndex would be worse than the bug');

console.log('\n4c. The freshness dot means ONE thing: age');
t('a missing challan count no longer forces amber', !/: !chCount \? 'hyellow'/.test(src),
  'a 34-minute-old index showed the same colour as a 5-day-old one');
t('green under 2 days, amber at 2, red at 5', /days >= IDX_STALE_DAYS \? 'hred'\s*\n\s*: days >= IDX_WARN_DAYS \? 'hyellow'\s*\n\s*: 'hgreen'/.test(src));
t('missing challans are stated in words instead of a colour', /no challan copies indexed yet/.test(src));

console.log('\n5. Adopting is still guarded — no thrashing, no going backwards');
t('an index of the same age or older is refused', /if \(piIndexBuiltAt && incoming <= piIndexBuiltAt\) return false;/.test(src));
t('an index from an older app version is refused', /data\.version !== PI_CACHE_VERSION/.test(src));
t('remarks and acceptances still merge even from an older-format file', /overlays merge before ANY other gate|Notes merge before ANY other gate/i.test(src));
t('the "new PIs in Drive" warning clears when a fresh index is adopted', /clearNewWorkOrderFlag\(\);\s*\n\s*if \(typeof renderIndexStamp/.test(src));

console.log('\n'+(fail?'FAILED':'PASSED')+` — ${pass} passed, ${fail} failed`);
process.exit(fail?1:0);
