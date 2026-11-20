const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'playwright');
const assert=require('node:assert/strict'),path=require('node:path');
const P=require('../planner.js');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROMIUM_PATH||'/usr/bin/chromium',args:['--no-sandbox']});
 const context=await browser.newContext({timezoneId:'America/Los_Angeles',viewport:{width:1440,height:1000},acceptDownloads:true});
 const page=await context.newPage(),errors=[],network=[];page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(/^https?:/.test(r.url()))network.push(r.url());});
 const url='file://'+path.resolve(__dirname,'../index.html');
 await context.setOffline(true);await page.goto(url);assert.equal(await page.locator('#sourceStatus').count(),0);
 const raw=P.demo(),first=P.workdays(raw)[0];
 const upload=async data=>page.locator('#import').setInputFiles({name:'calendar.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(data))});
 await upload(raw);await page.waitForFunction(()=>document.querySelectorAll('.day').length===30);
 assert.equal(await page.locator('#exportIcs').isEnabled(),false);
 let state=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));assert.equal(state.snapshot.source,'demo');
 assert.ok(state.sessions.length>0);await page.locator('#approveAll').click();state=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));assert.ok(state.sessions.some(s=>s.status==='approved'));
 assert.equal(await page.locator('#exportIcs').isEnabled(),true);
 // Test the save-picker contract; native OS dialogs still need a local Windows check.
 await page.evaluate(()=>{window.showSaveFilePicker=async options=>{window.savedCalendarOptions=options;return {createWritable:async()=>({write:async blob=>{window.savedCalendarText=await blob.text();},close:async()=>{window.calendarClosed=true;}})};};});
 await page.locator('#exportIcs').click();await page.waitForFunction(()=>window.calendarClosed);
 const calendar=await page.evaluate(()=>({text:window.savedCalendarText,options:window.savedCalendarOptions}));
 assert.equal((calendar.text.match(/BEGIN:VEVENT/g)||[]).length,state.sessions.filter(s=>s.status==='approved').length);
 assert.match(calendar.text,/SUMMARY:Focus Time/);assert.match(calendar.text,/TRANSP:OPAQUE/);assert.ok(calendar.options.suggestedName.endsWith('.ics'));
 await page.evaluate(()=>{window.showSaveFilePicker=async()=>{throw new DOMException('Canceled','AbortError');};});
 await page.locator('#exportIcs').click();await page.waitForFunction(()=>document.querySelector('#notice').textContent.includes('canceled'));
 assert.equal(await page.locator('#exportIcs').isEnabled(),true);
 await page.evaluate(()=>{window.showSaveFilePicker=undefined;});
 const icsPromise=page.waitForEvent('download');await page.locator('#exportIcs').click();const ics=await icsPromise;assert.ok(ics.suggestedFilename().endsWith('.ics'));
 await page.reload();assert.ok(parseFloat(await page.locator('#reserved').innerText())>0);
 const s=state.sessions.find(s=>s.status==='approved'&&s.date>first);await page.locator(`[data-date="${s.date}"]`).click();
 assert.equal(await page.getByRole('button',{name:'Completed',exact:true}).count(),0);assert.equal(await page.getByRole('button',{name:'Missed',exact:true}).count(),0);
 const row=page.locator(`[data-session-id="${s.id}"]`);await row.getByRole('button',{name:'Extend 30 min',exact:true}).click();
 state=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));assert.equal(state.sessions.find(x=>x.id===s.id).end,s.end+30);assert.equal(state.sessions.find(x=>x.id===s.id).status,'proposed');
 assert.equal(await page.locator('#exportIcs').isEnabled(),false);
 await row.getByRole('button',{name:'Approve',exact:true}).click();await row.getByRole('button',{name:'Remove from plan',exact:true}).click();
 const before=await page.evaluate(()=>localStorage.getItem('focus-planner-v0.4'));await upload({invalid:true});assert.match(await page.locator('#notice').innerText(),/Import failed/);assert.equal(await page.evaluate(()=>localStorage.getItem('focus-planner-v0.4')),before);
 assert.equal(await page.locator('#export').count(),0);await page.locator('#save').click();const backup=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));P.validatePlan(backup);assert.ok(backup.sessions.some(s=>s.status==='rejected'));
 page.on('dialog',d=>d.accept());await page.locator('#reset').click();assert.equal(await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')).sessions.some(s=>s.status==='rejected')),false);
 await upload(backup);assert.ok(await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')).sessions.some(s=>s.status==='rejected')));
 await page.reload();assert.ok(await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')).sessions.some(s=>s.status==='rejected')));
 await page.locator('#reset').click();state=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));const proposed=state.sessions.find(s=>s.status==='proposed');await page.locator(`[data-date="${proposed.date}"]`).click();await page.locator(`[data-session-id="${proposed.id}"]`).getByRole('button',{name:'Reject',exact:true}).click();await page.locator('#generate').click();state=await page.evaluate(()=>JSON.parse(localStorage.getItem('focus-planner-v0.4')));assert.equal(state.sessions.find(s=>s.id===proposed.id).status,'rejected');
 await page.setViewportSize({width:390,height:844});assert.ok(await page.evaluate(()=>document.documentElement.scrollWidth<=window.innerWidth));
 await page.evaluate(()=>{Storage.prototype.setItem=function(){throw Error('Quota exceeded');};});await page.locator('#save').click();assert.match(await page.locator('#notice').innerText(),/could not save/);assert.match(await page.locator('#notice').innerText(),/Changes may be lost/);
 assert.deepEqual(errors,[]);assert.deepEqual(network,[]);await context.close();await browser.close();console.log('PASS: offline browser import, decisions, extension approval, removal, reload, local save/restore, invalid input, rejected blocks, mobile layout, storage failure; no external requests or page errors.');
})().catch(e=>{console.error(e);process.exit(1);});


