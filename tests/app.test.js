"use strict";
/* DOM harness for app control/persistence behavior. This is not a browser renderer. */
const {test}=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs'),P=require('../planner.js');
class Element {
 constructor(tag='div'){this.tag=tag;this.children=[];this.listeners={};this.dataset={};this.textContent='';this.disabled=false;this.checked=false;this.value='';this.files=[];this.style={};this.attributes={};}
 append(...items){this.children.push(...items);}
 replaceChildren(...items){this.children=[...items];this.textContent='';}
 addEventListener(type,fn){this.listeners[type]=fn;}
 setAttribute(key,value){this.attributes[key]=value;}
 click(){if(!this.disabled)return this.listeners.click?.({target:this});}
}
function boot(storage=new Map(),bootstrap=null,localCalendar=null,windowOptions={}){
 const nodes=Object.fromEntries([...fs.readFileSync(require.resolve('../index.html'),'utf8').matchAll(/id="([^"]+)"/g)].map(m=>[m[1],new Element()]));
 if(bootstrap){nodes.bootstrap=new Element('script');nodes.bootstrap.textContent=JSON.stringify(bootstrap);}
 const localStorage={getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v)};
 const document={getElementById:id=>nodes[id]||null,createElement:tag=>new Element(tag),addEventListener(){}};
 const downloads=[];
 const context={window:{FocusPlanner:P,FocusICalendar:require('../ical.js'),FocusLocalCalendar:localCalendar,showDirectoryPicker:localCalendar?()=>{}:undefined,setInterval(){},addEventListener(){},confirm:()=>true,...windowOptions},document,localStorage,Blob,URL:{createObjectURL:blob=>{downloads.push(blob);return "blob:test";},revokeObjectURL(){}},setTimeout:fn=>fn(),Date,console};
 vm.runInNewContext(fs.readFileSync(require.resolve('../app.js'),'utf8'),context);
 return {nodes,storage,localStorage,downloads,state:()=>JSON.parse(storage.get('focus-planner-v0.4')),import:async raw=>{nodes.import.files=[{size:100,text:async()=>JSON.stringify(raw)}];await nodes.import.listeners.change({target:nodes.import});}};
}
test('app imports, generates, approves, saves and restores actual user decisions',async()=>{
 const app=boot();assert.equal(app.nodes.generate.disabled,true);await app.import(P.demo());assert.equal(app.nodes.generate.disabled,false);assert.equal(app.nodes.weeks.children.flatMap(x=>x.children).length,30);
 const before=app.state().sessions.length;app.nodes.approveAll.click();assert.ok(app.state().sessions.some(s=>s.status==='approved'));app.nodes.generate.click();assert.equal(app.state().sessions.length,before);
 const restored=boot(app.storage);assert.deepEqual(restored.state(),app.state());assert.equal(restored.nodes.saveStatus.textContent,'Restored saved plan');assert.match(restored.nodes.reserved.textContent,/h$/);
});
test('malformed input cannot overwrite saved state, and denied storage remains visible',async()=>{
 const app=boot();await app.import(P.demo());const before=JSON.stringify(app.state());await app.import({schema:'bad'});assert.equal(JSON.stringify(app.state()),before);assert.match(app.nodes.notice.textContent,/Import failed/);
 app.localStorage.setItem=()=>{throw Error('quota')};app.nodes.save.click();assert.match(app.nodes.notice.textContent,/could not save/);assert.equal(app.nodes.export,undefined);assert.match(app.nodes.notice.textContent,/Changes may be lost/);
});
test('bundled saved plan initializes offline, then existing browser decisions take precedence',()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='approved';plan.savedAt=new Date().toISOString();const app=boot(new Map(),plan);assert.equal(app.state().sessions[0].status,'approved');
 app.nodes.approveAll.click();const restored=boot(app.storage,P.createPlan(P.demo()));assert.deepEqual(restored.state(),app.state());
});
test('legacy outcome statuses migrate to planning decisions without outcome controls',async()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='completed';plan.sessions[1].status='rejected';plan.savedAt=new Date().toISOString();const app=boot();await app.import(plan);assert.equal(app.state().sessions[0].status,'approved');assert.equal(app.state().sessions[1].status,'rejected');assert.equal(app.nodes.insights,undefined);
});
const descendants=node=>[node,...node.children.filter(e=>typeof e==='object').flatMap(descendants)];
test('day calendar retains early and evening meetings, omits overnight invites and removes card time ticks',async()=>{
 const raw=P.demo(),date=P.workdays(raw)[0];
 raw.events=[['overnight','02:05','03:00'],['early','06:00','07:00'],['evening','21:00','22:00'],['late','22:00','23:00']].map(([id,start,end])=>({id,title:id,start:date+'T'+start+':00-05:00',end:date+'T'+end+':00-05:00',status:'busy',kind:'busy'}));
 const app=boot();await app.import(raw);
 assert.equal(app.nodes.sourceStatus,undefined);
 assert.deepEqual(app.state().snapshot.events.map(e=>e.id),['early','evening']);
 assert.equal(descendants(app.nodes.weeks).some(e=>e.className==='timeline-ticks'),false);
 const titles=descendants(app.nodes.detailBody).filter(e=>e.className==='calendar-event cal-busy').map(e=>e.children[0].textContent);
 assert.deepEqual(titles,['early','evening']);
});
test('coverage notes distinguish missing focus, secured focus and full-day PTO',()=>{
 const raw=P.demo(),dates=P.workdays(raw),[empty,focused,off]=dates.slice(1,4);
 raw.events=[{id:'away',start:off+'T00:00:00Z',end:P.addDays(off,1)+'T00:00:00Z',status:'oof',kind:'busy',allDay:true,title:'PTO',isOrganizer:true}];
 const plan=P.createPlan(raw);P.addBlock(plan,focused,600,720);plan.savedAt=new Date().toISOString();
 const app=boot(new Map([['focus-planner-v0.4',JSON.stringify(plan)]]));
 const note=date=>descendants(descendants(app.nodes.weeks).find(e=>e.dataset.date===date)).find(e=>e.className?.split(' ').includes('day-note'));
 assert.equal(note(empty).textContent,'No focus planned');assert.equal(note(empty).className,'day-note no-focus');
 assert.equal(note(focused).textContent,'✅ Focus Time Secured');assert.equal(note(focused).className,'day-note focus-secured');assert.equal(note(off).className,'day-note');
 const card=date=>descendants(app.nodes.weeks).find(e=>e.dataset.date===date);
 assert.equal(descendants(card(off)).some(e=>e.className==='meeting-caption'),false);
 assert.equal(descendants(card(empty)).find(e=>e.className==='meeting-caption').textContent,'in meetings');
});
test('teammate PTO does not appear in coverage, all-day banners or the timed day calendar',async()=>{
 const raw=P.demo(),date=P.workdays(raw)[0],end=P.addDays(date,1);
 raw.events=[true,false].map((allDay,i)=>({id:'teammate-'+i,start:date+'T00:00:00Z',end:end+'T00:00:00Z',status:'oof',kind:'busy',allDay,title:'Colleague PTO',isOrganizer:false}));
 const app=boot();await app.import(raw);
 const card=descendants(app.nodes.weeks).find(e=>e.dataset.date===date);card.click();
 assert.equal(descendants(card).some(e=>/PTO|OOO/.test(e.textContent)),false);
 assert.equal(descendants(app.nodes.detailBody).some(e=>/Colleague/.test(e.textContent)),false);
 assert.ok(app.state().sessions.some(s=>s.date===date));
 const restored=boot(app.storage);assert.deepEqual(restored.state(),app.state());
});
test('calendar export waits for decisions, prompts for location, and preserves local approvals',async()=>{
 let options,written,closed=false;
 const app=boot(new Map(),null,null,{showSaveFilePicker:async opts=>{options=opts;return {createWritable:async()=>({write:async blob=>{written=await blob.text();},close:async()=>{closed=true;}})};}});
 assert.equal(app.nodes.exportIcs.disabled,true);assert.equal(app.nodes.icsStatus.className,'hint');
 await app.import(P.demo());assert.equal(app.nodes.exportIcs.disabled,true);assert.equal(app.nodes.icsStatus.className,'hint export-warning');
 app.nodes.approveAll.click();assert.equal(app.nodes.exportIcs.disabled,false);assert.equal(app.nodes.icsStatus.className,'hint');const before=JSON.stringify(app.state());
 await app.nodes.exportIcs.click();assert.equal(closed,true);assert.equal(options.id,'focus-time-ics');assert.ok(options.suggestedName.endsWith('.ics'));
 assert.deepEqual(Array.from(options.types[0].accept['text/calendar']),['.ics']);assert.match(written,/BEGIN:VCALENDAR/);
 assert.equal(JSON.stringify(app.state()),before);assert.match(app.nodes.notice.textContent,/Saved .*focus blocks/);assert.equal(app.downloads.length,0);
 // Extending any approved block requires approval again before the next export.
 const state=app.state(),extendable=state.sessions.find(s=>s.status==='approved'&&P.available(state,s.date,undefined,s.id).some(g=>g.start<=s.start&&g.end>=s.end+30));
 assert.ok(extendable);descendants(app.nodes.weeks).find(e=>e.dataset.date===extendable.date).click();
 const extend=descendants(app.nodes.detailBody).find(e=>e.textContent==='Extend 30 min'&&!e.disabled);
 assert.ok(extend);extend.click();assert.equal(app.nodes.exportIcs.disabled,true);assert.equal(app.nodes.icsStatus.className,'hint export-warning');
});
test('canceling or failing calendar save reports the result and keeps export available',async()=>{
 for(const failure of ['cancel','write','close']){
   let aborted=false;
   const app=boot(new Map(),null,null,{showSaveFilePicker:async()=>{
     if(failure==='cancel')throw Object.assign(Error('canceled'),{name:'AbortError'});
     return {createWritable:async()=>({write:async()=>{if(failure==='write')throw Error('Disk full');},close:async()=>{throw Error('Close failed');},abort:async()=>{aborted=true;}})};
   }});
   await app.import(P.demo());app.nodes.approveAll.click();const before=JSON.stringify(app.state());
   await app.nodes.exportIcs.click();assert.equal(JSON.stringify(app.state()),before);assert.equal(app.nodes.exportIcs.disabled,false);
   assert.match(app.nodes.notice.textContent,failure==='cancel'?/canceled/:/export failed/);assert.equal(aborted,failure!=='cancel');assert.equal(app.downloads.length,0);
 }
});
test('browsers without a save picker download an ICS with an explicit location limitation',async()=>{
 const app=boot();await app.import(P.demo());app.nodes.approveAll.click();
 assert.match(app.nodes.icsStatus.textContent,/Ask where to save/);await app.nodes.exportIcs.click();
 assert.equal(app.downloads.length,1);assert.equal(app.downloads[0].type,'text/calendar;charset=utf-8');assert.match(await app.downloads[0].text(),/SUMMARY:Focus Time/);
 assert.match(app.nodes.notice.textContent,/browser controls the save location/);
});
test('day grid shows titles and discards descriptions and organizers on import and restore',async()=>{
 const raw=P.demo();raw.events[0].title='Planning review';raw.events[0].description='<img src=x onerror=alert(1)> Agenda';raw.events[0].organizer={name:'Organizer Name',email:'person@example.test'};
 const app=boot();await app.import(raw);
 const nodes=descendants(app.nodes.detailBody),box=nodes.find(e=>e.className==='calendar-event cal-busy');assert.ok(box);assert.equal(box.children[0].textContent,'Planning review');
 assert.equal(nodes.some(e=>e.tag==='details'),false);assert.equal('description' in app.state().snapshot.events[0],false);assert.equal('organizer' in app.state().snapshot.events[0],false);
 const saved=app.state();saved.snapshot.events[0].description='Legacy body';saved.snapshot.events[0].organizer={name:'Legacy organizer'};app.storage.set('focus-planner-v0.4',JSON.stringify(saved));
 const restored=boot(app.storage);assert.equal('description' in restored.state().snapshot.events[0],false);assert.equal('organizer' in restored.state().snapshot.events[0],false);assert.deepEqual(restored.state().sessions,saved.sessions);
});
test('all-day flags are hidden for every availability; timed multi-day events remain in the grid',async()=>{
 const raw=P.demo();const date=P.workdays(raw)[0],next=P.addDays(date,1);
 raw.events=['busy','tentative','free','oof','unknown'].map((status,i)=>({id:'all-'+i,start:date+'T05:00:00Z',end:next+'T05:00:00Z',status,kind:'busy',allDay:true,title:'All-day '+status}));
 raw.events.push({id:'timed',start:date+'T05:00:00Z',end:next+'T05:00:00Z',status:'busy',kind:'busy',allDay:false,title:'Timed overnight shift'});
 const app=boot();await app.import(raw);const boxes=descendants(app.nodes.detailBody).filter(e=>e.attributes.role==='listitem');
 assert.equal(boxes.length,1);assert.equal(boxes[0].children[0].textContent,'Timed overnight shift');assert.equal(boxes[0].style.height,1440*1.6+'px');
 assert.equal(P.meetingLoad(app.state().snapshot,date).workingMinutes,0);
});
test('overlaps get separate lanes; touching events use the full width and times determine box size',async()=>{
 const raw=P.demo();const date=P.workdays(raw)[0];raw.events=[['a','14:00','15:00'],['b','14:30','15:30'],['c','15:30','16:00']].map(([id,start,end])=>({id,start:date+'T'+start+':00Z',end:date+'T'+end+':00Z',kind:'busy',status:'busy',title:id}));
 const app=boot();await app.import(raw);const boxes=descendants(app.nodes.detailBody).filter(e=>e.className==='calendar-event cal-busy');
 assert.equal(boxes[0].style.width,'calc(50% - 6px)');assert.equal(boxes[1].style.left,'calc(50% + 3px)');assert.equal(boxes[2].style.width,'calc(100% - 6px)');
 assert.equal(boxes[0].style.height,'96px');assert.equal(boxes[2].style.height,'48px');assert.equal(boxes[0].style.top,'0px');
});


test('connected folder refresh updates app decisions and exposes conflicts without losing valid approvals',async()=>{
 const L=require('../local-calendar.js');let content,modified=Date.now(),handleSaved=null;
 const handle={name:'SuperlistSync',queryPermission:async()=>"granted",getFileHandle:async()=>({getFile:async()=>({size:100,lastModified:modified,text:async()=>content})})};
 const module={...L,Connection:class extends L.Connection{
   constructor(options){super({...options,pick:async()=>handle,storage:async(action,value)=>{if(action==='put')handleSaved=value;if(action==='delete')handleSaved=null;return handleSaved;}});}
 }};
 const app=boot(new Map(),null,module),raw={...P.demo(),source:'outlook-extract'};content=JSON.stringify(raw);
 await app.nodes.connectCalendar.click();app.nodes.approveAll.click();const before=app.state(),block=before.sessions.find(s=>s.date>P.localParts(new Date().toISOString(),raw.timeZone).date);
 const stamp=m=>block.date+'T'+String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0')+':00-05:00';
 content=JSON.stringify({...raw,capturedAt:new Date(Date.parse(raw.capturedAt)+60000).toISOString(),events:[...raw.events,{id:'new-meeting',start:stamp(block.start),end:stamp(block.end),status:'busy',kind:'busy'}]});modified++;
 await app.nodes.refreshCalendar.click();const state=app.state();assert.equal(state.sessions.find(s=>s.id===block.id).status,'conflicted');assert.ok(state.sessions.some(s=>s.status==='approved'));assert.match(app.nodes.coverageSummary.textContent,/conflict/);
 const day=descendants(app.nodes.weeks).find(e=>e.dataset.date===block.date);day.click();const dismiss=descendants(app.nodes.detailBody).find(e=>e.textContent==='Dismiss conflict');assert.ok(dismiss);dismiss.click();assert.equal(app.state().sessions.find(s=>s.id===block.id).status,'rejected');
 await app.nodes.disconnectCalendar.click();assert.equal(app.nodes.refreshCalendar.disabled,true);assert.ok(app.state());
});


test('coverage cards show month labels, exclude days off from targets and explain all-day time off in the sidebar',async()=>{
 const raw=P.demo(),date=P.workdays(raw)[0],end=P.addDays(date,1);
 raw.events=[{id:'away',start:date+'T00:00:00Z',end:end+'T00:00:00Z',status:'tentative',kind:'busy',allDay:true,title:'Office Closed - Holiday'}];
 const app=boot();await app.import(raw);
 const cards=descendants(app.nodes.weeks).filter(e=>e.dataset.date);
 assert.equal(cards.length,30);
 for(const card of cards){
  const month=descendants(card).find(e=>e.className==='month');
  assert.equal(month.textContent,new Date(card.dataset.date+'T12:00:00Z').toLocaleDateString('en-US',{timeZone:'UTC',month:'short'}).toUpperCase());
 }
 const off=cards.find(e=>e.dataset.date===date);off.click();
 assert.ok(descendants(off).some(e=>e.textContent==='PTO / OOO'));
 assert.ok(descendants(off).some(e=>e.textContent==='No focus planning needed'));
 assert.ok(descendants(off).some(e=>e.className==='segment off'));
 assert.equal(app.nodes.target.textContent,'58h');assert.equal(app.nodes.focusDays.textContent,'29/29');
 assert.equal(app.nodes.meetingGoal.textContent,'29/29');assert.match(app.nodes.coverageSummary.textContent,/1 PTO\/OOO days excluded/);
 const details=descendants(app.nodes.detailBody);
 assert.ok(details.some(e=>e.className==='time-off-banner'&&e.textContent.includes('Office Closed - Holiday')));
 assert.equal(details.some(e=>e.textContent==='Add more focus time'),false);
 assert.equal(app.state().sessions.some(s=>s.date===date&&P.ACTIVE.has(s.status)),false);
});


