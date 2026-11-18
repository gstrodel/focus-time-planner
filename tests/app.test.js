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
function boot(storage=new Map(),bootstrap=null,localCalendar=null){
 const nodes=Object.fromEntries([...fs.readFileSync(require.resolve('../index.html'),'utf8').matchAll(/id="([^"]+)"/g)].map(m=>[m[1],new Element()]));
 if(bootstrap){nodes.bootstrap=new Element('script');nodes.bootstrap.textContent=JSON.stringify(bootstrap);}
 const localStorage={getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v)};
 const document={getElementById:id=>nodes[id]||null,createElement:tag=>new Element(tag),addEventListener(){}};
 const context={window:{FocusPlanner:P,FocusLocalCalendar:localCalendar,showDirectoryPicker:localCalendar?()=>{}:undefined,setInterval(){},addEventListener(){},confirm:()=>true},document,localStorage,Blob,URL:{createObjectURL:()=>"blob:test",revokeObjectURL(){}},setTimeout:fn=>fn(),Date,console};
 vm.runInNewContext(fs.readFileSync(require.resolve('../app.js'),'utf8'),context);
 return {nodes,storage,localStorage,state:()=>JSON.parse(storage.get('focus-planner-v0.4')),import:async raw=>{nodes.import.files=[{size:100,text:async()=>JSON.stringify(raw)}];await nodes.import.listeners.change({target:nodes.import});}};
}
test('app imports, generates, approves, saves and restores actual user decisions',async()=>{
 const app=boot();assert.equal(app.nodes.generate.disabled,true);await app.import(P.demo());assert.equal(app.nodes.generate.disabled,false);assert.equal(app.nodes.weeks.children.flatMap(x=>x.children).length,30);
 const before=app.state().sessions.length;app.nodes.approveAll.click();assert.ok(app.state().sessions.some(s=>s.status==='approved'));app.nodes.generate.click();assert.equal(app.state().sessions.length,before);
 const restored=boot(app.storage);assert.deepEqual(restored.state(),app.state());assert.equal(restored.nodes.saveStatus.textContent,'Restored saved plan');assert.match(restored.nodes.reserved.textContent,/h$/);
});
test('malformed input cannot overwrite saved state, and denied storage remains visible',async()=>{
 const app=boot();await app.import(P.demo());const before=JSON.stringify(app.state());await app.import({schema:'bad'});assert.equal(JSON.stringify(app.state()),before);assert.match(app.nodes.notice.textContent,/Import failed/);
 app.localStorage.setItem=()=>{throw Error('quota')};app.nodes.save.click();assert.match(app.nodes.notice.textContent,/could not save/);assert.equal(app.nodes.export.disabled,false);
});
test('bundled saved plan initializes offline, then existing browser decisions take precedence',()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='approved';plan.savedAt=new Date().toISOString();const app=boot(new Map(),plan);assert.equal(app.state().sessions[0].status,'approved');
 app.nodes.approveAll.click();const restored=boot(app.storage,P.createPlan(P.demo()));assert.deepEqual(restored.state(),app.state());
});
test('legacy outcome statuses migrate to planning decisions without outcome controls',async()=>{
 const plan=P.createPlan(P.demo());P.generate(plan);plan.sessions[0].status='completed';plan.sessions[1].status='rejected';plan.savedAt=new Date().toISOString();const app=boot();await app.import(plan);assert.equal(app.state().sessions[0].status,'approved');assert.equal(app.state().sessions[1].status,'rejected');assert.equal(app.nodes.insights,undefined);
});
const descendants=node=>[node,...node.children.filter(e=>typeof e==='object').flatMap(descendants)];
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
