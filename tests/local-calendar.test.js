"use strict";
const {test}=require("node:test"),assert=require("node:assert/strict");
const P=require("../planner.js"),L=require("../local-calendar.js");
const NOW="2026-11-18T12:00:00Z";
const snapshot=()=>({...P.demo(NOW),source:"outlook-extract"});
test("Power Automate arrays and body/value wrappers normalize offsets and discard private details",()=>{
  const events=[{id:"a",subject:"Planning",start:"wrong",end:"wrong",startWithTimeZone:"2026-11-18T09:30:00-05:00",endWithTimeZone:"2026-11-18T10:30:00-05:00",showAs:"Busy",isAllDay:false,body:"private",organizer:"private"},
    {id:"b",isCancelled:true},{id:"c",showAs:"workingElsewhere"},{id:"d",responseStatus:{response:"declined"}}];
  for(const raw of [events,{value:events},{body:{value:events}},{body:JSON.stringify({value:events})}]){
    const result=L.snapshotFromJSON(raw,NOW);
    assert.equal(result.events.length,1);assert.equal(result.events[0].status,"busy");
    assert.equal(P.localParts(result.events[0].start,result.timeZone).minute,570);
    assert.equal("body" in result.events[0],false);assert.equal("organizer" in result.events[0],false);
    assert.equal(result.windowStart,"2026-11-18");assert.equal(result.windowEnd,"2026-12-30");
  }
});
test("Graph UTC timestamps, all-day flags, existing focus and explicit window metadata are supported",()=>{
  const result=L.snapshotFromJSON({windowStart:"2026-11-17",windowEnd:"2026-12-29",timeZone:"Eastern Standard Time",value:[{
    id:"x",subject:"Focus time",start:{dateTime:"2026-11-18T14:00:00",timeZone:"UTC"},end:{dateTime:"2026-11-18T16:00:00",timeZone:"UTC"},showAs:"busy",isAllDay:"false"
  }]},NOW);
  assert.equal(result.windowStart,"2026-11-17");assert.equal(result.timeZone,"America/New_York");
  assert.equal(result.events[0].kind,"existing-focus");assert.equal(result.events[0].allDay,false);
});
test("ambiguous times, pagination, duplicates, wrong schemas and unexpanded recurring series fail closed",()=>{
  const event={id:"x",start:"2026-11-18T14:00:00Z",end:"2026-11-18T15:00:00Z"};
  for(const raw of [{value:[{...event,start:"2026-11-18T14:00:00"}]},{value:[],"@odata.nextLink":"next"},
    {complete:false,value:[]},{value:[event,event]},{schema:"focus-planner-plan"},{value:[{...event,type:"seriesMaster"}]}])assert.throws(()=>L.snapshotFromJSON(raw,NOW));
});
test("refresh retains valid approvals and rejections, flags conflicts and proposes safe replacements",()=>{
  const raw=snapshot(),plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
  const approved=plan.sessions[0],safe=plan.sessions[1],rejected=plan.sessions[2];
  approved.status="approved";safe.status="approved";rejected.status="rejected";
  const next={...raw,capturedAt:"2026-11-18T12:30:00Z",events:[...raw.events,{id:"new",start:approved.date+"T"+String(Math.floor(approved.start/60)).padStart(2,"0")+":00:00-05:00",end:approved.date+"T"+String(Math.floor(approved.end/60)).padStart(2,"0")+":00:00-05:00",status:"busy",kind:"busy"}]};
  const result=L.reconcile(plan,next,NOW);
  assert.equal(result.conflicts,1);assert.equal(result.plan.sessions.find(s=>s.id===approved.id).status,"conflicted");
  assert.equal(result.plan.sessions.find(s=>s.id===safe.id).status,"approved");
  assert.equal(result.plan.sessions.find(s=>s.id===rejected.id).status,"rejected");
  assert.ok(result.plan.sessions.some(s=>s.date===approved.date&&s.status==="proposed"));
  assert.equal(result.plan.sessions.some(s=>s.date===approved.date&&P.ACTIVE.has(s.status)&&P.intervals(next)[s.date].some(e=>P.overlap(e,s))),false);
  assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(result.plan))),result.plan);
});
test("rolling windows keep in-range decisions with unique IDs and recognize blocks now on Outlook",()=>{
  const raw=snapshot(),plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
  const booked=plan.sessions.find(s=>s.date>raw.windowStart);booked.status="approved";
  const hhmm=m=>String(Math.floor(m/60)).padStart(2,"0")+":"+String(m%60).padStart(2,"0")+":00-05:00";
  const next={...raw,windowStart:P.addDays(raw.windowStart,1),windowEnd:P.addDays(raw.windowEnd,1),capturedAt:"2026-11-19T12:00:00Z",events:[...raw.events,{
    id:"booked",start:booked.date+"T"+hhmm(booked.start),end:booked.date+"T"+hhmm(booked.end),status:"busy",kind:"existing-focus"}]};
  const result=L.reconcile(plan,next,"2026-11-19T12:00:00Z").plan;
  assert.equal(result.sessions.some(s=>s.date<next.windowStart),false);
  assert.equal(result.sessions.filter(s=>s.date===booked.date&&s.start===booked.start&&s.end===booked.end&&P.ACTIVE.has(s.status)).length,1);
  assert.ok(result.sessions.some(s=>s.status==="existing"));assert.equal(new Set(result.sessions.map(s=>s.id)).size,result.sessions.length);
});
test("stale exports and timezone changes never replace an existing plan",()=>{
  const plan=P.createPlan(snapshot(),NOW),before=JSON.stringify(plan);
  assert.throws(()=>L.reconcile(plan,{...snapshot(),capturedAt:"2026-11-18T11:00:00Z"},NOW),/older/);
  assert.throws(()=>L.reconcile(plan,{...snapshot(),timeZone:"Europe/London"},NOW),/timezone/);
  assert.equal(JSON.stringify(plan),before);
});
function fixture(){
  let text=JSON.stringify(snapshot()),modified=Date.parse(NOW),permission="granted",remembered=null,reads=0,applied=[];
  const messages=[],handle={name:"SuperlistSync",queryPermission:async()=>permission,requestPermission:async()=>{permission="granted";return permission;},
    getFileHandle:async name=>{assert.equal(name,L.FILE);reads++;return {getFile:async()=>({size:Buffer.byteLength(text),lastModified:modified,text:async()=>text})};}};
  const storage=async(action,value)=>{if(action==="put")remembered=value;if(action==="delete")remembered=null;return remembered;};
  const connection=new L.Connection({pick:async()=>handle,storage,status:message=>messages.push(message),apply:async value=>applied.push(value)});
  return {connection,handle,storage,messages,applied,get reads(){return reads;},setText:value=>{text=value;modified++;},setPermission:value=>{permission=value;}};
}
test("connection re-resolves filename, skips unchanged imports and survives a partial JSON write",async()=>{
  const f=fixture();await f.connection.connect();assert.equal(f.applied.length,1);
  await f.connection.refresh();assert.equal(f.applied.length,1);assert.equal(f.reads,2);
  f.setText("{");await f.connection.refresh();assert.equal(f.applied.length,1);assert.match(f.messages.at(-1),/last valid plan/);
  f.setText(JSON.stringify({...snapshot(),capturedAt:"2026-11-18T13:00:00Z"}));await f.connection.refresh();
  assert.equal(f.applied.length,2);assert.equal(f.reads,4);
  await f.connection.disconnect();await f.connection.refresh();assert.equal(f.reads,4);assert.equal(await f.storage("get"),null);
});
test("restoration never prompts; a user refresh can reauthorize folder access",async()=>{
  const f=fixture();await f.connection.connect();f.setPermission("prompt");
  const restored=new L.Connection({storage:f.storage,apply:value=>f.applied.push(value),status:message=>f.messages.push(message)});
  await restored.restore();assert.equal(f.applied.length,1);assert.match(f.messages.at(-1),/Refresh now/);
  await restored.refresh(true);assert.equal(f.applied.length,2);
});
test("disconnect cancels an in-flight read before it can change the plan",async()=>{
  const f=fixture();await f.connection.connect();let release;
  f.handle.getFileHandle=()=>new Promise(resolve=>{release=resolve;});
  const pending=f.connection.refresh();await new Promise(resolve=>setImmediate(resolve));
  await f.connection.disconnect();release({getFile:async()=>({size:2,lastModified:Date.parse(NOW),text:async()=>"[]"})});await pending;
  assert.equal(f.applied.length,1);assert.match(f.messages.at(-1),/disconnected/);
});

test('Outlook OOO availability aliases and all-day fields survive import and prevent focus planning',()=>{
 for(const status of ['oof','OOO','OutOfOffice','out of office']){
  const raw={value:[{id:'away',subject:'Shutdown',startWithTimeZone:'2026-11-19T00:00:00Z',endWithTimeZone:'2026-11-21T00:00:00Z',showAs:status,isAllDay:true}]};
  const snapshot=L.snapshotFromJSON(raw,NOW),plan=P.createPlan(snapshot,NOW);P.generate(plan,NOW);
  assert.equal(snapshot.events[0].status,'oof');assert.equal(P.scheduled(plan,'2026-11-19'),0);
  assert.equal(P.meetingLoad(snapshot,'2026-11-20').workingMinutes,0);
  assert.equal(P.meetingLoad(snapshot,'2026-11-18').workingMinutes,480);
  assert.equal(P.meetingLoad(snapshot,'2026-11-21').workingMinutes,480);
 }
 const snapshot=L.snapshotFromJSON([{id:'away',subject:'Time off',start:{dateTime:'2026-11-19T00:00:00',timeZone:'UTC'},end:{dateTime:'2026-11-20T00:00:00',timeZone:'UTC'},show_as:'oof',is_all_day:true}],NOW);
 assert.equal(snapshot.events[0].allDay,true);assert.equal(snapshot.events[0].status,'oof');
});
test('a newly imported all-day closure removes approved blocks from active planning without replacements on days off',()=>{
 const raw=snapshot(),plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
 const block=plan.sessions.find(s=>s.date==='2026-11-19');block.status='approved';
 const next={...raw,capturedAt:'2026-11-18T13:00:00Z',events:[...raw.events,{id:'closure',title:'Holiday',allDay:true,status:'tentative',kind:'busy',start:'2026-11-19T00:00:00Z',end:'2026-11-20T00:00:00Z'}]};
 const result=L.reconcile(plan,next,NOW);
 assert.equal(result.plan.sessions.find(s=>s.id===block.id).status,'conflicted');
 assert.equal(P.scheduled(result.plan,block.date),0);assert.equal(P.dailyFocusTarget(next,block.date),0);
 assert.deepEqual(P.validatePlan(result.plan),result.plan);
});

test('Power Automate and Graph ownership survive import without retaining organizer identities',()=>{
 const base={id:'pto',subject:'PTO',startWithTimeZone:'2026-11-20T00:00:00-05:00',endWithTimeZone:'2026-11-21T00:00:00-05:00',showAs:'busy',isAllDay:true,organizer:{emailAddress:{address:'private@example.test'}}};
 for(const fields of [{responseType:'accepted'},{responseType:'tentativelyAccepted'},{responseType:'notResponded'},{responseStatus:{response:'accepted'}},{response_status:{response:'accepted'}},{isOrganizer:false},{is_organizer:'false'},{isOrganizer:false,responseType:'organizer'}]){
  const raw=L.snapshotFromJSON([ {...base,...fields} ],NOW);
  assert.equal(raw.events[0].isOrganizer,false);assert.equal('organizer' in raw.events[0],false);
  assert.equal(P.meetingLoad(raw,'2026-11-20').workingMinutes,480);
 }
 for(const fields of [{responseType:'organizer'},{responseStatus:{response:'organizer'}},{isOrganizer:true},{isOrganizer:'true',responseType:'accepted'}]){
  const raw=L.snapshotFromJSON([{...base,...fields}],NOW);
  assert.equal(raw.events[0].isOrganizer,true);assert.equal(P.meetingLoad(raw,'2026-11-20').workingMinutes,0);
 }
 for(const fields of [{},{responseType:null},{responseType:'none'},{responseType:'unexpected'}]){
  const raw=L.snapshotFromJSON([{...base,...fields}],NOW);
  assert.equal('isOrganizer' in raw.events[0],false);assert.equal(P.meetingLoad(raw,'2026-11-20').workingMinutes,0);
 }
});

test('teammate PTO arriving on refresh preserves approved focus instead of creating conflicts',()=>{
 const raw=snapshot(),plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
 const block=plan.sessions.find(s=>s.date==='2026-11-20');block.status='approved';
 const away={id:'colleague',title:'Colleague PTO',allDay:true,status:'oof',kind:'busy',isOrganizer:false,start:'2026-11-20T00:00:00-05:00',end:'2026-11-21T00:00:00-05:00'};
 const result=L.reconcile(plan,{...raw,events:[...raw.events,away]},NOW);
 assert.equal(result.conflicts,0);assert.equal(result.plan.sessions.find(s=>s.id===block.id).status,'approved');
 assert.equal(P.scheduled(result.plan,block.date),240);assert.deepEqual(P.validatePlan(result.plan),result.plan);
});

