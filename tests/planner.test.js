"use strict";
const {test}=require("node:test"),assert=require("node:assert/strict"),P=require("../planner.js");
const NOW="2026-11-18T13:00:00Z";
const snapshot=(events=[])=>({schema:"focus-planner-snapshot",version:1,source:"outlook-extract",timeZone:"America/New_York",capturedAt:NOW,windowStart:"2026-11-18",windowEnd:"2026-12-30",complete:true,events});
test('calendar import keeps 6 AM–10 PM Eastern meetings across offsets and daylight saving',()=>{
 for(const [date,offset] of [['2026-12-07','-05:00'],['2026-07-07','-04:00']]){
  const meeting=(id,start,end,endDate=date)=>({id,start:date+'T'+start+offset,end:endDate+'T'+end+offset,status:'busy',kind:'busy'});
  const events=[meeting('overnight','02:05:00','03:00:00'),meeting('ends-at-six','05:00:00','06:00:00'),
   meeting('six','06:00:00','07:00:00'),meeting('early-overlap','05:30:00','06:00:00.001'),
   meeting('late','21:00:00','22:00:00'),meeting('late-overlap','21:30:00','22:30:00'),meeting('starts-at-ten','22:00:00','23:00:00'),
   meeting('cross-midnight','23:00:00','05:00:00',P.addDays(date,1)),meeting('cross-midnight-overlap','23:00:00','07:00:00',P.addDays(date,1)),
   {...meeting('holiday','00:00:00','00:00:00',P.addDays(date,1)),allDay:true,status:'oof',title:'PTO'}];
  const raw={...snapshot(events),windowStart:date,windowEnd:P.addDays(date,42)};
  const expected=['six','early-overlap','late','late-overlap','cross-midnight-overlap','holiday'];
  assert.deepEqual(P.validateSnapshot(raw).events.map(e=>e.id),expected);
  assert.deepEqual(P.validateSnapshot({...raw,timeZone:'America/Los_Angeles'}).events.map(e=>e.id),expected);
  const utc={...raw,events:events.map(e=>({...e,start:new Date(e.start).toISOString(),end:new Date(e.end).toISOString()}))};
  assert.deepEqual(P.validateSnapshot(utc).events.map(e=>e.id),expected);
  assert.equal(Object.values(P.intervals(raw)).flat().some(e=>['overnight','starts-at-ten','cross-midnight'].includes(e.id)),false);
  const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
  const legacy={...plan,snapshot:raw};assert.deepEqual(P.validatePlan(legacy).snapshot.events.map(e=>e.id),expected);
 }
});
const event=(id,start,end,kind="busy",status="busy")=>({id,start,end,kind,status});
test("meeting limit uses unioned meeting time, excludes focus, and removes time off from both sides",()=>{
 const raw=snapshot([event("a","2026-11-18T14:00:00Z","2026-11-18T17:00:00Z"),event("b","2026-11-18T16:00:00Z","2026-11-18T18:00:00Z"),event("f","2026-11-18T18:00:00Z","2026-11-18T19:00:00Z","existing-focus"),event("away","2026-11-18T17:00:00Z","2026-11-18T19:00:00Z","busy","oof"),event("pto","2026-11-19T05:00:00Z","2026-11-20T05:00:00Z","busy","oof")]);
 assert.deepEqual(P.meetingLoad(raw,"2026-11-18"),{workingMinutes:360,meetingMinutes:150,ratio:150/360});
 assert.deepEqual(P.meetingLoad(raw,"2026-11-19"),{workingMinutes:0,meetingMinutes:0,ratio:null});
});
test("all-day reminders do not consume meetings, block focus or count as booked focus",()=>{
 for(const status of ["busy","tentative","unknown","free"])for(const kind of ["busy","existing-focus"]){
  const banner={...event("banner","2026-11-20T00:00:00Z","2026-11-21T00:00:00Z",kind,status),allDay:true,title:"Focus Friday"};
  const raw=snapshot([banner,event("meeting","2026-11-20T15:00:00Z","2026-11-20T16:00:00Z")]);
  const plan=P.createPlan(raw,NOW);
  assert.equal(P.scheduled(plan,"2026-11-20"),0);
  assert.deepEqual(P.meetingLoad(raw,"2026-11-20"),{workingMinutes:480,meetingMinutes:60,ratio:1/8});
  assert.deepEqual(P.intervals(raw,{includeFree:true})["2026-11-20"].map(e=>e.id),["meeting"]);
  P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-20"),240);
  assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(plan))),plan);
 }
});
test("manual additions can exceed two hours and round-trip through saved plans",()=>{
 const plan=P.createPlan(snapshot(),NOW);
 P.addBlock(plan,"2026-11-18",540,720,NOW);P.addBlock(plan,"2026-11-18",750,930,NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),360);assert.ok(P.hasContinuousFocus(plan,"2026-11-18"));
 assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(plan))),plan);
 assert.throws(()=>P.addBlock(plan,"2026-11-18",690,810,NOW));
});
test("manual additions respect calendar events, time off, work hours, commute and past time",()=>{
 const plan=P.createPlan(snapshot([event("m","2026-11-18T15:00:00Z","2026-11-18T16:00:00Z"),event("away","2026-11-19T05:00:00Z","2026-11-20T05:00:00Z","busy","oof")]),NOW,{protectCommute:true});
 for(const [date,start,end] of [["2026-11-18",540,660],["2026-11-18",570,690],["2026-11-18",1000,1120],["2026-11-19",570,750],["2026-11-21",570,750]])assert.throws(()=>P.addBlock(plan,date,start,end,NOW));
 assert.throws(()=>P.addBlock(plan,"2026-11-18",570,750,"2026-11-18T17:00:00Z"));
 assert.throws(()=>P.addBlock(plan,"2026-11-18",660,690,NOW));
});
test("extensions reset approval and stop at calendar conflicts or past starts",()=>{
 const plan=P.createPlan(snapshot([event("m","2026-11-18T18:00:00Z","2026-11-18T19:00:00Z")]),NOW);
 const s=P.addBlock(plan,"2026-11-18",600,720,NOW);s.status="approved";
 P.extendBlock(plan,s.id,30,NOW);assert.equal(s.end,750);assert.equal(s.status,"proposed");
 P.extendBlock(plan,s.id,30,NOW);assert.equal(s.end,780);assert.throws(()=>P.extendBlock(plan,s.id,30,NOW));
 assert.throws(()=>P.extendBlock(plan,s.id,30,"2026-11-18T16:00:00Z"));
});
test("legacy completed and missed statuses restore as planning decisions",()=>{
 const plan=P.createPlan(snapshot(),NOW);const s=P.addBlock(plan,"2026-11-18",540,660,NOW);
 s.status="completed";assert.equal(P.validatePlan(plan).sessions[0].status,"approved");
 s.status="missed";assert.equal(P.validatePlan(plan).sessions[0].status,"rejected");
});
test("42-day extract: weekday plans meet targets, stay in work hours and never overlap calendar events",()=>{
 const raw=snapshot([event("m","2026-11-18T14:30:00Z","2026-11-18T15:30:00Z"),event("oof","2026-11-19T05:00:00Z","2026-11-20T05:00:00Z","busy","oof")]);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);const busy=P.intervals(raw);
 assert.equal(P.workdays(raw).length,30);assert.equal(P.scheduled(plan,"2026-11-19"),0);
 for(const s of plan.sessions){assert.ok(s.start>=540&&s.end<=1050&&s.end-s.start>=60);assert.equal((busy[s.date]||[]).some(b=>P.overlap(b,s)),false);assert.equal(s.status,"proposed");}
 assert.equal(P.scheduled(plan,"2026-11-18"),240);
 const n=plan.sessions.length;assert.equal(P.generate(plan,NOW),0);assert.equal(plan.sessions.length,n);
});
test("continuous afternoon block takes priority over fragmented morning gaps",()=>{
 const plan=P.createPlan(snapshot([event("meeting","2026-11-18T15:30:00Z","2026-11-18T17:00:00Z")]),NOW);
 P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),240);
 const blocks=plan.sessions.filter(s=>s.date==="2026-11-18");
 assert.equal(blocks[0].end-blocks[0].start,120);assert.ok(blocks[0].start>=720);
 assert.ok(blocks.every(s=>s.end-s.start>=60));
});
test("existing focus counts once, overlapping meetings take priority, no calendar duplication",()=>{
 const plan=P.createPlan(snapshot([event("f","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z","existing-focus"),event("f2","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z","existing-focus"),event("m","2026-11-18T15:00:00Z","2026-11-18T16:00:00Z")]),NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),60);P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),240);assert.ok(P.hasContinuousFocus(plan,"2026-11-18"));
 assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(plan))),plan);
});
test("multi-day time off blocks every covered weekday and explicit offsets cross DST correctly",()=>{
 const raw=snapshot([event("away","2026-11-18T05:00:00Z","2026-11-21T05:00:00Z","busy","oof")]);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);for(const d of ["2026-11-18","2026-11-19","2026-11-20"])assert.equal(P.scheduled(plan,d),0);
 assert.equal(P.localParts("2026-11-01T05:30:00Z","America/New_York").minute,90);assert.equal(P.localParts("2026-11-01T06:30:00Z","America/New_York").minute,90);
});
test("snapshot timezone is independent of the machine timezone; fractional end minutes are blocked",()=>{
 const raw=snapshot([event("m","2026-11-18T14:00:00Z","2026-11-18T15:00:00.001Z")]);assert.equal(P.intervals(raw)["2026-11-18"][0].end,601);
 const plan=P.createPlan(raw,NOW,{protectCommute:true});P.generate(plan,NOW);assert.ok(plan.sessions.every(s=>s.start>=570));
});
test("past time is never proposed and rejected proposals are not silently regenerated",()=>{
 const plan=P.createPlan(snapshot(),NOW);P.generate(plan,"2026-11-19T17:03:00Z");assert.ok(plan.sessions.every(s=>s.date>"2026-11-19"||s.date==="2026-11-19"&&s.start>=735));
 const rejected=plan.sessions[0];rejected.status="rejected";P.generate(plan,"2026-11-19T17:03:00Z");assert.equal(plan.sessions.some(s=>s.id!==rejected.id&&s.date===rejected.date&&P.overlap(s,rejected)),false);
});
test("saved planning decisions and snapshot survive serialization; invalid imports fail closed",()=>{
 const plan=P.createPlan(snapshot(),NOW);P.generate(plan,NOW);plan.sessions[0].status="approved";plan.sessions[1].status="approved";plan.savedAt=NOW;
 assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(plan))),plan);
 assert.throws(()=>P.validateSnapshot({...snapshot(),complete:false}));assert.throws(()=>P.validateSnapshot({...snapshot(),windowEnd:"2027-01-01"}));
 assert.throws(()=>P.validateSnapshot(snapshot([event("x","bad","2026-11-18T15:00:00Z")])));
 const bad=JSON.parse(JSON.stringify(plan));bad.sessions[1].start=bad.sessions[0].start;bad.sessions[1].date=bad.sessions[0].date;assert.throws(()=>P.validatePlan(bad));
 const collision=P.createPlan(snapshot([event("m","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z")]),NOW);collision.sessions.push({id:1,date:"2026-11-18",start:540,end:600,status:"approved"});collision.nextId=2;assert.throws(()=>P.validatePlan(collision));
});
test("imports retain titles and all-day flags but discard legacy meeting details",()=>{
 const raw=snapshot([{...event("m","2026-11-18T14:00:00Z","2026-11-18T15:00:00Z"),allDay:true,title:"Planning review",description:{html:"legacy details"},organizer:{name:"Meeting Organizer",email:"organizer@example.test"}}]);
 const plan=P.createPlan(raw,NOW),restored=P.validatePlan(JSON.parse(JSON.stringify(plan)));
 assert.deepEqual(restored.snapshot.events[0],plan.snapshot.events[0]);assert.equal(restored.snapshot.events[0].title,"Planning review");
 const e=restored.snapshot.events[0];assert.equal(e.allDay,true);assert.equal('description' in e,false);assert.equal('organizer' in e,false);
 assert.deepEqual(P.intervals(restored.snapshot),{});
 assert.throws(()=>P.validateSnapshot(snapshot([{...raw.events[0],title:{html:"invalid title"}}])));
});
test("free invites do not block proposals or count as meetings",()=>{
 const raw=snapshot([{...event('free','2026-11-18T05:00:00Z','2026-11-19T05:00:00Z','busy','free'),allDay:true}]);
 const plan=P.createPlan(raw,NOW);assert.equal(P.intervals(plan.snapshot)['2026-11-18'],undefined);assert.deepEqual(P.intervals(plan.snapshot,{includeFree:true}),{});
 assert.equal(P.meetingLoad(plan.snapshot,'2026-11-18').meetingMinutes,0);P.addBlock(plan,'2026-11-18',540,660,NOW);assert.equal(P.scheduled(plan,'2026-11-18'),120);
});

test("light meeting days target four hours, inclusive of already scheduled focus",()=>{
 const raw=snapshot([event("m","2026-11-18T16:00:00Z","2026-11-18T18:00:00Z")]);
 const plan=P.createPlan(raw,NOW);assert.equal(P.dailyFocusTarget(raw,"2026-11-18"),240);
 P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),240);
 assert.ok(P.hasContinuousFocus(plan,"2026-11-18"));
 const n=plan.sessions.length;assert.equal(P.generate(plan,NOW),0);assert.equal(plan.sessions.length,n);
});
test("exactly 50% meetings still requires four hours of focus",()=>{
 const raw=snapshot([event("half","2026-11-18T14:30:00Z","2026-11-18T18:30:00Z")]);
 assert.equal(P.meetingLoad(raw,"2026-11-18").ratio,0.5);
 assert.equal(P.dailyFocusTarget(raw,"2026-11-18"),240);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),240);
 assert.equal(P.generate(plan,NOW),0);
});
test("over 50% meetings retains two-hour focus target",()=>{
 const raw=snapshot([event("heavy","2026-11-18T14:30:00Z","2026-11-18T19:30:00Z")]);
 assert.ok(P.meetingLoad(raw,"2026-11-18").ratio>0.5);
 assert.equal(P.dailyFocusTarget(raw,"2026-11-18"),120);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),120);
});
test("PTO removes work minutes and does not get focus proposals",()=>{
 const raw=snapshot([event("pto","2026-11-18T05:00:00Z","2026-11-19T05:00:00Z","busy","oof")]);
 assert.equal(P.dailyFocusTarget(raw,"2026-11-18"),0);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),0);
});
test("existing focus contributes to four-hour target without duplication",()=>{
 const raw=snapshot([event("existing","2026-11-18T15:00:00Z","2026-11-18T17:00:00Z","existing-focus")]);
 const plan=P.createPlan(raw,NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),120);
 P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),240);
 assert.equal(plan.sessions.filter(s=>s.date==="2026-11-18"&&s.status==="proposed").reduce((n,s)=>n+s.end-s.start,0),120);
});

test('December 29 shutdown and year-end closure exclude all covered days, with exclusive all-day ends',()=>{
 const raw={...snapshot(),windowStart:'2026-12-01',windowEnd:'2027-01-12',events:[
  {...event('closure','2026-12-25T00:00:00Z','2027-01-01T00:00:00Z','busy','tentative'),allDay:true,title:'Office Closed - Year-End Closing'},
  {...event('shutdown','2026-12-28T00:00:00Z','2027-01-02T00:00:00Z','busy','oof'),allDay:true,title:'Shutdown'},
  event('meeting1','2026-12-29T16:05:00Z','2026-12-29T17:00:00Z'),
  event('meeting2','2026-12-29T18:05:00Z','2026-12-29T19:00:00Z'),
  event('free','2026-12-29T14:35:00Z','2026-12-29T15:00:00Z','busy','free')
 ]};
 assert.deepEqual(P.meetingLoad(raw,'2026-12-29'),{workingMinutes:0,meetingMinutes:0,ratio:null});
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
 for(const date of ['2026-12-25','2026-12-28','2026-12-29','2026-12-30','2026-12-31','2027-01-01']){
  assert.equal(P.dailyFocusTarget(raw,date),0);assert.deepEqual(P.available(plan,date,NOW),[]);assert.equal(P.scheduled(plan,date),0);
  assert.throws(()=>P.addBlock(plan,date,540,660,NOW));
 }
 assert.equal(P.meetingLoad(raw,'2026-12-24').workingMinutes,480);
 assert.equal(P.meetingLoad(raw,'2027-01-02').workingMinutes,480);
 const meetings={...raw,events:raw.events.slice(2)};
 assert.equal(P.meetingLoad(meetings,'2026-12-29').meetingMinutes,110);
 assert.equal(P.meetingLoad(meetings,'2026-12-29').ratio,110/480);
 const closureOnly={...raw,events:raw.events.slice(0,1)};
 assert.equal(P.meetingLoad(closureOnly,'2026-12-29').workingMinutes,0);
 assert.equal(P.meetingLoad(closureOnly,'2027-01-01').workingMinutes,480);
});
test('all-day time-off titles require non-free availability and timed title mentions remain meetings',()=>{
 for(const title of ['PTO','OOO','Out of Office','Vacation','Holiday','Shutdown','Office Closed','Year-End Closing']){
  const away={...event('away','2026-11-19T00:00:00-05:00','2026-11-20T00:00:00-05:00'),allDay:true,title};
  assert.equal(P.meetingLoad(snapshot([away]),'2026-11-19').workingMinutes,0);
  assert.equal(P.meetingLoad(snapshot([{...away,status:'free'}]),'2026-11-19').workingMinutes,480);
 }
 const timed={...event('review','2026-11-19T15:00:00Z','2026-11-19T16:00:00Z'),title:'Holiday planning meeting'};
 assert.equal(P.meetingLoad(snapshot([timed]),'2026-11-19').meetingMinutes,60);
});

test('Nov 20 teammate PTO is ignored for full and partial days, including busy and OOO invites',()=>{
 for(const allDay of [true,false])for(const status of ['busy','tentative','oof','unknown','free']){
  const away={...event('colleague','2026-11-20T00:00:00-05:00',allDay?'2026-11-21T00:00:00-05:00':'2026-11-20T13:00:00-05:00','busy',status),allDay,title:'Colleague PTO',isOrganizer:false};
  const meeting=event('meeting','2026-11-20T15:00:00Z','2026-11-20T16:00:00Z');
  const raw=snapshot([away,meeting]),plan=P.createPlan(raw,NOW);P.generate(plan,NOW);
  assert.equal(P.isTimeOff(away),false);
  assert.deepEqual(P.intervals(raw,{includeFree:true})['2026-11-20'].map(e=>e.id),['meeting']);
  assert.deepEqual(P.meetingLoad(plan.snapshot,'2026-11-20'),{workingMinutes:480,meetingMinutes:60,ratio:1/8});
  assert.equal(P.scheduled(plan,'2026-11-20'),240);
  const restored=P.validatePlan(JSON.parse(JSON.stringify(plan)));
  assert.deepEqual(restored,plan);assert.equal(restored.snapshot.events[0].isOrganizer,false);
 }
});

test('own PTO and shared company closures still remove working time; timed ordinary invites still block',()=>{
 const base={...event('away','2026-11-20T00:00:00-05:00','2026-11-21T00:00:00-05:00'),allDay:true};
 for(const fields of [{title:'PTO',isOrganizer:true},{title:'Holiday',isOrganizer:false},{title:'Office Closed - Year-End Closing',isOrganizer:false},{title:'Shutdown',status:'oof',isOrganizer:false},{title:'PTO'}]){
  const raw=snapshot([{...base,...fields}]);
  assert.equal(P.meetingLoad(raw,'2026-11-20').workingMinutes,0);
 }
 const ownPartial={...base,allDay:false,start:'2026-11-20T09:30:00-05:00',end:'2026-11-20T13:00:00-05:00',title:'PTO',status:'oof',isOrganizer:true};
 assert.equal(P.meetingLoad(snapshot([ownPartial]),'2026-11-20').workingMinutes,270);
 const meeting={...ownPartial,title:'Planning review',status:'busy',isOrganizer:false};
 assert.equal(P.meetingLoad(snapshot([meeting]),'2026-11-20').meetingMinutes,210);
 // Personal absence wins over a holiday word in the same title.
 assert.deepEqual(P.intervals(snapshot([{...base,title:'Colleague PTO over the holiday',isOrganizer:false}]),{includeFree:true}),{});
 assert.deepEqual(P.intervals(snapshot([{...ownPartial,title:'Away',isOrganizer:false}])),{});
 assert.throws(()=>P.validateSnapshot(snapshot([{...base,isOrganizer:'false'}])),/ownership/);
});
