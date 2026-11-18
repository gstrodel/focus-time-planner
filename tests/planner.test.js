"use strict";
const {test}=require("node:test"),assert=require("node:assert/strict"),P=require("../planner.js");
const NOW="2026-11-18T13:00:00Z";
const snapshot=(events=[])=>({schema:"focus-planner-snapshot",version:1,source:"outlook-extract",timeZone:"America/New_York",capturedAt:NOW,windowStart:"2026-11-18",windowEnd:"2026-12-30",complete:true,events});
const event=(id,start,end,kind="busy",status="busy")=>({id,start,end,kind,status});
test("42-day extract: weekday plans meet targets, stay in work hours and never overlap calendar events",()=>{
 const raw=snapshot([event("m","2026-11-18T14:30:00Z","2026-11-18T15:30:00Z"),event("oof","2026-11-19T05:00:00Z","2026-11-20T05:00:00Z","busy","oof")]);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);const busy=P.intervals(raw);
 assert.equal(P.workdays(raw).length,30);assert.equal(P.scheduled(plan,"2026-11-19"),0);
 for(const s of plan.sessions){assert.ok(s.start>=540&&s.end<=1050&&s.end-s.start>=60);assert.equal((busy[s.date]||[]).some(b=>P.overlap(b,s)),false);assert.equal(s.status,"proposed");}
 assert.equal(P.scheduled(plan,"2026-11-18"),120);
 const n=plan.sessions.length;assert.equal(P.generate(plan,NOW),0);assert.equal(plan.sessions.length,n);
});
test("90-minute morning gap is split to allow another minimum block and reach two hours",()=>{
 const plan=P.createPlan(snapshot([event("meeting","2026-11-18T15:30:00Z","2026-11-18T17:00:00Z")]),NOW);
 P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),120);assert.deepEqual(plan.sessions.filter(s=>s.date==="2026-11-18").map(s=>s.end-s.start),[60,60]);
});
test("existing focus counts once, overlapping meetings take priority, no calendar duplication",()=>{
 const plan=P.createPlan(snapshot([event("f","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z","existing-focus"),event("f2","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z","existing-focus"),event("m","2026-11-18T15:00:00Z","2026-11-18T16:00:00Z")]),NOW);
 assert.equal(P.scheduled(plan,"2026-11-18"),60);P.generate(plan,NOW);assert.equal(P.scheduled(plan,"2026-11-18"),120);
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
test("saved decisions, feedback and snapshot survive serialization; invalid imports fail closed",()=>{
 const plan=P.createPlan(snapshot(),NOW);P.generate(plan,NOW);plan.sessions[0].status="approved";plan.sessions[1].status="completed";plan.savedAt=NOW;
 assert.deepEqual(P.validatePlan(JSON.parse(JSON.stringify(plan))),plan);
 assert.throws(()=>P.validateSnapshot({...snapshot(),complete:false}));assert.throws(()=>P.validateSnapshot({...snapshot(),windowEnd:"2027-01-01"}));
 assert.throws(()=>P.validateSnapshot(snapshot([event("x","bad","2026-11-18T15:00:00Z")])));
 const bad=JSON.parse(JSON.stringify(plan));bad.sessions[1].start=bad.sessions[0].start;bad.sessions[1].date=bad.sessions[0].date;assert.throws(()=>P.validatePlan(bad));
 const collision=P.createPlan(snapshot([event("m","2026-11-18T14:00:00Z","2026-11-18T16:00:00Z")]),NOW);collision.sessions.push({id:1,date:"2026-11-18",start:540,end:600,status:"approved"});collision.nextId=2;assert.throws(()=>P.validatePlan(collision));
});
test("meeting titles, literal descriptions and organizers survive import and saved-plan restore",()=>{
 const raw=snapshot([{...event("m","2026-11-18T14:00:00Z","2026-11-18T15:00:00Z"),title:"Planning review",description:"Agenda\n<script>alert('literal text')</script>",organizer:{name:"Meeting Organizer",email:"organizer@example.test"}}]);
 const plan=P.createPlan(raw,NOW);P.generate(plan,NOW);const restored=P.validatePlan(JSON.parse(JSON.stringify(plan)));
 assert.deepEqual(restored.snapshot.events[0],plan.snapshot.events[0]);assert.equal(P.intervals(restored.snapshot)["2026-11-18"][0].description,raw.events[0].description);
 assert.throws(()=>P.validateSnapshot(snapshot([{...raw.events[0],description:{html:"not plain text"}}])));
});
