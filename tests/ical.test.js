"use strict";
const {test}=require('node:test'),assert=require('node:assert/strict');
const P=require('../planner.js'),I=require('../ical.js');
const now='2026-03-02T12:00:00Z';
function fixture(){
 const snapshot={schema:'focus-planner-snapshot',version:1,complete:true,source:'outlook-extract',
   timeZone:'America/New_York',capturedAt:now,windowStart:'2026-03-02',windowEnd:'2026-04-13',events:[
     {id:'already-on-outlook',start:'2026-03-10T13:00:00Z',end:'2026-03-10T15:00:00Z',status:'busy',kind:'existing-focus',title:'Focus Time'}]};
 const plan=P.createPlan(snapshot,now);
 for(const date of ['2026-03-06','2026-03-09'])P.addBlock(plan,date,540,690,now).status='approved';
 return plan;
}
const unfold=ics=>ics.replace(/\r\n /g,'');
test('exports only approved future reservations as private busy events with standard framing',()=>{
 const plan=fixture(),before=JSON.stringify(plan);
 P.addBlock(plan,'2026-03-04',540,660,now).status='rejected';
 const expected=JSON.stringify(plan),file=I.create(plan,now),text=unfold(file.text);
 assert.equal(file.count,2);assert.equal(file.filename,'focus-time-2026-03-06-to-2026-03-09.ics');
 assert.equal((text.match(/BEGIN:VEVENT/g)||[]).length,2);assert.equal((text.match(/SUMMARY:Focus Time/g)||[]).length,2);
 for(const line of ['VERSION:2.0','METHOD:PUBLISH','STATUS:CONFIRMED','TRANSP:OPAQUE','CLASS:PRIVATE','X-MICROSOFT-CDO-BUSYSTATUS:BUSY'])assert.ok(text.includes(line+'\r\n'));
 assert.ok(text.startsWith('BEGIN:VCALENDAR\r\n'));assert.ok(text.endsWith('END:VCALENDAR\r\n'));
 assert.equal(/(?<!\r)\n/.test(text),false);assert.equal(/ATTENDEE|ORGANIZER|BEGIN:VALARM|already-on-outlook/.test(text),false);
 assert.equal(JSON.stringify(plan),expected);assert.notEqual(before,expected);
});
test('uses the snapshot zone across DST independently of the host timezone',()=>{
 const file=unfold(I.create(fixture(),now).text);
 assert.ok(file.includes('DTSTART:20260306T140000Z\r\nDTEND:20260306T163000Z'));
 assert.ok(file.includes('DTSTART:20260309T130000Z\r\nDTEND:20260309T153000Z'));
 assert.equal(I.toUTC('2026-03-09',570,'Asia/Kathmandu'),'20260309T034500Z');
 assert.equal(I.toUTC('2026-03-09',540,'Australia/Adelaide'),'20260308T223000Z');
 assert.throws(()=>I.toUTC('2026-03-08',150,'America/New_York'),/ambiguous or nonexistent/);
 assert.throws(()=>I.toUTC('2026-11-01',90,'America/New_York'),/ambiguous or nonexistent/);
});
test('blocks partial approval and unresolved conflicts, including direct export calls',()=>{
 const plan=fixture();plan.sessions.at(-1).status='proposed';
 assert.equal(I.readiness(plan,now).ready,false);assert.throws(()=>I.create(plan,now),/approve or reject 1/);
 plan.sessions.at(-1).status='conflicted';assert.throws(()=>I.create(plan,now),/1 unresolved conflicts/);
 plan.sessions.at(-1).status='rejected';assert.equal(I.create(plan,now).count,1);
});
test('excludes started/past blocks and does not let stale proposals block future exports',()=>{
 const plan=fixture();P.addBlock(plan,'2026-03-02',540,660,now).status='approved';
 P.addBlock(plan,'2026-03-03',540,660,now);
 assert.equal(I.create(plan,'2026-03-03T14:00:00Z').count,2);
 assert.equal(I.readiness(plan,'2026-04-14T12:00:00Z').ready,false);
 assert.throws(()=>I.create(plan,'2026-04-14T12:00:00Z'),/No new approved future blocks/);
 assert.equal(I.readiness(null,now).ready,false);
});
test('stable unique event UIDs survive JSON restore and renumbering on refresh',()=>{
 const plan=fixture(),uids=p=>unfold(I.create(p,now).text).split('\r\n').filter(l=>l.startsWith('UID:'));
 const expected=uids(plan);assert.equal(new Set(expected).size,2);
 const restored=P.validatePlan(JSON.parse(JSON.stringify(plan)));
 restored.sessions.forEach(s=>s.id+=100);restored.nextId+=100;
 assert.deepEqual(uids(restored),expected);
 assert.deepEqual(unfold(I.create(plan,'2026-03-02T12:01:00Z').text).split('\r\n').filter(l=>l.startsWith('UID:')),expected);
});
test('invalid or newly overlapping saved plans cannot be exported',()=>{
 const plan=fixture();plan.sessions.at(-1).start=600;plan.sessions.at(-1).end=599;
 assert.throws(()=>I.create(plan,now),/Invalid saved focus block/);
 const overlapping=fixture();overlapping.snapshot.events.push({id:'new-meeting',start:'2026-03-06T14:00:00Z',end:'2026-03-06T15:00:00Z',kind:'busy',status:'busy'});
 assert.throws(()=>I.create(overlapping,now),/overlaps calendar availability/);
});
test('folds lines at 75 UTF-8 octets without splitting characters and escapes calendar text',()=>{
 const line='DESCRIPTION:'+('é😀,;\\\n'.repeat(30)),escaped='DESCRIPTION:'+I.escapeText(line.slice(12));
 const folded=I.foldLine(escaped);
 for(const physical of folded.split('\r\n'))assert.ok(Buffer.byteLength(physical,'utf8')<=75);
 assert.equal(folded.replace(/\r\n /g,''),escaped);
 assert.equal(I.escapeText('a\\b,c;d\r\ne\rf\ng'),'a\\\\b\\,c\\;d\\ne\\nf\\ng');
});
