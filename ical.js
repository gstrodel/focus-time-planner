"use strict";
/* Local iCalendar export. Dates in the plan are wall times in its snapshot timezone. */
(function(root,factory){
  const api=factory(typeof module==="object"&&module.exports?require("./planner.js"):root.FocusPlanner);
  if(typeof module==="object"&&module.exports)module.exports=api;else root.FocusICalendar=api;
})(typeof window==="object"?window:globalThis,function(P){
  function readiness(plan,now=new Date().toISOString()){
    if(!plan)return {ready:false,sessions:[],message:"Import a calendar extract, then approve your focus proposals to export."};
    const today=P.localParts(now,plan.snapshot.timeZone);
    const future=s=>s.date>today.date||s.date===today.date&&s.start>today.minute;
    const upcoming=plan.sessions.filter(future);
    const pending=upcoming.filter(s=>s.status==="proposed").length;
    const conflicts=upcoming.filter(s=>s.status==="conflicted").length;
    const sessions=upcoming.filter(s=>s.status==="approved").sort((a,b)=>a.date.localeCompare(b.date)||a.start-b.start);
    if(pending||conflicts)return {ready:false,needsReview:true,sessions:[],message:`Before exporting, approve or reject ${pending} upcoming proposals and dismiss ${conflicts} unresolved conflicts.`};
    if(!sessions.length)return {ready:false,sessions:[],message:"No new approved future blocks to export. Existing Outlook focus is already on your calendar."};
    return {ready:true,sessions,message:`Ready to export ${sessions.length} approved future focus ${sessions.length===1?"block":"blocks"} as Busy appointments titled Focus Time.`};
  }
  const utcStamp=value=>new Date(value).toISOString().replace(/[-:]/g,"").replace(/\.\d{3}Z$/,"Z");
  function toUTC(date,minute,timeZone){
    const wall=Date.parse(date+"T00:00:00Z")+minute*60000;
    const offsets=new Set();
    // Sample both sides of DST changes, including zones with half-hour offsets.
    for(let hours=-36;hours<=36;hours+=12){
      const sample=wall+hours*3600000,p=P.localParts(sample,timeZone);
      offsets.add(Date.parse(p.date+"T00:00:00Z")+p.minute*60000-sample);
    }
    const candidates=[...offsets].map(offset=>wall-offset).filter(candidate=>{
      const p=P.localParts(candidate,timeZone);return p.date===date&&p.minute===minute;
    });
    if(candidates.length!==1)throw Error(`Cannot export an ambiguous or nonexistent local time on ${date} in ${timeZone}.`);
    return utcStamp(candidates[0]);
  }
  const escapeText=value=>String(value).replace(/\\/g,"\\\\").replace(/\r\n|\r|\n/g,"\\n").replace(/;/g,"\\;").replace(/,/g,"\\,");
  function foldLine(line){
    const encoder=new TextEncoder();let result="",bytes=0;
    for(const character of line){
      const size=encoder.encode(character).length;
      if(bytes+size>75){result+="\r\n ";bytes=1;}
      result+=character;bytes+=size;
    }
    return result;
  }
  function create(plan,now=new Date().toISOString()){
    const valid=P.validatePlan(plan),state=readiness(valid,now);
    if(!state.ready)throw Error(state.message);
    const lines=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Focus Time Planner//Calendar Export//EN","CALSCALE:GREGORIAN","METHOD:PUBLISH"];
    for(const s of state.sessions){
      // Stable across saving, restoring and refreshing a plan; numeric session IDs can be reassigned.
      const uid=`focus-${encodeURIComponent(valid.createdAt)}-${encodeURIComponent(valid.snapshot.timeZone)}-${s.date}-${s.start}-${s.end}@focus-time-planner.local`;
      lines.push("BEGIN:VEVENT","UID:"+uid,"DTSTAMP:"+utcStamp(now),
        "DTSTART:"+toUTC(s.date,s.start,valid.snapshot.timeZone),"DTEND:"+toUTC(s.date,s.end,valid.snapshot.timeZone),
        "SUMMARY:Focus Time","DESCRIPTION:"+escapeText("Approved focus reservation from Focus Planner. Planning timezone: "+valid.snapshot.timeZone+"."),
        "STATUS:CONFIRMED","TRANSP:OPAQUE","CLASS:PRIVATE","X-MICROSOFT-CDO-BUSYSTATUS:BUSY","END:VEVENT");
    }
    lines.push("END:VCALENDAR");
    return {text:lines.map(foldLine).join("\r\n")+"\r\n",count:state.sessions.length,
      filename:`focus-time-${state.sessions[0].date}-to-${state.sessions.at(-1).date}.ics`};
  }
  return {readiness,create,toUTC,foldLine,escapeText};
});

