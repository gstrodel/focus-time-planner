"use strict";
/* Pure snapshot and planning functions; every date and minute is in the snapshot timezone. */
(function(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.FocusPlanner = api;
})(typeof window === "object" ? window : globalThis, function() {
  const CONFIG = Object.freeze({horizonDays:42, dailyTargetMinutes:120, minBlockMinutes:60,
    startMinute:540, endMinute:1050, preferMorning:true, meetingStartMinute:570,
    protectCommute:false, approvalRequired:true});
  const ACTIVE = new Set(["proposed", "approved", "existing"]);
  const STATUSES = new Set(["proposed", "approved", "existing", "rejected"]);
  const pad = n => String(n).padStart(2, "0");
  const addDays = (key, n) => new Date(Date.parse(key+"T12:00:00Z")+n*86400000).toISOString().slice(0,10);
  const isDate = key => typeof key === "string" && /^\d{4}-\d{2}-\d{2}$/.test(key) && !isNaN(Date.parse(key+"T12:00:00Z")) && addDays(key,0)===key;
  const isISO = value => typeof value === "string" && /^\d{4}-\d{2}-\d{2}T.*(?:Z|[+-]\d{2}:\d{2})$/.test(value) && Number.isFinite(Date.parse(value));
  const overlap = (a,b) => a.start < b.end && b.start < a.end;
  function localParts(value, timeZone) {
    const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", {timeZone,year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",second:"2-digit",hourCycle:"h23"}).formatToParts(new Date(value)).map(x=>[x.type,x.value]));
    return {date:`${p.year}-${p.month}-${p.day}`, minute:+p.hour*60 + +p.minute, seconds:+p.second};
  }
  function validateSnapshot(input) {
    if (!input || input.schema!=="focus-planner-snapshot" || input.version!==1 || input.complete!==true) throw Error("Use a complete Focus Planner snapshot (version 1). No data has been replaced.");
    if (!isDate(input.windowStart) || !isDate(input.windowEnd) || addDays(input.windowStart,42)!==input.windowEnd) throw Error("The extract must cover exactly 42 calendar days, with an exclusive end date.");
    if (!isISO(input.capturedAt)) throw Error("The extract needs a valid capture timestamp.");
    try {localParts(input.capturedAt,input.timeZone);} catch (_) {throw Error("The extract has an invalid timezone.");}
    if (typeof input.timeZone!=="string" || !input.timeZone) throw Error("The extract needs an explicit timezone.");
    if (!Array.isArray(input.events) || input.events.length>10000) throw Error("Invalid event list.");
    const ids = new Set();
    const events = input.events.map(e=>{
      if (!e || typeof e.id!=="string" || !e.id || ids.has(e.id) || !isISO(e.start) || !isISO(e.end) || Date.parse(e.end)<=Date.parse(e.start) || !["busy","tentative","oof","unknown"].includes(e.status) || !["busy","existing-focus"].includes(e.kind)) throw Error("An event is invalid or duplicated. Import stopped; your saved plan is unchanged.");
      ids.add(e.id);
      const text=(value,limit)=>{
        if(value===undefined || value===null)return "";
        if(typeof value!=="string" || value.length>limit)throw Error("Invalid meeting description or organizer metadata.");
        return value;
      };
      return {id:e.id,start:e.start,end:e.end,status:e.status,kind:e.kind,allDay:!!e.allDay,
        title:text(e.title,1000),description:text(e.description,1000000),
        organizer:{name:text(e.organizer?.name,500),email:text(e.organizer?.email,500)}};
    });
    return {schema:input.schema,version:1,source:input.source==="demo"?"demo":"outlook-extract",timeZone:input.timeZone,capturedAt:input.capturedAt,windowStart:input.windowStart,windowEnd:input.windowEnd,complete:true,events};
  }
  function workdays(snapshot) {
    return Array.from({length:42},(_,i)=>addDays(snapshot.windowStart,i)).filter(d=>![0,6].includes(new Date(d+"T12:00:00Z").getUTCDay()));
  }
  function intervals(snapshot) {
    const byDate = {};
    for (const event of snapshot.events) {
      const a=localParts(event.start,snapshot.timeZone), b=localParts(event.end,snapshot.timeZone);
      for(let date=a.date<snapshot.windowStart?snapshot.windowStart:a.date;date<=b.date && date<snapshot.windowEnd;date=addDays(date,1)) {
        const start=date===a.date?a.minute:0;
        const end=date===b.date?b.minute+((b.seconds || new Date(event.end).getUTCMilliseconds())?1:0):1440;
        if(end>start)(byDate[date] ||= []).push({id:event.id,start,end,kind:event.kind,status:event.status,title:event.title,description:event.description,organizer:event.organizer});
      }
    }
    return byDate;
  }
  function merge(items, start=0, end=1440) {
    const sorted=items.map(x=>({start:Math.max(start,x.start),end:Math.min(end,x.end)})).filter(x=>x.start<x.end).sort((a,b)=>a.start-b.start);
    const result=[];
    for(const item of sorted){const last=result.at(-1);if(last && item.start<=last.end)last.end=Math.max(last.end,item.end);else result.push({...item});}
    return result;
  }
  function gaps(items,start,end,min=60) {
    const free=[];let cursor=start;
    for(const item of merge(items,start,end)){if(item.start-cursor>=min)free.push({start:cursor,end:item.start});cursor=Math.max(cursor,item.end);}
    if(end-cursor>=min)free.push({start:cursor,end});
    return free;
  }
  function createPlan(raw, now=new Date().toISOString(), config={}) {
    const snapshot=validateSnapshot(raw);
    const plan={schema:"focus-planner-plan",version:1,snapshot,config:{...CONFIG,protectCommute:config.protectCommute===true},createdAt:now,savedAt:null,sessions:[],nextId:1};
    const busy=intervals(snapshot);
    // Count existing calendar focus only where it has no competing meeting or time off.
    for(const date of workdays(snapshot)) {
      const day=busy[date]||[];
      const other=day.filter(e=>e.kind!=="existing-focus");
      for(const e of merge(day.filter(e=>e.kind==="existing-focus"),CONFIG.startMinute,CONFIG.endMinute)) {
        for(const slot of gaps(other,e.start,e.end,CONFIG.minBlockMinutes))plan.sessions.push({id:plan.nextId++,date,...slot,status:"existing"});
      }
    }
    return plan;
  }
  const scheduled = (plan,date) => merge(plan.sessions.filter(s=>s.date===date && ACTIVE.has(s.status))).reduce((n,s)=>n+s.end-s.start,0);
  const hasContinuousFocus = (plan,date) => merge(plan.sessions.filter(s=>s.date===date && ACTIVE.has(s.status))).some(s=>s.end-s.start>=CONFIG.dailyTargetMinutes);
  function available(plan,date,now=new Date().toISOString(),excludeId=null,respectRejected=false) {
    const today=localParts(now,plan.snapshot.timeZone);
    if(!workdays(plan.snapshot).includes(date)||date<today.date)return [];
    const lower=Math.max(plan.config.protectCommute?570:540,date===today.date?Math.ceil((today.minute+1)/15)*15:0);
    const occupied=[...(intervals(plan.snapshot)[date]||[]),...plan.sessions.filter(s=>s.date===date&&s.id!==excludeId&&(ACTIVE.has(s.status)||(respectRejected&&s.status==="rejected")))];
    return gaps(occupied,lower,CONFIG.endMinute,CONFIG.minBlockMinutes);
  }
  function addBlock(plan,date,start,end,now=new Date().toISOString()) {
    if(!Number.isInteger(start)||!Number.isInteger(end)||end-start<CONFIG.minBlockMinutes||!available(plan,date,now).some(g=>g.start<=start&&g.end>=end))throw Error("Choose at least 60 minutes within an available future window.");
    const session={id:plan.nextId++,date,start,end,status:"proposed"};
    plan.sessions.push(session);return session;
  }
  function extendBlock(plan,id,minutes=30,now=new Date().toISOString()) {
    const s=plan.sessions.find(s=>s.id===id);
    const today=localParts(now,plan.snapshot.timeZone);
    if(!s||!["proposed","approved"].includes(s.status)||s.date<today.date||s.date===today.date&&s.start<=today.minute||!Number.isInteger(minutes)||minutes<=0||!available(plan,s.date,now,s.id).some(g=>g.start<=s.start&&g.end>=s.end+minutes))throw Error("This block cannot be extended into the next 30 minutes.");
    s.end+=minutes;s.status="proposed";return s;
  }
  function generate(plan, now=new Date().toISOString()) {
    let count=0;
    for(const date of workdays(plan.snapshot)) {
      let need=Math.max(0,CONFIG.dailyTargetMinutes-scheduled(plan,date));
      const options=()=>available(plan,date,now,null,true).sort((a,b)=>(a.start<720?0:1)-(b.start<720?0:1)||a.start-b.start);
      // Secure a continuous two-hour block before filling smaller gaps.
      if(!hasContinuousFocus(plan,date)) {
        const whole=options().find(g=>g.end-g.start>=CONFIG.dailyTargetMinutes);
        if(whole){addBlock(plan,date,whole.start,whole.start+CONFIG.dailyTargetMinutes,now);count++;continue;}
      }
      while(need>=CONFIG.minBlockMinutes) {
        const best=options()[0];if(!best)break;
        let length=Math.min(best.end-best.start,need);
        if(need-length>0&&need-length<CONFIG.minBlockMinutes&&need>=2*CONFIG.minBlockMinutes)length=need-CONFIG.minBlockMinutes;
        addBlock(plan,date,best.start,best.start+length,now);need-=length;count++;
      }
    }
    return count;
  }
  function validatePlan(raw) {
    if(!raw || raw.schema!=="focus-planner-plan" || raw.version!==1 || !isISO(raw.createdAt) || (raw.savedAt!==null && !isISO(raw.savedAt)))throw Error("Invalid saved plan.");
    const plan=createPlan(raw.snapshot,raw.createdAt,raw.config);
    const dates=new Set(workdays(plan.snapshot)),byDate=intervals(plan.snapshot),ids=new Set();
    if(!Array.isArray(raw.sessions) || raw.sessions.length>10000 || !Number.isSafeInteger(raw.nextId))throw Error("Invalid saved sessions.");
    const imported=raw.sessions.map(s=>{
      if(s?.status==="completed")s={...s,status:"approved"};
      if(s?.status==="missed")s={...s,status:"rejected"};
      if(!s || !Number.isSafeInteger(s.id)||s.id<1||ids.has(s.id)||!dates.has(s.date)||!STATUSES.has(s.status)||!Number.isInteger(s.start)||!Number.isInteger(s.end)||s.start<540||s.end>1050||s.end-s.start<60)throw Error("Invalid saved focus block.");
      ids.add(s.id);
      if(ACTIVE.has(s.status)) {
        if(s.status==="existing") {
          if(!plan.sessions.some(e=>e.date===s.date&&e.start===s.start&&e.end===s.end))throw Error("Existing focus does not match the calendar extract.");
        } else if((byDate[s.date]||[]).some(e=>overlap(e,s)))throw Error("Saved focus overlaps calendar availability.");
      }
      return {id:s.id,date:s.date,start:s.start,end:s.end,status:s.status};
    });
    for(let i=0;i<imported.length;i++)if(ACTIVE.has(imported[i].status)&&imported.slice(i+1).some(s=>s.date===imported[i].date&&ACTIVE.has(s.status)&&overlap(s,imported[i])))throw Error("Saved focus blocks overlap.");
    if(plan.sessions.some(e=>!imported.some(s=>s.status==="existing"&&s.date===e.date&&s.start===e.start&&s.end===e.end)))throw Error("Saved plan is missing existing calendar focus.");
    if(raw.nextId<=Math.max(0,...ids))throw Error("Invalid next session ID.");
    return {...plan,sessions:imported,nextId:raw.nextId,savedAt:raw.savedAt};
  }
  function demo(now=new Date().toISOString()) {
    const timeZone="America/New_York",start=localParts(now,timeZone).date,events=[];
    for(let i=0;i<42;i++){const date=addDays(start,i);if([0,6].includes(new Date(date+"T12:00:00Z").getUTCDay()))continue;
      // Synthetic winter example, only used as explicitly labeled demo data.
      events.push({id:"demo-"+i,start:date+"T15:00:00Z",end:date+"T16:00:00Z",status:"busy",kind:"busy",allDay:false});}
    return {schema:"focus-planner-snapshot",version:1,source:"demo",timeZone,capturedAt:now,windowStart:start,windowEnd:addDays(start,42),complete:true,events};
  }
  return {CONFIG,ACTIVE,addDays,localParts,overlap,validateSnapshot,validatePlan,workdays,intervals,merge,gaps,createPlan,scheduled,hasContinuousFocus,available,addBlock,extendBlock,generate,demo};
});
