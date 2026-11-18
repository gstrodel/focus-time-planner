"use strict";
(() => {
  const KEY = "focus-planner-v0.1";
  const CONFIG = Object.freeze({ horizonDays: 42, dailyTargetMinutes: 120, minBlockMinutes: 60, startMinute: 9*60, endMinute: 17*60, preferMorning: true, conflictRecovery: true, approvalRequired: true });
  const $ = id => document.getElementById(id);
  const pad = n => String(n).padStart(2, "0");
  const dayKey = d => [d.getFullYear(), pad(d.getMonth()+1), pad(d.getDate())].join("-");
  const dateFromKey = key => { const [y,m,d]=key.split("-").map(Number); return new Date(y,m-1,d); };
  const plusDays = (d,n) => { const x = new Date(d); x.setDate(x.getDate()+n); return x; };
  const displayTime = m => `${Math.floor(m/60)%12||12}:${pad(m%60)} ${m<720?"AM":"PM"}`;
  const weekdays = dates => dates.filter(d=>{const w=d.getDay();return w!==0&&w!==6;});
  const overlap=(a,b)=>a.start<b.end&&b.start<a.end;
  const demoMeetings = date => {
    if(liveMode)return liveBusy?.[date]||[];
    const n=Math.floor((dateFromKey(date)-dateFromKey(anchor))/86400000);
    const result=[];
    // Stable simulated meetings, varied by day and week.
    if(n%3!==1)result.push({start:9*60+30,end:10*60+30,subject:"Team standup"});
    if(n%4!==2)result.push({start:11*60,end:12*60,subject:"Working session"});
    if(n%5===0)result.push({start:13*60,end:15*60,subject:"Project review"});
    if(n%7===3)result.push({start:9*60,end:11*60,subject:"Planning meeting"});
    if(n%9===5)result.push({start:14*60,end:16*60,subject:"Partner check-in"});
    return result;
  };
  const anchor=dayKey(new Date());
  let liveBusy = null;
  let liveMode = false;
  const createState=()=>({anchor, sessions:[], extraBusy:[], nextId:1});
  let state;
  try{state=JSON.parse(localStorage.getItem(KEY))||createState();if(state.anchor!==anchor||!Array.isArray(state.sessions)||!Array.isArray(state.extraBusy))state=createState();}
  catch(_){state=createState();}
  let selected=anchor;
  const dates=Array.from({length:CONFIG.horizonDays},(_,i)=>plusDays(dateFromKey(anchor),i));
  const workdays=weekdays(dates);
  const inHorizon = date => workdays.some(d=>dayKey(d)===date);
  const store=()=>{ if(!liveMode)localStorage.setItem(KEY,JSON.stringify(state)); };
  const sessionsOn=date=>state.sessions.filter(s=>s.date===date);
  const active=s=>["proposed","approved","completed"].includes(s.status);
  const scheduledMinutes=date=>sessionsOn(date).filter(active).reduce((sum,s)=>sum+s.end-s.start,0);
  const busyOn=date=>[...demoMeetings(date),...state.extraBusy.filter(e=>e.date===date),...sessionsOn(date).filter(active)];
  function gaps(busy){
    const blocked=busy.map(e=>({start:Math.max(CONFIG.startMinute,e.start),end:Math.min(CONFIG.endMinute,e.end)})).filter(e=>e.start<e.end).sort((a,b)=>a.start-b.start);
    const merged=[];
    for(const e of blocked){const last=merged[merged.length-1];if(last&&e.start<=last.end)last.end=Math.max(last.end,e.end);else merged.push({...e});}
    const free=[];let cur=CONFIG.startMinute;
    for(const e of merged){if(e.start-cur>=CONFIG.minBlockMinutes)free.push({start:cur,end:e.start});cur=Math.max(cur,e.end);}
    if(CONFIG.endMinute-cur>=CONFIG.minBlockMinutes)free.push({start:cur,end:CONFIG.endMinute});
    return free;
  }
  function proposeForDate(date){
    let needed=CONFIG.dailyTargetMinutes-scheduledMinutes(date);
    if(needed<CONFIG.minBlockMinutes)return 0;
    let added=0;
    while(needed>=CONFIG.minBlockMinutes){
      const options=gaps(busyOn(date)).map(g=>({...g,length:g.end-g.start})).filter(g=>g.length>=CONFIG.minBlockMinutes);
      if(!options.length)break;
      options.sort((a,b)=>{
        const aMorning=a.start<12*60?0:1,bMorning=b.start<12*60?0:1;
        if(aMorning!==bMorning)return aMorning-bMorning;
        // Prefer a continuous slot if one meets the entire outstanding target.
        const aFull=a.length>=needed?0:1,bFull=b.length>=needed?0:1;
        return aFull-bFull||a.start-b.start;
      });
      const best=options[0];
      let length=Math.min(best.length,needed);
      // If no single block can satisfy the remainder, split into at least 60-minute blocks.
      if(length<CONFIG.minBlockMinutes)break;
      const session={id:state.nextId++,date,start:best.start,end:best.start+length,status:"proposed",recovered:false};
      state.sessions.push(session);added++;needed-=length;
    }
    return added;
  }
  function generate(){
    if(liveMode&&!liveBusy){say("Sync Outlook calendar before generating proposals.");return;}
    let added=0;
    for(const d of workdays)added+=proposeForDate(dayKey(d));
    store();render();say(added?`Generated ${added} proposals. Approve before reserving.`:"No new eligible gaps found.");
  }
  function say(msg){$("notice").textContent=msg;}
  function setStatus(id,status){
    const s=state.sessions.find(s=>s.id===id);
    if(!s||!["proposed","approved"].includes(s.status))return;
    if(status==="approved"&&s.status!=="proposed")return;
    if((status==="completed"||status==="missed")&&s.status!=="approved")return;
    s.status=status;store();render();
  }
  function simulate(){
    if(liveMode){say("Conflict simulation is disabled for live Outlook data. Refresh Outlook to reassess availability.");return;}
    const approved=state.sessions.find(s=>s.status==="approved"&&inHorizon(s.date));
    if(!approved){say("Approve at least one proposed session first.");return;}
    const collision={date:approved.date,start:approved.start,end:approved.end,subject:"New priority meeting"};
    state.extraBusy.push(collision);
    const collisions=state.sessions.filter(s=>s.date===collision.date&&s.status==="approved"&&overlap(s,collision));
    collisions.forEach(s=>{s.status="conflicted";s.displacedAt=new Date().toISOString();});
    let replaced=0;
    if(CONFIG.conflictRecovery){
      replaced=proposeForDate(collision.date);
      // Mark newly generated candidates as recovery proposals.
      state.sessions.filter(s=>s.date===collision.date&&s.status==="proposed"&&!s.recovered).forEach(s=>{s.recovered=true;});
    }
    selected=collision.date;store();render();
    say(`Meeting displaced ${collisions.length} approved block(s). ${replaced} replacement proposal(s) require approval.`);
  }
  function render(){
    $("modeBadge").textContent=liveMode?"Outlook · read-only":"Demo mode · No Outlook connection";
    $("simulate").disabled=liveMode;
    $("target").textContent=(workdays.length*2)+"h";
    const sum=status=>state.sessions.filter(s=>s.status===status).reduce((n,s)=>n+s.end-s.start,0);
    $("reserved").textContent=(sum("approved")/60).toFixed(1)+"h";
    $("pending").textContent=(sum("proposed")/60).toFixed(1)+"h";
    $("completed").textContent=state.sessions.filter(s=>s.status==="completed").length;
    $("approveAll").disabled=!state.sessions.some(s=>s.status==="proposed");
    $("weeks").replaceChildren();
    for(let i=0;i<CONFIG.horizonDays;i+=7){
      const week=document.createElement("div");week.className="week";
      for(const d of weekdays(dates.slice(i,i+7))){
        const key=dayKey(d), ss=sessionsOn(key);
        const minutes=ss.filter(active).reduce((n,s)=>n+s.end-s.start,0);
        const btn=document.createElement("button");btn.type="button";
        btn.className="day"+(key===selected?" selected":"")+(minutes<120&&ss.length?" unavailable":"");
        const date=document.createElement("div");date.className="date";date.textContent=d.toLocaleDateString(undefined,{weekday:"short",month:"short",day:"numeric"});
        const hours=document.createElement("div");hours.className="hours";hours.textContent=(minutes/60).toFixed(minutes%60?1:0)+"/2h";
        const sub=document.createElement("div");sub.className="sub";sub.textContent=ss.some(s=>s.status==="conflicted")?"Conflict detected":minutes<120?"Needs focus time":"Focus target covered";
        const chips=document.createElement("div");chips.className="chips";
        for(const s of ss.filter(s=>["proposed","approved","conflicted","completed"].includes(s.status))){
          const chip=document.createElement("span");chip.className="chip "+(s.status==="conflicted"?"conflict":s.status==="proposed"?"pending":"booked");chip.title=s.status;chips.append(chip);
        }
        btn.append(date,hours,sub,chips);btn.addEventListener("click",()=>{selected=key;render();});week.append(btn);
      }
      $("weeks").append(week);
    }
    const day=dateFromKey(selected);
    $("detailTitle").textContent=day.toLocaleDateString(undefined,{weekday:"long",month:"long",day:"numeric"});
    const details=$("detailBody");details.replaceChildren();
    const ss=sessionsOn(selected).sort((a,b)=>a.start-b.start);
    const meetings=[...demoMeetings(selected),...(liveMode?[]:state.extraBusy.filter(e=>e.date===selected))].sort((a,b)=>a.start-b.start);
    if(!ss.length){const p=document.createElement("p");p.className="hint";p.textContent="No focus sessions yet. Use Generate plan to find available blocks.";details.append(p);}
    for(const s of ss){
      const row=document.createElement("div");row.className="row";
      const title=document.createElement("div");title.className="title";title.textContent=`${displayTime(s.start)} – ${displayTime(s.end)}`;
      const tag=document.createElement("span");tag.className="tag"+(["approved","completed"].includes(s.status)?" good":s.status==="conflicted"?" warn":"");tag.textContent=s.status+(s.recovered?" · recovery":"");title.append(tag);
      const meta=document.createElement("small");meta.textContent=`${s.end-s.start} minutes of focus`;title.append(meta);row.append(title);
      const controls=document.createElement("div");controls.className="feedback";
      const add=(label,status)=>{const b=document.createElement("button");b.className="btn";b.textContent=label;b.addEventListener("click",()=>setStatus(s.id,status));controls.append(b);};
      if(s.status==="proposed"){add("Approve","approved");add("Reject","rejected");}
      if(s.status==="approved"){add("Completed","completed");add("Missed","missed");}
      row.append(controls);details.append(row);
    }
    const p=document.createElement("p");p.className="hint";p.textContent="Simulated busy: "+(meetings.length?meetings.map(m=>`${displayTime(m.start)}–${displayTime(m.end)} ${m.subject}`).join("; "):"none");details.append(p);
    const reported=state.sessions.filter(s=>["completed","missed"].includes(s.status));
    const complete=reported.filter(s=>s.status==="completed").length;
    const displaced=state.sessions.filter(s=>s.status==="conflicted").length;
    $("insights").textContent=reported.length?`${complete} of ${reported.length} reported sessions completed (${Math.round(100*complete/reported.length)}%). ${displaced} disrupted block(s). Completion is self-reported.`:`No sessions reported yet. ${displaced} blocks displaced. Approve a session, then mark it completed or missed.`;
  }
  $("connectOutlook").addEventListener("click",async()=>{
    const button=$("connectOutlook");button.disabled=true;
    $("outlookStatus").textContent="Connecting and reading calendar…";
    try {
      const tenantId=$("tenantId").value.trim(),clientId=$("clientId").value.trim();
      if(location.protocol==="file:")throw Error("Open on http://localhost:8000/ instead of double-clicking index.html.");
      await FocusOutlook.setup(tenantId,clientId);
      const start=new Date(dateFromKey(anchor));const end=plusDays(start,CONFIG.horizonDays);
      const result=await FocusOutlook.sync(start,end);
      const byDate=FocusOutlook.toBusyByDate(result.events,start,end);
      // Distinct plan for the live calendar: do not carry simulated approvals across.
      liveMode=true;liveBusy=byDate;state=createState();selected=anchor;localStorage.removeItem(KEY);
      render();$("outlookStatus").textContent=`Connected: ${result.account}. Read ${result.count} Outlook events through ${dayKey(plusDays(end,-1))}. No calendar writes.`;
      say("Outlook synced. Generate plan to propose available slots.");
    }catch(e){$("outlookStatus").textContent="Connection failed: "+(e.message||e);}
    finally{button.disabled=false;}
  });
  $("demoMode").addEventListener("click",()=>{
    liveMode=false;liveBusy=null;state=createState();selected=anchor;store();render();
    $("outlookStatus").textContent="Demo mode. Live calendar data cleared from memory.";
  });
  $("generate").addEventListener("click",generate);
  $("approveAll").addEventListener("click",()=>{
    let count=0;for(const s of state.sessions)if(s.status==="proposed"){s.status="approved";count++;}
    store();render();say(`Approved ${count} demo blocks; no Outlook changes made.`);
  });
  $("simulate").addEventListener("click",simulate);
  $("reset").addEventListener("click",()=>{if(!window.confirm("Reset all demo sessions and feedback?"))return;state=createState();selected=anchor;store();render();say("Demo reset.");});
  $("export").addEventListener("click",()=>{
    const blob=new Blob([JSON.stringify({config:CONFIG,source:liveMode?"outlook-readonly":"demo",...state},null,2)],{type:"application/json"});
    const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="focus-planner-demo.json";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  });
  render();
})();