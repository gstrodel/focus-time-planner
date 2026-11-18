"use strict";
(() => {
  const P=window.FocusPlanner, KEY="focus-planner-v0.3", $=id=>document.getElementById(id);
  let plan=null,selected=null;
  const time=m=>`${Math.floor(m/60)%12||12}:${String(m%60).padStart(2,"0")} ${m%1440<720?"AM":"PM"}`;
  const label=date=>new Date(date+"T12:00:00Z").toLocaleDateString("en-US",{timeZone:"UTC",weekday:"short",month:"short",day:"numeric"});
  const hours=n=>(n/60).toFixed(n%60?1:0)+"h";
  function say(message){$("notice").textContent=message;}
  function persist(message="Plan saved in this browser. Outlook is unchanged.") {
    if(!plan)return;
    const previous=plan.savedAt;plan.savedAt=new Date().toISOString();
    try{localStorage.setItem(KEY,JSON.stringify(plan));$("saveStatus").textContent="Saved "+new Date(plan.savedAt).toLocaleTimeString();if(message)say(message);return true;}
    catch(_){plan.savedAt=previous;$("saveStatus").textContent="Browser storage unavailable — download a backup";say("The plan is open, but this browser could not save it. Use Download plan to keep your results.");return false;}
  }
  function approveable(s) {
    const now=P.localParts(new Date().toISOString(),plan.snapshot.timeZone);
    return s.date>now.date || s.date===now.date && s.start>now.minute;
  }
  function update(id,status) {
    const s=plan.sessions.find(s=>s.id===id);if(!s)return;
    if(s.status==="proposed" && ["approved","rejected"].includes(status)) {
      if(status==="approved"&&!approveable(s)){say("That block has already started. Generate new proposals for the remaining time.");return;}
    } else if(!(s.status==="approved" && ["completed","missed"].includes(status)))return;
    s.status=status;persist();render();
  }
  function render() {
    const loaded=!!plan;
    for(const id of ["generate","save","export","approveAll","protectCommute","reset"])$(id).disabled=!loaded;
    $("weeks").replaceChildren();$("detailBody").replaceChildren();
    if(!loaded){$("modeBadge").textContent="Calendar extract · Offline prototype";$("sourceStatus").textContent="Import a six-week calendar extract or a saved plan. No sign-in needed.";$("weeks").textContent="Your six-week plan will appear here after import.";return;}
    const snapshot=plan.snapshot,zoneLabel=snapshot.timeZone==="America/New_York"?"Eastern Time":snapshot.timeZone,days=P.workdays(snapshot),busy=P.intervals(snapshot),today=P.localParts(new Date().toISOString(),snapshot.timeZone).date;
    $("modeBadge").textContent=snapshot.source==="demo"?"Demo · Synthetic availability":"Outlook extract · No live connection";
    $("sourceStatus").textContent=`${label(snapshot.windowStart)} – ${label(P.addDays(snapshot.windowEnd,-1))}, ${snapshot.windowEnd.slice(0,4)}. Captured ${new Date(snapshot.capturedAt).toLocaleString("en-US",{timeZone:snapshot.timeZone})} (${snapshot.timeZone}). ${snapshot.events.length} calendar intervals. Availability is a snapshot; changes after capture are not reflected.`;
    $("protectCommute").checked=plan.config.protectCommute;
    $("target").textContent=hours(days.length*120);
    const sum=status=>plan.sessions.filter(s=>s.status===status).reduce((n,s)=>n+s.end-s.start,0);
    $("reserved").textContent=hours(sum("approved"));$("pending").textContent=hours(sum("proposed"));$("existing").textContent=hours(sum("existing"));
    $("approveAll").disabled=!plan.sessions.some(s=>s.status==="proposed"&&approveable(s));
    const covered=days.filter(d=>P.scheduled(plan,d)>=120).length;
    const shortfall=days.reduce((n,d)=>n+Math.max(0,120-P.scheduled(plan,d)),0);
    $("coverageSummary").textContent=`${covered} of ${days.length} weekdays covered · ${hours(shortfall)} below target · ${zoneLabel}`;
    let group=null;
    for(const date of days) {
      if(!group || new Date(date+"T12:00:00Z").getUTCDay()===1){group=document.createElement("div");group.className="week";$("weeks").append(group);}
      const minutes=P.scheduled(plan,date),sessions=plan.sessions.filter(s=>s.date===date);
      const btn=document.createElement("button");btn.type="button";btn.dataset.date=date;
      btn.className="day"+(date===selected?" selected":"")+(minutes<120?" unavailable":"")+(date<today?" past":"");
      btn.setAttribute("aria-label",`${label(date)}: ${hours(minutes)} focus, ${hours(Math.max(0,120-minutes))} shortfall`);
      const name=document.createElement("div");name.className="date";name.textContent=label(date);
      const total=document.createElement("div");total.className="hours";total.textContent=hours(minutes)+" / 2h";
      const sub=document.createElement("div");sub.className="sub";sub.textContent=minutes>=120?"Target covered":hours(120-minutes)+" below target";
      const chips=document.createElement("div");chips.className="chips";
      for(const s of sessions.filter(s=>P.ACTIVE.has(s.status))){const chip=document.createElement("span");chip.className="chip "+(s.status==="proposed"?"pending":s.status==="existing"?"calendar":"booked");chip.title=s.status;chips.append(chip);}
      btn.append(name,total,sub,chips);btn.addEventListener("click",()=>{selected=date;render();});group.append(btn);
    }
    $("detailTitle").textContent=label(selected)+" · "+zoneLabel;
    const details=$("detailBody"),sessions=plan.sessions.filter(s=>s.date===selected).sort((a,b)=>a.start-b.start);
    if(!sessions.length){const p=document.createElement("p");p.textContent="No focus blocks yet. Generate proposals to find available time.";details.append(p);}
    for(const s of sessions){
      const row=document.createElement("div");row.className="row";row.dataset.sessionId=s.id;
      const title=document.createElement("div");title.className="title";title.textContent=`${time(s.start)} – ${time(s.end)}`;
      const tag=document.createElement("span");tag.className="tag"+(["approved","completed","existing"].includes(s.status)?" good":"");tag.textContent=s.status==="existing"?"On your calendar":s.status==="approved"?"Approved locally":s.status;title.append(tag);
      const meta=document.createElement("small");meta.textContent=`${s.end-s.start} minutes`+(s.status==="proposed"&&!approveable(s)?" · Already started":"");title.append(meta);
      const controls=document.createElement("div");controls.className="feedback";
      const action=(text,status,disabled=false)=>{const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=text;b.disabled=disabled;b.addEventListener("click",()=>update(s.id,status));controls.append(b);};
      if(s.status==="proposed"){action("Approve","approved",!approveable(s));action("Reject","rejected");}
      if(s.status==="approved"){action("Completed","completed");action("Missed","missed");}
      row.append(title,controls);details.append(row);
    }
    const heading=document.createElement("h4");heading.textContent="Calendar availability";details.append(heading);
    for(const event of (busy[selected]||[]).sort((a,b)=>a.start-b.start)) {
      const row=document.createElement("p");row.className="hint";row.textContent=`${event.start===0&&event.end===1440?"All day":time(event.start)+" – "+time(event.end)} · ${event.kind==="existing-focus"?"Existing focus":event.status==="oof"?"Out of office":event.status==="tentative"?"Tentative (held as busy)":"Busy"}`;details.append(row);
    }
    if(!(busy[selected]||[]).length){const p=document.createElement("p");p.className="hint";p.textContent="No blocking calendar events in this extract for this day.";details.append(p);}
    const reported=plan.sessions.filter(s=>["completed","missed"].includes(s.status));
    $("insights").textContent=reported.length?`${reported.filter(s=>s.status==="completed").length} of ${reported.length} reported sessions completed. Feedback is self-reported and saved with the plan.`:"Mark approved sessions completed or missed to start a feedback history.";
  }
  function open(raw, autoGenerate=false) {
    const next=raw.schema==="focus-planner-plan"?P.validatePlan(raw):P.createPlan(raw);
    if(autoGenerate)P.generate(next);
    plan=next;selected=P.workdays(plan.snapshot)[0];persist("Calendar extract loaded. Your plan is saved in this browser.");render();
  }
  $("import").addEventListener("change",async event=>{
    const file=event.target.files[0];if(!file)return;
    try{
      if(file.size>5*1024*1024)throw Error("Choose a JSON file smaller than 5 MB.");
      const raw=JSON.parse(await file.text());
      // Validate before asking to replace anything or touching saved state.
      if(raw.schema==="focus-planner-plan")P.validatePlan(raw);else P.validateSnapshot(raw);
      if(plan&&!window.confirm("Replace this browser's current plan? Download a backup first if you want to keep it."))return;
      open(raw,raw.schema!=="focus-planner-plan");
    }catch(error){say("Import failed: "+error.message);}
    finally{event.target.value="";}
  });
  $("generate").addEventListener("click",()=>{const n=P.generate(plan);persist(n?`Added ${n} proposed blocks. Approvals save locally; Outlook is unchanged.`:"No additional blocks fit the target and minimum length. Existing decisions were kept.");render();});
  $("approveAll").addEventListener("click",()=>{let n=0;for(const s of plan.sessions)if(s.status==="proposed"&&approveable(s)){s.status="approved";n++;}persist(`Approved ${n} blocks locally. They are not reserved in Outlook.`);render();});
  $("save").addEventListener("click",()=>persist());
  $("export").addEventListener("click",()=>{
    persist();const url=URL.createObjectURL(new Blob([JSON.stringify(plan,null,2)],{type:"application/json"}));
    const a=document.createElement("a");a.href=url;a.download=`focus-plan-${plan.snapshot.windowStart}.json`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  });
  $("demoMode").addEventListener("click",()=>{if(plan&&!window.confirm("Replace the current plan with synthetic demo data?"))return;open(P.demo(),true);say("Demo loaded. This is synthetic data.");});
  $("reset").addEventListener("click",()=>{if(!plan||!window.confirm("Clear approvals and feedback and regenerate from the same calendar extract?"))return;const next=P.createPlan(plan.snapshot,new Date().toISOString(),plan.config);P.generate(next);plan=next;persist();render();});
  $("protectCommute").addEventListener("change",event=>{
    const protect=event.target.checked;
    if(!window.confirm("Rebuild the plan with this commuting preference? Approvals and feedback will be cleared.")){event.target.checked=plan.config.protectCommute;return;}
    plan=P.createPlan(plan.snapshot,new Date().toISOString(),{protectCommute:protect});P.generate(plan);persist();render();
  });
  let saved=null;
  try{saved=localStorage.getItem(KEY);}catch(_){say("Browser storage is unavailable. Download your plan to keep your results.");}
  if(saved){try{plan=P.validatePlan(JSON.parse(saved));selected=P.workdays(plan.snapshot)[0];$("saveStatus").textContent="Restored saved plan";}catch(_){say("The saved plan could not be restored. It has not been overwritten; import a valid backup to continue.");}}
  else if($("bootstrap")){
    try{open(JSON.parse($("bootstrap").textContent),true);}catch(e){say("Could not load the bundled extract: "+e.message);}
  }
  render();
})();
