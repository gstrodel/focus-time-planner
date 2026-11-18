"use strict";
(() => {
  const P=window.FocusPlanner, KEY="focus-planner-v0.4", $=id=>document.getElementById(id);
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
    } else if(!(s.status==="approved" && status==="rejected" && approveable(s)))return;
    s.status=status;persist();render();
  }
  function render() {
    const loaded=!!plan;
    for(const id of ["generate","save","export","approveAll","protectCommute","reset"])$(id).disabled=!loaded;
    $("weeks").replaceChildren();$("detailBody").replaceChildren();
    if(!loaded){$("modeBadge").textContent="Calendar extract · Offline prototype";$("sourceStatus").textContent="Import a six-week calendar extract or a saved plan. No sign-in needed.";$("weeks").textContent="Your six-week plan will appear here after import.";return;}
    const snapshot=plan.snapshot,zoneLabel=snapshot.timeZone==="America/New_York"?"Eastern Time":snapshot.timeZone,today=P.localParts(new Date().toISOString(),snapshot.timeZone).date,days=P.workdays(snapshot).filter(d=>d>=today),busy=P.intervals(snapshot);
    if(!days.includes(selected))selected=days[0]||null;
    $("modeBadge").textContent=snapshot.source==="demo"?"Demo · Synthetic availability":"Outlook extract · No live connection";
    $("sourceStatus").textContent=`${label(snapshot.windowStart)} – ${label(P.addDays(snapshot.windowEnd,-1))}, ${snapshot.windowEnd.slice(0,4)}. Captured ${new Date(snapshot.capturedAt).toLocaleString("en-US",{timeZone:snapshot.timeZone})} (${snapshot.timeZone}). ${snapshot.events.length} calendar intervals. Availability is a snapshot; changes after capture are not reflected.`;
    $("protectCommute").checked=plan.config.protectCommute;
    $("target").textContent=hours(days.length*120);
    const sum=status=>plan.sessions.filter(s=>s.date>=today&&s.status===status).reduce((n,s)=>n+s.end-s.start,0);
    $("reserved").textContent=hours(sum("approved"));$("pending").textContent=hours(sum("proposed"));$("existing").textContent=hours(sum("existing"));
    $("approveAll").disabled=!plan.sessions.some(s=>s.status==="proposed"&&approveable(s));
    const covered=days.filter(d=>P.hasContinuousFocus(plan,d)).length;
    const shortfall=days.reduce((n,d)=>n+Math.max(0,120-P.scheduled(plan,d)),0);
    $("coverageSummary").textContent=`${covered} of ${days.length} upcoming weekdays have 2h continuous focus planned · ${hours(shortfall)} below the daily minimum in total · ${zoneLabel}`;
    let group=null;
    for(const date of days) {
      if(!group || new Date(date+"T12:00:00Z").getUTCDay()===1){group=document.createElement("div");group.className="week";$("weeks").append(group);}
      const continuous=P.hasContinuousFocus(plan,date),minutes=P.scheduled(plan,date),sessions=plan.sessions.filter(s=>s.date===date);
      const btn=document.createElement("button");btn.type="button";btn.dataset.date=date;
      btn.className="day"+(date===selected?" selected":"")+(!continuous?" unavailable":"")+(date<today?" past":"");
      btn.setAttribute("aria-label",`${label(date)}: ${hours(minutes)} focus, ${hours(Math.max(0,120-minutes))} shortfall`);
      const name=document.createElement("div");name.className="date";name.textContent=label(date);
      const total=document.createElement("div");total.className="hours";total.textContent=hours(minutes)+" planned";
      const sub=document.createElement("div");sub.className="sub";sub.textContent=continuous?"2h minimum covered":minutes>=120?"Needs 2h continuous":hours(120-minutes)+" below minimum";
      const chips=document.createElement("div");chips.className="chips";
      for(const s of sessions.filter(s=>P.ACTIVE.has(s.status))){const chip=document.createElement("span");chip.className="chip "+(s.status==="proposed"?"pending":s.status==="existing"?"calendar":"booked");chip.title=s.status;chips.append(chip);}
      btn.append(name,total,sub,chips);btn.addEventListener("click",()=>{selected=date;render();});group.append(btn);
    }
    if(!selected){$("detailTitle").textContent="Calendar extract ended";$("detailBody").textContent="Import a new six-week extract to plan upcoming focus time.";return;}
    $("detailTitle").textContent=label(selected)+" · "+zoneLabel;
    const details=$("detailBody"),sessions=plan.sessions.filter(s=>s.date===selected&&P.ACTIVE.has(s.status)).sort((a,b)=>a.start-b.start);
    if(!sessions.length){const p=document.createElement("p");p.textContent="No focus blocks yet. Generate proposals to find available time.";details.append(p);}
    for(const s of sessions){
      const row=document.createElement("div");row.className="row";row.dataset.sessionId=s.id;
      const title=document.createElement("div");title.className="title";title.textContent=`${time(s.start)} – ${time(s.end)}`;
      const tag=document.createElement("span");tag.className="tag"+(["approved","existing"].includes(s.status)?" good":"");tag.textContent=s.status==="existing"?"On your calendar":s.status==="approved"?"Approved locally":s.status;title.append(tag);
      const meta=document.createElement("small");meta.textContent=`${s.end-s.start} minutes`+(s.status==="proposed"&&!approveable(s)?" · Already started":"");title.append(meta);
      const controls=document.createElement("div");controls.className="block-actions";
      const action=(text,status,disabled=false)=>{const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=text;b.disabled=disabled;b.addEventListener("click",()=>update(s.id,status));controls.append(b);};
      if(s.status==="proposed"){action("Approve","approved",!approveable(s));action("Reject","rejected");}
      if(s.status==="approved"&&approveable(s))action("Remove from plan","rejected");
      if(["proposed","approved"].includes(s.status)&&approveable(s)){
        const extend=document.createElement("button");extend.type="button";extend.className="btn";extend.textContent="Extend 30 min";
        extend.disabled=!P.available(plan,s.date,undefined,s.id).some(g=>g.start<=s.start&&g.end>=s.end+30);
        extend.addEventListener("click",()=>{try{P.extendBlock(plan,s.id);persist("Extended by 30 minutes. Review and approve the updated block locally.");render();}catch(e){say(e.message);}});controls.append(extend);
      }
      row.append(title,controls);details.append(row);
    }
    renderBookingForm(details);
    const heading=document.createElement("h4");heading.textContent="Calendar availability";details.append(heading);
    for(const event of (busy[selected]||[]).sort((a,b)=>a.start-b.start)) {
      const row=document.createElement("details");row.className="meeting";
      const summary=document.createElement("summary");
      const eventTime=event.start===0&&event.end===1440?"All day":time(event.start)+" – "+time(event.end);
      const status=event.kind==="existing-focus"?"Existing focus":event.status==="oof"?"Out of office":event.status==="tentative"?"Tentative":"Busy";
      summary.textContent=`${eventTime} · ${event.title||status}`;row.append(summary);
      const organizer=document.createElement("p");organizer.className="hint";
      const person=event.organizer||{};
      organizer.textContent="Organizer: "+(person.name?person.name+(person.email?" <"+person.email+">":""):person.email||"Not included in this extract");row.append(organizer);
      const description=document.createElement("p");description.className="hint meeting-description";
      description.textContent=event.description||"No description included in this extract.";row.append(description);
      const availability=document.createElement("p");availability.className="hint";availability.textContent="Availability: "+status;row.append(availability);details.append(row);
    }
    if(!(busy[selected]||[]).length){const p=document.createElement("p");p.className="hint";p.textContent="No blocking calendar events in this extract for this day.";details.append(p);}


  }
  function renderBookingForm(parent) {
    const heading=document.createElement("h4");heading.textContent="Add more focus time";parent.append(heading);
    const windows=P.available(plan,selected);
    if(!windows.length){const p=document.createElement("p");p.textContent="No additional future window of 60 minutes or more is available on this day.";parent.append(p);return;}
    const form=document.createElement("form");form.className="booking-form";
    const note=document.createElement("p");note.className="hint";note.textContent="Two hours is a minimum. Add a longer block or another block, including after the minimum is covered. Times are in "+plan.snapshot.timeZone+".";form.append(note);
    const labelFor=(text,id)=>{const l=document.createElement("label");l.htmlFor=id;l.textContent=text;form.append(l);};
    labelFor("Available window","extraWindow");const choice=document.createElement("select");choice.id="extraWindow";
    windows.forEach((g,i)=>{const o=document.createElement("option");o.value=i;o.textContent=time(g.start)+" – "+time(g.end);choice.append(o);});form.append(choice);
    labelFor("Start time","extraStart");const start=document.createElement("select");start.id="extraStart";form.append(start);
    labelFor("Duration","extraDuration");const duration=document.createElement("select");duration.id="extraDuration";form.append(duration);
    const addOption=(select,value,text)=>{const o=document.createElement("option");o.value=value;o.textContent=text;select.append(o);};
    const fillDurations=()=>{duration.replaceChildren();const limit=windows[Number(choice.value)].end-Number(start.value);for(let n=60;n<=limit;n+=15)addOption(duration,n,hours(n));if(limit%15)addOption(duration,limit,hours(limit)+" (full remaining window)");duration.value=String(Math.min(120,Math.floor(limit/15)*15));};
    const fillStarts=()=>{start.replaceChildren();const g=windows[Number(choice.value)];for(let n=g.start;n<=g.end-60;n+=15)addOption(start,n,time(n));fillDurations();};
    choice.addEventListener("change",fillStarts);start.addEventListener("change",fillDurations);fillStarts();
    const submit=document.createElement("button");submit.type="submit";submit.className="btn primary";submit.textContent="Add focus proposal";form.append(submit);
    form.addEventListener("submit",event=>{event.preventDefault();try{const a=Number(start.value);P.addBlock(plan,selected,a,a+Number(duration.value));persist("Added a focus proposal. Review and approve it locally; Outlook is unchanged.");render();}catch(e){say(e.message);}});
    parent.append(form);
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
  $("reset").addEventListener("click",()=>{if(!plan||!window.confirm("Clear local planning decisions and regenerate from the same calendar extract?"))return;const next=P.createPlan(plan.snapshot,new Date().toISOString(),plan.config);P.generate(next);plan=next;persist();render();});
  $("protectCommute").addEventListener("change",event=>{
    const protect=event.target.checked;
    if(!window.confirm("Rebuild the plan with this commuting preference? Local planning decisions will be cleared.")){event.target.checked=plan.config.protectCommute;return;}
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
