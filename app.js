"use strict";
(() => {
  const P=window.FocusPlanner, I=window.FocusICalendar, KEY="focus-planner-v0.4", $=id=>document.getElementById(id);
  let plan=null,selected=null,calendarConnection=null,exportingCalendar=false;
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
    } else if(!(s.status==="conflicted"&&status==="rejected")&&!(s.status==="approved" && status==="rejected" && approveable(s)))return;
    s.status=status;persist();render();
  }
  function render() {
    const loaded=!!plan;
    renderCalendarExport();
    for(const id of ["generate","save","export","approveAll","protectCommute","reset"])$(id).disabled=!loaded;
    $("weeks").replaceChildren();$("detailBody").replaceChildren();
    if(!loaded){$("modeBadge").textContent="Calendar extract · Offline prototype";$("weeks").textContent="Your six-week plan will appear here after import.";$("detailTitle").textContent="Select a day";$("detailBody").textContent="Import your calendar, then select a day to see its calendar and focus time.";return;}
    const snapshot=plan.snapshot,zoneLabel=snapshot.timeZone==="America/New_York"?"Eastern Time":snapshot.timeZone,today=P.localParts(new Date().toISOString(),snapshot.timeZone).date,days=P.workdays(snapshot).filter(d=>d>=today),busy=P.intervals(snapshot);
    if(!days.includes(selected))selected=days[0]||null;
    $("modeBadge").textContent=snapshot.source==="demo"?"Demo · Synthetic availability":calendarConnection?.handle?"Outlook · Synced local file":"Outlook extract · No live connection";
    $("protectCommute").checked=plan.config.protectCommute;
    const sum=status=>plan.sessions.filter(s=>s.date>=today&&s.status===status).reduce((n,s)=>n+s.end-s.start,0);
    $("reserved").textContent=hours(sum("approved"));$("pending").textContent=hours(sum("proposed"));$("existing").textContent=hours(sum("existing"));
    $("approveAll").disabled=!plan.sessions.some(s=>s.status==="proposed"&&approveable(s));
    const loads=Object.fromEntries(days.map(d=>[d,P.meetingLoad(snapshot,d,busy)]));
    const workingDays=days.filter(d=>loads[d].workingMinutes>0);
    const covered=workingDays.filter(d=>P.hasContinuousFocus(plan,d)).length;
    $("target").textContent=hours(workingDays.length*120);
    $("meetingGoal").textContent=workingDays.filter(d=>loads[d].ratio<=P.CONFIG.maxMeetingRatio).length+"/"+workingDays.length;
    $("focusDays").textContent=covered+"/"+workingDays.length;
    const shortfall=workingDays.reduce((n,d)=>n+Math.max(0,120-P.scheduled(plan,d)),0);
    const conflicts=plan.sessions.filter(s=>s.status==="conflicted"&&s.date>=today).length;
    $("coverageSummary").textContent=`${covered} of ${workingDays.length} upcoming working days have 2h continuous focus planned · ${hours(shortfall)} below the daily minimum in total · ${days.length-workingDays.length} PTO/OOO days excluded · ${zoneLabel}`+(conflicts?` · ${conflicts} previous blocks conflict with the updated calendar; select their days to review.`:"");

    // Each weekday card renders a real time-scaled strip, not event-count chips.
    const stripStart=P.CONFIG.meetingStartMinute,stripEnd=P.CONFIG.endMinute;
    const timelineFor=(date,sessions)=>{
      if(!loads[date].workingMinutes)return [{start:stripStart,end:stripEnd,visual:"off"}];
      const base=(busy[date]||[]).filter(e=>(!e.allDay || e.status==="oof") && e.status!=="free");
      const blocks=[
        ...base.map(e=>({...e,visual:e.status==="oof"?"off":e.kind==="existing-focus"?"booked":"meeting"})),
        ...sessions.filter(s=>P.ACTIVE.has(s.status)&&s.status!=="existing").map(s=>({...s,visual:s.status==="proposed"?"proposed":"booked"}))
      ];
      const bounds=new Set([stripStart,stripEnd]);
      for(const b of blocks){bounds.add(Math.max(stripStart,Math.min(stripEnd,b.start)));bounds.add(Math.max(stripStart,Math.min(stripEnd,b.end)));}
      const points=[...bounds].sort((a,b)=>a-b);
      const parts=[];
      for(let i=0;i<points.length-1;i++){
        const a=points[i],b=points[i+1];if(b<=a)continue;
        const active=blocks.filter(e=>e.start<b&&e.end>a);
        const visual=active.some(e=>e.visual==="off")?"off":active.some(e=>e.visual==="meeting")?"meeting":active.some(e=>e.visual==="booked")?"booked":active.some(e=>e.visual==="proposed")?"proposed":"free";
        const previous=parts.at(-1);
        if(previous&&previous.visual===visual)previous.end=b;else parts.push({start:a,end:b,visual});
      }
      return parts;
    };
    let group=null;
    for(const date of days) {
      if(!group || new Date(date+"T12:00:00Z").getUTCDay()===1){group=document.createElement("div");group.className="week";$("weeks").append(group);}
      const sessions=plan.sessions.filter(s=>s.date===date),minutes=P.scheduled(plan,date);
      const load=loads[date],pct=load.ratio===null?null:Math.round(load.ratio*100);
      const btn=document.createElement("button");btn.type="button";btn.dataset.date=date;
      btn.className="day"+(pct===null?" unavailable":"")+(date===selected?" selected":"");
      btn.setAttribute("aria-pressed",date===selected?"true":"false");
      const status=pct===null?"PTO/OOO; no focus planning needed":pct>50?"Above 50% meeting goal":pct===50?"At 50% meeting goal":"Below 50% meeting goal";
      btn.setAttribute("aria-label",`${label(date)}: ${pct===null?"N/A":pct+"%"} in meetings, ${hours(minutes)} of focus; ${status}`);
      const name=document.createElement("div");name.className="date";
      const number=document.createElement("span");number.className="day-number";number.textContent=String(Number(date.slice(-2)));
      const weekday=document.createElement("span");weekday.className="weekday";weekday.textContent=new Date(date+"T12:00:00Z").toLocaleDateString("en-US",{timeZone:"UTC",weekday:"short"}).toUpperCase();
      const month=document.createElement("span");month.className="month";month.textContent=new Date(date+"T12:00:00Z").toLocaleDateString("en-US",{timeZone:"UTC",month:"short"}).toUpperCase();
      name.append(number,month,weekday);
      const badge=document.createElement("span");badge.className="meeting-badge "+(pct===null?"na":pct>50?"over":pct===50?"limit":"good");
      badge.textContent=pct===null?"PTO / OOO":pct+"%";
      const strip=document.createElement("div");strip.className="day-timeline";
      strip.setAttribute("aria-label","Workday distribution, 9:30 AM to 5:30 PM");
      for(const part of timelineFor(date,sessions)){
        const segment=document.createElement("span");segment.className="segment "+part.visual;
        segment.style.flex=`${part.end-part.start} 0 0%`;
        segment.title=`${time(part.start)} – ${time(part.end)}: ${({off:"PTO / Out of office",meeting:"Meeting or unavailable",booked:"Booked focus",proposed:"Proposed focus",free:"Available"})[part.visual]}`;
        strip.append(segment);
      }
      const note=document.createElement("div");note.className="day-note";
      note.textContent=pct===null?"No focus planning needed":sessions.some(s=>s.status==="conflicted")?"Review focus conflict":P.hasContinuousFocus(plan,date)?"✅ Focus Time Secured":minutes?"Focus planned; 2h block needed":"No focus planned";
      if(note.textContent==="No focus planned")note.className+=" no-focus";
      if(note.textContent==="✅ Focus Time Secured")note.className+=" focus-secured";
      btn.append(name,badge);
      if(pct!==null){const caption=document.createElement("div");caption.className="meeting-caption";caption.textContent="in meetings";btn.append(caption);}
      btn.append(strip,note);
      btn.addEventListener("click",()=>{
        selected=date;render();
        $("daySidebar").scrollTop=0;
        // Rendering replaces the day buttons; keep keyboard focus on the selected day.
        $("weeks").querySelector?.(`[data-date="${date}"]`)?.focus({preventScroll:true});
      });
      group.append(btn);
    }
    if(!selected){$("detailTitle").textContent="Calendar extract ended";$("detailBody").textContent="Import a new six-week extract to plan upcoming focus time.";return;}
    $("detailTitle").textContent=label(selected)+" · "+zoneLabel;
    const details=$("detailBody"),sessions=plan.sessions.filter(s=>s.date===selected&&(P.ACTIVE.has(s.status)||s.status==="conflicted")).sort((a,b)=>a.start-b.start);
    const load=loads[selected],summary=document.createElement("p");summary.className="hint";
    summary.textContent=load.workingMinutes?`Meetings: ${hours(load.meetingMinutes)} / ${hours(load.workingMinutes)} (${Math.round(load.ratio*100)}%; target ≤50%). Focus target: ${hours(P.dailyFocusTarget(snapshot,selected,busy))}. Continuous 2h focus: ${P.hasContinuousFocus(plan,selected)?"planned":"not yet secured"}.` : "PTO / Out of office — no focus planning needed. Excluded from meeting and focus coverage targets.";
    details.append(summary);
    renderDayCalendar(details);
    const focusHeading=document.createElement("h4");focusHeading.textContent="Focus blocks";details.append(focusHeading);
    if(!sessions.length && load.workingMinutes){const p=document.createElement("p");p.textContent="No focus blocks yet. Generate proposals to find available time.";details.append(p);}
    for(const s of sessions){
      const row=document.createElement("div");row.className="row";row.dataset.sessionId=s.id;
      const title=document.createElement("div");title.className="title";title.textContent=`${time(s.start)} – ${time(s.end)}`;
      const tag=document.createElement("span");tag.className="tag"+(["approved","existing"].includes(s.status)?" good":s.status==="conflicted"?" warn":"");tag.textContent=s.status==="existing"?"On your calendar":s.status==="approved"?"Approved locally":s.status;title.append(tag);
      const meta=document.createElement("small");meta.textContent=`${s.end-s.start} minutes`+(s.status==="conflicted"?" · Conflicts with the latest calendar. Review available replacements.":"")+(s.status==="proposed"&&!approveable(s)?" · Already started":"");title.append(meta);
      const controls=document.createElement("div");controls.className="block-actions";
      const action=(text,status,disabled=false)=>{const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=text;b.disabled=disabled;b.addEventListener("click",()=>update(s.id,status));controls.append(b);};
      if(s.status==="conflicted")action("Dismiss conflict","rejected");
      if(s.status==="proposed"){action("Approve","approved",!approveable(s));action("Reject","rejected");}
      if(s.status==="approved"&&approveable(s))action("Remove from plan","rejected");
      if(["proposed","approved"].includes(s.status)&&approveable(s)){
        const extend=document.createElement("button");extend.type="button";extend.className="btn";extend.textContent="Extend 30 min";
        extend.disabled=!P.available(plan,s.date,undefined,s.id).some(g=>g.start<=s.start&&g.end>=s.end+30);
        extend.addEventListener("click",()=>{try{P.extendBlock(plan,s.id);persist("Extended by 30 minutes. Review and approve the updated block locally.");render();}catch(e){say(e.message);}});controls.append(extend);
      }
      row.append(title,controls);details.append(row);
    }
    if(load.workingMinutes)renderBookingForm(details);
  }
  function renderPlanCompletion(){
    let title="",detail="";
    if(plan){
      const today=P.localParts(new Date().toISOString(),plan.snapshot.timeZone).date,busy=P.intervals(plan.snapshot);
      const days=P.workdays(plan.snapshot).filter(date=>date>=today&&P.meetingLoad(plan.snapshot,date,busy).workingMinutes>0);
      const upcoming=plan.sessions.filter(s=>s.date>=today);
      const confirmed={...plan,sessions:plan.sessions.filter(s=>s.status==="approved"||s.status==="existing")};
      const complete=days.length>0&&!upcoming.some(s=>s.status==="proposed"||s.status==="conflicted")&&days.every(date=>
        P.scheduled(confirmed,date)>=P.dailyFocusTarget(plan.snapshot,date,busy)&&P.hasContinuousFocus(confirmed,date));
      if(complete){
        title="✅ All applicable focus days are fully planned";
        detail=`All ${days.length} applicable working ${days.length===1?"day meets":"days meet"} the daily focus target, including a continuous 2-hour block. `+
          (upcoming.some(s=>s.status==="approved")?"Export approved blocks and import them into Outlook to reserve the remaining time.":
            plan.snapshot.source==="demo"?"This is a synthetic demo plan.":"Focus time is already reserved in your latest Outlook calendar file. No additional suggestions are needed.");
      }
    }
    $("planCompletion").hidden=!title;
    if($("planCompletionTitle").textContent!==title)$("planCompletionTitle").textContent=title;
    if($("planCompletionDetail").textContent!==detail)$("planCompletionDetail").textContent=detail;
  }
  function renderCalendarExport(){
    renderPlanCompletion();
    const state=I.readiness(plan);
    $("exportIcs").disabled=exportingCalendar||!state.ready;
    $("exportIcs").textContent=exportingCalendar?"Saving .ics…":"Export to Outlook (.ics)";
    $("icsStatus").className="hint"+(state.needsReview?" export-warning":"");
    $("icsStatus").textContent=state.message+(typeof window.showSaveFilePicker==="function"?
      " Choose where to save when prompted.":" This browser uses downloads. To choose a location, enable ‘Ask where to save each file’ in browser settings or open in Edge or Chrome.");
  }
  $("exportIcs").addEventListener("click",async()=>{
    if(exportingCalendar)return;
    let writable=null;
    try{
      // Capture and revalidate the approved plan at the click, before any asynchronous work.
      const file=I.create(plan),blob=new Blob([file.text],{type:"text/calendar;charset=utf-8"});
      exportingCalendar=true;renderCalendarExport();
      if(typeof window.showSaveFilePicker==="function"){
        const handle=await window.showSaveFilePicker({id:"focus-time-ics",suggestedName:file.filename,
          types:[{description:"iCalendar file (.ics)",accept:{"text/calendar":[".ics"]}}],excludeAcceptAllOption:true});
        writable=await handle.createWritable();await writable.write(blob);await writable.close();writable=null;
        say(`Saved ${file.count} focus blocks. In Outlook, use Calendar → Add calendar → Upload from file and choose your main work calendar. Saving did not book them. Reimporting may create duplicates.`);
      }else{
        const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=file.filename;
        try{a.click();}finally{setTimeout(()=>URL.revokeObjectURL(url),1000);}
        say(`Download started for ${file.count} focus blocks. Your browser controls the save location. Import the .ics into your main Outlook work calendar to reserve them. Reimporting may create duplicates.`);
      }
    }catch(error){
      if(writable){try{await writable.abort();}catch(_){}}
      say(error.name==="AbortError"?"Calendar export canceled. No Outlook reservations were made.":"Calendar export failed: "+error.message);
    }finally{exportingCalendar=false;renderCalendarExport();}
  });
  function renderDayCalendar(parent) {
    const heading=document.createElement("h4");heading.textContent="Day calendar";parent.append(heading);
    const note=document.createElement("p");note.className="hint";
    note.textContent="Timed events and planned focus · All-day time off shown above the grid · Times in "+plan.snapshot.timeZone;
    parent.append(note);
    const timeOff=(P.intervals(plan.snapshot)[selected]||[]).filter(e=>e.allDay&&e.status==="oof");
    for(const event of timeOff){
      const banner=document.createElement("p");banner.className="time-off-banner";
      banner.textContent="All day · PTO / Out of office · "+(event.title||"Time off");parent.append(banner);
    }
    const legend=document.createElement("div");legend.className="legend calendar-legend";
    for(const [kind,label] of [["busy","Meeting"],["tentative","Tentative"],["oof","Out of office"],["free","Free invite"],["existing","Calendar focus"],["proposed","Proposed focus"],["approved","Approved focus"]]){
      const item=document.createElement("span"),dot=document.createElement("i");dot.className="dot cal-"+kind;item.append(dot,label);legend.append(item);
    }
    parent.append(legend);
    const events=(P.intervals(plan.snapshot,{includeFree:true})[selected]||[]).filter(e=>!e.allDay).map(e=>({
      ...e,calendarKind:e.status==="free"?"free":e.kind==="existing-focus"?"existing":e.status==="tentative"?"tentative":e.status==="oof"?"oof":"busy",
      calendarTitle:e.title||(e.kind==="existing-focus"?"Calendar focus":e.status==="oof"?"Out of office":"Meeting")
    }));
    for(const s of plan.sessions.filter(s=>s.date===selected&&["proposed","approved"].includes(s.status)))events.push({...s,calendarKind:s.status,calendarTitle:s.status==="proposed"?"Proposed focus":"Approved focus"});
    const start=Math.min(P.CONFIG.startMinute,...events.map(e=>Math.floor(e.start/60)*60));
    const end=Math.max(P.CONFIG.endMinute,...events.map(e=>Math.ceil(e.end/60)*60));
    const pixelsPerMinute=1.6,height=(end-start)*pixelsPerMinute;
    const scroll=document.createElement("div");scroll.className="day-calendar-scroll";scroll.tabIndex=0;scroll.setAttribute("aria-label","Day calendar, scroll for more times");
    const calendar=document.createElement("div");calendar.className="day-calendar";calendar.style.height=height+"px";
    const grid=document.createElement("div");grid.className="calendar-grid";grid.setAttribute("aria-hidden","true");calendar.append(grid);
    const shade=(a,b)=>{if(b<=a)return;const block=document.createElement("div");block.className="calendar-outside-work";block.style.top=(a-start)*pixelsPerMinute+"px";block.style.height=(b-a)*pixelsPerMinute+"px";grid.append(block);};
    shade(start,P.CONFIG.startMinute);shade(P.CONFIG.endMinute,end);
    for(let minute=start;minute<=end;minute+=30){
      const tick=document.createElement("div");tick.className="calendar-tick"+(minute%60===0?" hour":"");tick.style.top=(minute-start)*pixelsPerMinute+"px";
      const label=document.createElement("span");label.textContent=time(minute);tick.append(label);calendar.append(tick);
    }
    const boxes=document.createElement("div");boxes.className="calendar-boxes";boxes.setAttribute("role","list");boxes.setAttribute("aria-label","Timed events and focus for "+label(selected));calendar.append(boxes);
    // Partition each overlapping group into lanes. Events that touch do not overlap.
    const sorted=events.sort((a,b)=>a.start-b.start||b.end-a.end);
    let group=[],groupEnd=-1;
    const draw=()=>{
      const laneEnds=[];
      for(const e of group){let lane=laneEnds.findIndex(end=>end<=e.start);if(lane<0)lane=laneEnds.length;laneEnds[lane]=e.end;e.lane=lane;}
      for(const e of group){
        const status={busy:"Busy",tentative:"Tentative",oof:"Out of office",free:"Free",existing:"On calendar",proposed:"Proposed locally",approved:"Approved locally"}[e.calendarKind];
        const fullLabel=`${time(e.start)} – ${time(e.end)} · ${e.calendarTitle} · ${status}`;
        const box=document.createElement("div");box.className="calendar-event cal-"+e.calendarKind;box.setAttribute("role","listitem");box.setAttribute("aria-label",fullLabel);box.title=fullLabel;box.tabIndex=0;
        box.style.top=(e.start-start)*pixelsPerMinute+"px";box.style.height=(e.end-e.start)*pixelsPerMinute+"px";
        box.style.left=`calc(${e.lane*100/laneEnds.length}% + 3px)`;box.style.width=`calc(${100/laneEnds.length}% - 6px)`;
        const title=document.createElement("strong");title.textContent=e.calendarTitle;
        const timing=document.createElement("span");timing.textContent=time(e.start)+" – "+time(e.end);box.append(title,timing);boxes.append(box);
      }
    };
    for(const e of sorted){if(group.length&&e.start>=groupEnd){draw();group=[];groupEnd=-1;}group.push(e);groupEnd=Math.max(groupEnd,e.end);}
    if(group.length)draw();
    scroll.append(calendar);parent.append(scroll);
    if(!events.length){const empty=document.createElement("p");empty.className="hint";empty.textContent="No timed events or focus blocks to show for this day.";parent.append(empty);}
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
      let raw=JSON.parse((await file.text()).replace(/^\uFEFF/,""));
      if(window.FocusLocalCalendar&&raw?.schema!=="focus-planner-plan")raw=window.FocusLocalCalendar.snapshotFromJSON(raw,new Date(file.lastModified||Date.now()).toISOString());
      // Validate before asking to replace anything or touching saved state.
      if(raw.schema==="focus-planner-plan")P.validatePlan(raw);else P.validateSnapshot(raw);
      if(plan&&!window.confirm("Replace this browser's current plan? Download a backup first if you want to keep it."))return;
      if(calendarConnection)await calendarConnection.disconnect();
      open(raw,raw.schema!=="focus-planner-plan");
    }catch(error){say("Import failed: "+error.message);}
    finally{event.target.value="";}
  });
  $("preferencesButton").addEventListener("click",()=>$("preferencesDialog").showModal());
  $("generate").addEventListener("click",()=>{const n=P.generate(plan);if(persist(n?`Added ${n} proposed blocks. Approvals save locally; Outlook is unchanged.`:"")&&!n)say("");render();});
  $("approveAll").addEventListener("click",()=>{let n=0;for(const s of plan.sessions)if(s.status==="proposed"&&approveable(s)){s.status="approved";n++;}persist(`Approved ${n} blocks locally. They are not reserved in Outlook.`);render();});
  $("save").addEventListener("click",()=>persist());
  $("export").addEventListener("click",()=>{
    persist();const url=URL.createObjectURL(new Blob([JSON.stringify(plan,null,2)],{type:"application/json"}));
    const a=document.createElement("a");a.href=url;a.download=`focus-plan-${plan.snapshot.windowStart}.json`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  });
  $("reset").addEventListener("click",()=>{if(!plan||!window.confirm("Clear local planning decisions and regenerate from the same calendar extract?"))return;const next=P.createPlan(plan.snapshot,new Date().toISOString(),plan.config);P.generate(next);plan=next;persist();render();});
  $("protectCommute").addEventListener("change",event=>{
    const protect=event.target.checked;
    if(!window.confirm("Rebuild the plan with this commuting preference? Local planning decisions will be cleared.")){event.target.checked=plan.config.protectCommute;return;}
    plan=P.createPlan(plan.snapshot,new Date().toISOString(),{protectCommute:protect});P.generate(plan);persist();render();
  });
  let saved=null;
  try{saved=localStorage.getItem(KEY);}catch(_){say("Browser storage is unavailable. Download your plan to keep your results.");}
  if(saved){try{plan=P.validatePlan(JSON.parse(saved));selected=P.workdays(plan.snapshot)[0];if(persist(""))$("saveStatus").textContent="Restored saved plan";}catch(_){say("The saved plan could not be restored. It has not been overwritten; import a valid backup to continue.");}}
  else if($("bootstrap")){
    try{open(JSON.parse($("bootstrap").textContent),true);}catch(e){say("Could not load the bundled extract: "+e.message);}
  }
  render();
  // Keep time-based eligibility current without rebuilding an in-progress booking form.
  window.setInterval(renderCalendarExport,60000);
  window.addEventListener("focus",renderCalendarExport);
  if(window.FocusLocalCalendar){
    const supported=typeof window.showDirectoryPicker==="function";
    const status=(message,{stale=false}={})=>{
      $("calendarFileStatus").textContent=message;
      $("calendarFileStatus").style.color=stale?"#b42318":"";
      $("refreshCalendar").disabled=!calendarConnection?.handle;
      $("disconnectCalendar").disabled=!calendarConnection?.handle;
    };
    if(supported){
      calendarConnection=new window.FocusLocalCalendar.Connection({status,apply:snapshot=>{
        const result=window.FocusLocalCalendar.reconcile(plan,snapshot);
        plan=result.plan;
        persist(result.conflicts?`Calendar updated. ${result.conflicts} previous blocks now conflict and no longer count toward focus time. Review available replacements.`:"Calendar file updated. Valid planning decisions were kept; new blocks need local approval.");
        render();
      }});
      $("connectCalendar").addEventListener("click",()=>calendarConnection.connect());
      $("refreshCalendar").addEventListener("click",()=>calendarConnection.refresh(true));
      $("disconnectCalendar").addEventListener("click",async()=>{await calendarConnection.disconnect();render();});
      window.setInterval(()=>calendarConnection.refresh(),60000);
      window.addEventListener("focus",()=>calendarConnection.refresh());
      document.addEventListener("visibilitychange",()=>{if(!document.hidden)calendarConnection.refresh();});
      calendarConnection.restore();
    }else{
      $("connectCalendar").disabled=true;
      status("Automatic file updates need Edge or Chrome with folder access enabled. Use Import for a one-time load, or reopen in Edge or Chrome.");
    }
  }
})();

