"use strict";
/* Local, read-only folder connection. No network requests or calendar writes. */
(function(root,factory){
  const api=factory(typeof module==="object"&&module.exports?require("./planner.js"):root.FocusPlanner);
  if(typeof module==="object"&&module.exports)module.exports=api;else root.FocusLocalCalendar=api;
})(typeof window==="object"?window:globalThis,function(P){
  const FILE="calendar_events.json", MAX_BYTES=5*1024*1024;
  function snapshotFromJSON(raw,modified=new Date().toISOString()){
    if(raw?.schema==="focus-planner-snapshot")return P.validateSnapshot(raw);
    if(raw?.schema)throw Error("Connect a calendar export, not a saved focus plan. Use Import for saved plans.");
    const outer=raw;
    if(raw?.body!==undefined)raw=typeof raw.body==="string"?JSON.parse(raw.body):raw.body;
    if(raw?.["@odata.nextLink"]||raw?.nextLink||outer?.["@odata.nextLink"]||raw?.complete===false||outer?.complete===false)throw Error("This export is incomplete. Include every page of the 42-day calendar view.");
    const events=Array.isArray(raw)?raw:(raw?.value??raw?.events);
    if(!Array.isArray(events)||events.length>10000)throw Error("Expected an event array, a value/events array, or a Focus Planner snapshot.");
    const capturedAt=raw?.capturedAt??outer?.capturedAt??modified;
    const zone=raw?.timeZone??outer?.timeZone??"America/New_York";
    const timeZone=zone==="Eastern Standard Time"?"America/New_York":zone;
    const windowStart=raw?.windowStart??outer?.windowStart??P.localParts(capturedAt,timeZone).date;
    const windowEnd=raw?.windowEnd??outer?.windowEnd??P.addDays(windowStart,42);
    const stamp=(e,key)=>{
      const value=e[key+"WithTimeZone"]??e[key];
      let text=typeof value==="object"?value?.dateTime:value;
      if(typeof text!=="string")throw Error("An event is missing its start or end time.");
      if(!/(?:Z|[+-]\d{2}:\d{2})$/i.test(text)){
        const zone=typeof value==="object"?value?.timeZone:e.timeZone;
        if(!["UTC","Etc/UTC"].includes(zone))throw Error("Event times need offsets or UTC. Export startWithTimeZone and endWithTimeZone from Power Automate.");
        text+="Z";
      }
      return text;
    };
    const truth=v=>v===true||v==="true";
    const ownership=e=>{
      const explicit=e?.isOrganizer??e?.is_organizer;
      if(explicit===true||explicit==="true")return {isOrganizer:true};
      if(explicit===false||explicit==="false")return {isOrganizer:false};
      const response=String(e?.responseStatus?.response??e?.response_status?.response??e?.responseType??"").trim().toLowerCase();
      if(response==="organizer")return {isOrganizer:true};
      if(["accepted","tentativelyaccepted","tentative","declined","notresponded"].includes(response))return {isOrganizer:false};
      return {};
    };
    return P.validateSnapshot({schema:"focus-planner-snapshot",version:1,source:"outlook-extract",complete:true,
      capturedAt,timeZone,windowStart,windowEnd,
      events:events.filter(e=>!truth(e?.isCancelled??e?.is_cancelled)&&!truth(e?.isCanceled)&&
        !["declined"].includes(String(e?.responseStatus?.response??e?.response_status?.response??e?.responseType??"").toLowerCase())&&
        String(e?.showAs??e?.show_as??e?.status??"").toLowerCase()!=="workingelsewhere").map(e=>{
        if(e?.type==="seriesMaster")throw Error("Export the calendar view with recurring occurrences expanded.");
        const value=String(e?.showAs??e?.show_as??e?.status??"unknown").trim().toLowerCase();
        const status=["ooo","outofoffice","out of office"].includes(value)?"oof":value;
        return {id:e?.id,start:stamp(e,"start"),end:stamp(e,"end"),
          status:["busy","tentative","oof","free","unknown"].includes(status)?status:"unknown",
          kind:e?.kind==="existing-focus"||/^focus time$/i.test(String(e?.subject??e?.title??"").trim())?"existing-focus":"busy",
          allDay:truth(e?.isAllDay??e?.is_all_day??e?.allDay),title:e?.subject??e?.title??"",...ownership(e)};
      })});
  }
  function reconcile(previous,snapshot,now=new Date().toISOString()){
    if(!previous||previous.snapshot.source==="demo"){
      const plan=P.createPlan(snapshot,now);P.generate(plan,now);return {plan,conflicts:0};
    }
    if(snapshot.timeZone!==previous.snapshot.timeZone)throw Error("The calendar timezone changed. Import this export manually to start a new plan.");
    if(Date.parse(snapshot.capturedAt)<Date.parse(previous.snapshot.capturedAt))throw Error("This file is older than your loaded calendar. Waiting for a newer export.");
    const plan=P.createPlan(snapshot,previous.createdAt,previous.config),dates=new Set(P.workdays(snapshot)),busy=P.intervals(snapshot);
    plan.nextId=Math.max(plan.nextId,previous.nextId);
    for(const s of plan.sessions)s.id=plan.nextId++;
    let conflicts=0;
    for(const old of previous.sessions){
      if(old.status==="existing"||!dates.has(old.date))continue;
      const s={...old};
      if(P.ACTIVE.has(s.status)){
        if(plan.sessions.some(e=>e.status==="existing"&&e.date===s.date&&e.start<=s.start&&e.end>=s.end))continue;
        if((busy[s.date]||[]).some(e=>P.overlap(e,s))){s.status="conflicted";conflicts++;}
      }
      plan.sessions.push(s);
    }
    P.generate(plan,now);
    return {plan:P.validatePlan(plan),conflicts};
  }
  function remember(action,handle){
    return new Promise((resolve,reject)=>{
      const request=indexedDB.open("focus-planner-calendar",1);
      request.onupgradeneeded=()=>request.result.createObjectStore("handles");
      request.onerror=()=>reject(request.error);
      request.onblocked=()=>reject(Error("Browser storage is busy. The folder can be used for this session."));
      request.onsuccess=()=>{
        const db=request.result;
        try{
          const tx=db.transaction("handles",action==="get"?"readonly":"readwrite"),store=tx.objectStore("handles");
          const op=action==="get"?store.get("calendar-folder"):action==="put"?store.put(handle,"calendar-folder"):store.delete("calendar-folder");
          tx.oncomplete=()=>{db.close();resolve(op.result);};
          tx.onabort=tx.onerror=()=>{db.close();reject(tx.error||Error("Cannot remember this folder."));};
        }catch(error){db.close();reject(error);}
      };
    });
  }
  class Connection{
    constructor({apply,status,storage=remember,pick=options=>window.showDirectoryPicker(options)}){
      this.apply=apply;this.status=status;this.storage=storage;this.pick=pick;
      this.handle=null;this.generation=0;this.busy=false;this.lastText=null;this.lastModified=null;this.loadedAt=null;this.remembered=true;
    }
    async restore(){
      const generation=this.generation;
      try{
        const handle=await this.storage("get");
        if(this.generation!==generation||!handle)return;
        this.handle=handle;await this.refresh();
      }catch(_){if(this.generation===generation)this.status("Choose your calendar folder to enable automatic updates.");}
    }
    async connect(){
      try{
        // The picker must run directly from a user gesture.
        const handle=await this.pick({id:"focus-calendar",mode:"read"});
        this.generation++;this.handle=handle;this.lastText=null;this.lastModified=null;this.loadedAt=null;
        try{await this.storage("put",handle);this.remembered=true;}catch(_){this.remembered=false;}
        await this.refresh();
      }catch(e){if(e.name!=="AbortError")this.status("Could not connect: "+e.message+" Use Import below if folder access is blocked.");}
    }
    async refresh(allowPermission=false){
      if(!this.handle||this.busy)return;
      this.busy=true;
      const generation=this.generation,handle=this.handle;
      try{
        let permission=await handle.queryPermission({mode:"read"});
        if(permission!=="granted"&&allowPermission)permission=await handle.requestPermission({mode:"read"});
        if(permission!=="granted")throw Error("Folder access needs approval. Click Refresh now to allow reading again.");
        // Resolve by filename each time: OneDrive may replace rather than edit the file.
        const file=await (await handle.getFileHandle(FILE)).getFile();
        if(file.size>MAX_BYTES)throw Error("Calendar JSON is larger than 5 MB.");
        const text=await file.text();
        if(this.generation!==generation)return;
        if(text!==this.lastText||file.lastModified!==this.lastModified){
          const snapshot=snapshotFromJSON(JSON.parse(text.replace(/^\uFEFF/,"")),new Date(file.lastModified).toISOString());
          await this.apply(snapshot);
          if(this.generation!==generation)return;
          this.lastText=text;this.lastModified=file.lastModified;this.loadedAt=new Date();
        }
        const stale=Date.now()-file.lastModified>75*60*1000;
        this.status(`${handle.name} / ${FILE} · File updated ${new Date(file.lastModified).toLocaleString()}. Last checked ${new Date().toLocaleTimeString()}. Checking every minute while open.${stale?" Export is over 75 minutes old; check Power Automate and OneDrive.":""}${this.remembered?"":" Choose the folder again next time; this browser could not remember it."}`);
      }catch(e){
        if(this.generation===generation)this.status("Calendar refresh paused: "+(e.name==="NotFoundError"?"calendar_events.json was not found. Check OneDrive sync and the selected folder.":e.message)+" Your last valid plan is kept; retries continue every minute.");
      }finally{this.busy=false;}
    }
    async disconnect(){
      this.generation++;this.handle=null;this.lastText=null;this.lastModified=null;this.loadedAt=null;
      try{await this.storage("delete");this.status("Automatic updates disconnected. Your loaded plan is kept.");}
      catch(_){this.status("Disconnected for this session. The browser could not forget the saved folder; it may reconnect when reopened.");}
    }
  }
  return {FILE,snapshotFromJSON,reconcile,Connection};
});

