"use strict";
/* Read-only Microsoft Graph adapter. No client secret and no write scopes. */
window.FocusOutlook = (() => {
  const SCOPES = ["Calendars.Read"];
  let client = null;
  let currentAccount = null;
  const isValidGuid = s => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
  async function setup(tenantId,clientId) {
    if (!isValidGuid(tenantId)||!isValidGuid(clientId)) throw Error("Use the Application (client) ID and Directory (tenant) ID from Entra app registration (both GUIDs).");
    if (!window.msal) throw Error("Microsoft authentication library did not load. Check internet/CDN access.");
    client = new msal.PublicClientApplication({
      auth:{clientId,authority:"https://login.microsoftonline.com/"+tenantId,redirectUri:location.origin+location.pathname},
      cache:{cacheLocation:"sessionStorage",storeAuthStateInCookie:false}
    });
    await client.handleRedirectPromise();
    currentAccount=client.getAllAccounts()[0]||null;
  }
  async function getToken() {
    if(!client)throw Error("Configure Microsoft Entra IDs first.");
    if(!currentAccount){
      const login=await client.loginPopup({scopes:SCOPES,prompt:"select_account"});
      currentAccount=login.account;
    }
    try {
      return (await client.acquireTokenSilent({scopes:SCOPES,account:currentAccount})).accessToken;
    } catch(e) {
      const result=await client.acquireTokenPopup({scopes:SCOPES,account:currentAccount});
      return result.accessToken;
    }
  }
  async function request(url,token) {
    const response=await fetch(url,{headers:{Authorization:"Bearer "+token,Prefer:'outlook.timezone="UTC"'}});
    if(!response.ok){
      let reason;try{const data=await response.json();reason=data.error?.message||data.error?.code;}catch(_){}
      throw Error("Microsoft Graph "+response.status+": "+(reason||response.statusText));
    }
    return response.json();
  }
  async function sync(start,end){
    if(location.protocol==="file:") throw Error("Microsoft sign-in needs http://localhost:8000, not file://. See README.");
    const token=await getToken();
    // Bounded calendarView expands recurring event occurrences.
    const params=new URLSearchParams({startDateTime:start.toISOString(),endDateTime:end.toISOString(),$select:"id,start,end,isAllDay,showAs,isCancelled,subject,responseStatus,isOrganizer", $top:"200"});
    let url="https://graph.microsoft.com/v1.0/me/calendarView?"+params.toString();
    const events=[];let pages=0;
    while(url){
      if(++pages>100)throw Error("Unexpected calendar pagination; scan stopped.");
      const data=await request(url,token);
      events.push(...(data.value||[]));
      url=data["@odata.nextLink"]||null;
      if(url&&!url.startsWith("https://graph.microsoft.com/"))throw Error("Unexpected Graph pagination URL.");
    }
    return {events,account:currentAccount?.username||"Connected",count:events.length};
  }
  function toBusyByDate(events,start,end){
    const busy={};
    const utcToDate = v => {
      // Prefer: outlook.timezone=UTC returns 'UTC', generally without a trailing Z.
      const value=v?.dateTime;
      if(!value) return null;
      if(/[zZ]$|[+-]\d\d:\d\d$/.test(value)) return new Date(value);
      if(["UTC","Etc/UTC"].includes(v.timeZone))return new Date(value+"Z");
      // Fail closed for unsupported timezone encodings instead of incorrect placements.
      throw Error("Calendar returned non-UTC event times. Cannot safely schedule.");
    };
    for(const e of events){
      if(e.isCancelled||["free","workingElsewhere"].includes(e.showAs)||e.responseStatus?.response==="declined")continue;
      const a=utcToDate(e.start),b=utcToDate(e.end);
      if(!a||!b||Number.isNaN(+a)||Number.isNaN(+b)||b<=a)continue;
      let day=new Date(a.getFullYear(),a.getMonth(),a.getDate());
      while(day<b&&day<end){
        const next=new Date(day.getFullYear(),day.getMonth(),day.getDate()+1);
        const key=[day.getFullYear(),String(day.getMonth()+1).padStart(2,"0"),String(day.getDate()).padStart(2,"0")].join("-");
        if(next>start) {
          const from=Math.max(+a,+day),to=Math.min(+b,+next);
          if(to>from){
            const begin=new Date(from),finish=new Date(to);
            const startMin=begin.getHours()*60+begin.getMinutes();
            const endMin=finish>=next?1440:finish.getHours()*60+finish.getMinutes()+((finish.getSeconds()||finish.getMilliseconds())?1:0);
            (busy[key] ||= []).push({start:startMin,end:endMin,subject:"Outlook busy"});
          }
        }
        day=next;
      }
    }
    return busy;
  }
  return {setup,sync,toBusyByDate};
})();
