<h1>
  <img src="focus-planner-icon.svg" alt="Focus Time App icon" width="48">
  Focus Time Planner App
</h1>

An offline, browser-based six-week focus planner using an Outlook calendar **extract**. No Microsoft app registration or sign-in is needed for this release. Import your extract, review proposals, approve or reject them, and export an `.ics` file to import into Outlook.
I built this because I was tired of spending 60-80% of my day in meetings and the Microsoft Viva Insights functionality to reserve focus time only worked for 2 weeks into the future. I need the ability to reserve with more control and longer into the future.

## Recreate the Power Automate calendar export

The **current five-step flow** below combines the trigger/action Code view definitions Gordon supplied with the Select mapping updates he made on November 19, 2026. It reads Outlook every 30 minutes and replaces `/SuperlistSync/calendar_events.json` in OneDrive for Business. The OneDrive desktop client syncs that file to Windows for Focus Planner to read. The revised export was checked with the updated importer: event timestamps had explicit timezone offsets, and import, plan generation, and saved-plan validation succeeded. The flow's latest full definition was not re-inspected. The later **Recommended flow for complete six-week imports** section remains a proposed revision, not the deployed configuration.

### Current flow: prerequisites and conventions

Sign in to [Power Automate](https://make.powerautomate.com/) with your work account and connect **Office 365 Outlook** and **OneDrive for Business**. Create the cloud folder `/SuperlistSync` and an existing UTF-8 file named `calendar_events.json` (a new file can initially contain `{}`; this placeholder is not importable). Do not replace an existing working extract with the placeholder. `Update file` needs a file that already exists.

For example, the file syncs on Windows to `C:\Users\<username>\OneDrive - <organization>\SuperlistSync\calendar_events.json`. Other users must use their own OneDrive location. Cloud actions use the OneDrive picker, not a Windows path; keep OneDrive running and the file available locally.

Build the steps in the order below. Add expressions through **Expression / fx** without the leading `@` shown in Code view. Keep the internal action names `Get_calendar_view_of_events_(V3)`, `Select`, and `Compose`, because later expressions reference them. If the designer generates different names, update the expressions accordingly. Each action after the calendar retrieval runs only when its preceding action **Succeeded**.

Calendar and OneDrive file identifiers are specific to the account. Select your own calendar and file in the designer; the opaque IDs supplied from Gordon's flow are deliberately omitted here.

### Current flow, step 1: Recurrence

Create a **Scheduled cloud flow** and configure its Recurrence trigger:

| Setting | Value in the supplied definition |
| --- | --- |
| Frequency | `Minute` |
| Interval | `30` |
| Time zone | `Eastern Standard Time` |

The supplied trigger contains no start-time or concurrency configuration. This is a 30-minute interval; it does not establish exact wall-clock execution at `:00` and `:30`. The trigger's timezone setting does not change the meaning of `utcNow()` in subsequent expressions.

### Current flow, step 2: Get calendar view of events (V3)

Add **Office 365 Outlook → Get calendar view of events (V3)** (`GetEventsCalendarViewV3`):

| Field | Value / Expression |
| --- | --- |
| Calendar Id | Select the work calendar to export |
| Start Time | `startOfDay(utcNow())` |
| End Time | `addDays(utcNow(),42)` |
| Other query fields | Unset in the supplied definition |

This action expands recurring events into occurrences. The supplied definition performs one calendar-view request; it has no explicit Top Count / Skip Count loop or pagination runtime configuration.

These bounds mean **midnight today in UTC through the current UTC time 42 days later**. They are not two Eastern-midnight boundaries exactly 42 local calendar days apart. For example, at `2026-11-19T14:00:00Z` the query starts at `2026-11-19T00:00:00Z` (7 PM Eastern on November 18) and ends at `2026-12-31T14:00:00Z` (9 AM Eastern on December 31). The two `utcNow()` calls are also evaluated independently.

### Current flow, step 3: Select

Add **Data Operation → Select**, retaining the name `Select`. Set **From** to:

```text
body('Get_calendar_view_of_events_(V3)')?['value']
```

Enter these eight map keys as text and their values as expressions. Keep the output keys `start` and `end`, but use the connector's timezone-bearing fields as their values:

| JSON key | Value expression |
| --- | --- |
| `id` | `item()?['id']` |
| `start` | `item()?['startWithTimeZone']` |
| `end` | `item()?['endWithTimeZone']` |
| `subject` | `item()?['subject']` |
| `showAs` | `item()?['showAs']` |
| `isAllDay` | `item()?['isAllDay']` |
| `isCancelled` | `item()?['isCancelled']` |
| `responseType` | `item()?['responseType']` |

This preserves titles, availability, all-day flags, cancellation flags, explicit timezone offsets, and response status while omitting descriptions, organizer identities, attendees, and meeting links. `responseType` distinguishes the calendar owner's events (`organizer`) from received invitations. Do not default a missing response to an attendee value. Keep all-day events in the extract: the app uses the owner's PTO and shared closures for availability, ignores teammates' personal time off, and hides other all-day entries from the day-calendar grid.

### Current flow, step 4: Compose

Add **Data Operation → Compose**, retaining the name `Compose`. Paste this expression into its Inputs:

```text
setProperty(
  setProperty(
    json('{}'),
    'generatedAt',
    utcNow()
  ),
  'events',
  body('Select')
)
```

The output is an object with `generatedAt` and an `events` array. This synthetic example shows the corrected mapping, including explicit UTC offsets and response status:

```json
{
  "generatedAt": "2026-11-19T14:00:00Z",
  "events": [
    {
      "id": "synthetic-current-event-1",
      "start": "2026-11-19T15:00:00+00:00",
      "end": "2026-11-19T16:00:00+00:00",
      "subject": "Example meeting",
      "showAs": "busy",
      "isAllDay": false,
      "isCancelled": false,
      "responseType": "accepted"
    }
  ]
}
```

### Current flow, step 5: Update file

Add **OneDrive for Business → Update file** (`UpdateFile`), select `/SuperlistSync/calendar_events.json`, and set **File Content** to:

```text
string(outputs('Compose'))
```

This replaces the existing file contents on each successful run. The supplied action has **Content transfer → Transfer mode: Chunked** in its runtime configuration; retain that setting where exposed in the action settings. The file picker stores an account-specific file ID. If the file is deleted and recreated, reselect it. No **Create file** action or Outlook write is present in the supplied flow.

Save, run a manual test, and check each action in run history. Confirm the intended calendar was selected, Compose produces an object with an event array, and the file updates in both OneDrive and the local synced folder. A successful cloud run does not by itself establish compatibility or complete six-week coverage; review the differences below before connecting current main.

### Current flow: differences from the app's import requirements

| Area | Current flow / observed output | Consequence / recommended revision |
| --- | --- | --- |
| Event timestamps | Select now maps `startWithTimeZone` and `endWithTimeZone` into `start` and `end` | Verified explicit offsets throughout the revised export. Keep these mappings; never append a guessed `Z` or offset to timezone-free values. |
| Capture time | Compose writes `generatedAt` | The folder importer recognizes `capturedAt`, not `generatedAt`, and falls back to file modification time when absent. Use `capturedAt` in the revised wrapper. |
| Window | UTC midnight to `utcNow() + 42 days`; no explicit bounds exported | The folder importer otherwise assumes 42 Eastern calendar days from the file date. Freeze Eastern local dates and export `windowStart`, exclusive `windowEnd`, and `timeZone`. |
| Completeness | Single request; Select retains only `value` | Microsoft documents a 256-event request limit. Later pages can be lost, and Select also discards any next-link metadata the importer could use to detect this. Retrieve all pages before publishing; never assert completeness for the single response without verification. |
| Invitation ownership and declined invitations | Select now includes `responseType` | Verified throughout the revised export. The updated PTO filter ignores teammates' full-day and partial-day leave while retaining the owner's leave and shared closures. Older exports without ownership metadata retain their previous behavior. |
| Overlapping runs | No concurrency setting in the supplied trigger | The recommended design serializes runs to reduce stale overwrites. |

These findings combine the supplied definitions, the connector reference, and a read-only check of the updated OneDrive export with the repository importer. No live flow was edited or triggered by this check. The observed payload imported successfully; that result alone does not prove complete upstream pagination or exact six-week query bounds.

## Recommended flow for complete six-week imports

The following design extends the current five-step flow with explicit bounds, pagination, and completeness checks. It is **not deployed or tenant-tested**. Build and test it before connecting the app: it requests a complete, rolling **42-calendar-day** extract every **30 minutes**, starting at midnight today in Eastern Time.

The instructions use one primary work calendar. Holidays/PTO must exist in that calendar to be included; separate holiday or shared calendars are not automatically combined. When revising the original, replace its retrieval/mapping/wrapper steps with this design rather than running both versions against the same output file.

### 1. Prepare the connections and output file

You need access to [Power Automate](https://make.powerautomate.com/) and working **Office 365 Outlook** and **OneDrive for Business** connections using your work account. These are standard connectors; this design does not require a custom Entra app registration, HTTP action, desktop flow, or local gateway. Your organization's licensing and connector policies still apply.

In OneDrive for Business, create a folder named `SuperlistSync`. For a new installation, create a UTF-8 text file named `calendar_events.json` inside it with `{}` as its initial content; ensure the extension is `.json`, not `.json.txt`. This is only a placeholder and cannot be imported until the first successful flow run replaces it. If the file already contains a working extract, keep it.

| Location | Value |
| --- | --- |
| Cloud folder, relative to your OneDrive root | `/SuperlistSync` |
| Output filename, reused on every run | `calendar_events.json` |
| Example synced Windows file | `C:\Users\<username>\OneDrive - <organization>\SuperlistSync\calendar_events.json` |

Use the cloud file picker in Power Automate. The Windows `C:\...` path belongs only in the local setup; a cloud flow cannot write to that path directly. Other users should use their own synced OneDrive location. Keep OneDrive running and mark the file **Always keep on this device**.

### 2. Create the scheduled flow

Choose **Create → Scheduled cloud flow**, name it `Focus Planner - Export calendar`, and repeat every **30 minutes**. On the **Recurrence** trigger's settings, enable **Concurrency control** with degree **1** to prevent overlapping runs from overwriting each other. Designer labels may vary.

Rename each action to the exact name shown below **before** entering expressions. Add expressions using the designer's **Expression / fx** editor, without a leading `@` or `@{...}`. Names inside `outputs(...)` and `body(...)` refer to actions; names inside `variables(...)` refer to variables.

### 3. Freeze the date window for this run

Add five **Data Operation → Compose** actions in this order. Freezing the values avoids mixing two different days if a run crosses midnight.

| Action name | Inputs (Expression) |
| --- | --- |
| `CapturedAt` | `utcNow()` |
| `WindowStart` | `convertTimeZone(outputs('CapturedAt'), 'UTC', 'Eastern Standard Time', 'yyyy-MM-dd')` |
| `WindowEnd` | `addDays(concat(outputs('WindowStart'), 'T00:00:00'), 42, 'yyyy-MM-dd')` |
| `QueryStartUtc` | `convertToUtc(concat(outputs('WindowStart'), 'T00:00:00'), 'Eastern Standard Time')` |
| `QueryEndUtc` | `convertToUtc(concat(outputs('WindowEnd'), 'T00:00:00'), 'Eastern Standard Time')` |

`WindowEnd` is exclusive. For example, a run on November 19, 2026 covers November 19 through December 30, with an end boundary of December 31. Convert both local midnights independently to UTC: adding 42 × 24 hours to the UTC start can shift the end boundary across daylight-saving changes. Power Automate uses the Windows zone name `Eastern Standard Time` for both EST and EDT; the JSON uses `America/New_York`, which the app understands.

### 4. Retrieve all pages of the calendar view

Use **Office 365 Outlook → Get calendar view of events (V3)**, which expands recurring meetings into occurrences. **Get events (V4)** is not a substitute for an expanded calendar view. Microsoft documents a maximum of **256 events per request** and recommends `Top Count` / `Skip Count` for additional pages. This design uses an explicit loop rather than relying on an automatic pagination setting.

First add these **Initialize variable** actions at the top level, outside any loop:

| Variable name | Type | Initial value |
| --- | --- | --- |
| `AllEvents` | Array | `[]` (empty array) |
| `SkipCount` | Integer | `0` |
| `PageCount` | Integer | `1` |

Add a **Control → Do until** action named `FetchAllPages`. Set its stopping expression to `equals(variables('PageCount'), 0)`. Under its limits, set **Count** to `50` and **Timeout** to `PT20M`. An empty final page proves the loop reached the end, including when the total is an exact multiple of 256.

Inside `FetchAllPages`, add the following actions in order:

1. **Get calendar view of events (V3)**, renamed `GetCalendarPage`, with these settings:

   | Field | Value |
   | --- | --- |
   | Calendar Id | Select your primary work calendar, usually **Calendar** |
   | Start Time | Expression: `outputs('QueryStartUtc')` |
   | End Time | Expression: `outputs('QueryEndUtc')` |
   | Top Count | `256` |
   | Skip Count | Expression: `variables('SkipCount')` |
   | Order By | `start/dateTime asc` |
   | Filter Query / Search | Leave empty |

   Keep automatic pagination off if the designer exposes it, so `SkipCount` controls every request. Use the same calendar, bounds, and ordering for every iteration. Do not filter to work hours, weekdays, busy-only events, or meeting subjects: all-day PTO, existing focus, exceptions, and events overlapping the window boundaries must remain available to the importer.

2. **Set variable** `PageCount` to the expression `length(body('GetCalendarPage')?['value'])`. A missing event array must fail the run; do not replace a missing `value` with an empty array, which would incorrectly declare completion.
3. **Data Operation → Select**, renamed `SelectPage`, with **From** set to `body('GetCalendarPage')?['value']`. In its key/value map, type the keys as plain text and enter each value through **Expression**:

   | JSON key | Value expression |
   | --- | --- |
   | `id` | `item()?['id']` |
   | `subject` | `coalesce(item()?['subject'], '')` |
   | `startWithTimeZone` | `item()?['startWithTimeZone']` |
   | `endWithTimeZone` | `item()?['endWithTimeZone']` |
   | `showAs` | `coalesce(item()?['showAs'], 'unknown')` |
   | `isAllDay` | `coalesce(item()?['isAllDay'], false)` |
   | `isCancelled` | `coalesce(item()?['isCancelled'], false)` |
   | `responseType` | `item()?['responseType']` |

   Preserve the timestamps with their explicit `Z` or numeric offset. Do not substitute the timezone-free `start` / `end` fields or append a guessed offset. The Select map omits descriptions, organizer addresses, attendee lists, and meeting links. The app uses the subject `Focus Time` (case insensitive) to recognize existing focus reservations, and uses availability/response fields to handle free, cancelled, declined, and working-elsewhere events. Keep all-day events; hiding them in the day view is the app's responsibility.

4. **Control → Apply to each**, renamed `AppendPage`, over `body('SelectPage')`. Keep concurrency off (sequential). Inside it, add **Append to array variable**: variable `AllEvents`, value expression `item()`. Append each object individually; appending the entire Select result would create nested arrays.
5. **Increment variable** `SkipCount` by the expression `variables('PageCount')`. Use the returned count, so a shorter nonempty page still leads to another request.

After `FetchAllPages`, add a **Condition** with this expression:

```text
and(equals(variables('PageCount'), 0), lessOrEquals(length(variables('AllEvents')), 10000))
```

In the **If no** branch, use **Terminate** with status **Failed** and a message such as `Calendar retrieval incomplete or over the app's 10000-event limit; previous export retained.` Put **all remaining actions below inside the If yes branch**. Leave their run-after settings at **is successful**. A failed request, loop timeout, exhausted loop count, or failed guard must never publish a partial file with `complete: true`.

Paging is not a transaction against a frozen Outlook calendar. If meetings change during retrieval, pages can shift. Run again if counts or dates look wrong; the importer rejects duplicate IDs rather than accepting them silently. Do not change page ordering or discard a next-link-bearing first response and label it complete.

### 5. Build and replace the JSON file

In the successful branch, add **Compose** named `CalendarExport`. Paste this complete expression into its **Expression** editor:

```text
json(concat(
  '{"capturedAt":"', outputs('CapturedAt'),
  '","timeZone":"America/New_York","windowStart":"', outputs('WindowStart'),
  '","windowEnd":"', outputs('WindowEnd'),
  '","complete":true,"events":', string(variables('AllEvents')), '}'
))
```

The result is an object containing an **array**, not a quoted JSON document. Deliberately omit `schema`: this is the raw Outlook-event wrapper supported by `local-calendar.js`. Adding `schema: "focus-planner-snapshot"` would select a different import format requiring normalized event fields described in **Extract format** below.

Add **OneDrive for Business → Update file**, select the existing `/SuperlistSync/calendar_events.json`, and set **File Content** to the expression `string(outputs('CalendarExport'))`. Reuse this same file each run; do not generate timestamped filenames or append events to yesterday's export. If the selected file is later deleted/recreated, reselect it in this action because its OneDrive identifier can change. Keep the existing export untouched whenever retrieval fails.

This synthetic example shows the resulting shape; the real array contains every retrieved event:

```json
{
  "capturedAt": "2026-11-19T14:00:00Z",
  "timeZone": "America/New_York",
  "windowStart": "2026-11-19",
  "windowEnd": "2026-12-31",
  "complete": true,
  "events": [
    {
      "id": "synthetic-occurrence-1",
      "subject": "Focus Time",
      "startWithTimeZone": "2026-11-19T10:00:00-05:00",
      "endWithTimeZone": "2026-11-19T12:00:00-05:00",
      "showAs": "busy",
      "isAllDay": false,
      "isCancelled": false,
      "responseType": "organizer"
    }
  ]
}
```

### 6. Test the flow before using the planner

Save the flow, choose **Test → Manually**, and inspect its run history and the synced file:

- Confirm both connections use the intended account/calendar, each action succeeds, and the output is valid JSON under the app's 5 MB file limit. `events` must be an array of objects, including `[]` when the calendar is genuinely empty.
- Confirm `windowStart` is today's Eastern date, `windowEnd` is exactly 42 calendar days later, and UTC query boundaries correspond to those local midnights. Include a window crossing a daylight-saving change when validating the conversions.
- Check a recurring meeting near the end of week six, a moved/cancelled recurrence exception, an all-day PTO item, and an existing `Focus Time` reservation against Outlook. Each retained occurrence must have its own unique `id` and offset-bearing start/end timestamps. Free all-day items should remain free, not become PTO.
- For a calendar with more than 256 events, inspect successive `SkipCount` values and verify events beyond the first page are present. Even for fewer events, expect one nonempty page followed by an empty page. Confirm the final `PageCount` is zero and the sum of page lengths matches `AllEvents`.
- In a temporary test copy of the flow, force a calendar-action failure or lower the loop Count so the guard cannot pass. Verify the run fails and the previous output file is unchanged. Restore the settings before enabling the recurring flow.
- Verify the file arrives in the local OneDrive folder. Then follow **Test it** below and use **Connect calendar folder**. The app polls locally every minute; **Refresh now** does not trigger Power Automate. A calendar edit appears after the next flow run, OneDrive sync, and app refresh.

| Symptom | Check |
| --- | --- |
| Only the first portion of six weeks appears | Correct calendar view action, page loop, Skip Count, and completion guard |
| App reports missing offsets or missing times | Select mapping uses `startWithTimeZone` / `endWithTimeZone`; inspect actual action output |
| Duplicate event ID error | Retry the full export after calendar activity settles; verify recurrence expansion and page settings |
| Output is a quoted string or nested arrays | Use the `CalendarExport` expression and append each Select object individually |
| File does not update locally | Flow run history, OneDrive connection/file selection, desktop sync, and local availability |
| File is over 75 minutes old | Failed/disabled flow, expired connection, or stalled OneDrive sync |

The flow only reads Outlook; booking still follows the app's approval and `.ics` import process. Keep real extracts, run-history payloads, and connection secrets out of this repository.

### Can the flow be stored as code?

Yes. Power Automate represents cloud flows as JSON definitions. A non-solution flow can be exported as a **Package (.zip)**; solution-aware flows are exported with their solution and can be versioned with Power Platform tooling. If the flow's Export menu only offers **Learn how** and **Get flow identifier**, check the **Solutions** card on its details page; a solution-aware flow is exported through its solution. The menu alone does not establish an organization policy restriction. To inspect individual steps without a package, use **Edit → select trigger/action → Code view** in the new designer, or **… → Peek code** in the classic designer. Those per-step definitions are the source for the current-flow documentation above; they do not include a complete importable package or establish a successful live run. This README provides manual recreation instructions. Imported flows still need the appropriate connections, calendar/file selections, environment configuration, and permissions; a bare JSON snippet is not a ready-to-run authenticated integration.

Microsoft references: [Outlook connector and paging limits](https://learn.microsoft.com/en-us/connectors/office365/), [OneDrive for Business connector](https://learn.microsoft.com/en-us/connectors/onedriveforbusiness/), [workflow expressions](https://learn.microsoft.com/en-us/azure/logic-apps/expression-functions-reference), [Power Automate timezone conversion](https://learn.microsoft.com/en-us/power-automate/convert-time-zone), [Code view and Peek code](https://learn.microsoft.com/en-us/power-automate/guidance/coding-guidelines/use-peekcode-addnotes), [export/import a non-solution flow](https://learn.microsoft.com/en-us/power-automate/export-import-flow-non-solution), and [export solution-aware flows](https://learn.microsoft.com/en-us/power-automate/export-flow-solution).

## Test it

1. Download the app files together (`index.html`, `planner.js`, `local-calendar.js`, `ical.js`, `app.js`). Open `index.html` in Edge or Chrome, or serve the folder with `py -m http.server 8000` and open `http://localhost:8000/`.
2. Click **Connect calendar folder** and choose `C:\Users\<username>\OneDrive - <organization>\SuperlistSync`. The app reads `calendar_events.json` immediately and checks every minute while open. You can also use **Import** for a one-time calendar extract or saved plan. Proposals are generated automatically.
3. Select a day and approve or reject blocks. **Approve all proposals** approves future blocks locally. Use **Add more focus time** to select an available window, start time and duration, including longer than two hours. Use **Extend 30 min** to lengthen a future local block; the revised block returns to proposed status for approval.
4. Every change saves automatically to this browser. **Save plan** explicitly saves the calendar extract, planning preferences, and decisions again in this browser's local storage on this computer. The plan is restored when you reopen the app at the same origin or local file path. Saving does not create a file, update OneDrive or Outlook, or sync across browsers/devices. There is no plan JSON download; previously saved plan JSON files can still be imported.
5. When upcoming proposals are approved or rejected and upcoming conflicts dismissed, click **Export to Outlook (.ics)**. Choose a save location, then import the file into your main Outlook work calendar as described below.

## Export focus reservations to Outlook

**Export to Outlook (.ics)** becomes available when at least one new approved future block exists and no upcoming proposals or conflicts remain. Reject unwanted proposals and dismiss old conflicts; days without available focus time do not prevent exporting. Eligibility updates each minute, on window focus, and on click. Existing calendar focus, rejected/conflicted blocks, and blocks that have already started are excluded.

In Edge or Chrome with File System Access enabled, every export opens a **Save As** dialog so you can choose the filename and folder. Canceling leaves the plan unchanged. File writing errors are reported and do not mark the plan as booked. When the browser lacks a save picker, the app downloads the file and explains that the browser controls its location; enable **Ask where to save each file** in browser download settings to get a location prompt. Enterprise browser policy can block file pickers even in a supported browser.

In **Outlook on the web**, open **Calendar → Add calendar → Upload from file**, browse to the `.ics`, choose your **main work calendar**, and import. Import into the calendar colleagues use for your availability. See [Microsoft's import instructions](https://support.microsoft.com/en-us/outlook/add-a-calendar-in-outlook-on-the-web). Appointments are titled **Focus Time**, private, and Busy; there are no attendees or meeting invitations. UTC start/end timestamps preserve the snapshot timezone's local times across daylight-saving changes and computers in other timezones.

This is a one-time import, not calendar synchronization. Exporting never writes to Microsoft 365. Reimporting a file may create duplicate appointments, even though event UIDs are stable within a plan. The app cannot verify an Outlook import succeeded. Later edits, extensions, rejections, rebuilds or removals in the app do not update or cancel previously imported appointments; change those in Outlook before importing a revised block. Refresh your Power Automate export after importing so reservations appear as existing calendar focus and are excluded from future exports. Demo exports contain synthetic reservations; import your real calendar before booking.

A privately supplied, self-contained HTML test copy can include an extract already loaded. Open that file in your browser; all code is bundled and no network requests are needed. Browser storage may be cleared or may not be available for local files; local planning decisions are lost if that storage is cleared. Keep the same local filename/path when reopening. A saved plan already in the browser takes precedence over bundled data.

No real calendar extract or generated personal plan belongs in this repository. The app starts empty. To try synthetic data, import `demo-calendar.json` using the normal JSON import; there is no demo button. The example has a fixed six-week window beginning November 19, 2026; update its window and event dates when using it later. The previous v0.2 read-only adapter remains in `outlook.js` for a later release, but this release does not load it or its Microsoft authentication dependency.

## Automatic updates from OneDrive

For Gordon's Power Automate flow, select this **folder** in Edge or Chrome:

```text
C:\Users\<username>\OneDrive - <organization>\SuperlistSync
```

The app reads `calendar_events.json` in that folder. Browsers require the folder picker; JavaScript cannot silently open a Windows path. Only read permission is requested. The file is looked up by name on every check, so atomic OneDrive replacements work. Keep OneDrive running and the file available locally. A local HTML file or localhost may be used; enterprise browser policies can disable folder access. Manual import remains available.

The app remembers the folder handle in IndexedDB where supported. When reopening, it checks permission without displaying an automatic prompt. If permission has expired, click **Refresh now** to allow it again. If the browser cannot remember a folder, select it again. Keep the app at the same local filename/path for consistent storage. **Disconnect** stops reading and forgets the handle while keeping the current plan. Importing a calendar extract or saved plan also disconnects.

The file is checked every 60 seconds, when the tab becomes visible, and when the window regains focus. **Refresh now** checks immediately; it does not run Power Automate. The file's modification time and last successful check are displayed. Exports over 75 minutes old are flagged. A missing file, partial JSON write, revoked access, older snapshot, or invalid data leaves the last valid plan intact and is retried. No polling occurs while the page is closed, and background tabs may delay checks.

Valid local approvals, proposals and rejections survive a refresh within the new window. Blocks that overlap new calendar commitments become **conflicted** and stop counting as approved or available focus. The app proposes replacements when capacity permits; every replacement needs approval. Select the affected day to review or dismiss the conflict. Blocks that have appeared as existing Outlook focus are counted once. New days in the rolling window receive proposals; expired days drop out. Timezone changes require a manual import to avoid shifting earlier decisions.

### Power Automate JSON

The connected file can contain the existing Focus Planner snapshot, an event array, `{ "value": [...] }`, `{ "events": [...] }`, or an action-output `{ "body": { "value": [...] } }` wrapper. A stringified `body` is accepted. Saved Focus Planner plans use the separate **Import** control.

For raw Outlook events, export `id`, `subject`, `startWithTimeZone`, `endWithTimeZone`, `showAs`, `isAllDay`, `isCancelled`, and `responseType`. In Power Automate's **Select**, add `responseType` with the expression `item()?['responseType']`. The response distinguishes the calendar owner's events (`organizer`) from received invitations and excludes declined invitations. Do not default missing responses to an attendee value. Microsoft Graph `responseStatus.response` and an explicit boolean `isOrganizer` are also supported; `isOrganizer` takes precedence. Only the resulting optional `isOrganizer` boolean is retained, not organizer names or addresses.

Teammates' full-day and partial-day PTO/OOO/vacation/time-off invitations are ignored for meeting load, availability, focus proposals, and the day calendar. Shared company holidays and office closures continue to remove working time. Existing exports and saved plans without ownership metadata retain their previous behavior: titles alone cannot reliably identify whose leave an invitation describes. Add the field, run the flow, and refresh the calendar to apply ownership filtering to existing entries.

All-day entries that are not recognized as your PTO/OOO or a company closure are ignored for meeting load, focus availability, existing-focus totals, and conflict checks, regardless of their Busy, Tentative, Free, or Unknown status. For example, an all-day Busy `Focus Friday` or `All-day hold` reminder does not consume the workday or count as booked focus. Timed focus reservations still count normally. Your time off and shared office closures continue to remove working time, while teammates' personal leave is ignored. This applies to existing snapshots and saved plans as well as new imports; **Generate plan** fills newly available focus time.

Microsoft Graph `start`/`end` objects with explicit UTC timezones also work. Timestamps must carry an offset or explicitly declare UTC. Ambiguous timezone-free timestamps are rejected with instructions to include `startWithTimeZone`/`endWithTimeZone`. Titles matching `Focus time` or explicit `kind: "existing-focus"` identify existing focus. Descriptions, organizer and attendee fields are discarded.

The upstream flow must export the complete 42-day **calendar view**, expanding recurring occurrences and collecting every page. Pagination links or `complete: false` are rejected. Without explicit snapshot metadata, the app assumes the flow's asserted full 42-day coverage beginning on the file modification date in Eastern Time. Include `capturedAt`, `timeZone`, `windowStart`, and exclusive `windowEnd` around an `events`/`value` array if the flow uses different bounds. A bare event array cannot independently prove coverage or freshness; the file modification timestamp is only a fallback. Explicit metadata is preferred.

### Single-file copy

Run `python scripts/build_standalone.py` to produce `dist/Focus_Planner_v0.8.html`. This bundles the same local code and requires no web server, packages or network requests. Open the HTML in Edge or Chrome, then connect the folder. No personal calendar data is bundled.

## Planning preferences

A green **All applicable focus days are fully planned** banner appears near the plan controls when every upcoming working day meets its daily focus target using existing calendar focus or approved local blocks, includes a continuous two-hour block, and has no upcoming proposals or unresolved conflicts. Full-day PTO/OOO and past days are excluded. The banner confirms when the latest calendar extract contains the reservations; if approved local blocks still need booking, it directs you to export and import them into Outlook. Having no suggestions because there is insufficient space does not count as fully planned.

Open **Planning preferences** in the header. The plan controls place **Export to Outlook (.ics)** on the second row, beside its status message. Outstanding proposal decisions or conflicts make that message red; it returns to the normal color when those are resolved. Coverage cards show **✅ Focus Time Secured** in green and **No focus planned** in red. Full-day PTO/OOO cards omit the workday-exclusion caption. Generating again with no new blocks keeps existing decisions without displaying a no-change notice.

| Setting | Value |
| --- | --- |
| Horizon | 42 calendar days, exclusive end boundary; weekdays only |
| Daily minimum | One continuous 120-minute block; additional time is optional |
| Additional blocks | At least 60 minutes; no two-hour daily cap |
| Meeting goal | At most 50% of each workday (preference; no meeting edits) |
| Working hours | 9:00 AM–5:30 PM Eastern |
| Preferred focus time | Morning |
| Meeting preference | Avoid before 9:30 AM |
| Commute buffer | Optional checkbox excludes 9:00–9:30 from new focus proposals |
| Approval | Required; saved locally only |

Timed events entirely outside **6:00 AM–10:00 PM Eastern** are excluded during calendar import, refresh, and saved-plan restore. This includes overnight invites from shared mailboxes. Events that overlap this range are retained, including early-morning and evening meetings. The filter uses `America/New_York`, including daylight-saving changes, independently of the computer or snapshot timezone. All-day PTO/OOO and holiday handling is preserved. Focus-planning hours and meeting-load targets are unchanged.

Meeting preference and focus availability are separate. Calendar events already before 9:30 still block scheduling. No new proposals are placed in past time; today's proposals begin at the next quarter hour. Existing calendar focus counts toward the target only in uninterrupted, unconflicted blocks of at least 60 minutes inside work hours. Overlapping existing focus is counted once. Busy, tentative, unknown, and out-of-office events block planning; working-elsewhere, cancelled, and declined events are excluded when preparing the extract. Free events may be included for display but do not block planning. All-day and multi-day events are expanded across covered dates.

Generate is idempotent: it retains decisions and first seeks a continuous two-hour block, preferring mornings among qualifying windows. If no such window is available, it fills up to the daily total in blocks of at least 60 minutes and flags the missing continuous block. Rejected slots are avoided by automatic generation until rebuilding the plan. Manual additions can deliberately reuse a rejected window. Days without sufficient capacity show their shortfall. Full days off are excluded from focus targets, shortfalls, and meeting/focus coverage totals. Time off must be present in the imported work calendar; separate holiday calendars are not fetched.

Select a day to add longer or additional focus proposals, even after the minimum is covered. Available windows exclude calendar events and active local blocks, respect the commute preference, and never include past time. Extend a future proposed or approved block by 30 minutes when free space follows it. Extensions reset approval. **Remove from plan** withdraws a future local approval; it does not remove calendar events. Existing Outlook focus remains read-only.

The meeting-limit and continuous-focus indicators are retained from v0.2. The meeting limit uses the 9:30 AM–5:30 PM window, counts overlapping meetings once, and excludes existing focus. Explicit out-of-office intervals are removed from both meeting time and working time; fully out-of-office days are excluded from the meeting-limit day count. Holidays must be represented in the extract to be recognized. This goal is informational and does not move meetings.

The interface shows upcoming dates and planning coverage. Completed/missed controls and outcome reporting have been removed. Older completed statuses restore as approved; older missed statuses restore as rejected. The existing browser storage key is retained to preserve saved plans.

Changing the commute checkbox or using **Rebuild plan** explicitly clears local planning decisions after an in-app confirmation. Manually importing a file replaces the current plan after confirmation and disconnects automatic updates. Connecting the folder preserves valid decisions when refreshing the calendar. Invalid imports never replace it.

## Extract format

The upstream extract must include **all** relevant events in the window, including expanded recurring occurrences and exceptions. Fetch all pages; do not label a truncated response complete. If using date partitions, fetch every partition and deduplicate events overlapping boundaries. Dates are local calendar dates in `timeZone`; `windowEnd` is exclusive and exactly 42 days after `windowStart`.

```json
{
  "schema": "focus-planner-snapshot",
  "version": 1,
  "source": "outlook-extract",
  "timeZone": "America/New_York",
  "capturedAt": "2026-11-18T16:00:00Z",
  "windowStart": "2026-11-18",
  "windowEnd": "2026-12-30",
  "complete": true,
  "events": [
    {
      "id": "local-event-1",
      "start": "2026-11-18T14:30:00Z",
      "end": "2026-11-18T15:30:00Z",
      "status": "busy",
      "kind": "busy",
      "allDay": false
    }
  ]
}
```

Event timestamps require explicit UTC or numeric offsets. `status` is `busy`, `tentative`, `oof`, `unknown`, or `free`; `kind` is `busy` or `existing-focus`. Local pseudonymous IDs suffice. Optional `title`, `allDay`, and boolean `isOrganizer` are retained. Omit `isOrganizer` when ownership is unknown; do not encode it as a string in normalized snapshots. Meeting descriptions and organizer identities are discarded when importing an extract or saved plan, including older backups. Restored browser plans are saved back without those fields while retaining planning decisions. New extracts should omit meeting descriptions, organizers and attendee lists.

The selected day's **Day calendar** shows timed events and local focus proposals/approvals against a half-hour grid. Boxes match event durations; overlaps appear in separate columns. Titles and times remain readable as literal text, with full labels available on hover or keyboard focus. All-day PTO/OOO appears in banners above the grid and as a striped PTO/OOO card with no focus planning required. Other all-day invites remain hidden in the grid. Out-of-office availability is treated as time off; non-free all-day items titled PTO, OOO, out of office, vacation, time off, holiday, shutdown, office closed, or year-end closing/closure also remove their dates from working time. Free all-day items remain free. All-day midnight dates retain their calendar dates and exclusive end date, including UTC-midnight exports. Free events never block planning or count as meetings. Timed events crossing midnight remain visible on each affected day. The grid starts with work hours and expands for earlier/later timed events; it scrolls on long days and narrow screens.

Working days with no active focus blocks show **No focus planned** in red on their coverage card. PTO/OOO days retain **No focus planning needed** in the normal note color.

The planner uses the snapshot's timezone rather than the computer's timezone, including daylight-saving transitions. An extract cannot prove its own upstream completeness; the producer is responsible for the `complete` assertion.

## Persistence and limits

- The extract and plan stay in memory and localStorage under `focus-planner-v0.4`. They persist across midnight and reloads, with the snapshot's fixed date window. Imported files are not uploaded by this app.
- Results reflect the most recently read export. Folder connections check every minute and when the page regains focus; background tabs may be throttled. Changes appear after Power Automate runs and OneDrive syncs. The capture timestamp is always visible. Conflict recovery uses the latest export, with no direct Outlook connection.
- Recheck current availability before importing reservations. Outlook booking requires a manual `.ics` import; direct Outlook writes, notifications, closed-page background jobs and cloud plan storage are not implemented.
- Browser-local data is specific to that browser/origin and may be cleared. There is no plan JSON download. If localStorage fails, the app reports that changes may be lost when closing or reloading; the current plan stays open in memory.
- Legacy v0.1–v0.3 storage is left untouched and not migrated. The scheduler is deterministic, not machine learning.

## Validation

Run the pure planning tests with Node 18+:

```sh
node --test tests/*.test.js
```

The optional browser test requires Playwright, Chromium, and an environment that permits Chromium to open local sockets:

```sh
node tests/browser.cjs
```

It covers import, approvals, rejection, reload, local save/restore, malformed imports, storage failure, an offline load and a narrow viewport. `PLAYWRIGHT_MODULE` and `CHROMIUM_PATH` can point to an existing local installation.




Validation: 41 Node checks pass, covering planning, local-file refresh, export approval gating, private/Busy appointments, timezone/DST conversion, stable UIDs, iCalendar escaping/folding, save success, cancellation, write failures and the download fallback. The optional Playwright suite was extended for export but could not launch Chromium here because local browser sockets are blocked. Native Windows file dialogs and Outlook import remain to be checked on the target machine.


Calendar coverage cards show the day number, three-letter month, and three-letter weekday. Timed meetings on a day off remain visible for reference but do not count toward workday targets.

