# Focus Planner v0.7

An offline, browser-based six-week focus planner using an Outlook calendar **extract**. No Microsoft app registration or sign-in is needed for this release. Import your extract, review proposals, approve or reject them, and save the results. Approvals do not create Outlook appointments.

## Test it

1. Download the app files together (`index.html`, `planner.js`, `local-calendar.js`, `app.js`). Open `index.html` in Edge or Chrome, or serve the folder with `py -m http.server 8000` and open `http://localhost:8000/`.
2. Click **Connect calendar folder** and choose `C:\Users\UBC3554\OneDrive - Takeda\SuperlistSync`. The app reads `calendar_events.json` immediately and checks every minute while open. You can also use **Import** for a one-time calendar extract or saved plan. Proposals are generated automatically.
3. Select a day and approve or reject blocks. **Approve all proposals** approves future blocks locally. Use **Add more focus time** to select an available window, start time and duration, including longer than two hours. Use **Extend 30 min** to lengthen a future local block; the revised block returns to proposed status for approval.
4. Every change saves automatically to this browser. **Save plan** explicitly saves it again. **Download plan** exports the extract, planning decisions together as JSON. Import that file to resume in another browser or device.

A privately supplied, self-contained HTML test copy can include an extract already loaded. Open that file in your browser; all code is bundled and no network requests are needed. Download your plan as a backup, since browser storage may be cleared or may not be available for local files. Keep the same local filename/path when reopening. A saved plan already in the browser takes precedence over bundled data.

No real calendar extract or generated personal plan belongs in this repository. The app starts empty; **Try demo data** is explicitly synthetic. The previous v0.2 read-only adapter remains in `outlook.js` for a later release, but this release does not load it or its Microsoft authentication dependency.

## Automatic updates from OneDrive

For Gordon's Power Automate flow, select this **folder** in Edge or Chrome:

```text
C:\Users\UBC3554\OneDrive - Takeda\SuperlistSync
```

The app reads `calendar_events.json` in that folder. Browsers require the folder picker; JavaScript cannot silently open a Windows path. Only read permission is requested. The file is looked up by name on every check, so atomic OneDrive replacements work. Keep OneDrive running and the file available locally. A local HTML file or localhost may be used; enterprise browser policies can disable folder access. Manual import remains available.

The app remembers the folder handle in IndexedDB where supported. When reopening, it checks permission without displaying an automatic prompt. If permission has expired, click **Refresh now** to allow it again. If the browser cannot remember a folder, select it again. Keep the app at the same local filename/path for consistent storage. **Disconnect** stops reading and forgets the handle while keeping the current plan. Importing a saved plan or switching to demo also disconnects.

The file is checked every 60 seconds, when the tab becomes visible, and when the window regains focus. **Refresh now** checks immediately; it does not run Power Automate. The file's modification time and last successful check are displayed. Exports over 75 minutes old are flagged. A missing file, partial JSON write, revoked access, older snapshot, or invalid data leaves the last valid plan intact and is retried. No polling occurs while the page is closed, and background tabs may delay checks.

Valid local approvals, proposals and rejections survive a refresh within the new window. Blocks that overlap new calendar commitments become **conflicted** and stop counting as approved or available focus. The app proposes replacements when capacity permits; every replacement needs approval. Select the affected day to review or dismiss the conflict. Blocks that have appeared as existing Outlook focus are counted once. New days in the rolling window receive proposals; expired days drop out. Timezone changes require a manual import to avoid shifting earlier decisions.

### Power Automate JSON

The connected file can contain the existing Focus Planner snapshot, an event array, `{ "value": [...] }`, `{ "events": [...] }`, or an action-output `{ "body": { "value": [...] } }` wrapper. A stringified `body` is accepted. Saved Focus Planner plans use the separate **Import** control.

For raw Outlook events, export `id`, `subject`, `startWithTimeZone`, `endWithTimeZone`, `showAs`, `isAllDay`, and `isCancelled`; response status helps exclude declined invitations. Microsoft Graph `start`/`end` objects with explicit UTC timezones also work. Timestamps must carry an offset or explicitly declare UTC. Ambiguous timezone-free timestamps are rejected with instructions to include `startWithTimeZone`/`endWithTimeZone`. Titles matching `Focus time` or explicit `kind: "existing-focus"` identify existing focus. Descriptions, organizer and attendee fields are discarded.

The upstream flow must export the complete 42-day **calendar view**, expanding recurring occurrences and collecting every page. Pagination links or `complete: false` are rejected. Without explicit snapshot metadata, the app assumes the flow's asserted full 42-day coverage beginning on the file modification date in Eastern Time. Include `capturedAt`, `timeZone`, `windowStart`, and exclusive `windowEnd` around an `events`/`value` array if the flow uses different bounds. A bare event array cannot independently prove coverage or freshness; the file modification timestamp is only a fallback. Explicit metadata is preferred.

### Single-file copy

Run `python scripts/build_standalone.py` to produce `dist/Focus_Planner_v0.7.html`. This bundles the same local code and requires no web server, packages or network requests. Open the HTML in Edge or Chrome, then connect the folder. No personal calendar data is bundled.

## Planning preferences

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

Meeting preference and focus availability are separate. Calendar events already before 9:30 still block scheduling. No new proposals are placed in past time; today's proposals begin at the next quarter hour. Existing calendar focus counts toward the target only in uninterrupted, unconflicted blocks of at least 60 minutes inside work hours. Overlapping existing focus is counted once. Busy, tentative, unknown, and out-of-office events block planning; working-elsewhere, cancelled, and declined events are excluded when preparing the extract. Free events may be included for display but do not block planning. All-day and multi-day events are expanded across covered dates.

Generate is idempotent: it retains decisions and first seeks a continuous two-hour block, preferring mornings among qualifying windows. If no such window is available, it fills up to the daily total in blocks of at least 60 minutes and flags the missing continuous block. Rejected slots are avoided by automatic generation until rebuilding the plan. Manual additions can deliberately reuse a rejected window. Days without sufficient capacity show their shortfall. Time off still appears as a target shortfall; holiday calendars and adjusted weekday targets are not inferred.

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

Event timestamps require explicit UTC or numeric offsets. `status` is `busy`, `tentative`, `oof`, `unknown`, or `free`; `kind` is `busy` or `existing-focus`. Local pseudonymous IDs suffice. Optional `title` and `allDay` are retained. Meeting descriptions and organizer metadata are discarded when importing an extract or saved plan, including older backups. Restored browser plans are saved back without those fields while retaining planning decisions. New extracts should omit meeting descriptions, organizers and attendee lists.

The selected day's **Day calendar** shows timed events and local focus proposals/approvals against a half-hour grid. Boxes match event durations; overlaps appear in separate columns. Titles and times remain readable as literal text, with full labels available on hover or keyboard focus. All-day events are hidden regardless of availability status. This is a display filter: busy/tentative all-day events and all-day PTO still affect scheduling and meeting calculations. Free events never block planning or count as meetings. Timed events crossing midnight remain visible on each affected day. The grid starts with work hours and expands for earlier/later timed events; it scrolls on long days and narrow screens.

The planner uses the snapshot's timezone rather than the computer's timezone, including daylight-saving transitions. An extract cannot prove its own upstream completeness; the producer is responsible for the `complete` assertion.

## Persistence and limits

- The extract and plan stay in memory and localStorage under `focus-planner-v0.4`. They persist across midnight and reloads, with the snapshot's fixed date window. Imported files are not uploaded by this app.
- Results reflect the most recently read export. Folder connections check every minute and when the page regains focus; background tabs may be throttled. Changes appear after Power Automate runs and OneDrive syncs. The capture timestamp is always visible. Conflict recovery uses the latest export, with no direct Outlook connection.
- Recheck current availability before making actual calendar reservations in a later workflow. No Outlook writes, notifications, closed-page background jobs or cloud plan storage are implemented.
- Browser-local data is specific to that browser/origin and may be cleared. Export JSON for a portable backup. If localStorage fails, the app reports it and keeps download available.
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

It covers import, approvals, rejection, reload, download/restore, malformed imports, storage failure, an offline load and a narrow viewport. `PLAYWRIGHT_MODULE` and `CHROMIUM_PATH` can point to an existing local installation.




Latest validation: 31 Node checks pass, including the connected-folder UI harness. The optional Playwright browser run was attempted here but Chromium could not start because this environment blocks its local sockets. Native folder selection and OneDrive synchronization must be verified on the target Windows machine.
