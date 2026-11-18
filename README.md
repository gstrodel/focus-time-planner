# Focus Planner v0.4

An offline, browser-based six-week focus planner using an Outlook calendar **extract**. No Microsoft app registration or sign-in is needed for this release. Import your extract, review proposals, approve or reject them, and save the results. Approvals do not create Outlook appointments.

## Test it

1. Download the app files together (`index.html`, `planner.js`, `app.js`). Open `index.html` in Edge or Chrome, or serve the folder with `py -m http.server 8000` and open `http://localhost:8000/`.
2. Under **Your calendar extract**, select the six-week extract JSON provided separately. Proposals are generated automatically.
3. Select a day and approve or reject blocks. **Approve all proposals** approves future blocks locally. Record completed or missed sessions when reviewing your results.
4. Every change saves automatically to this browser. **Save plan** explicitly saves it again. **Download plan** exports the extract, decisions and feedback together as JSON. Import that file to resume in another browser or device.

A privately supplied, self-contained HTML test copy can include an extract already loaded. Open that file in your browser; all code is bundled and no network requests are needed. Download your plan as a backup, since browser storage may be cleared or may not be available for local files. Keep the same local filename/path when reopening. A saved plan already in the browser takes precedence over bundled data.

No real calendar extract or generated personal plan belongs in this repository. The public app starts empty; **Try demo data** is explicitly synthetic. The previous v0.2 read-only adapter remains in `outlook.js` for a later release, but v0.4 does not load it or its Microsoft authentication dependency.

## Planning preferences

| Setting | Value |
| --- | --- |
| Horizon | 42 calendar days, exclusive end boundary; weekdays only |
| Daily target | 120 minutes |
| Minimum block | 60 minutes |
| Working hours | 9:00 AM–5:30 PM Eastern |
| Preferred focus time | Morning |
| Meeting preference | Avoid before 9:30 AM |
| Commute buffer | Optional checkbox excludes 9:00–9:30 from new focus proposals |
| Approval | Required; saved locally only |

Meeting preference and focus availability are separate. Calendar events already before 9:30 still block scheduling. No new proposals are placed in past time; today's proposals begin at the next quarter hour. Existing calendar focus counts toward the target only in uninterrupted, unconflicted blocks of at least 60 minutes inside work hours. Overlapping existing focus is counted once. Busy, tentative, unknown, and out-of-office events block planning; free, working-elsewhere, cancelled, and declined events are excluded when preparing the extract. All-day and multi-day events are expanded across covered dates.

Generate is idempotent: it retains decisions and adds proposals only where the daily target remains unmet. Rejected slots are avoided until rebuilding the plan. The scheduler favors mornings, then continuous blocks, and keeps enough time for a second minimum block when splitting the target. Days without sufficient capacity show their shortfall. Time off still appears as a target shortfall; v0.4 does not infer holiday calendars or adjust the weekday target.

Changing the commute checkbox or using **Rebuild plan** explicitly clears approvals and feedback after an in-app confirmation. Importing a new file replaces the current plan after confirmation. Invalid imports never replace it.

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

Event timestamps require explicit UTC or numeric offsets. `status` is `busy`, `tentative`, `oof`, or `unknown`; `kind` is `busy` or `existing-focus`. Local pseudonymous IDs suffice. Optional `title`, plain-text `description`, and `organizer` (`name` and `email`) are retained for meeting review. Expand a meeting under Calendar availability to see its description and organizer. Outlook organizer may differ from a delegate who sent the invitation; this release does not retrieve the original invitation message sender. Bodies render as literal text, never executable markup. Attendee lists and original Outlook IDs are omitted. These details remain in your private extract, browser storage and downloaded plan, and are excluded from git. The planner uses the snapshot's timezone rather than the computer's timezone, including daylight-saving transitions. An extract cannot prove its own upstream completeness; the producer is responsible for the `complete` assertion.

## Persistence and limits

- The extract and plan stay in memory and localStorage under `focus-planner-v0.4`. They persist across midnight and reloads, with the snapshot's fixed date window. Imported files are not uploaded by this app.
- Results reflect availability when captured, not later cancellations or new meetings. The capture timestamp is always visible. Dynamic connectivity and automatic conflict recovery against Outlook are deferred.
- Recheck current availability before making actual calendar reservations in a later workflow. No Outlook writes, notifications, background jobs or cloud plan storage are implemented.
- Browser-local data is specific to that browser/origin and may be cleared. Export JSON for a portable backup. If localStorage fails, the app reports it and keeps download available.
- Legacy v0.1–v0.3 storage is left untouched and not migrated. The scheduler is deterministic, not machine learning.

## Validation

Run the pure planning tests with Node 18+:

```sh
node --test tests/planner.test.js tests/app.test.js
```

The optional browser test requires Playwright, Chromium, and an environment that permits Chromium to open local sockets:

```sh
node tests/browser.cjs
```

It covers import, approvals, rejection, reload, download/restore, malformed imports, storage failure, an offline load and a narrow viewport. `PLAYWRIGHT_MODULE` and `CHROMIUM_PATH` can point to an existing local installation.
