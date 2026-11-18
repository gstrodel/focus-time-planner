# Focus Planner v0.1

A browser-based prototype for planning and protecting focus time with a rolling six-week horizon.

## Initial configuration

- Rolling horizon: **6 weeks** (42 calendar days, weekdays scheduled)
- Daily focus target: **2 hours**
- Minimum block: **60 minutes**
- Preferred time: **morning**
- Automatic conflict recovery: **on** (creates replacement proposals)
- Approval required before calendar changes: **on**

## Run

Open `index.html` in a modern browser. No build process, server, dependencies, or credentials are required.

This is a **demo calendar**. It does **not** connect to Outlook or make changes to your Microsoft 365 calendar. Scheduling approvals update local browser state only. Changes and feedback persist in localStorage.

## Features

1. Generate 42 days of candidate focus blocks around simulated meetings (weekdays only).
2. View six weeks of scheduled coverage; inspect daily proposed, approved, and unmet focus time.
3. Approve all proposed blocks or approve individual suggestions. Reject recommendations.
4. Simulate a newly introduced meeting collision with an approved focus block and propose replacement time automatically; replacement still requires approval.
5. Record completed or missed focus sessions. See basic completion and disruption metrics.
6. Export current demo state as JSON; reset the demo.

## Scheduling algorithm

For each weekday, the engine starts with a 9 AM–5 PM work window and excludes simulated busy events and previously approved focus reservations. Available 60-minute+ intervals are ranked by preferred morning placement, then duration/earliness. It attempts to reserve 120 minutes per day with blocks of at least 60 minutes. The engine does not overbook an unavailable day and shows a shortfall rather than pretending capacity exists.

Conflicts trigger **replacement proposals**, not unapproved calendar edits. Previously approved blocks that are displaced are marked as conflicted, and the UI requests a new approval. All actions are local demo state.

## Outlook integration design (not implemented)

For production, register an organization-approved Microsoft Entra ID application; use MSAL delegated authorization and Microsoft Graph calendar APIs. Read `calendarView` across the six-week horizon (including expanded recurring meetings), interpret Outlook free/busy and mailbox timezone/working hours, preserve event IDs and etags, and re-check availability before each write. Write only approved Focus Planner-owned events using `Calendars.ReadWrite`; never modify unrelated meetings. Store ownership markers and avoid duplicates. Run reconciliation on a schedule or notifications, but queue proposed relocations for approval. Persist activity history securely, with retention and least privilege. Coordinate tenant policies and admin approval with IT.

## Limitations

- Simulated calendar events only; no live Microsoft Graph or notifications.
- Work hours fixed at 9 AM–5 PM and local browser time.
- Uses deterministic rules, not machine learning.
- Single user/browser; localStorage is not an enterprise data store.
- Demo conflict generation is synthetic; not an Outlook webhook.

## Next milestone

Read-only Microsoft Graph connection and timezone-safe recurring event expansion, followed by an explicit approval-gated write pipeline and durable audit history.
