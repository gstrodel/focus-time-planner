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


## v0.2 — Read-only Outlook calendar (requires your Microsoft 365 tenant approval)

The browser prototype now supports **read-only Outlook synchronization** directly through Microsoft Graph. It requests only the delegated `Calendars.Read` scope; it **cannot create, update, or delete Outlook events**, regardless of which focus blocks you approve in the prototype.

### One-time Microsoft Entra setup

1. Ask your Microsoft 365 administrator whether custom app registrations and delegated `Calendars.Read` are permitted in your organization. Do not bypass corporate app policies.
2. Register a **single-tenant** application in Microsoft Entra ID. Under **Authentication**, add platform **Single-page application** with redirect URI `http://localhost:8000/` (including the final slash).
3. Under **API permissions**, request delegated Microsoft Graph `Calendars.Read`, and obtain administrator consent if required by policy. No client secret, write scope, or application permission is needed.
4. Copy the **Directory (tenant) ID** and **Application (client) ID** from the app overview. These are public identifiers, not secrets.

### Launch locally (do not double-click the HTML for Outlook sign-in)

Open a command prompt in the repository folder:

```powershell
py -m http.server 8000
```

Then open `http://localhost:8000/` in Microsoft Edge/Chrome. Enter the two Entra IDs in the Outlook panel and click **Connect & sync**. Authenticate with your work account and approve the requested read-only calendar scope if your organization permits it. The app reads up to 42 days of `/me/calendarView` including recurring occurrences, excludes free/working-elsewhere/cancelled/declined events, and generates local proposals from remaining gaps.

The Microsoft authentication library is loaded from Microsoft's CDN, so internet access is required.

**Security notes:** The access token is handled by MSAL in browser session storage; the app never commits credentials or calendar content to GitHub. No calendar event titles are stored by the app. Only derived busy intervals are held in memory. Note that approved/proposed **focus slots** are persisted in localStorage; using the JSON Export button creates an explicitly downloaded file containing focus session times. Browser storage is not suitable for regulated sensitive information.

**Important:** Each sync resets the local prototype plan to avoid treating stale suggestions as valid. The browser-side app has no server, and therefore cannot continually monitor conflicts while closed. Approval remains strictly local; it does **not** reserve time in Outlook. A subsequent phase requires an authorized write workflow and refresh/revalidation before event creation.
