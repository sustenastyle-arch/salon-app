# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page React app ("Spa Daily Sheet") used by a Hawaii spa/salon to run its daily
operations sheet: therapist schedule, appointments, retail sales, deposits, ticket-package
purchases, staff purchases (社販), refunds, payroll, and monthly sales reporting. Data is
synced from Square (bookings + gift card activity) and persisted per day in a shared Upstash
Redis store, so the same data shows up on every computer the site is opened from (previously
`localStorage`-only, which kept each browser's data separate — see "Data model and persistence").

## Commands

- `npm run dev` — start Vite dev server (includes dev-only Square API middleware, see below)
- `npm run build` — production build
- `npm run lint` — ESLint
- `npm run preview` — preview a production build locally

There is no test suite configured in this repo.

## Environment

Copy `.env.example` to `.env` and set:
- `SQUARE_ACCESS_TOKEN` / `SQUARE_LOCATION_ID` — Square Developer Dashboard credentials, kept server-side only
- `VITE_MANAGER_PIN` — PIN required to unlock a "確定済み" (locked/submitted) day's sheet
- `VITE_APP_PASSWORD` — optional shared password gate (`src/App.jsx`) for when the app is hosted publicly; skipped entirely if unset (local dev). Doubles as the shared-secret header for the cloud data API (see below) — also skipped if unset.
- `KV_REST_API_URL` / `KV_REST_API_TOKEN` — Upstash Redis REST credentials (Vercel Storage integration); copy these by hand from the Vercel dashboard, since Vercel only auto-injects them into deployed environments, not local `.env`
- `GMAIL_USER` / `GMAIL_APP_PASSWORD` — salon Gmail account + Google App Password used to save the automatic sales-report draft (see "Sales-report Gmail draft on lock"); also set in Vercel's env vars (Production). `REPORT_EMAIL_TO` optionally overrides the recipients.

## Architecture

### Square API proxy (dual implementation)

Square API calls must not expose `SQUARE_ACCESS_TOKEN` to the browser, so every Square
integration is implemented **twice** with matching logic:
- `vite.config.js` — Vite dev-server middleware (`squareGiftCardApi`, `squareBookingsApi`), used by `npm run dev`
- `api/giftcard-activities.js`, `api/square-bookings.js` — Vercel serverless functions, used in production

When changing how a Square call works (date range, response shape, filtering, etc.), update
**both** the middleware in `vite.config.js` and the matching file in `api/`, or dev and prod
will silently diverge. Both compute the day boundary as Hawaii time (`UTC-10`, no DST) rather
than using `Intl`/local time.

`squareBookingsApi` hardcodes a `TEAM_MEMBER_MAP` (Square team member ID → therapist first
name) — this must be updated by hand if staff change or Square IDs are regenerated.

### Data model and persistence

Almost the entire app lives in `src/spa-daily-sheet.jsx` (~5200 lines, one file). There is no
router:
- One day's full state (`appointments`, `retails`, `deposits`, `ticketPurchases`,
  `staffPurchases`, `refunds`, `forgottenTips`, `locked`, `workingStaff`) is one JSON blob keyed
  by date (`YYYY-MM-DD`) in a shared Upstash Redis store (`api/_lib/dayDataStore.js`) — reached
  through `apiFetch()` calls to `/api/day-data` etc., never `localStorage` directly. This used to
  be `localStorage`-only (one key per browser), which meant data entered on one computer never
  appeared on another; moving it server-side fixed that. Any change to the per-day shape must
  stay backward-compatible with already-saved days' JSON, or aggregate views (and old data) will
  break silently — same constraint as before, just against Redis blobs instead of localStorage.
- The monthly sales report (`exportSalesReportXlsx`) and `PayrollTab` both fetch every day in
  their target range in one call to `/api/day-data-range` (bounded `MGET`, not a per-day
  round trip) — there is no separate aggregate store.
- A deposit/gift-card prepayment can be recorded on one day for an appointment on a different
  (often future) day, so finding "deposits for date X" or "this client's deposit history" needs
  to scan every day ever saved — `/api/deposits` (`mode=date|client`) does this server-side via
  Redis `SCAN`, deliberately unbounded (data volume is tiny for a single small spa; a lookback
  window would silently drop older deposits with no error).
- `📤 Backup Data` downloads the full store as one JSON file (`/api/export-all`). Restoring
  (from that file via `/api/import-all`, or from an automatic snapshot via
  `/api/auto-backup-restore`) is deliberately **not** a button anymore — a full-store restore
  overwrites everything, so it's done by calling those endpoints directly. They have no UI
  caller but are the emergency recovery path, so don't delete them as "unused".
- Every `/api/*` day-data endpoint checks a shared-secret header against `VITE_APP_PASSWORD`
  (`checkAuth` in `dayDataStore.js`) — skipped if unset, same as the client-side password gate.
- Excel export uses `xlsx` (SheetJS).

### Automatic daily backup

A Vercel Cron job (`vercel.json`, `0 13 * * *` = 3am Hawaii time, safely after closing) hits
`/api/auto-backup` once a day, which snapshots every saved day (same data `/api/export-all`
would return) under its own Redis key (`auto-backup-YYYY-MM-DD`, separate from the day-sheet
keys themselves so an accidental delete/overwrite of a day via the app can't also wipe out its
backup) and prunes anything older than the most recent 30 snapshots. This exists so a day's data
can be recovered even if nobody remembered to click the manual backup button. `/api/auto-backup`
optionally checks a `CRON_SECRET` bearer token (set in Vercel's env vars; skipped if unset) to
confirm the request actually came from Vercel's scheduler — Cron jobs don't run under `vite dev`,
so `saveAutoBackup()` must be triggered manually in local dev. `/api/auto-backup-list` lists
available snapshot dates and `/api/auto-backup-restore` restores the full store to one via
`setDays()` — the same bulk-upsert helper `import-all` uses (no UI for either; see above).

### Sales-report Gmail draft on lock

Locking a day (`handleLockToggle` → `createSalesReportDraft` in `spa-daily-sheet.jsx`) builds
that day's 売上報告 email in the browser (`src/lib/salesReportEmail.js`: day breakdown,
month-to-date 累計, gift-card ※ notes — numbers from `computeDayTotals`, same as the Excel
export) plus the monthly Excel (`buildSalesReportXlsx`, the same file as ⬇️ Download Excel), and
POSTs both to `/api/sales-report-draft`. The server saves it to the salon Gmail's **Drafts**
via IMAP + App Password (`api/_lib/gmailDraft.js`) — it is never sent; the owner reviews and
presses Send. Re-locking only drafts again if the day's numbers changed: a fingerprint of the
day's figures is kept under its own Redis key (`report-draft-YYYY-MM-DD`, outside the
`spa-sheet-*` prefix so export/backup scans never see it). A draft failure only shows a toast;
it never undoes the lock.

### Square check before lock

A day with any appointment card that has no amounts entered (price, tip, package and new-ticket
amount all 0) can't be locked at all, and there is no PIN override. Such a card is a Square-synced
booking that either someone forgot to enter or that was cancelled and left on the sheet, and it
still counts as a client. The card has to be filled in or deleted.

A day can only be locked once it matches Square (`checkSquareBeforeLock` in
`spa-daily-sheet.jsx`): locking first calls `/api/square-payments` and requires all of
`squareReconcileRows` (cash total, card total, cash tip, card tip) to match. The 🔍 Square
Reconcile panel shows the same rows. Square gift card redemptions count as card payments on
Square's side but are never revenue on the sheet (`giftCardUsed`), so the API also returns
`giftCardTotal`/`giftCardTip` (card brand `SQUARE_GIFT_CARD`), and these are subtracted from
Square's card figures before comparing. A genuine exception can still be locked with the manager
PIN: for example, a refund of an earlier day's sale, which Square nets against the original day,
or Square being unreachable.

### Vercel Hobby plan: 12-function cap

Vercel's Hobby plan allows at most **12 serverless functions** per deployment, and `api/`
already has exactly 12 route files — adding a 13th fails the deploy. New endpoints must not be
new `api/*.js` files: put the handler under `api/_lib/` and route to it through an existing
function with a `vercel.json` rewrite + `action` query param. `/api/sales-report-draft` works this
way (rewritten to `/api/import-all?action=sales-report-draft` → `api/_lib/salesReportDraftHandler.js`);
the Vite dev middleware still registers the plain path. Production deploys from the `master` branch.

### Cloud data API (dual implementation, shared logic)

Like the Square proxy below, the day-data API exists once as Vite dev middleware
(`vite.config.js`'s `dayDataApi`) and once as Vercel serverless functions (`api/day-data.js`,
`api/day-data-range.js`, `api/deposits.js`, `api/export-all.js`, `api/import-all.js`,
`api/auto-backup.js`, `api/auto-backup-list.js`, `api/auto-backup-restore.js`, plus
`/api/sales-report-draft` via rewrite) — but
unlike Square, the actual Redis logic isn't duplicated: both import the same
`api/_lib/dayDataStore.js` (files under `api/_lib/` are helpers, not routes — Vercel's
zero-config routing skips `_`-prefixed paths). Only the request/response plumbing differs
(manual `url.searchParams`/`res.end(JSON.stringify(...))` in the Vite middleware vs.
`req.query`/`res.status(n).json(...)` in the Vercel functions). `dayDataStore.js`'s Redis client
is lazily constructed on first use, not at module load — under `vite dev`, `vite.config.js`'s
top-level imports run before its `defineConfig` callback calls `loadEnv()` and copies values into
`process.env`, so a module-load-time client would read undefined credentials.

### Pricing and staff capability tables

Near the top of `spa-daily-sheet.jsx` are the tables that drive validation and pricing — these
are the ones most likely to need edits when the business changes menu/pricing/staff:
- `PRICE_TABLE` — per-service `{ body_service, body_tip, cav_service, cav_tip }` (cav = the
  body-contouring machine; `cav_*` is `null` if the service has no machine option)
- `TICKET_PACKAGE_PRICES`, `MENU_OPTIONS`, `SQUARE_SERVICES`, `RETAIL_PRODUCTS`
- `CAV_CAPABLE` / `BODY_CAPABLE` / `DUAL_LICENSE` — which therapists can run the cav machine,
  do body massage, or both; used to validate appointment assignments (e.g. a cav service needs
  either a dual-license therapist or a body therapist + a separate cav therapist)
- `THERAPISTS` is the canonical staff list; keep it in sync with `TEAM_MEMBER_MAP` in
  `vite.config.js`/`api/square-bookings.js` when staff change

### Treatment minutes (Maki is paid by the minute)

Maki is on call and is paid by the minutes she actually treated, so a body therapist's saved
`duration` is real money for her. Other staff are paid by the service/tip amounts, so their
minutes don't change their pay. The payroll CSV shows a Minutes column only for Maki.

When a separate therapist runs the machine (cav), the body therapist's minutes are the course
length minus the machine time:

| Course | Machine | Body therapist's minutes |
|---|---|---|
| Weight Loss 75 | 40 | **35** |
| Weight Loss 90 | 40 | **50** |
| Weight Loss 115 | 40 | **75** |
| Improving Posture 60 / 75 | (none) | 60 / 75 |
| Improving Posture 90 | 15 | **75** (90 when the same therapist also runs the machine) |
| Improving Posture 120 | 15 | **105** |

Weight Loss always uses 40 minutes of machine time, whether or not the visit is a ticket.
`saveAppt` forces this rule. A Weight Loss ticket once fell back to the 15-minute default:
10/09/2026 Vickie Miao was saved as Maki 75min instead of 50, and Maki was overpaid 25 minutes.
Any change to how `duration`, `bodyMins` or `cavMins` are saved must keep these numbers. Check
Maki's minutes against this table in every payroll audit.

### UI structure

`SpaDailySheet` (default export of `spa-daily-sheet.jsx`) is the top-level component with all
top-level state; it renders tabs (schedule/payroll/etc.) and a set of per-record modals
(`ApptModal`, `RetailModal`, `DepositModal`, `TicketPurchaseModal`, `StaffPurchaseModal`,
`RefundModal`) that are opened by setting `editingX` state to a record (or `null` to close).
Each modal owns its own form state and calls `onSave`/`onDelete` back into the parent, which is
what calls `save()` (an async `POST /api/day-data`).

`App.jsx` wraps everything in `PasswordGate`, a client-side-only check against
`VITE_APP_PASSWORD` (a soft deterrent, not real auth — the password ships in the bundle).

## Business facts for sales analysis

The monthly statistics workbook (Desktop `Dr.Body 月次統計.xlsx`, built outside this repo by
`C:\Users\suste\koguchi-auto\monthly-stats.mjs`) ends each month with a written analysis the owner
uses in staff 1-on-1s. Any analysis of this salon's numbers has to take these facts from the owner
into account:
- Every new customer, local or traveler, gets $20 off the first visit. No first visit is at list
  price, so a shift between new travelers and new locals doesn't explain a change in sales.
- Travelers come mostly from the US mainland. Only about 3–5 Japanese customers come a month, so
  Japanese holidays don't explain anything; US holidays (Thanksgiving, Christmas–New Year, spring
  break, summer) do.
- About 40% of visits are ticket sessions (Sep 2026: 102 of 237, Oct 2026: 98 of 235). A ticket
  session fills a slot but its revenue was booked when the ticket was sold, so the month's sales
  depend heavily on how many tickets are *sold* that month (`ticketPurchases`, `newTicket`,
  same-day tickets), and on retail. Oct 2026 fell $5,183 from Sep almost entirely from fewer
  ticket sales (−$3,690) and less retail (−$1,180); the owner confirmed that reading.
- When ticket holders fill the book there is no room for new customers, so the owner lowers the
  Instagram ad spend on purpose. Fewer new customers or lower ad spend can be a deliberate choice,
  not a problem.
- The owner wants each staff member to build their own regulars, so that their bookings don't
  depend on ads or the season. The concrete goal is 2–3 new regulars a month each: clients who ask
  for them, rebook on the spot, or keep coming back. Charging a nomination fee (指名料) is each
  staff member's own choice; some don't want to, and that's fine. Staff comments name one
  strength and one goal each and never blame anyone.
- Travelers almost never come back, so return rates (overall, by staff, by source) count new
  locals only. Target: 40% of new locals come back (4 of 10).
- Aya is the model the owner wants everyone to aim for: she already charges a $10 nomination fee
  and most of her clients ask for her by name. Her new-customer numbers are low because her book is
  filled by her own regulars, so that is a strength, not a gap.
- Whether a weak month comes from the season or from staff skill/service is often unclear; say so
  rather than pick one.
- Sales goal: $50,000 a month, tip-included.
