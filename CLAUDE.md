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
- `📤 データバックアップ` / `📥 データ復元` export/import the full store as one JSON file
  (`/api/export-all` / `/api/import-all`); `☁️ この端末のデータをクラウドへ移行` is a one-time
  button that reads whatever's still in a browser's local `localStorage` (pre-migration data) and
  bulk-uploads it via the same `import-all` endpoint — the only place the app still reads
  `localStorage` on purpose.
- Every `/api/*` day-data endpoint checks a shared-secret header against `VITE_APP_PASSWORD`
  (`checkAuth` in `dayDataStore.js`) — skipped if unset, same as the client-side password gate.
- Excel export uses `xlsx` (SheetJS).

### Cloud data API (dual implementation, shared logic)

Like the Square proxy below, the day-data API exists once as Vite dev middleware
(`vite.config.js`'s `dayDataApi`) and once as Vercel serverless functions (`api/day-data.js`,
`api/day-data-range.js`, `api/deposits.js`, `api/export-all.js`, `api/import-all.js`) — but
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

### UI structure

`SpaDailySheet` (default export of `spa-daily-sheet.jsx`) is the top-level component with all
top-level state; it renders tabs (schedule/payroll/etc.) and a set of per-record modals
(`ApptModal`, `RetailModal`, `DepositModal`, `TicketPurchaseModal`, `StaffPurchaseModal`,
`RefundModal`) that are opened by setting `editingX` state to a record (or `null` to close).
Each modal owns its own form state and calls `onSave`/`onDelete` back into the parent, which is
what calls `save()` (an async `POST /api/day-data`).

`App.jsx` wraps everything in `PasswordGate`, a client-side-only check against
`VITE_APP_PASSWORD` (a soft deterrent, not real auth — the password ships in the bundle).
