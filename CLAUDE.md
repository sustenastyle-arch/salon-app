# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page React app ("Spa Daily Sheet") used by a Hawaii spa/salon to run its daily
operations sheet: therapist schedule, appointments, retail sales, deposits, ticket-package
purchases, staff purchases (社販), refunds, payroll, and monthly sales reporting. Data is
synced from Square (bookings + gift card activity) and persisted to `localStorage` per day —
there is no backend database.

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
- `VITE_APP_PASSWORD` — optional shared password gate (`src/App.jsx`) for when the app is hosted publicly; skipped entirely if unset (local dev)

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

Almost the entire app lives in `src/spa-daily-sheet.jsx` (~4700 lines, one file). There is no
router and no backend database:
- One day's full state (`appointments`, `retails`, `deposits`, `ticketPurchases`,
  `staffPurchases`, `refunds`, `locked`) is stored under the `localStorage` key
  `spa-sheet-YYYY-MM-DD`.
- The monthly sales report (`exportSalesReportXlsx`) and `PayrollTab` both work by iterating
  every day of the target month and reading each `spa-sheet-YYYY-MM-DD` key directly — there is
  no aggregate store. Any change to the per-day shape must stay backward-compatible with
  already-saved days' JSON, or these aggregate views (and old data) will break silently.
- Excel export uses `xlsx` (SheetJS).

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
what writes to `localStorage`.

`App.jsx` wraps everything in `PasswordGate`, a client-side-only check against
`VITE_APP_PASSWORD` (a soft deterrent, not real auth — the password ships in the bundle).
