import { Redis } from "@upstash/redis";

// Shared by both the Vite dev middleware (vite.config.js) and the Vercel serverless
// functions (api/day-data.js etc) — unlike the Square integrations, there's no
// framework-specific divergence in this logic, so it lives in one place instead of
// being hand-duplicated. Both runtimes read the same env vars (Vercel's Upstash
// Marketplace integration injects KV_REST_API_URL/KV_REST_API_TOKEN in production;
// they're copied into local .env by hand for `npm run dev`, same as SQUARE_ACCESS_TOKEN).
//
// Lazily constructed (not at module load) because under `vite dev`, vite.config.js's
// top-level imports run before its defineConfig callback calls loadEnv() and copies
// the values into process.env — a module-load-time `new Redis(...)` here would read
// undefined. By the time any function below is actually called (an incoming request),
// vite.config.js has already populated process.env.
let _redis = null;
function getRedis() {
  if (!_redis) {
    _redis = new Redis({ url: process.env.KV_REST_API_URL, token: process.env.KV_REST_API_TOKEN });
  }
  return _redis;
}

const KEY_PREFIX = "spa-sheet-";
const keyFor = (date) => `${KEY_PREFIX}${date}`;

export async function getDay(date) {
  const data = await getRedis().get(keyFor(date));
  return data ?? null;
}

export async function setDay(date, payload) {
  await getRedis().set(keyFor(date), payload);
}

// Inclusive date range (e.g. one calendar month, or a half-month pay period) — bounded
// by the caller, so a single MGET over the computed key list is enough; no SCAN needed.
export async function getDaysInRange(startDate, endDate) {
  const dates = [];
  const d = new Date(`${startDate}T00:00:00Z`);
  const end = new Date(`${endDate}T00:00:00Z`);
  while (d <= end) {
    dates.push(d.toISOString().slice(0, 10));
    d.setUTCDate(d.getUTCDate() + 1);
  }
  if (dates.length === 0) return {};
  const values = await getRedis().mget(...dates.map(keyFor));
  const result = {};
  dates.forEach((date, i) => {
    if (values[i] != null) result[date] = values[i];
  });
  return result;
}

// Every day ever saved — used for the deposit scans (a deposit can be recorded for a
// future/other date so it can't be scoped to a bounded range), the full backup export,
// and the migration bulk-import. Data volume is tiny for a single small spa (dozens of
// saves/day), so a full SCAN is fine — deliberately not adding a lookback window, since
// a deposit older than an arbitrary cutoff would silently stop appearing with no error.
export async function getAllDayBlobs() {
  const result = {};
  let cursor = "0";
  do {
    const [nextCursor, keys] = await getRedis().scan(cursor, { match: `${KEY_PREFIX}*`, count: 100 });
    cursor = nextCursor;
    if (keys.length > 0) {
      const values = await getRedis().mget(...keys);
      keys.forEach((key, i) => {
        if (values[i] != null) result[key.slice(KEY_PREFIX.length)] = values[i];
      });
    }
  } while (cursor !== "0");
  return result;
}

// Bulk upsert — used by both the 📥 restore button and the one-time local→cloud
// migration button. Accepts either raw JSON strings (the format older manual backup
// files use, since they came straight from localStorage.getItem) or already-parsed
// objects, so it can't fail on an old backup file created before this migration.
export async function setDays(dateToPayload) {
  const entries = Object.entries(dateToPayload);
  await Promise.all(entries.map(([date, payload]) => {
    const parsed = typeof payload === "string" ? JSON.parse(payload) : payload;
    return setDay(date, parsed);
  }));
  return entries.length;
}

// Shared-secret check reused across every endpoint below — the app already ships
// VITE_APP_PASSWORD to the browser bundle for the client-side PasswordGate, so sending
// it again as a header adds no new exposure. If it's unset (e.g. local dev without the
// site password configured), skip the check — same "soft deterrent, skipped if unset"
// behavior as the existing PasswordGate.
export function checkAuth(req) {
  const expected = process.env.VITE_APP_PASSWORD;
  if (!expected) return true;
  return req.headers["x-app-password"] === expected;
}
