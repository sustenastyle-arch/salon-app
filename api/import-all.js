// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Backs both the 📥 データ復元 button and the one-time "migrate this
// device's local data to the cloud" button — bulk-upserts many days in one request.
import { setDays, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "method not allowed" });
    return;
  }

  const days = req.body || {};
  if (typeof days !== "object" || Object.keys(days).length === 0) {
    res.status(400).json({ error: "body must be a non-empty {date: data} object" });
    return;
  }

  try {
    const count = await setDays(days);
    res.status(200).json({ count });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
