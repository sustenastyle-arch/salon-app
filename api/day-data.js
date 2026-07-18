// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Reads/writes one day's blob in the shared cloud store so the same
// data shows up on every computer, not just the browser it was entered on.
import { getDay, setDay, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  try {
    if (req.method === "GET") {
      const date = req.query.date;
      if (!date) {
        res.status(400).json({ error: "date query param required (YYYY-MM-DD)" });
        return;
      }
      const data = await getDay(date);
      res.status(200).json({ data });
      return;
    }

    if (req.method === "POST") {
      const { date, data } = req.body || {};
      if (!date || data === undefined) {
        res.status(400).json({ error: "date and data required" });
        return;
      }
      await setDay(date, data);
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: "method not allowed" });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
