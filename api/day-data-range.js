// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Returns every existing day's blob within an inclusive date range in
// one round trip, replacing the monthly sales report / payroll tab's old per-day loop.
import { getDaysInRange, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  const { start, end } = req.query;
  if (!start || !end) {
    res.status(400).json({ error: "start and end query params required (YYYY-MM-DD)" });
    return;
  }

  try {
    const days = await getDaysInRange(start, end);
    res.status(200).json({ days });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
