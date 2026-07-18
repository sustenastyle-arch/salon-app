// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. A deposit/gift-card prepayment can be recorded on one day for an
// appointment on a different (often future) day, so finding "deposits for date X" or
// "this client's deposit history" needs to scan every day ever saved — that scan now
// happens server-side instead of the browser enumerating every localStorage key.
import { getAllDayBlobs, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  const { mode } = req.query;

  try {
    const allDays = await getAllDayBlobs();

    if (mode === "date") {
      const date = req.query.date;
      if (!date) {
        res.status(400).json({ error: "date query param required" });
        return;
      }
      const found = [];
      for (const [recordedDate, d] of Object.entries(allDays)) {
        (d.deposits || []).forEach((dep) => {
          if (dep.appointmentDate === date && (dep.type === "deposit" || dep.type === "giftcard")) {
            found.push({ ...dep, recordedDate });
          }
        });
      }
      found.sort((a, b) => (a.appointmentTime || "").localeCompare(b.appointmentTime || ""));
      res.status(200).json({ deposits: found });
      return;
    }

    if (mode === "client") {
      const name = (req.query.name || "").toLowerCase().trim();
      if (!name) {
        res.status(200).json({ deposits: [] });
        return;
      }
      const found = [];
      for (const [sheetDate, d] of Object.entries(allDays)) {
        (d.deposits || []).forEach((dep) => {
          if ((dep.clientName || "").toLowerCase().trim() === name && (dep.type === "deposit" || dep.type === "giftcard")) {
            found.push({ ...dep, sheetDate });
          }
        });
      }
      found.sort((a, b) => a.sheetDate.localeCompare(b.sheetDate));
      res.status(200).json({ deposits: found });
      return;
    }

    res.status(400).json({ error: "mode query param required (date|client)" });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
