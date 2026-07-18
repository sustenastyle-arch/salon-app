// Vercel serverless function — lists the dates that have an automatic daily snapshot
// available (see api/auto-backup.js), for the 🕐 Restore from Auto-Backup picker.
import { listAutoBackups, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  try {
    const dates = await listAutoBackups();
    res.status(200).json({ dates });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
