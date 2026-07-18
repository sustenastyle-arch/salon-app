// Vercel serverless function — restores every day's data to the state captured in a past
// automatic daily snapshot (see api/auto-backup.js), overwriting whatever's currently saved.
// Used by the 🕐 Restore from Auto-Backup picker for recovering from an accidental delete or
// bad edit without needing a manually downloaded backup file.
import { restoreAutoBackup, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "method not allowed" });
    return;
  }

  const { date } = req.body || {};
  if (!date) {
    res.status(400).json({ error: "date required" });
    return;
  }

  try {
    const count = await restoreAutoBackup(date);
    if (count == null) {
      res.status(404).json({ error: "backup not found" });
      return;
    }
    res.status(200).json({ count });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
