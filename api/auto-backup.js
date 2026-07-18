// Vercel serverless function — triggered once a day by the Cron Job defined in vercel.json.
// Takes a full snapshot of every saved day so a day's data can still be recovered even if
// nobody remembered to click the manual 📤 Backup Data button (see saveAutoBackup in
// api/_lib/dayDataStore.js).
import { saveAutoBackup } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  // Vercel signs Cron requests with this header when CRON_SECRET is set in the project's env
  // vars — verifies the request actually came from Vercel's scheduler, not an arbitrary caller.
  // Skipped if unset, same "soft deterrent" pattern as checkAuth in dayDataStore.js.
  const expected = process.env.CRON_SECRET;
  if (expected && req.headers["authorization"] !== `Bearer ${expected}`) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  try {
    const result = await saveAutoBackup();
    res.status(200).json(result);
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
