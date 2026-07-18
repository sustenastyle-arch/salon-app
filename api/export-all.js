// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Backs the 📤 データバックアップ button: returns every saved day as
// one object, downloaded client-side as a single JSON file.
import { getAllDayBlobs, checkAuth } from "./_lib/dayDataStore.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  try {
    const days = await getAllDayBlobs();
    res.status(200).json({ days });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
