// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Saves the sales-report email built in the browser (when a day is locked)
// as a Gmail draft, with the monthly Excel attached. Never sends it.
import { checkAuth } from "./_lib/dayDataStore.js";
import { createSalesReportDraftIfChanged } from "./_lib/gmailDraft.js";

export default async function handler(req, res) {
  if (!checkAuth(req)) {
    res.status(401).json({ error: "unauthorized" });
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "method not allowed" });
    return;
  }

  const { date, fingerprint, subject, body, attachments } = req.body || {};
  if (!subject || !body) {
    res.status(400).json({ error: "subject and body required" });
    return;
  }

  try {
    const result = await createSalesReportDraftIfChanged({ date, fingerprint, subject, body, attachments: attachments || [] });
    res.status(200).json(result);
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
