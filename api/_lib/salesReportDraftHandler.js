// Request handler for POST /api/sales-report-draft — saves the sales-report email built in the
// browser (when a day is locked) as a Gmail draft with the monthly Excel attached. Never sends it.
//
// Not its own file under api/: Vercel's Hobby plan allows at most 12 serverless functions per
// deployment and api/ already has 12, so vercel.json rewrites /api/sales-report-draft to
// /api/import-all?action=sales-report-draft and import-all.js hands the request off to this.
// gmailDraft.js (nodemailer + imapflow) is imported lazily so a plain import doesn't load it.
export async function salesReportDraftHandler(req, res) {
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
    const { createSalesReportDraftIfChanged } = await import("./gmailDraft.js");
    const result = await createSalesReportDraftIfChanged({ date, fingerprint, subject, body, attachments: attachments || [] });
    res.status(200).json(result);
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
