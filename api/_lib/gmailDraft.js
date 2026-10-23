// Saves an email into the salon Gmail account's Drafts folder — never sends it. Backs the
// automatic sales-report draft created when a day is locked: the owner reviews the draft in
// Gmail and presses Send themselves.
//
// Uses IMAP APPEND with a Gmail App Password (GMAIL_USER / GMAIL_APP_PASSWORD) instead of the
// Gmail API, since the API would need a Google Cloud OAuth app whose refresh tokens expire
// every 7 days while it's in "testing" mode. Shared by the Vite dev middleware
// (vite.config.js) and the Vercel function (api/_lib/salesReportDraftHandler.js), same pattern as
// dayDataStore.js.

import nodemailer from "nodemailer";
import { ImapFlow } from "imapflow";
import { getReportFingerprint, setReportFingerprint } from "./dayDataStore.js";

// Same people the hand-written reports always went to. REPORT_EMAIL_TO (comma-separated)
// overrides this without a code change if the list ever changes.
const DEFAULT_RECIPIENTS = ["dr.bodyinc@gmail.com", "hiroki@pt-w.jp", "sustenastyle@gmail.com"];

export function reportRecipients() {
  const fromEnv = (process.env.REPORT_EMAIL_TO || "").split(",").map(s => s.trim()).filter(Boolean);
  return fromEnv.length > 0 ? fromEnv : DEFAULT_RECIPIENTS;
}

// attachments: [{ filename, contentBase64 }]
export async function createGmailDraft({ subject, body, attachments = [] }) {
  const user = process.env.GMAIL_USER;
  const pass = (process.env.GMAIL_APP_PASSWORD || "").replace(/\s+/g, ""); // Google shows it as 4 groups of 4
  if (!user || !pass) throw new Error("GMAIL_USER / GMAIL_APP_PASSWORD not set");

  // streamTransport only builds the MIME message (UTF-8 subject, attachment encoding) — nothing is sent.
  const composer = nodemailer.createTransport({ streamTransport: true, buffer: true, newline: "windows" });
  const { message } = await composer.sendMail({
    from: user,
    to: reportRecipients(),
    subject,
    text: body,
    attachments: attachments.map(a => ({ filename: a.filename, content: Buffer.from(a.contentBase64, "base64") })),
  });

  const client = new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user, pass }, logger: false });
  await client.connect();
  try {
    // The Drafts folder's name depends on the account's language ("[Gmail]/Drafts" vs
    // "[Gmail]/下書き"), so find it by its special-use flag instead of by name.
    const mailboxes = await client.list();
    const drafts = mailboxes.find(mb => mb.specialUse === "\\Drafts");
    if (!drafts) throw new Error("Gmail Drafts folder not found");
    await client.append(drafts.path, message, ["\\Draft", "\\Seen"]);
  } finally {
    await client.logout();
  }
}

// Locking a day calls this every time, but the owner sometimes unlocks a day to fix something
// that doesn't touch the money (a name, a note) — only draft again when the day's numbers
// (`fingerprint`, built by buildSalesReportEmail) differ from the last draft for that date.
// Returns { skipped: true } when nothing changed.
export async function createSalesReportDraftIfChanged({ date, fingerprint, subject, body, attachments }) {
  if (date && fingerprint && (await getReportFingerprint(date)) === fingerprint) {
    return { skipped: true };
  }
  await createGmailDraft({ subject, body, attachments });
  if (date && fingerprint) await setReportFingerprint(date, fingerprint);
  return { skipped: false };
}
