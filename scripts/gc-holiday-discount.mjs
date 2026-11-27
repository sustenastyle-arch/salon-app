// Turns Square's "Holiday Gift Card 10% OFF ($100+)" discount on or off at the register.
// Run by Windows Task Scheduler (tasks "DrBody GC Discount ON/OFF") for the post-Thanksgiving
// sale: on at 11/27/2026 00:00, off at 11/30/2026 23:59 (Hawaii time).
//
//   node scripts/gc-holiday-discount.mjs on|off
//
// "Off" hides the discount from every location (present_at_all_locations: false) instead of
// deleting it, so the same discount can be turned on again for the next sale.
import fs from "fs";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";

const DISCOUNT_ID = "7ESFCH55IS2IELK5FP3OBTSZ";
const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const env = Object.fromEntries(fs.readFileSync(path.join(REPO_ROOT, ".env"), "utf8").split(/\r?\n/)
  .filter(l => l.includes("=") && !l.trim().startsWith("#"))
  .map(l => { const i = l.indexOf("="); return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^["']|["']$/g, "")]; }));
const LOG = path.join(os.homedir(), "OneDrive", "デスクトップ", "ギフトカード割引ログ.txt");
const log = (msg) => { const line = `${new Date().toLocaleString("ja-JP", { timeZone: "Pacific/Honolulu" })}  ${msg}`; console.log(line); fs.appendFileSync(LOG, line + "\r\n"); };

const mode = process.argv[2];
if (!["on", "off"].includes(mode)) { console.error("usage: node scripts/gc-holiday-discount.mjs on|off"); process.exit(1); }

const h = { Authorization: `Bearer ${env.SQUARE_ACCESS_TOKEN}`, "Square-Version": "2024-10-17", "Content-Type": "application/json" };
try {
  const obj = (await (await fetch(`https://connect.squareup.com/v2/catalog/object/${DISCOUNT_ID}`, { headers: h })).json()).object;
  if (!obj) throw new Error("discount not found in Square");
  obj.present_at_all_locations = mode === "on";
  delete obj.present_at_location_ids;
  delete obj.absent_at_location_ids;
  const up = await (await fetch("https://connect.squareup.com/v2/catalog/object", { method: "POST", headers: h, body: JSON.stringify({ idempotency_key: `gc-discount-${mode}-${Date.now()}`, object: obj }) })).json();
  if (up.errors) throw new Error(JSON.stringify(up.errors));
  log(`${obj.discount_data.name}: ${mode === "on" ? "ON (レジで使えます)" : "OFF (レジに出ません)"}`);
} catch (e) {
  log(`失敗 (${mode}): ${e.message}`);
  process.exit(1);
}
