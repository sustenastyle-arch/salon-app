// Patches confirmed ("locked") days straight into the owner's hand-formatted Excel sales
// report on the Desktop, without recreating the file. Run by double-clicking
// sync-day-to-report.bat in the project root (or `node scripts/sync-day-to-report.mjs
// [YYYY-MM]` from a terminal — defaults to the current month).
//
// Only ever writes B..M of the day's own row (row = 4 + day-of-month, matching the sheet's
// fixed header layout), leaving every other byte of the workbook (styles, formulas, other
// months, hand-typed notes) untouched — same technique validated by hand for July 16 in
// scratch_patch_day16.cjs. Re-running is safe: values are set absolutely, not added, and
// only locked days are touched, so it can be re-run any time to pick up newly-confirmed
// days. If a locked day's numbers are later corrected by hand in Redis, re-running will
// overwrite this file's cell with the corrected value — but a manual edit made directly in
// the Excel file itself for an already-locked day will be silently overwritten on the next
// run, since this script has no way to tell "corrected by hand" apart from "stale".

import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import yauzl from "yauzl";
import yazl from "yazl";
import { Redis } from "@upstash/redis";
import { computeDayTotals } from "../src/lib/reportTotals.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..");
const REPORT_PATH = "C:/Users/Dr. Body/Desktop/2026 July Sales Report -/Sales Report2026 NEW.xlsx";

const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

function loadEnv() {
  const text = fs.readFileSync(path.join(REPO_ROOT, ".env"), "utf8");
  const env = {};
  for (const line of text.split("\n")) {
    if (!line.includes("=") || line.trim().startsWith("#")) continue;
    const i = line.indexOf("=");
    let value = line.slice(i + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    env[line.slice(0, i).trim()] = value;
  }
  return env;
}

function readZipEntries(zipPath) {
  return new Promise((resolve, reject) => {
    const entries = [];
    yauzl.open(zipPath, { lazyEntries: true }, (err, zipfile) => {
      if (err) return reject(err);
      zipfile.readEntry();
      zipfile.on("entry", (entry) => {
        if (/\/$/.test(entry.fileName)) {
          entries.push({ name: entry.fileName, dir: true });
          zipfile.readEntry();
          return;
        }
        zipfile.openReadStream(entry, (err, stream) => {
          if (err) return reject(err);
          const chunks = [];
          stream.on("data", (c) => chunks.push(c));
          stream.on("end", () => {
            entries.push({ name: entry.fileName, data: Buffer.concat(chunks) });
            zipfile.readEntry();
          });
          stream.on("error", reject);
        });
      });
      zipfile.on("end", () => resolve(entries));
      zipfile.on("error", reject);
    });
  });
}

function writeZipEntries(zipPath, entries) {
  return new Promise((resolve, reject) => {
    const zip = new yazl.ZipFile();
    entries.forEach((e) => {
      if (e.dir) return; // yazl doesn't need explicit directory entries for a valid xlsx
      zip.addBuffer(e.data, e.name);
    });
    const chunks = [];
    zip.outputStream.on("data", (c) => chunks.push(c));
    zip.outputStream.on("end", () => {
      fs.writeFileSync(zipPath, Buffer.concat(chunks));
      resolve();
    });
    zip.outputStream.on("error", reject);
    zip.end();
  });
}

// Resolve which xl/worksheets/sheetN.xml corresponds to a given "July 2026"-style sheet
// name, by following workbook.xml's <sheet name r:id> list through workbook.xml.rels'
// r:id -> target mapping. Sheet tab names/order can change if the owner reorders tabs in
// Excel, so this can't just assume sheet1.xml is always the current month.
function resolveSheetFile(entries, monthLabel) {
  const workbookXml = entries.find(e => e.name === "xl/workbook.xml").data.toString("utf8");
  const relsXml = entries.find(e => e.name === "xl/_rels/workbook.xml.rels").data.toString("utf8");

  // Sheet tab names in this workbook have inconsistent stray spaces (e.g. "July  2026",
  // " April 2026", "May 2026 ") — normalize whitespace on both sides before comparing.
  const normalize = (s) => s.trim().replace(/\s+/g, " ");
  const sheetRe = /<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"/g;
  let m;
  let rId = null;
  while ((m = sheetRe.exec(workbookXml))) {
    if (normalize(m[1]) === normalize(monthLabel)) { rId = m[2]; break; }
  }
  if (!rId) throw new Error(`No sheet tab named "${monthLabel}" found in workbook.xml`);

  const relRe = new RegExp(`<Relationship Id="${rId}"[^>]*Target="([^"]*)"`);
  const relMatch = relsXml.match(relRe);
  if (!relMatch) throw new Error(`No relationship found for ${rId}`);
  return `xl/${relMatch[1]}`;
}

function setCellValue(rowXml, col, rowNum, value) {
  const cellRe = new RegExp(`<c r="${col}${rowNum}"([^>]*?)(/>|>([\\s\\S]*?)<\\/c>)`);
  const m = rowXml.match(cellRe);
  if (!m) throw new Error(`Cell ${col}${rowNum} not found`);
  const attrs = m[1];
  const inner = m[3] || "";

  if (value === 0 || value === undefined || value === null || value === "") {
    // Leave blank (self-closing), matching the sheet's own "0 => no <v>" convention — but
    // keep any <f> (formula) tag intact if present.
    const fMatch = inner.match(/<f[^>]*\/>|<f[^>]*>.*?<\/f>/);
    const newInner = fMatch ? fMatch[0] : "";
    const replacement = newInner
      ? `<c r="${col}${rowNum}"${attrs}>${newInner}</c>`
      : `<c r="${col}${rowNum}"${attrs}/>`;
    return rowXml.replace(cellRe, replacement);
  }

  const fMatch = inner.match(/<f[^>]*\/>|<f[^>]*>.*?<\/f>/);
  const fTag = fMatch ? fMatch[0] : "";
  const replacement = `<c r="${col}${rowNum}"${attrs}>${fTag}<v>${value}</v></c>`;
  return rowXml.replace(cellRe, replacement);
}

function patchRow(sheetXml, row, totals) {
  const cellValues = {
    B: totals.totalSales,
    C: totals.clients,
    D: totals.cashTreatment,
    E: totals.cashProduct,
    F: totals.totalCash,
    G: totals.cashTip,
    H: totals.cardTreatment,
    I: totals.cardProduct,
    J: totals.totalCard,
    K: totals.cardTip,
    L: totals.totalTip,
    M: totals.grandTotal,
  };
  const rowRe = new RegExp(`<row r="${row}"[^>]*>[\\s\\S]*?<\\/row>`);
  const rowMatch = sheetXml.match(rowRe);
  if (!rowMatch) throw new Error(`Row ${row} not found in sheet`);
  let rowXml = rowMatch[0];
  Object.entries(cellValues).forEach(([col, val]) => {
    rowXml = setCellValue(rowXml, col, row, val);
  });
  return sheetXml.replace(rowRe, rowXml);
}

async function main() {
  const monthArg = process.argv[2]; // e.g. "2026-07"
  const now = new Date();
  const [y, m] = (monthArg || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`)
    .split("-").map(Number);
  const monthLabel = `${MONTH_NAMES[m - 1]} ${y}`;
  const monthStr = `${y}-${String(m).padStart(2, "0")}`;
  const lastDay = new Date(y, m, 0).getDate();

  console.log(`Syncing locked days for ${monthLabel} into: ${REPORT_PATH}`);

  const env = loadEnv();
  const redis = new Redis({ url: env.KV_REST_API_URL, token: env.KV_REST_API_TOKEN });

  const entries = await readZipEntries(REPORT_PATH);
  const sheetEntryName = resolveSheetFile(entries, monthLabel);
  const sheetEntry = entries.find(e => e.name === sheetEntryName);
  if (!sheetEntry) throw new Error(`${sheetEntryName} not found in workbook`);
  let sheetXml = sheetEntry.data.toString("utf8");

  let written = 0;
  const skipped = [];
  for (let d = 1; d <= lastDay; d++) {
    const dateStr = `${monthStr}-${String(d).padStart(2, "0")}`;
    const data = await redis.get(`spa-sheet-${dateStr}`);
    if (!data) continue;
    if (!data.locked) { skipped.push(dateStr); continue; }
    const totals = computeDayTotals(data);
    const row = 4 + d;
    sheetXml = patchRow(sheetXml, row, totals);
    console.log(`  ${dateStr} (row ${row}): clients=${totals.clients} totalSales=${totals.totalSales} grandTotal=${totals.grandTotal}`);
    written++;
  }

  if (written === 0) {
    console.log("No locked days to sync. Nothing written.");
    return;
  }

  sheetEntry.data = Buffer.from(sheetXml, "utf8");
  await writeZipEntries(REPORT_PATH, entries);
  console.log(`Done. Wrote ${written} day(s) into "${monthLabel}".`);
  if (skipped.length > 0) console.log(`Skipped (not yet confirmed/locked): ${skipped.join(", ")}`);
}

main().catch((e) => {
  console.error("FAILED:", e.message);
  process.exitCode = 1;
});
