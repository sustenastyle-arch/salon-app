// New-client ticket sale incentive Excel for the payroll office (ゆきさん). The rule and the
// math live in src/lib/ticketIncentive.js, shared with the app's Payroll tab (which shows the
// same amounts on each staff member's own sheet), so the two always agree.
//
//   node scripts/ticket-incentive.mjs 2026-11-01 2026-11-15
//
// Writes to the Desktop: one sheet with the staff totals, every sale's split and the repeat
// clients' sales that didn't count, plus one sheet per staff member with only their lines.
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import ExcelJS from "exceljs";
import { computeTicketIncentives } from "../src/lib/ticketIncentive.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..");
const env = Object.fromEntries(fs.readFileSync(path.join(REPO_ROOT, ".env"), "utf8").split(/\r?\n/)
  .filter(l => l.includes("=") && !l.trim().startsWith("#"))
  .map(l => { const i = l.indexOf("="); return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^"|"$/g, "")]; }));

const [start, end] = process.argv.slice(2);
if (!start || !end) { console.error("usage: node scripts/ticket-incentive.mjs YYYY-MM-DD YYYY-MM-DD"); process.exit(1); }
const ymd = (d) => d.toISOString().slice(0, 10);
const r2 = (x) => Math.round(x * 100) / 100;
const norm = (s) => (s || "").toLowerCase().trim().replace(/\s+/g, " ");

// History from 2026-07-01 (start of real use), to say when a repeat client first came.
const dates = [];
for (let d = new Date("2026-07-01T12:00:00Z"); d <= new Date(end + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + 1)) dates.push(ymd(d));
const res = await (await fetch(env.KV_REST_API_URL, { method: "POST", headers: { Authorization: `Bearer ${env.KV_REST_API_TOKEN}` }, body: JSON.stringify(["MGET", ...dates.map(d => "spa-sheet-" + d)]) })).json();
const days = {};
dates.forEach((d, i) => { if (res.result[i]) days[d] = JSON.parse(res.result[i]); });
const firstVisit = {};
for (const d of dates) for (const a of days[d]?.appointments || []) if (!a.isCavSlot && a.clientName) firstVisit[norm(a.clientName)] ??= d;

const fmtDate = (d) => `${+d.slice(5, 7)}/${+d.slice(8, 10)}`;
const partJa = (p) => ({ body: `ボディ(${p.name})`, machine: "マシン", service: p.name, addon: `追加 ${p.name}`, addonMachine: `追加 ${p.name} マシン` }[p.kind]) + ` ${p.mins}分`;

const result = computeTicketIncentives(days, start, end);
const sales = result.sales.map(s => ({
  date: s.date, client: s.client, type: s.customerType, ticket: s.ticket, amount: s.amount, pot: s.pot, total: s.totalMins,
  staff: s.staff.map(x => ({ t: x.t, m: x.mins, share: x.share, amount: x.amount, what: x.parts.map(partJa).join(" + ") })),
}));
const excluded = result.excluded.map(e => {
  const fv = firstVisit[norm(e.client)];
  return { ...e, type: e.customerType, note: fv === e.date ? "リピーター(7月以降の記録ではこの日が初来店。オーナー確認済みならこのまま)" : `リピーター(初来店 ${fmtDate(fv)})` };
});
// ---- Excel ----
const wb = new ExcelJS.Workbook();
const money = '"$"#,##0.00';
const head = (row) => { row.font = { bold: true }; row.eachCell(c => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F0E8" } }; c.border = { bottom: { style: "thin" } }; }); };
const title = `新規チケット販売インセンティブ(5%) ${fmtDate(start)}〜${fmtDate(end)}`;
const totals = {};
for (const s of sales) for (const x of s.staff) {
  const t = (totals[x.t] ||= { amount: 0, count: 0 });
  t.amount = r2(t.amount + x.amount); t.count++;
}

const ws = wb.addWorksheet("まとめ");
ws.columns = [{ width: 8 }, { width: 22 }, { width: 6 }, { width: 11 }, { width: 12 }, { width: 11 }, { width: 10 }, { width: 8 }, { width: 8 }, { width: 11 }, { width: 46 }];
ws.addRow([title]).font = { bold: true, size: 14 };
ws.addRow(["新規のお客様(NL/NT)がチケットを買ったとき、チップを除いたチケット代の5%を、その日にそのお客様を担当した時間(分)の割合で分けます。"]);
ws.addRow([]);
ws.addRow(["■ スタッフ別 合計(お給料に足す金額)"]).font = { bold: true, size: 12 };
head(ws.addRow(["スタッフ", "インセンティブ", "件数"]));
for (const [t, v] of Object.entries(totals).sort((a, b) => b[1].amount - a[1].amount)) {
  const row = ws.addRow([t, v.amount, v.count]); row.getCell(2).numFmt = money;
}
const sumRow = ws.addRow(["合計", r2(Object.values(totals).reduce((s, v) => s + v.amount, 0)), sales.length + "件の販売"]);
sumRow.font = { bold: true }; sumRow.getCell(2).numFmt = money;
ws.addRow([]);
ws.addRow(["■ 販売ごとの明細"]).font = { bold: true, size: 12 };
head(ws.addRow(["日付", "お客様", "種類", "チケット", "チケット代(チップ除く)", "5%", "スタッフ", "分", "割合", "インセンティブ", "担当した内容"]));
for (const s of sales) {
  s.staff.forEach((x, i) => {
    const row = ws.addRow(i === 0 ? [fmtDate(s.date), s.client, s.type, s.ticket, s.amount, s.pot, x.t, x.m, x.share, x.amount, x.what] : ["", "", "", "", "", "", x.t, x.m, x.share, x.amount, x.what]);
    row.getCell(5).numFmt = money; row.getCell(6).numFmt = money; row.getCell(9).numFmt = "0.0%"; row.getCell(10).numFmt = money;
    if (i === s.staff.length - 1) row.eachCell({ includeEmpty: true }, c => { c.border = { bottom: { style: "hair" } }; });
  });
}
ws.addRow([]);
ws.addRow(["■ 対象外(リピーターのお客様のチケット購入)"]).font = { bold: true, size: 12 };
head(ws.addRow(["日付", "お客様", "種類", "チケット", "チケット代(チップ除く)", "", "", "", "", "", "メモ"]));
for (const e of excluded) {
  const row = ws.addRow([fmtDate(e.date), e.client, e.type, e.ticket, e.amount, "", "", "", "", "", e.note]);
  row.getCell(5).numFmt = money;
  if (e.note.startsWith("⚠️")) row.getCell(11).font = { color: { argb: "FFC00000" }, bold: true };
}

for (const t of Object.keys(totals).sort()) {
  const s2 = wb.addWorksheet(t);
  s2.columns = [{ width: 8 }, { width: 22 }, { width: 11 }, { width: 12 }, { width: 8 }, { width: 8 }, { width: 11 }, { width: 46 }];
  s2.addRow([`${t}さん ${title}`]).font = { bold: true, size: 13 };
  s2.addRow([]);
  head(s2.addRow(["日付", "お客様", "チケット", "チケット代", "分", "割合", "インセンティブ", "担当した内容"]));
  for (const s of sales) for (const x of s.staff.filter(x => x.t === t)) {
    const row = s2.addRow([fmtDate(s.date), s.client, s.ticket, s.amount, `${x.m}/${s.total}`, x.share, x.amount, x.what]);
    row.getCell(4).numFmt = money; row.getCell(6).numFmt = "0.0%"; row.getCell(7).numFmt = money;
  }
  const tr = s2.addRow(["", "", "", "", "", "合計", totals[t].amount]); tr.font = { bold: true }; tr.getCell(7).numFmt = money;
}

const file = path.join(REPO_ROOT, "..", `チケットインセンティブ_${start}_${end}.xlsx`);
await wb.xlsx.writeFile(file);
console.log("SALES"); for (const s of sales) console.log(fmtDate(s.date), s.client, s.type, s.ticket, "$" + s.amount, "5%=$" + s.pot, s.staff.map(x => `${x.t} ${x.m}/${s.total}分 $${x.amount}`).join(" | "));
console.log("TOTALS", JSON.stringify(totals));
console.log("EXCLUDED"); for (const e of excluded) console.log(fmtDate(e.date), e.client, e.type, e.ticket, "$" + e.amount, e.note);
console.log("Saved:", file);
