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
import { TICKET_PACKAGE_PRICES, NEW_CLIENT_DISCOUNT } from "../src/lib/ticketPrices.js";

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
// Every new client gets $20 off list, ticket packages included. A different price is most
// likely the $20 forgotten, so the office can check it with the staff member.
const priceCheck = (ticket, amount) => {
  const list = TICKET_PACKAGE_PRICES[ticket]?.service;
  if (!list) return "(定価の表にないチケット)";
  const off = r2(list - amount);
  if (off === NEW_CLIENT_DISCOUNT) return `✅ 定価$${list} − $${NEW_CLIENT_DISCOUNT}`;
  if (off === 0) return `⚠️ 定価$${list}のまま($${NEW_CLIENT_DISCOUNT}オフし忘れ?)`;
  return `⚠️ 定価$${list}から$${off}引き($${NEW_CLIENT_DISCOUNT}オフと違う)`;
};

// Split one sale's pot by minutes, same rounding as src/lib/ticketIncentive.js (any leftover
// cent goes to the larger share).
const splitPot = (pot, mins) => {
  const total = mins.reduce((s, m) => s + m, 0);
  const amounts = mins.map(m => r2(pot * m / total));
  const diff = r2(pot - amounts.reduce((s, a) => s + a, 0));
  if (diff) { const i = mins.indexOf(Math.max(...mins)); amounts[i] = r2(amounts[i] + diff); }
  return amounts;
};

// Course → minutes when a second staff member runs the machine (null = no machine in the course).
const COURSES = [
  { key: "IP-60", name: "Improving Posture 60分(リンパ60分も同じ)", body: 60, machine: null },
  { key: "IP-75", name: "Improving Posture 75分(リンパ75分も同じ)", body: 75, machine: null },
  { key: "IP-90", name: "Improving Posture 90分(リンパ90分も同じ)", body: 75, machine: 15 },
  { key: "IP-120", name: "Improving Posture 120分(リンパ120分も同じ)", body: 105, machine: 15 },
  { key: "WL-75", name: "Weight Loss 75分", body: 35, machine: 40 },
  { key: "WL-90", name: "Weight Loss 90分", body: 50, machine: 40 },
  { key: "WL-115", name: "Weight Loss 115分", body: 75, machine: 40 },
  { key: "SF-45", name: "Sculpted Facial 45分", body: 45, machine: null },
  { key: "SF-60", name: "Sculpted Facial 60分", body: 60, machine: null },
  { key: "MC-60", name: "ProCell Microchanneling 60分", body: 60, machine: null },
];

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
ws.columns = [{ width: 8 }, { width: 22 }, { width: 6 }, { width: 11 }, { width: 12 }, { width: 11 }, { width: 10 }, { width: 8 }, { width: 8 }, { width: 11 }, { width: 46 }, { width: 34 }];
ws.addRow([title]).font = { bold: true, size: 14 };
ws.addRow(["新規のお客様(NL/NT)がチケットを買ったとき、チップを除いたチケット代の5%を、その日にそのお客様を担当した時間(分)の割合で分けます。"]);
ws.addRow(["コースごとの金額と割合は「コース別早見表」のシートにあります。「価格チェック」に⚠️が出ているものは、$20オフのし忘れの可能性があるので、スタッフに確認してください。"]);
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
head(ws.addRow(["日付", "お客様", "種類", "チケット", "チケット代(チップ除く)", "5%", "スタッフ", "分", "割合", "インセンティブ", "担当した内容", "価格チェック(新規は$20オフ)"]));
for (const s of sales) {
  s.staff.forEach((x, i) => {
    const check = priceCheck(s.ticket, s.amount);
    const row = ws.addRow(i === 0 ? [fmtDate(s.date), s.client, s.type, s.ticket, s.amount, s.pot, x.t, x.m, x.share, x.amount, x.what, check] : ["", "", "", "", "", "", x.t, x.m, x.share, x.amount, x.what, ""]);
    row.getCell(5).numFmt = money; row.getCell(6).numFmt = money; row.getCell(9).numFmt = "0.0%"; row.getCell(10).numFmt = money;
    if (i === 0 && check.startsWith("⚠️")) row.getCell(12).font = { color: { argb: "FFC00000" }, bold: true };
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

// Reference table: every ticket course at the new-client price, alone vs. split with a machine
// therapist, so the office can check any line in the summary at a glance.
{
  const ref = wb.addWorksheet("コース別早見表");
  ref.columns = [{ width: 40 }, { width: 6 }, { width: 10 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 22 }, { width: 22 }];
  ref.addRow(["コース別 早見表:新規チケットのインセンティブ(5%)"]).font = { bold: true, size: 14 };
  for (const line of [
    `・新規のお客様は、ツーリスト・ローカル関係なく全員$${NEW_CLIENT_DISCOUNT}オフです。チケット代は「新規価格(定価−$${NEW_CLIENT_DISCOUNT})」になるはずです。`,
    "・実際のチケット代がこの表の新規価格と違うときは、スタッフが$20オフをし忘れている可能性があります。そのときはスタッフに確認してください。インセンティブは実際のチケット代の5%で計算しています。",
    "・チップは含めません。チケットの施術代だけの5%です。",
    "・ボディとマシンを1人で担当した場合や、マシンを使わなかった場合は、そのスタッフが100%です。",
    "・Weight Lossのマシンはいつも40分です。Improving Posture・リンパの90分と120分は、マシンが15分です。",
    "・同じ日にほかのスタッフのフェイシャルなども受けた場合は、全員の分数を合計した割合で分けます。",
    "  例:IP 90分(ボディ75分+マシン15分)+フェイシャル60分=合計150分 → ボディ50%/マシン10%/フェイシャル40%",
  ]) ref.addRow([line]);
  ref.addRow([]);
  head(ref.addRow(["コース", "回数", "定価", `新規価格(−$${NEW_CLIENT_DISCOUNT})`, "インセンティブ5%", "1人で担当(100%)", "2人のとき ボディの人", "2人のとき マシンの人"]));
  for (const c of COURSES) for (const times of [3, 5]) {
    const list = TICKET_PACKAGE_PRICES[`${c.key}-${times}`]?.service;
    if (!list) continue;
    const price = list - NEW_CLIENT_DISCOUNT;
    const pot = r2(price * 0.05);
    let bodyCell = "(マシンなし)", machineCell = "(マシンなし)";
    if (c.machine) {
      const [b, m] = splitPot(pot, [c.body, c.machine]);
      const total = c.body + c.machine;
      bodyCell = `$${b.toFixed(2)}(${c.body}分・${(c.body / total * 100).toFixed(1)}%)`;
      machineCell = `$${m.toFixed(2)}(${c.machine}分・${(c.machine / total * 100).toFixed(1)}%)`;
    }
    const row = ref.addRow([c.name, `${times}回`, list, price, pot, pot, bodyCell, machineCell]);
    [3, 4, 5, 6].forEach(i => { row.getCell(i).numFmt = money; });
    if (times === 5) row.eachCell({ includeEmpty: true }, cell => { cell.border = { bottom: { style: "hair" } }; });
  }
}

for (const t of Object.keys(totals).sort()) {
  const s2 = wb.addWorksheet(t);
  s2.columns = [{ width: 8 }, { width: 22 }, { width: 11 }, { width: 12 }, { width: 30 }, { width: 8 }, { width: 8 }, { width: 11 }, { width: 46 }];
  s2.addRow([`${t}さん ${title}`]).font = { bold: true, size: 13 };
  s2.addRow([]);
  head(s2.addRow(["日付", "お客様", "チケット", "チケット代", "一緒に担当した人", "分", "割合", "インセンティブ", `${t}さんが担当した内容`]));
  for (const s of sales) for (const x of s.staff.filter(x => x.t === t)) {
    // "with Mikiko(マシン 15分)" so the office can find the other half of a split sale.
    const others = s.staff.filter(y => y.t !== t).map(y => `with ${y.t}(${y.what})`).join("、") || "(1人で担当・100%)";
    const row = s2.addRow([fmtDate(s.date), s.client, s.ticket, s.amount, others, `${x.m}/${s.total}`, x.share, x.amount, x.what]);
    row.getCell(4).numFmt = money; row.getCell(7).numFmt = "0.0%"; row.getCell(8).numFmt = money;
  }
  const tr = s2.addRow(["", "", "", "", "", "", "合計", totals[t].amount]); tr.font = { bold: true }; tr.getCell(8).numFmt = money;
}

const file = path.join(REPO_ROOT, "..", `チケットインセンティブ_${start}_${end}.xlsx`);
await wb.xlsx.writeFile(file);
console.log("SALES"); for (const s of sales) console.log(fmtDate(s.date), s.client, s.type, s.ticket, "$" + s.amount, "5%=$" + s.pot, s.staff.map(x => `${x.t} ${x.m}/${s.total}分 $${x.amount}`).join(" | "), "|", priceCheck(s.ticket, s.amount));
console.log("TOTALS", JSON.stringify(totals));
console.log("EXCLUDED"); for (const e of excluded) console.log(fmtDate(e.date), e.client, e.type, e.ticket, "$" + e.amount, e.note);
console.log("Saved:", file);
