// Builds the "Dr.Body 売上報告" email (subject + plain-text body) that's drafted into Gmail
// automatically whenever a day is locked (see createSalesReportDraft in spa-daily-sheet.jsx).
// Same per-day numbers as the monthly Excel export and the Desktop Sales Report, since all
// three go through computeDayTotals — the format mirrors the reports the owner used to write
// by hand (one day's breakdown, then the month-to-date 累計, then ※ notes).

import { computeDayTotals } from "./reportTotals.js";

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
const RULE = "──────────────────────────────────────";

const money = (n) => "$" + (Math.round(Number(n || 0) * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const mdLabel = (dateStr) => {
  const [, m, d] = dateStr.split("-").map(Number);
  return `${m}/${d}`;
};
const weekday = (dateStr) => WEEKDAYS[new Date(`${dateStr}T00:00:00Z`).getUTCDay()];

// Gift card balance used today — same rule as the daily sheet's totalGCUsed (regular visits,
// same-day ticket purchases, and add-ons paid today). Excluded from the sales total because it
// was already counted on the day the gift card was bought.
function giftCardUsedItems(data) {
  const appts = (data.appointments || []).filter(a => !a.isCavSlot);
  const items = [];
  appts.forEach(a => {
    const isRegular = !a.isTicket && !a.isGiftCard && !a.isPromo;
    const isSameDay = a.isTicket && a.isSameDayTicket;
    let amount = (isRegular || isSameDay) ? Number(a.giftCardUsed || 0) : 0;
    (a.addons || []).filter(ad => ad.countsAsRevenue === true).forEach(ad => { amount += Number(ad.giftCardUsed || 0); });
    if (amount > 0) items.push({ clientName: a.clientName || "(名前なし)", amount });
  });
  return items;
}

function dayBlock(totals) {
  return [
    `現金: 施術 ${money(totals.cashTreatment)} / 物販 ${money(totals.cashProduct)} / チップ ${money(totals.cashTip)}`,
    `カード: 施術 ${money(totals.cardTreatment)} / 物販 ${money(totals.cardProduct)} / チップ ${money(totals.cardTip)}`,
    `売上 ${money(totals.totalSales)} / チップ ${money(totals.totalTip)} / 売上+チップ ${money(totals.grandTotal)}`,
  ];
}

// `date` is the day just locked (YYYY-MM-DD); `monthDays` is {dateStr: dayData} for the 1st of
// that month through `date` (as returned by /api/day-data-range).
export function buildSalesReportEmail(date, monthDays) {
  const [y, m, d] = date.split("-").map(Number);
  const dayData = monthDays[date];
  const day = computeDayTotals(dayData);

  const mtd = { clients: 0, cashTreatment: 0, cashProduct: 0, totalCash: 0, cashTip: 0, cardTreatment: 0, cardProduct: 0, totalCard: 0, cardTip: 0, totalTip: 0, totalSales: 0, grandTotal: 0 };
  const unlockedDays = [];
  Object.keys(monthDays).filter(ds => ds <= date).sort().forEach(ds => {
    const t = computeDayTotals(monthDays[ds]);
    Object.keys(mtd).forEach(k => { mtd[k] += Number(t[k] || 0); });
    if (!monthDays[ds].locked && ds !== date) unlockedDays.push(ds);
  });

  const lines = [
    "Dr.Body, Inc. 売上報告",
    `日付: ${m}/${d}/${y} (${weekday(date)})`,
    RULE,
    "",
    `【本日】 客数 ${Number(day.clients || 0)}人`,
    ...dayBlock(day),
    "",
    RULE,
    `【${m}月 累計 (${m}/1〜${m}/${d})】`,
    `来客数: ${mtd.clients} 人`,
    "",
    "【現金】",
    `施術売上: ${money(mtd.cashTreatment)}`,
    `物販売上: ${money(mtd.cashProduct)}`,
    `現金合計: ${money(mtd.totalCash)}`,
    `チップ: ${money(mtd.cashTip)}`,
    "",
    "【クレジットカード】",
    `施術売上: ${money(mtd.cardTreatment)}`,
    `物販売上: ${money(mtd.cardProduct)}`,
    `カード合計: ${money(mtd.totalCard)}`,
    `チップ: ${money(mtd.cardTip)}`,
    "",
    RULE,
    `チップ合計: ${money(mtd.totalTip)}`,
    `総売上 (施術+物販): ${money(mtd.totalSales)}`,
    `総売上 + チップ: ${money(mtd.grandTotal)}`,
    RULE,
    "",
  ];

  const gcItems = giftCardUsedItems(dayData);
  const gcTotal = gcItems.reduce((s, it) => s + it.amount, 0);
  if (gcTotal > 0) {
    const who = gcItems.map(it => `${it.clientName}様 ${money(it.amount)}`).join("、");
    lines.push(`※ギフトカード使用分: 本日のギフトカード使用額は ${money(gcTotal)} です（${who}）。ギフトカード分は購入日に売上計上済みのため上記の売上には含まれていません。上記の売上に ${money(gcTotal)} を足すと、SquareのSales Reportの合計と一致します。`);
  } else {
    lines.push("※ギフトカード使用分: 本日のギフトカード使用額は $0.00 です。そのため、上記の売上はそのままSquareのSales Reportの合計と一致するはずです。");
  }

  const gcPurchases = (dayData.deposits || []).filter(dep => dep.type === "giftcard" && Number(dep.amount || 0) > 0);
  if (gcPurchases.length > 0) {
    const list = gcPurchases.map(dep => `${money(dep.amount)}（${dep.paymentType === "cash" ? "現金" : "カード"}）`).join("、");
    lines.push("", `※本日のギフトカード購入（${list}）は、上記の施術売上に含まれています。`);
  }

  if (unlockedDays.length > 0) {
    lines.push("", `※累計には、まだ確定（ロック）されていない日の数字も含まれています: ${unlockedDays.map(mdLabel).join("、")}`);
  }

  lines.push("", `※添付のExcelは、アプリの「Sales Report Excel Export」と同じ${m}月分のレポートです（${m}/${d}時点）。`);

  // The day's own money figures — re-locking a day only drafts a new email when this changes
  // (see createSalesReportDraftIfChanged in api/_lib/gmailDraft.js). Rounded to cents so a
  // floating-point artifact (e.g. 289.68000000000001) can't look like a change.
  const cents = (n) => Math.round(Number(n || 0) * 100);
  const fingerprint = JSON.stringify({
    totals: ["clients", "cashTreatment", "cashProduct", "cashTip", "cardTreatment", "cardProduct", "cardTip"].map(k => cents(day[k])),
    gcUsed: cents(gcTotal),
    gcPurchases: gcPurchases.map(dep => [cents(dep.amount), dep.paymentType]),
  });

  return {
    subject: `Dr.Body 売上報告 ${m}/${d}/${y}`,
    body: lines.join("\n"),
    fingerprint,
  };
}
