// Payroll audit for one pay period (1st–15th or 16th–end of month), run automatically by
// Windows Task Scheduler at 9pm on the 15th and the last day of the month (see payroll-audit.bat and CLAUDE.md
// "Payroll audit"). Every check here comes from a real mistake found in a past audit, so
// don't drop one as "never fires" — it never fires because it's being watched.
//
// Read-only: it only reads Redis day blobs and Square, never writes either. The report goes
// to a text file on the Desktop (next to the repo folder) and is opened in Notepad.
//
//   node scripts/payroll-audit.mjs                 # period that just ended (by today's date)
//   node scripts/payroll-audit.mjs 2026-11-01 2026-11-15
//   node scripts/payroll-audit.mjs ... --no-open   # don't open Notepad
import fs from "fs";
import path from "path";
import { spawn } from "child_process";
import { fileURLToPath } from "url";
import { computeDayTotals, getRetailItems, getStaffPurchaseItems } from "../src/lib/reportTotals.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..");
const OUT_DIR = path.resolve(REPO_ROOT, ".."); // the Desktop the repo folder lives on

const env = Object.fromEntries(fs.readFileSync(path.join(REPO_ROOT, ".env"), "utf8").split(/\r?\n/)
  .filter(l => l.includes("=") && !l.trim().startsWith("#"))
  .map(l => { const i = l.indexOf("="); return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^"|"$/g, "")]; }));

// Keep in sync with THERAPISTS / CAV_CAPABLE / BODY_CAPABLE / DUAL_LICENSE in spa-daily-sheet.jsx.
const THERAPISTS = ["Mami", "Aya", "Megumi", "Hiromi", "Maki", "Yuka", "Mai", "Betsy", "Mikiko", "Saori"];
const CAV = ["Mami", "Betsy", "Megumi", "Yuka", "Mai", "Mikiko"];
const BODY = ["Mami", "Betsy", "Megumi", "Aya", "Hiromi", "Mai", "Maki", "Mikiko", "Saori"];
const DUAL = ["Mami", "Betsy", "Megumi", "Mai", "Mikiko"];
// Square categories whose sales are retail (物販) or staff purchases (社販).
const RETAIL_CATEGORIES = ["NX5QDWW5G7QD35ZGZ26H5WA6", "5T5WCNAUETAZX7RMLXNJPMAL"];

const n = (x) => Number(x || 0);
const r2 = (x) => Math.round(x * 100) / 100;
const $ = (x) => `$${r2(x).toFixed(2)}`;

// ---- period ----
const args = process.argv.slice(2).filter(a => !a.startsWith("--"));
const noOpen = process.argv.includes("--no-open");
const ymd = (d) => d.toISOString().slice(0, 10);
const hawaiiToday = new Date(Date.now() - 10 * 3600e3);
let start, end;
if (args.length === 2) [start, end] = args;
else {
  // Scheduled for 9pm on the 15th and on the last day of the month, but a missed run fires at
  // the next start-up, so a late run (e.g. the 16th or the 1st) still audits the period just paid.
  const y = hawaiiToday.getUTCFullYear(), m = hawaiiToday.getUTCMonth(), day = hawaiiToday.getUTCDate();
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  if (day === lastDay) { start = ymd(new Date(Date.UTC(y, m, 16))); end = ymd(new Date(Date.UTC(y, m, lastDay))); }
  else if (day >= 15) { start = ymd(new Date(Date.UTC(y, m, 1))); end = ymd(new Date(Date.UTC(y, m, 15))); }
  else { start = ymd(new Date(Date.UTC(y, m - 1, 16))); end = ymd(new Date(Date.UTC(y, m, 0))); }
}
const dates = [];
for (let d = new Date(start + "T12:00:00Z"); d <= new Date(end + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + 1)) dates.push(ymd(d));

// ---- data ----
const redis = async (cmd) => {
  const r = await fetch(env.KV_REST_API_URL, { method: "POST", headers: { Authorization: `Bearer ${env.KV_REST_API_TOKEN}` }, body: JSON.stringify(cmd) });
  return (await r.json()).result;
};
// Also load the 30 days before the period, to catch a card copied in from an earlier day.
const lookback = [];
for (let d = new Date(start + "T12:00:00Z"), i = 0; i < 30; i++) { d.setUTCDate(d.getUTCDate() - 1); lookback.push(ymd(d)); }
const allDates = [...lookback.reverse(), ...dates];
const raw = await redis(["MGET", ...allDates.map(d => "spa-sheet-" + d)]);
const days = {};
allDates.forEach((d, i) => { if (raw[i]) days[d] = JSON.parse(raw[i]); });

const SQ = "https://connect.squareup.com/v2";
const sqH = { Authorization: `Bearer ${env.SQUARE_ACCESS_TOKEN}`, "Square-Version": "2024-10-17", "Content-Type": "application/json" };
const sq = async (p, body) => (await fetch(SQ + p, { method: body ? "POST" : "GET", headers: sqH, body: body && JSON.stringify(body) })).json();

// variation id → is it a retail/staff-purchase product?
const retailVariations = new Set();
{
  let cursor;
  do {
    const r = await sq(`/catalog/list?types=ITEM${cursor ? "&cursor=" + cursor : ""}`);
    for (const it of r.objects || []) {
      const cats = [it.item_data?.category_id, ...(it.item_data?.categories || []).map(c => c.id)].filter(Boolean);
      if (cats.some(c => RETAIL_CATEGORIES.includes(c))) (it.item_data.variations || []).forEach(v => retailVariations.add(v.id));
    }
    cursor = r.cursor;
  } while (cursor);
}

// Same math as api/square-payments.js (tip-prompt + "Tip" line items, refunds netted, gift
// card tenders separated) so the comparison matches what the 🔍 Square Reconcile panel shows.
async function squareDay(date) {
  const next = new Date(`${date}T00:00:00Z`); next.setUTCDate(next.getUTCDate() + 1);
  const orders = {};
  const getOrder = async (id) => (orders[id] ??= (await sq(`/orders/${id}`)).order || null);
  const tipLineDone = new Set();
  const out = { cashTotal: 0, cardTotal: 0, cashTip: 0, cardTip: 0, gcTotal: 0, gcTip: 0, retail: 0 };
  let cursor;
  do {
    const q = new URLSearchParams({ location_id: env.SQUARE_LOCATION_ID, begin_time: `${date}T00:00:00-10:00`, end_time: `${ymd(next)}T00:00:00-10:00`, limit: "100" });
    if (cursor) q.set("cursor", cursor);
    const r = await sq(`/payments?${q}`);
    for (const p of r.payments || []) {
      if (p.status !== "COMPLETED") continue;
      const net = Math.max(0, n(p.total_money?.amount) / 100 - n(p.refunded_money?.amount) / 100);
      let tip = n(p.tip_money?.amount) / 100;
      if (p.order_id && !tipLineDone.has(p.order_id)) {
        const o = await getOrder(p.order_id);
        const owner = (o?.tenders || []).find(t => t.type !== "SQUARE_GIFT_CARD") || (o?.tenders || [])[0];
        if (!owner || owner.id === p.id) {
          tipLineDone.add(p.order_id);
          tip += (o?.line_items || []).filter(li => (li.name || "").trim().toLowerCase() === "tip").reduce((s, li) => s + n(li.gross_sales_money?.amount) / 100, 0);
          if (net > 0) out.retail += (o?.line_items || []).filter(li => retailVariations.has(li.catalog_object_id)).reduce((s, li) => s + n(li.total_money?.amount) / 100, 0);
        }
      }
      if (net <= 0) tip = 0;
      if (p.source_type === "CASH") { out.cashTotal += net; out.cashTip += tip; } else { out.cardTotal += net; out.cardTip += tip; }
      if (p.card_details?.card?.card_brand === "SQUARE_GIFT_CARD") { out.gcTotal += net; out.gcTip += tip; }
    }
    cursor = r.cursor;
  } while (cursor);
  return out;
}

// ---- checks ----
const sections = {};
const add = (key, line) => (sections[key] ||= []).push(line);
const S = {
  sq: "1. Squareと合っていない日",
  lock: "2. まだ確定していない日",
  off: "3. 出勤にチェックがない人にお金が入っている",
  dupId: "4. 同じカードが2つの日に入っている(2重払いの可能性)",
  dupDay: "5. 同じ日・同じ担当・同じお客様のカードが2枚",
  dupTicket: "6. 同じお客様・同じチケット・同じ回数が2回",
  maki: "7. Makiさんの分数がルール・Square予約と違う(お給料に直接影響)",
  mins: "8. ほかのスタッフの分数がルールと違う",
  noMachine: "9. マシンが必要なのにマシン担当がいない",
  slot: "10. マシン担当の記録がおかしい",
  license: "11. ライセンスのない施術",
  retailSq: "12. Squareの物販とシートの物販が合わない",
  retailSeller: "13. 物販の担当・分け方を確認",
  empty: "14. 金額が$0のカード",
  tip: "15. チップが大きい・おかしい",
  forgot: "16. 忘れチップの記録に情報がない",
};

const periodDays = dates.filter(d => days[d]);
let cardCount = 0;
const idSeen = {};
for (const d of allDates) for (const a of days[d]?.appointments || []) (idSeen[a.id] ||= []).push(d);
const ticketSeen = {};

for (const d of periodDays) {
  const day = days[d];
  const appts = day.appointments || [];
  if (!appts.length && !(day.retails || []).length && !(day.ticketPurchases || []).length) continue;
  cardCount += appts.length;
  const ws = Array.isArray(day.workingStaff) && day.workingStaff.length ? day.workingStaff : null;
  const md = d.slice(5).replace("-", "/");
  if (!day.locked) add(S.lock, `${md}`);

  // Square
  const t = computeDayTotals(day);
  const sheet = { cash: r2(n(t.totalCash) + n(t.cashTip)), card: r2(n(t.totalCard) + n(t.cardTip)), cashTip: r2(n(t.cashTip)), cardTip: r2(n(t.cardTip)) };
  const s = await squareDay(d);
  const sqv = { cash: r2(s.cashTotal), card: r2(s.cardTotal - s.gcTotal), cashTip: r2(s.cashTip), cardTip: r2(s.cardTip - s.gcTip) };
  const labels = { cash: "現金合計", card: "カード合計", cashTip: "現金チップ", cardTip: "カードチップ" };
  const diffs = Object.keys(labels).filter(k => Math.abs(sheet[k] - sqv[k]) > 0.01).map(k => `${labels[k]} シート${$(sheet[k])} / Square${$(sqv[k])}`);
  if (diffs.length) add(S.sq, `${md}  ${diffs.join("、")}`);

  // Retail: Square retail line items vs sheet retail + 社販 (both tax-included)
  const sheetRetail = r2(appts.filter(a => !a.isCavSlot).flatMap(getRetailItems).reduce((x, it) => x + n(it.amount), 0)
    + (day.retails || []).reduce((x, rt) => x + n(rt.price), 0)
    + (day.staffPurchases || []).flatMap(getStaffPurchaseItems).reduce((x, it) => x + n(it.amount), 0));
  if (Math.abs(sheetRetail - r2(s.retail)) > 1) add(S.retailSq, `${md}  シート${$(sheetRetail)} / Square${$(s.retail)}(打ち忘れか、Squareで別の商品・キーパッドで打った可能性)`);

  const byId = Object.fromEntries(appts.map(a => [a.id, a]));
  const hasMoney = (a) => n(a.price) || n(a.tip) || a.isTicket;
  const off = (who, what) => { if (ws && who && !ws.includes(who)) add(S.off, `${md} ${who}  ${what}(出勤: ${ws.join(", ")})`); };

  for (const a of appts) {
    const who = a.therapist;
    const label = `${a.startTime || ""} ${a.clientName || "(名前なし)"}`.trim();
    if (hasMoney(a)) off(who, `${a.isCavSlot ? "マシン " : ""}${label} チップ${$(n(a.tip))}`);
    for (const x of a.addons || []) {
      if (x.therapist && n(x.price) + n(x.tip) > 0) off(x.therapist, `追加 ${x.serviceName} (${a.clientName})`);
      if (x.hasCav && x.cavTherapist) off(x.cavTherapist, `追加のマシン ${x.serviceName} (${a.clientName})`);
    }
    if (n(a.extraTip) > 0) off(a.extraTipTherapist || who, `エキストラチップ ${label} ${$(n(a.extraTip))}`);
    if (who && !THERAPISTS.includes(who)) add(S.off, `${md} "${who}" はスタッフ一覧にない名前 ${label}(お給料に入らない)`);
    if (!who && (n(a.price) || n(a.tip))) add(S.off, `${md} 担当なし ${label} ${$(n(a.price))} チップ${$(n(a.tip))}(だれのお給料にも入らない)`);

    const dupDays = (idSeen[a.id] || []).filter(x => x !== d);
    if (dupDays.length && !a.isCavSlot) add(S.dupId, `${md} ${who} ${label} ← ${dupDays.map(x => x.slice(5).replace("-", "/")).join(", ")} にも同じカード`);

    if (a.isTicket && a.ticketMenu && a.ticketCurrent != null) {
      const k = `${(a.clientName || "").toLowerCase().trim()}|${a.ticketMenu}|${a.ticketCurrent}/${a.ticketTotal}|${a.isCavSlot ? "machine" : "body"}`;
      (ticketSeen[k] ||= []).push(`${md} ${who}`);
    }

    if (a.isCavSlot) {
      const p = byId[a.parentId];
      if (!CAV.includes(who)) add(S.license, `${md} ${who} がマシン ${label}(マシンのライセンスなし)`);
      if (!p) add(S.slot, `${md} マシン枠 ${who} ${label} の元のカードがない(消し忘れ?)チップ${$(n(a.tip))}`);
      else if (p.cavTherapist !== who) add(S.slot, `${md} ${label} マシン枠は${who}、元のカードは${p.cavTherapist || "(マシン担当なし)"}`);
      if (!a.isTicket && !a.isGiftCard && !a.isPromo && !n(a.price) && !n(a.tip) && p && !n(p.depositApplied)) add(S.slot, `${md} ${who} マシン ${label} の金額が$0`);
      continue;
    }

    // Body card. A 🧪 Staff Practice card isn't a client and its minutes aren't paid, so the
    // minutes / machine / empty-card checks don't apply.
    if (a.isStaffPractice) continue;
    const svc = a.serviceName || "";
    const nameAll = `${svc} ${a.ticketMenu || ""}`;
    const isWL = /weight\s*loss|\bWL-/i.test(nameAll), isIP = /posture|\bIP-/i.test(nameAll), isLy = /lymph/i.test(nameAll);
    if (who && !BODY.includes(who) && (isWL || isIP || isLy || /deep|lomi|couple/i.test(svc))) add(S.license, `${md} ${who} がボディ ${label} ${svc}(マッサージのライセンスなし)`);
    const cav = a.cavTherapist;
    if (cav && !appts.some(x => x.isCavSlot && x.parentId === a.id)) add(S.slot, `${md} ${who}+${cav} ${label} マシン担当が書いてあるのにマシン枠がない(マシンの人に入っていない)`);

    if (isWL || isIP || isLy) {
      const m = svc.match(/(\d+)\s*min/i) || (a.ticketMenu || "").match(/-(\d+)-/);
      if (m) {
        const L = +m[1], dur = n(a.duration);
        let expect = null, why = "";
        if (isWL) {
          if (cav) { expect = L - 40; why = `マシン${cav} 40分`; }
          else if (DUAL.includes(who)) { expect = L; why = "一人で全部"; }
          else add(S.noMachine, `${md} ${who} ${label} ${svc}(${who}さんはマッサージだけ。マシンはだれ?)`);
        } else {
          if (cav && L >= 90) { expect = L - 15; why = `マシン${cav} 15分`; }
          else if (cav) add(S.slot, `${md} ${label} ${svc} はマシンのないコースなのにマシン担当${cav}`);
          else { expect = L; why = who === "Maki" ? "マシンなし=Makiさん一人" : "マシンなし"; }
        }
        if (expect != null && dur !== expect) add(who === "Maki" ? S.maki : S.mins, `${md} ${who} ${label} ${svc}  ${dur}分 → 正しくは${expect}分(${why})`);
      }
    }

    const money = n(a.price) + n(a.tip) + n(a.newTicketAmount) + n(a.packagePrice) + (a.addons || []).reduce((x, y) => x + n(y.price) + n(y.tip), 0) + getRetailItems(a).reduce((x, it) => x + n(it.amount), 0);
    if (!money && !a.isTicket && !a.isGiftCard && !a.isPromo && !n(a.depositApplied) && !n(a.giftCardUsed)) add(S.empty, `${md} ${who} ${label} ${svc}(キャンセルなら削除、来店なら金額を入力)`);
    if (n(a.tip) < 0 || n(a.price) < 0) add(S.tip, `${md} ${who} ${label} マイナスの金額`);
    else if (n(a.tip) >= 80 || (n(a.price) > 0 && n(a.tip) > n(a.price) * 0.5)) add(S.tip, `${md} ${who} ${label} 料金${$(n(a.price))} チップ${$(n(a.tip))}`);

    // Retail sold during a visit
    for (const it of getRetailItems(a)) {
      const sellers = (it.sellers || []).filter(x => x.therapist && n(x.amount) > 0);
      if (cav && sellers.length <= 1) add(S.retailSeller, `${md} ${label} ${it.productName || "物販"} ${$(n(it.amount))}:${who}+${cav}の2人の施術だけど、${sellers.length ? sellers[0].therapist + "さんだけ" : who + "さんに全部"}(分けなくていい?)`);
      for (const x of sellers) if (ws && !ws.includes(x.therapist)) add(S.off, `${md} ${x.therapist} 物販 ${it.productName} (${a.clientName})`);
    }
  }

  // Same client twice for the same therapist on one day
  const seen = {};
  for (const a of appts.filter(a => !a.isCavSlot && a.clientName)) (seen[`${a.clientName.toLowerCase().trim()}|${a.therapist}`] ||= []).push(a);
  for (const arr of Object.values(seen)) if (arr.length > 1) add(S.dupDay, `${md} ${arr[0].therapist} ${arr[0].clientName}: ${arr.map(a => `${a.startTime} ${a.serviceName || ""} ${$(n(a.price))}/チップ${$(n(a.tip))}`).join(" | ")}`);

  // Standalone retail (phone/walk-in)
  for (const rt of day.retails || []) {
    const sellers = rt.sellers || (rt.soldBy ? [{ therapist: rt.soldBy, amount: 1 }] : []);
    if (!sellers.some(x => x.therapist && n(x.amount) > 0)) add(S.retailSeller, `${md} 物販(来店なし) ${rt.item || ""} ${$(n(rt.price))} の担当がいない(だれの歩合にも入らない)`);
    for (const x of sellers) if (ws && x.therapist && !ws.includes(x.therapist)) add(S.off, `${md} ${x.therapist} 物販(来店なし) ${rt.item}`);
  }
  for (const ft of day.forgottenTips || []) {
    if (!ft.clientName || !ft.originalDate) add(S.forgot, `${md} ${ft.therapist} ${$(n(ft.tipAmount) + n(ft.serviceAmount))}  お客様名・来店日${!ft.clientName && !ft.originalDate ? "が両方" : !ft.clientName ? "(名前)が" : "(来店日)が"}空欄`);
  }
}
// Maki is paid by the minute, so also hold her minutes against Square's booked minutes. A
// booking changed on the spot (e.g. booked 75, charged 90) is fine but should be a conscious
// check; a Maki booking with no Maki card on the sheet may be a treatment she wasn't paid for,
// or one logged under someone else.
const MAKI_TM = "TMwBMik2I8Xpyog2"; // Maki's Square team member id (TEAM_MEMBER_MAP)
{
  const booked = {};
  let cursor;
  const next = new Date(end + "T00:00:00Z"); next.setUTCDate(next.getUTCDate() + 1);
  do {
    const q = new URLSearchParams({ location_id: env.SQUARE_LOCATION_ID, start_at_min: `${start}T00:00:00-10:00`, start_at_max: `${ymd(next)}T00:00:00-10:00`, limit: "100" });
    if (cursor) q.set("cursor", cursor);
    const b = await sq(`/bookings?${q}`);
    for (const bk of b.bookings || []) {
      const mins = (bk.appointment_segments || []).filter(s => s.team_member_id === MAKI_TM).reduce((x, s) => x + n(s.duration_minutes), 0);
      if (mins && !/CANCEL|DECLINED|NO_SHOW/.test(bk.status)) booked[bk.id] = { mins, start: new Date(new Date(bk.start_at).getTime() - 10 * 3600e3).toISOString(), customerId: bk.customer_id };
    }
    cursor = b.cursor;
  } while (cursor);
  const onSheet = {}; // booking id → therapists on its sheet cards
  const makiClients = {};
  for (const d of periodDays) for (const a of days[d].appointments || []) {
    if (a.isCavSlot) continue;
    const bid = (a.id.match(/^sq-(.+)-\d+$/) || [])[1];
    if (bid) (onSheet[bid] ||= []).push(a.therapist);
    if (a.therapist !== "Maki") continue;
    (makiClients[d] ||= []).push((a.clientName || "").toLowerCase().trim());
    const bkd = bid && booked[bid];
    if (!bkd) continue;
    const slot = (days[d].appointments || []).find(x => x.isCavSlot && x.parentId === a.id);
    const sheetCourse = n(a.duration) + (a.cavTherapist ? n(slot?.duration) || 15 : 0);
    if (sheetCourse !== bkd.mins) add(S.maki, `${d.slice(5).replace("-", "/")} ${a.startTime} ${a.clientName} ${a.serviceName || ""}  シート${a.duration}分${a.cavTherapist ? `+マシン${a.cavTherapist}` : ""} / Square予約${bkd.mins}分(当日コース変更なら会計の内容を確認)`);
  }
  for (const [id, bkd] of Object.entries(booked)) {
    if ((onSheet[id] || []).includes("Maki")) continue;
    const d = bkd.start.slice(0, 10);
    let name = "";
    if (bkd.customerId) { const c = (await sq(`/customers/${bkd.customerId}`)).customer; name = `${c?.given_name || ""} ${c?.family_name || ""}`.trim(); }
    // A manual card for the same client counts as on the sheet (e.g. entered by hand at another time).
    if (name && (makiClients[d] || []).includes(name.toLowerCase())) continue;
    add(S.maki, `${d.slice(5).replace("-", "/")} ${bkd.start.slice(11, 16)} ${name || "(名前なし)"} Square予約はMakiさん${bkd.mins}分なのに、シートにMakiさんのカードがない(ほかの人の名前で入っている?)`);
  }
}

for (const [k, list] of Object.entries(ticketSeen)) if (list.length > 1) {
  const [client, menu, num, side] = k.split("|");
  add(S.dupTicket, `${client} ${menu} ${num}${side === "machine" ? "(マシン)" : ""} → ${list.join(" / ")}(同じ来店の2重か、回数の書き間違い)`);
}

// ---- report ----
const lines = [];
lines.push(`Dr.Body 給料チェック  ${start} 〜 ${end}`);
lines.push(`作成: ${hawaiiToday.toISOString().slice(0, 16).replace("T", " ")}(ハワイ時間)  対象 ${periodDays.length}日・カード${cardCount}枚`);
lines.push("");
const found = Object.values(S).filter(k => sections[k]);
if (!found.length) lines.push("✅ 問題は見つかりませんでした。このままお給料を出して大丈夫です。");
else {
  lines.push(`⚠️ ${found.length}種類の確認があります。お給料に関係するのは 1, 3〜7, 9, 10, 12, 13 です。`);
  lines.push("   8・14・15・16 はお給料に影響しないか、念のための確認です。");
}
for (const k of Object.values(S)) {
  lines.push("");
  lines.push(`■ ${k}  ${sections[k] ? `(${sections[k].length}件)` : "… OK"}`);
  for (const l of sections[k] || []) lines.push(`   ・${l}`);
}
lines.push("");
lines.push("分からないところは、このファイルをClaudeに見せてください。");
const text = lines.join("\r\n");
const file = path.join(OUT_DIR, `給料チェック_${start}_${end}.txt`);
fs.writeFileSync(file, "﻿" + text, "utf8");
console.log(text);
console.log(`\nSaved: ${file}`);
if (!noOpen) spawn("notepad.exe", [file], { detached: true, stdio: "ignore" }).unref();
