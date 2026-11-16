// New-client ticket sale incentive, shared by PayrollTab (each staff member's payroll rows and
// Excel) and scripts/ticket-incentive.mjs (the summary Excel for the payroll office), so the two
// can never disagree.
//
// Rule (owner, 2026-11-15): when a NEW client (customerType NL or NT, traveler or local) buys a
// ticket package, the staff get `rate` × the package price (after discounts, tip excluded). If
// several staff treated that client that day, it's split by the minutes each one spent on the
// client across all of that day's cards: body, machine, a facial by someone else, add-ons. A card
// with one staff member and no machine therapist means that person did the whole course alone.

// Add a campaign here to turn the incentive on for another period.
export const TICKET_INCENTIVE_CAMPAIGNS = [
  { start: "2026-11-01", end: "2026-11-30", rate: 0.05, label: "November 2026 new-client ticket incentive" },
];

const n = (x) => Number(x || 0);
const r2 = (x) => Math.round(x * 100) / 100;
const norm = (s) => (s || "").toLowerCase().trim().replace(/\s+/g, " ");
const isWeightLoss = (s) => /weight\s*loss/i.test(s || "");

export const campaignFor = (date) => TICKET_INCENTIVE_CAMPAIGNS.find(c => date >= c.start && date <= c.end) || null;
export const periodHasCampaign = (start, end) => TICKET_INCENTIVE_CAMPAIGNS.some(c => c.start <= end && c.end >= start);

// Ticket package sales on one appointment card: a same-day ticket (bought and first session used
// on the spot) and/or a new ticket bought for later visits.
export function ticketSalesOf(a) {
  const out = [];
  if (a.isCavSlot) return out;
  if (a.isSameDayTicket && n(a.packagePrice) > 0) out.push({ ticket: a.ticketMenu || "", amount: n(a.packagePrice), tip: n(a.packageTip ?? a.tip) });
  if ((a.purchaseTags || []).includes("newTicket") && n(a.newTicketAmount) > 0) out.push({ ticket: a.newTicketMenu || "", amount: n(a.newTicketAmount), tip: n(a.newTicketTip) });
  return out;
}

// Minutes each staff member spent on this client on this day. Each part is
// { kind: "body" | "machine" | "service" | "addon" | "addonMachine", name, mins }.
export function minutesByStaff(day, clientName) {
  const client = norm(clientName);
  const appts = day.appointments || [];
  const byId = Object.fromEntries(appts.map(a => [a.id, a]));
  const staff = {};
  const add = (t, part) => { if (t && part.mins > 0) (staff[t] ||= []).push(part); };
  for (const a of appts) {
    if (norm(a.clientName) !== client) continue;
    if (a.isCavSlot) {
      const parent = byId[a.parentId];
      // Weight Loss machine time is always 40 min (older machine slots were saved as 15).
      add(a.therapist, { kind: "machine", name: parent?.serviceName || "", mins: parent && isWeightLoss(parent.serviceName) ? 40 : n(a.duration) || 15 });
      continue;
    }
    const name = (a.serviceName || "").replace(/\s*[(（].*$/, "").trim();
    add(a.therapist, { kind: a.cavTherapist ? "body" : "service", name, mins: n(a.duration) });
    for (const x of a.addons || []) {
      const t = x.therapist || a.therapist;
      if (x.hasCav && x.cavTherapist) {
        add(t, { kind: "addon", name: x.serviceName || "", mins: n(x.duration) - n(x.cavMins) });
        add(x.cavTherapist, { kind: "addonMachine", name: x.serviceName || "", mins: n(x.cavMins) });
      } else add(t, { kind: "addon", name: x.serviceName || "", mins: n(x.duration) });
    }
  }
  return Object.entries(staff).map(([t, parts]) => ({ t, parts, mins: parts.reduce((s, p) => s + p.mins, 0) }));
}

// Every ticket sale in [start, end] that falls inside a campaign, split per staff member.
// `days` maps "YYYY-MM-DD" → that day's saved blob. Sales by repeat clients come back in
// `excluded` (with the client's type) so a report can list what didn't count and why.
export function computeTicketIncentives(days, start, end) {
  const sales = [], excluded = [];
  for (const date of Object.keys(days).filter(d => d >= start && d <= end).sort()) {
    const campaign = campaignFor(date);
    if (!campaign) continue;
    for (const a of days[date]?.appointments || []) {
      for (const sale of ticketSalesOf(a)) {
        const base = { date, client: (a.clientName || "").trim(), customerType: a.customerType || "", ...sale };
        if (!["NL", "NT"].includes(a.customerType)) { excluded.push(base); continue; }
        const staff = minutesByStaff(days[date], a.clientName);
        const totalMins = staff.reduce((s, x) => s + x.mins, 0);
        const pot = r2(sale.amount * campaign.rate);
        const shares = staff.map(x => ({ ...x, share: totalMins ? x.mins / totalMins : 1, amount: totalMins ? r2(pot * x.mins / totalMins) : pot }));
        // Put any rounding cent on the largest share so the shares add up to exactly `pot`.
        const diff = r2(pot - shares.reduce((s, x) => s + x.amount, 0));
        if (diff && shares.length) { const top = [...shares].sort((p, q) => q.mins - p.mins)[0]; top.amount = r2(top.amount + diff); }
        sales.push({ ...base, rate: campaign.rate, pot, totalMins, staff: shares });
      }
    }
  }
  return { sales, excluded };
}

// "Hiromi body 75 + Mikiko machine 15 of 90 min" style summary of how one sale was split.
export function describeSplit(sale) {
  if (sale.staff.length === 1) return `${sale.staff[0].t} alone (100%)`;
  const kind = (x) => x.parts.map(p => p.kind === "machine" || p.kind === "addonMachine" ? "machine" : p.kind === "body" ? "body" : p.name).join(" + ");
  return `${sale.staff.map(x => `${x.t} ${kind(x)} ${x.mins}min`).join(" / ")} of ${sale.totalMins}min`;
}
