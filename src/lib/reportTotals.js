// Shared between the in-app monthly report export (exportSalesReportXlsx in
// spa-daily-sheet.jsx) and scripts/sync-day-to-report.mjs, which patches confirmed days
// directly into the owner's hand-formatted "Sales ReportYYYY NEW.xlsx" on the Desktop —
// both need the exact same per-day cash/card/tip breakdown, so it lives here once instead
// of being hand-duplicated (the export script had already drifted from the in-app 📊集計
// tab twice before this was consolidated — see git history on this file's predecessor).

export const REFERRAL_SOURCES = ["Google / Website", "Google Map", "Instagram", "Yelp", "紹介"];

// Inline 物販購入 (within an appointment) supports multiple products in the same visit — the first
// item lives directly on the appointment (retailProductName/retailPurchaseAmount/...), any further
// items live in extraRetailItems[] so existing saved appointments (single item) keep working as-is.
export const getRetailItems = (a) => {
  const items = [];
  if (Number(a.retailPurchaseAmount || 0) > 0 || a.retailProductName) {
    items.push({ productName: a.retailProductName || "", amount: Number(a.retailPurchaseAmount || 0), paymentType: a.retailPurchasePaymentType, sellers: a.retailSellers });
  }
  (a.extraRetailItems || []).forEach(it => {
    if (Number(it.amount || 0) > 0 || it.productName) items.push(it);
  });
  return items;
};

// 社販 (staff self-purchase) — same pattern as getRetailItems: first item lives directly on the
// record (productName/amount/paymentType), further items live in extraItems[].
export const getStaffPurchaseItems = (sp) => {
  const items = [];
  if (Number(sp.amount || 0) > 0 || sp.productName) {
    items.push({ productName: sp.productName || "", amount: Number(sp.amount || 0), paymentType: sp.paymentType });
  }
  (sp.extraItems || []).forEach(it => {
    if (Number(it.amount || 0) > 0 || it.productName) items.push(it);
  });
  return items;
};

// Per-day cash/card/tip/clients breakdown used by the monthly sales report — both the in-app
// export and the local Excel-patch script call this so the two can never silently diverge.
export function computeDayTotals(data) {
  const appts = (data.appointments || []).filter(a => !a.isCavSlot);
  // The machine (cav) therapist's own portion of a split visit lives on its own isCavSlot
  // row (price/tip fields hold the cav amounts) — must be folded into today's cash/card
  // totals same as the body side, or a cav-split visit's cav earnings go missing from the
  // monthly report entirely (this was a real bug: matched the in-app 📊集計 tab's own
  // cavSlotAppts handling, which the export here had never mirrored).
  const cavSlotAppts = (data.appointments || []).filter(a => a.isCavSlot && !a.isTicket && !a.isGiftCard && !a.isPromo);
  const retails = data.retails || [];
  const refunds = data.refunds || [];
  const forgottenTips = data.forgottenTips || [];
  // Staff self-purchases (社販) are a real sale (the salon is still paid for the product) and
  // the in-app 📊集計 tab already folds them into its cash/card product totals — this export
  // hadn't mirrored that, so a day with a staff purchase reported a lower total here than the
  // tab's own "Today's GRAND TOTAL" for the same day.
  const staffPurchaseItems = (data.staffPurchases || []).flatMap(sp => getStaffPurchaseItems(sp));
  const spCash = staffPurchaseItems.filter(it => it.paymentType === "cash").reduce((s, it) => s + Number(it.amount || 0), 0);
  const spCard = staffPurchaseItems.filter(it => it.paymentType === "card").reduce((s, it) => s + Number(it.amount || 0), 0);
  const ticketSaleEvents = [];
  const localVisitEvents = [];
  appts.forEach(a => {
    const therapists = a.cavTherapist ? [{ name: a.therapist, share: 0.5 }, { name: a.cavTherapist, share: 0.5 }] : [{ name: a.therapist, share: 1 }];
    const hasTicketSale = (a.isSameDayTicket && Number(a.packagePrice || 0) > 0) || ((a.purchaseTags || []).includes("newTicket") && Number(a.newTicketAmount || 0) > 0);
    if (a.isSameDayTicket && Number(a.packagePrice || 0) > 0) {
      ticketSaleEvents.push({ amount: Number(a.packagePrice), customerType: a.customerType, referralSource: a.referralSource, therapists });
    }
    if ((a.purchaseTags || []).includes("newTicket") && Number(a.newTicketAmount || 0) > 0) {
      ticketSaleEvents.push({ amount: Number(a.newTicketAmount), customerType: a.customerType, referralSource: a.referralSource, therapists });
    }
    if (a.customerType === "RL" || a.customerType === "NL") {
      localVisitEvents.push({ therapists: therapists.map(th => th.name), hasTicketSale });
    }
  });
  const deposits = data.deposits || [];
  const refundServiceCash = refunds.filter(rf => rf.paymentType === "cash").reduce((s,rf) => s + Number(rf.serviceAmount||0), 0);
  const refundServiceCard = refunds.filter(rf => rf.paymentType === "card").reduce((s,rf) => s + Number(rf.serviceAmount||0), 0);
  const refundTipCash = refunds.filter(rf => rf.paymentType === "cash").reduce((s,rf) => s + Number(rf.tipAmount||0), 0);
  const refundTipCard = refunds.filter(rf => rf.paymentType === "card").reduce((s,rf) => s + Number(rf.tipAmount||0), 0);
  const forgottenServiceCash = forgottenTips.filter(ft => ft.paymentType === "cash").reduce((s,ft) => s + Number(ft.serviceAmount||0), 0);
  const forgottenServiceCard = forgottenTips.filter(ft => ft.paymentType === "card").reduce((s,ft) => s + Number(ft.serviceAmount||0), 0);
  const forgottenTipCash = forgottenTips.filter(ft => ft.paymentType === "cash").reduce((s,ft) => s + Number(ft.tipAmount||0), 0);
  const forgottenTipCard = forgottenTips.filter(ft => ft.paymentType === "card").reduce((s,ft) => s + Number(ft.tipAmount||0), 0);

  // Only count money actually received today: regular visits and same-day ticket purchases.
  // Excludes GC消化/PR無料 (no money received today) and pure ticket redemptions (already paid
  // for when the ticket bundle was originally purchased) — same rule as the 📊集計 tab.
  const revenueAppts = appts.filter(a => !a.isTicket && !a.isGiftCard && !a.isPromo);
  const sameDayTicketAppts = appts.filter(a => a.isTicket && a.isSameDayTicket && !a.isGiftCard && !a.isPromo);
  const pureTicketAppts = appts.filter(a => a.isTicket && !a.isSameDayTicket);
  // Addons explicitly marked "本日のお支払い" count as revenue; "前回の未消化分" don't (paid earlier).
  const revenueAddons = appts.flatMap(a => (a.addons || []).filter(ad => ad.countsAsRevenue === true));
  // A giftCardUsed amount was already collected as revenue on the (possibly earlier) day the gift
  // card was purchased/loaded, so it's excluded here too — same rule as the 📊集計 tab.
  const gcAlloc = (a) => {
    const gc = Number(a.giftCardUsed || 0);
    const svc = Number(a.price || 0);
    const tip = Number(a.tip || 0);
    const retail = (a.purchaseTags?.includes("retail")) ? getRetailItems(a).reduce((s, it) => s + Number(it.amount || 0), 0) : 0;
    const gcSvc = Math.min(gc, svc);
    const gcTip = Math.min(gc - gcSvc, tip);
    const gcRetail = Math.min(gc - gcSvc - gcTip, retail);
    return { gcSvc, gcTip, gcRetail };
  };
  // An add-on can also be paid from an existing gift card balance — same exclusion rule as
  // gcAlloc above, just against the add-on's own price/tip instead of the main service's.
  const addonGcAlloc = (ad) => {
    const gc = Number(ad.giftCardUsed || 0);
    const svc = Number(ad.price || 0);
    const tip = Number(ad.tip || 0);
    const gcSvc = Math.min(gc, svc);
    const gcTip = Math.min(gc - gcSvc, tip);
    return { gcSvc, gcTip };
  };
  // Retail sold inline during a visit (物販 tag on the appointment, up to 3 items via
  // getRetailItems) lives on the appointment record itself, not in the standalone `retails`
  // array — missed here entirely before this fix (matches the in-app 📊集計 tab's own
  // inlineRetailCash/inlineRetailCard handling).
  const inlineRetailAppts = appts.filter(a => (a.purchaseTags || []).includes("retail"));
  const retailCashCardSplit = (a) => {
    const items = getRetailItems(a);
    const itemsTotal = items.reduce((s, it) => s + Number(it.amount || 0), 0);
    const gcRetail = gcAlloc(a).gcRetail;
    let cash = 0, card = 0;
    items.forEach(it => {
      const amt = Number(it.amount || 0);
      const gcShare = itemsTotal > 0 ? gcRetail * amt / itemsTotal : 0;
      const net = Math.max(0, amt - gcShare);
      if (it.paymentType === "cash") cash += net; else card += net;
    });
    return { cash, card };
  };
  const inlineRetailCash = inlineRetailAppts.reduce((s,a) => s + retailCashCardSplit(a).cash, 0);
  const inlineRetailCard = inlineRetailAppts.reduce((s,a) => s + retailCashCardSplit(a).card, 0);
  const gcAllocPackage = (a) => {
    const gc = Number(a.giftCardUsed || 0);
    const svc = Number(a.packagePrice || 0);
    const tip = Number(a.packageTip ?? a.tip ?? 0);
    const gcSvc = Math.min(gc, svc);
    const gcTip = Math.min(gc - gcSvc, tip);
    return { gcSvc, gcTip };
  };

  // Deposits, gift cards, and cancellation fees are money received today but not tied to a
  // service line item — the user wants them folded into the "Treatment" column, not "Product".
  const depositCash = deposits.filter(dep => dep.paymentType === "cash").reduce((s,dep) => s + Number(dep.amount||0), 0);
  const depositCard = deposits.filter(dep => dep.paymentType === "card").reduce((s,dep) => s + Number(dep.amount||0), 0);

  const cashTreatment = revenueAppts.filter(a => !a.svcSplitPayment && a.paymentType === "cash").reduce((s,a) => s + Number(a.price||0) - gcAlloc(a).gcSvc, 0)
    + revenueAppts.filter(a => a.svcSplitPayment).reduce((s,a) => s + Number(a.svcCashPortion||0), 0)
    + sameDayTicketAppts.filter(a => !a.packageSplitPayment && a.paymentType === "cash").reduce((s,a) => s + Number(a.packagePrice||0) - gcAllocPackage(a).gcSvc, 0)
    + sameDayTicketAppts.filter(a => a.packageSplitPayment).reduce((s,a) => s + Number(a.packageCashPortion||0), 0)
    + pureTicketAppts.filter(a => a.extraPricePaymentType === "cash").reduce((s,a) => s + Number(a.extraPrice||0), 0)
    + sameDayTicketAppts.filter(a => a.extraPricePaymentType === "cash").reduce((s,a) => s + Number(a.extraPrice||0), 0)
    + revenueAddons.filter(ad => ad.paymentType === "cash").reduce((s,ad) => s + Number(ad.price||0) - addonGcAlloc(ad).gcSvc, 0)
    + cavSlotAppts.filter(a => a.paymentType === "cash").reduce((s,a) => s + Number(a.price||0), 0)
    + depositCash
    + forgottenServiceCash
    - refundServiceCash;
  const cashProduct = retails.filter(r => r.paymentType === "cash").reduce((s,r) => s + Number(r.price||0), 0) + inlineRetailCash + spCash;
  const cashTip = revenueAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "cash").reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
    + revenueAppts.filter(a => a.tipSplitPayment).reduce((s,a) => s + Number(a.tipCashPortion||0), 0)
    + sameDayTicketAppts.filter(a => a.tipPaymentType === "cash").reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
    + pureTicketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s,a) => s + Number(a.extraTip||0), 0)
    + sameDayTicketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s,a) => s + Number(a.extraTip||0), 0)
    + revenueAddons.filter(ad => ad.tipPaymentType === "cash").reduce((s,ad) => s + Number(ad.tip||0) - addonGcAlloc(ad).gcTip, 0)
    + cavSlotAppts.filter(a => a.tipPaymentType === "cash").reduce((s,a) => s + Number(a.tip||0), 0)
    + forgottenTipCash
    - refundTipCash;
  const cardTreatment = revenueAppts.filter(a => !a.svcSplitPayment && a.paymentType === "card").reduce((s,a) => s + Number(a.price||0) - gcAlloc(a).gcSvc, 0)
    + revenueAppts.filter(a => a.svcSplitPayment).reduce((s,a) => s + Number(a.svcCardPortion||0), 0)
    + sameDayTicketAppts.filter(a => !a.packageSplitPayment && a.paymentType === "card").reduce((s,a) => s + Number(a.packagePrice||0) - gcAllocPackage(a).gcSvc, 0)
    + sameDayTicketAppts.filter(a => a.packageSplitPayment).reduce((s,a) => s + Number(a.packageCardPortion||0), 0)
    + pureTicketAppts.filter(a => a.extraPricePaymentType === "card").reduce((s,a) => s + Number(a.extraPrice||0), 0)
    + sameDayTicketAppts.filter(a => a.extraPricePaymentType === "card").reduce((s,a) => s + Number(a.extraPrice||0), 0)
    + revenueAddons.filter(ad => ad.paymentType !== "cash").reduce((s,ad) => s + Number(ad.price||0) - addonGcAlloc(ad).gcSvc, 0)
    + cavSlotAppts.filter(a => a.paymentType !== "cash").reduce((s,a) => s + Number(a.price||0), 0)
    + depositCard
    + forgottenServiceCard
    - refundServiceCard;
  const cardProduct = retails.filter(r => r.paymentType === "card").reduce((s,r) => s + Number(r.price||0), 0) + inlineRetailCard + spCard;
  const cardTip = revenueAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "card").reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
    + revenueAppts.filter(a => a.tipSplitPayment).reduce((s,a) => s + Number(a.tipCardPortion||0), 0)
    + sameDayTicketAppts.filter(a => a.tipPaymentType === "card").reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
    + pureTicketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s,a) => s + Number(a.extraTip||0), 0)
    + sameDayTicketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s,a) => s + Number(a.extraTip||0), 0)
    + revenueAddons.filter(ad => ad.tipPaymentType !== "cash").reduce((s,ad) => s + Number(ad.tip||0) - addonGcAlloc(ad).gcTip, 0)
    + cavSlotAppts.filter(a => a.tipPaymentType !== "cash").reduce((s,a) => s + Number(a.tip||0), 0)
    + forgottenTipCard
    - refundTipCard;
  const totalTip = revenueAppts.reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
    + sameDayTicketAppts.reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
    + pureTicketAppts.reduce((s,a) => s + Number(a.extraTip||0), 0)
    + sameDayTicketAppts.reduce((s,a) => s + Number(a.extraTip||0), 0)
    + revenueAddons.reduce((s,ad) => s + Number(ad.tip||0), 0)
    + cavSlotAppts.reduce((s,a) => s + Number(a.tip||0), 0)
    + forgottenTipCash + forgottenTipCard
    - refundTipCash - refundTipCard;
  const totalSales = cashTreatment + cashProduct + cardTreatment + cardProduct;
  const totalCash = cashTreatment + cashProduct;
  const totalCard = cardTreatment + cardProduct;
  const clients = appts.filter(a => !a.isTicket).length + appts.filter(a => a.isTicket).length;

  return {
    totalSales,
    clients,
    cashTreatment: cashTreatment || "",
    cashProduct: cashProduct || "",
    totalCash: totalCash || "",
    cashTip: cashTip || "",
    cardTreatment: cardTreatment || "",
    cardProduct: cardProduct || "",
    totalCard: totalCard || "",
    cardTip: cardTip || "",
    totalTip: totalTip || "",
    grandTotal: totalSales + totalTip,
    rl: appts.filter(a => a.customerType === "RL").length,
    rt: appts.filter(a => a.customerType === "RT").length,
    nl: appts.filter(a => a.customerType === "NL").length,
    nt: appts.filter(a => a.customerType === "NT").length,
    referrals: REFERRAL_SOURCES.map(src => appts.filter(a => a.referralSource === src).length),
  };
}
