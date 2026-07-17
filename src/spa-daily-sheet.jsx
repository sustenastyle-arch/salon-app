import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const THERAPISTS = ["Mami", "Aya", "Megumi", "Hitomi", "Maki", "Yuka", "Mai", "Betsy"];
const CUSTOMER_TYPES = ["RL", "RT", "NL", "NT"];
const REFERRAL_SOURCES = ["Google / Website", "Google Map", "Instagram", "Yelp", "紹介"];
const HOURS = Array.from({ length: 11 }, (_, i) => i + 9);
// Weight Loss (regular, non-ticket): machine/cavitation allocation is a fixed $116 regardless
// of duration or total price — the rest goes to the body therapist.
const REGULAR_WL_CAV_PRICE = 116;
const isWeightLossService = (name) => (name || "").toLowerCase().includes("weight loss");
// Inline 物販購入 (within an appointment) supports multiple products in the same visit — the first
// item lives directly on the appointment (retailProductName/retailPurchaseAmount/...), any further
// items live in extraRetailItems[] so existing saved appointments (single item) keep working as-is.
const getRetailItems = (a) => {
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
const getStaffPurchaseItems = (sp) => {
  const items = [];
  if (Number(sp.amount || 0) > 0 || sp.productName) {
    items.push({ productName: sp.productName || "", amount: Number(sp.amount || 0), paymentType: sp.paymentType });
  }
  (sp.extraItems || []).forEach(it => {
    if (Number(it.amount || 0) > 0 || it.productName) items.push(it);
  });
  return items;
};

// ============================================================
// PRICE TABLE
// Format: { body_service, body_tip, cav_service, cav_tip }
// cav_service/cav_tip = null if no machine
// ============================================================
const PRICE_TABLE = {
  new: {
    // Improving Posture — new 2026 pricing (Sheet 2)
    "IP-60-3":  { body: { service: 108, tip: 23 }, cav: null },
    "IP-60-5":  { body: { service: 108, tip: 22 }, cav: null },
    "IP-75-3":  { body: { service: 125, tip: 25 }, cav: null },
    "IP-75-5":  { body: { service: 125, tip: 24 }, cav: null },
    "IP-90-3":  { body: { service: 125, tip: 25 }, cav: { service: 25, tip: 5 }, combined: { service: 151 } },
    "IP-90-5":  { body: { service: 125, tip: 24 }, cav: { service: 25, tip: 5 }, combined: { service: 151 } },
    "IP-120-3": { body: { service: 177, tip: 36 }, cav: { service: 25, tip: 5 }, combined: { service: 204 } },
    "IP-120-5": { body: { service: 177, tip: 35 }, cav: { service: 25, tip: 5 }, combined: { service: 200 } },
    // Weight Loss — new 2026 pricing (Sheet 1); combined = single-therapist rate
    "WL-75-3":  { body: { service: 75, tip: 17 }, cav: { service: 110, tip: 21 }, combined: { service: 188 } },
    "WL-75-5":  { body: { service: 75, tip: 17 }, cav: { service: 110, tip: 20 }, combined: { service: 183 } },
    "WL-90-3":  { body: { service: 85, tip: 19 }, cav: { service: 110, tip: 21 }, combined: { service: 202 } },
    "WL-90-5":  { body: { service: 85, tip: 19 }, cav: { service: 110, tip: 20 }, combined: { service: 193 } },
    "WL-115-3": { body: { service: 125, tip: 25 }, cav: { service: 110, tip: 21 }, combined: { service: 234, tip: 47 } },
    "WL-115-5": { body: { service: 125, tip: 24 }, cav: { service: 110, tip: 20 }, combined: { service: 224, tip: 45 } },
    // ProCell (Micro Channeling) — new 2026 pricing
    "MC-60-3":  { body: { service: 279, tip: 53.6 }, cav: null },
    "MC-60-5":  { body: { service: 251, tip: 48.2 }, cav: null },
    // Sculpted Facial — replaces the old Glow Facial ticket; same price as the
    // equivalent-length Improving Posture course (45min = IP-75, 60min = IP-90)
    "SF-45-3":  { body: { service: 125, tip: 25 }, cav: null },
    "SF-45-5":  { body: { service: 125, tip: 24 }, cav: null },
    "SF-60-3":  { body: { service: 125, tip: 25 }, cav: { service: 25, tip: 5 }, combined: { service: 151 } },
    "SF-60-5":  { body: { service: 125, tip: 24 }, cav: { service: 25, tip: 5 }, combined: { service: 151 } },
  },
  old: {
    // Improving Posture — old pricing (Sheet 3)
    "IP-60-3":  { body: { service: 108, tip: 21 }, cav: null },
    "IP-60-5":  { body: { service: 105, tip: 20 }, cav: null },
    "IP-75-3":  { body: { service: 120, tip: 24 }, cav: null },
    "IP-75-5":  { body: { service: 120, tip: 23 }, cav: null },
    "IP-90-3":  { body: { service: 120, tip: 24 }, cav: { service: 31, tip: 5 } },
    "IP-90-5":  { body: { service: 120, tip: 23 }, cav: { service: 31, tip: 5 } },
    "IP-120-3": { body: { service: 169, tip: 32 }, cav: { service: 24, tip: 5 } },
    "IP-120-5": { body: { service: 164, tip: 31 }, cav: { service: 24, tip: 5 } },
    // Weight Loss — old pricing (Sheet 4)
    "WL-75-3":  { body: { service: 83, tip: 16 }, cav: { service: 105, tip: 17 } },
    "WL-75-5":  { body: { service: 78, tip: 16 }, cav: { service: 105, tip: 16 } },
    "WL-90-3":  { body: { service: 96, tip: 18 }, cav: { service: 105, tip: 17 } },
    "WL-90-5":  { body: { service: 88, tip: 17 }, cav: { service: 105, tip: 16 }, combined: { service: 193, tip: 34 } },
    "WL-115-3": { body: { service: 120, tip: 24 }, cav: { service: 105, tip: 17 }, combined: { service: 235 } },
    "WL-115-5": { body: { service: 120, tip: 23 }, cav: { service: 105, tip: 16 } },
  }
};

// Full-package (ticket bundle) prices — from "施術振り分け表" (2026 new pricing).
// These are the total price for the whole 3- or 5-session ticket, used when a customer
// buys a new ticket (not the per-visit price used above for redeeming one session).
const TICKET_PACKAGE_PRICES = {
  "IP-60-3":  { service: 335,  tip: 67 },
  "IP-60-5":  { service: 542,  tip: 108 },
  "IP-75-3":  { service: 375,  tip: 75 },
  "IP-75-5":  { service: 608,  tip: 122 },
  "IP-90-3":  { service: 452,  tip: 91 },
  "IP-90-5":  { service: 729,  tip: 146 },
  "IP-120-3": { service: 612,  tip: 123 },
  "IP-120-5": { service: 1000, tip: 200 },
  "WL-75-3":  { service: 565,  tip: 113 },
  "WL-75-5":  { service: 917,  tip: 183 },
  "WL-90-3":  { service: 605,  tip: 121 },
  "WL-90-5":  { service: 967,  tip: 193 },
  "WL-115-3": { service: 703,  tip: 140 },
  "WL-115-5": { service: 1121, tip: 224 },
  "MC-60-3":  { service: 839,  tip: 161 },
  "MC-60-5":  { service: 1259, tip: 241 },
  // Sculpted Facial — same package price as the equivalent-length Improving Posture course
  "SF-45-3":  { service: 375,  tip: 75 },
  "SF-45-5":  { service: 608,  tip: 122 },
  "SF-60-3":  { service: 452,  tip: 91 },
  "SF-60-5":  { service: 729,  tip: 146 },
};

const MENU_OPTIONS = [
  { group: "Improving Posture", prefix: "IP", durations: [60, 75, 90, 120] },
  { group: "Weight Loss", prefix: "WL", durations: [75, 90, 115] },
  { group: "Sculpted Facial", prefix: "SF", durations: [45, 60] },
  { group: "Micro Channeling", prefix: "MC", durations: [60] },
];

// Staff capabilities
const CAV_CAPABLE = ["Mami", "Betsy", "Megumi", "Yuka"]; // can operate machine
const BODY_CAPABLE = ["Mami", "Betsy", "Megumi", "Aya", "Hitomi", "Mai", "Maki"]; // can do body massage
const DUAL_LICENSE = ["Mami", "Betsy", "Megumi"]; // body + machine (handles both themselves)
// Yuka: machine/facial only (no body massage)
const isCavCapable = (name) => CAV_CAPABLE.includes(name);
const isDualLicense = (name) => DUAL_LICENSE.includes(name);
const isBodyOnly = (name) => BODY_CAPABLE.includes(name) && !DUAL_LICENSE.includes(name) && name !== "Yuka";

const EMPTY_APPOINTMENT = {
  id: null,
  clientName: "",
  therapist: "",
  startTime: "",
  duration: 60,
  isTicket: false,
  ticketMenu: "",       // e.g. "IP-90-3"
  ticketTotal: 3,       // 3 or 5 (fallback used for price lookup; ticketTotalChosen tracks explicit selection)
  ticketTotalChosen: false,
  ticketCurrent: null,  // which session (1/3, 2/3...) — must be explicitly chosen
  priceVersion: "new",  // new | old (fallback used for price lookup; priceVersionChosen tracks explicit selection)
  priceVersionChosen: false,
  cavTherapist: "",
  price: 0,
  paymentType: "",
  svcSplitPayment: false, svcCashPortion: 0, svcCardPortion: 0,
  tip: 0,
  tipPaymentType: "",
  tipSplitPayment: false, tipCashPortion: 0, tipCardPortion: 0,
  cavPrice: 0,
  cavTip: 0,
  customerType: "",
  serviceName: "",  // コース名（Squareから自動入力）
  notes: "",
  purchaseTags: [],
  // Inline purchase records (recorded on same appointment)
  newTicketMenu: "", newTicketTotal: 3, newTicketPackageName: "",
  newTicketAmount: 0, newTicketTip: 0, newTicketPaymentType: "", newTicketTipPaymentType: "",
  newTicketSplitPayment: false, newTicketCashPortion: 0, newTicketCardPortion: 0,
  retailPurchaseAmount: 0, retailPurchasePaymentType: "", retailProductName: "",
  giftCardUsed: 0,
  giftCardPurchaseAmount: 0, giftCardPurchasePaymentType: "",
  depositApplied: 0,
  isGiftCard: false,
  isPromo: false,
  addons: [],
  fromSquare: false,
};

// Full Square menu list - names match exactly what Square returns
const SQUARE_SERVICES = [
  // Improving Posture
  { name: "Improving Posture 60min", duration: 60 },
  { name: "Improving Posture 75min", duration: 75 },
  { name: "Improving Posture 90min", duration: 90 },
  { name: "Improving Posture 120min", duration: 120 },
  // Weight Loss
  { name: "Weight Loss 75min", duration: 75 },
  { name: "Weight Loss 90min", duration: 90 },
  { name: "Weight Loss 115min", duration: 115 },
  // Facial / HIFU
  { name: "Sculpted Face Lift 45min", duration: 45 },
  { name: "Sculpted Face Lift 60min", duration: 60 },
  { name: "Sculpted Face Lift 75min", duration: 75 },
  { name: "HIFU Full Face 60min", duration: 60 },
  { name: "HIFU Half Face 20min", duration: 20 },
  { name: "HIFU Double Chin 30min", duration: 30 },
  // Microchanneling
  { name: "ProCell Microchanneling 60min", duration: 60 },
  // Bridal
  { name: "Bridal Treatment 60min", duration: 60 },
  { name: "Bridal Treatment 90min", duration: 90 },
  { name: "Bridal Treatment 120min", duration: 120 },
  // Massage
  { name: "Deep Tissue / Lomi Lomi 60min", duration: 60 },
  { name: "Deep Tissue / Lomi Lomi 90min", duration: 90 },
  { name: "Deep Tissue / Lomi Lomi 120min", duration: 120 },
  { name: "Prenatal Massage 60min", duration: 60 },
  { name: "Prenatal Massage 75min", duration: 75 },
  { name: "Couple Massage 90min", duration: 90 },
  // Add-ons
  { name: "Cavitation 10min", duration: 10 },
  { name: "Cavitation 20min", duration: 20 },
  { name: "Cavitation 30min", duration: 30 },
  { name: "Cavitation 40min", duration: 40 },
  { name: "Electric Brush 15min", duration: 15 },
  { name: "Hot Stone 30min", duration: 30 },
  { name: "Shaving 30min", duration: 30 },
  { name: "Head Massage 15min", duration: 15 },
  { name: "Radio Frequency 10min", duration: 10 },
  { name: "Radio Frequency 15min", duration: 15 },
  { name: "Hydro Diet 50min", duration: 50 },
  { name: "Added Massage 15min", duration: 15 },
  { name: "Added Massage 30min", duration: 30 },
  // Combo (for Jannie-style double bookings)
  { name: "Improving Posture 90min + HIFU Full Face 60min", duration: 150 },
];

// Match serviceName from Square data to SQUARE_SERVICES (fuzzy)
const matchServiceName = (rawName) => {
  if (!rawName) return rawName;
  const exact = SQUARE_SERVICES.find(s => s.name === rawName);
  if (exact) return rawName;
  // Try case-insensitive match
  const lower = rawName.toLowerCase();
  const fuzzy = SQUARE_SERVICES.find(s => s.name.toLowerCase() === lower);
  if (fuzzy) return fuzzy.name;
  // Return as-is (will show in dropdown as custom value)
  return rawName;
};
// A standalone retail sale not tied to any appointment (e.g. phone/walk-in customer). `sellers`
// (up to 3) attributes each their own dollar share for payroll purposes — commission rates differ
// per staff member, so the split is entered manually rather than assumed equal (see loadPayroll's
// "Retail — standalone" block).
const EMPTY_RETAIL = { id: Date.now(), item: "", price: 0, sellers: [{ therapist: "", amount: 0 }], paymentType: "" };
const EMPTY_DEPOSIT = { id: Date.now(), type: "deposit", amount: 0, clientName: "", paymentType: "", appointmentDate: "", appointmentTime: "", tip: 0, tipPaymentType: "", notes: "" };
const EMPTY_TICKET_PURCHASE = { id: Date.now(), clientName: "", packageName: "", ticketMenu: "", priceVersion: "new", ticketTotal: 3, amount: 0, paymentType: "", splitPayment: false, cashPortion: 0, cardPortion: 0, tip: 0, tipPaymentType: "", notes: "" };
const EMPTY_STAFF_PURCHASE = { id: Date.now(), staffName: "", productName: "", amount: 0, paymentType: "", notes: "", extraItems: [] };
// A same-day refund for a past visit (e.g. yesterday's client). Recorded as its own entry so it
// deducts from *today's* revenue total without touching the original (possibly already-locked) day.
// Doesn't touch any therapist's payroll — if a therapist's allocation also needs correcting, that's
// done manually by unlocking and editing the original day's appointment.
const EMPTY_REFUND = { id: Date.now(), clientName: "", serviceAmount: 0, tipAmount: 0, paymentType: "", originalDate: "", notes: "" };

// The opposite case from a refund: a payment (usually a cash tip) that a client actually paid on
// a past visit but staff forgot to ring into the register that day. Recorded as its own entry so
// it's ADDED to *today's* revenue total (the day the cash physically gets counted/reconciled) and
// credited to the therapist who should get payroll credit for it — without re-creating the old
// appointment (which would wrongly count as a new visit toward today's customer-type totals).
const EMPTY_FORGOTTEN_TIP = { id: Date.now(), clientName: "", therapist: "", serviceAmount: 0, tipAmount: 0, paymentType: "", originalDate: "", notes: "" };

const PURCHASE_TAGS = [
  { id: "newTicket", label: "🎟️ チケット新規購入", color: "#B71C1C", bg: "#FFEBEE" },
  { id: "giftCard",  label: "🎁 ギフトカード購入",  color: "#00796B", bg: "#E0F2F1" },
  { id: "retail",    label: "🛍️ 物販購入",         color: "#6A1B9A", bg: "#F3E5F5" },
];

const ADDON_PRESETS = [
  "前回キャビ未消化分",
  "キャビ追加 +10分",
  "キャビ追加 +20分",
  "キャビ追加 +30分",
  "キャビ追加 +40分",
];

// Sales tax subtracted before splitting a retail sale's commission-eligible amount among sellers.
const RETAIL_TAX_RATE = 0.04712;

const RETAIL_PRODUCTS = [
  { name: "Sheet Mask", price: 10 },
  { name: "Epicutis Sheet Mask", price: 30 },
  { name: "Epicutis Sample Set", price: 42 },
  { name: "Koso Shot", price: 10 },
  { name: "Koso Shot Campaign", price: 5 },
  { name: "Koso Drink", price: 99 },
  { name: "Belly Bliss Gut Renewal Tea 4oz", price: 30 },
  { name: "Detox Herbal Tea 4oz", price: 30 },
  { name: "Botanical Beauty Tea 4oz", price: 30 },
  { name: "Lymph Love Herbal Tea 4oz", price: 30 },
  { name: "Muscle Rub 5oz", price: 35 },
  { name: "Muscle Rub 8oz", price: 55 },
  { name: "Oil Cleanser", price: 85 },
  { name: "Lipid Serum", price: 250 },
  { name: "Hyvia Creme", price: 195 },
  { name: "Arctigenin Brightening Treatment", price: 175 },
  { name: "Lipid Recovery Mask", price: 125 },
  { name: "Lipid Body Treatment", price: 225 },
  { name: "Post Procedure Set", price: 75 },
  { name: "Cleansing Essentials Set", price: 105 },
  { name: "Luxury Skin Care Set", price: 395 },
  { name: "Liftech Cream", price: 0 },
  { name: "電気バリブラシ", price: 2095 },
];

const formatCurrency = (n) => `$${Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 0 })}`;
const r2 = (n) => Math.round(n * 100) / 100;
const formatTime = (hour) => { const h = hour % 12 || 12; return `${h}:00 ${hour < 12 ? "AM" : "PM"}`; };

const PaymentToggle = ({ value, onChange, small }) => (
  <div style={{ display: "flex", gap: 6 }}>
    {["cash", "card"].map(pt => (
      <button key={pt} onClick={() => onChange(pt)} style={{
        flex: 1, padding: small ? "6px 8px" : "10px", borderRadius: 8,
        border: `2px solid ${value === pt ? (pt === "cash" ? "#4CAF50" : "#2196F3") : "#DDD"}`,
        background: value === pt ? (pt === "cash" ? "#E8F5E9" : "#E3F2FD") : "#fff",
        cursor: "pointer", fontWeight: 700, fontSize: small ? 12 : 14,
        color: value === pt ? (pt === "cash" ? "#2E7D32" : "#1565C0") : "#888"
      }}>
        {pt === "cash" ? "💵 現金" : "💳 カード"}
      </button>
    ))}
  </div>
);

const PayBadge = ({ type }) => (
  <span style={{ fontSize: 11, background: type === "cash" ? "#E8F5E9" : "#E3F2FD", color: type === "cash" ? "#2E7D32" : "#1565C0", padding: "2px 8px", borderRadius: 10, whiteSpace: "nowrap" }}>
    {type === "cash" ? "現金" : "カード"}
  </span>
);

export default function SpaDailySheet() {
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [appointments, setAppointments] = useState([]);
  const [retails, setRetails] = useState([]);
  const [deposits, setDeposits] = useState([]);
  const [ticketPurchases, setTicketPurchases] = useState([]);
  const [staffPurchases, setStaffPurchases] = useState([]);
  const [refunds, setRefunds] = useState([]);
  const [forgottenTips, setForgottenTips] = useState([]);
  const [locked, setLocked] = useState(false);
  const [editingStaffPurchase, setEditingStaffPurchase] = useState(null);
  const [editingAppt, setEditingAppt] = useState(null);
  const [editingRetail, setEditingRetail] = useState(null);
  const [editingDeposit, setEditingDeposit] = useState(null);
  const [editingTicketPurchase, setEditingTicketPurchase] = useState(null);
  const [editingRefund, setEditingRefund] = useState(null);
  const [editingForgottenTip, setEditingForgottenTip] = useState(null);
  const [squareLoading, setSquareLoading] = useState(false);
  const [squareStatus, setSquareStatus] = useState(null);
  const [activeTab, setActiveTab] = useState("schedule");
  const [selectedTherapist, setSelectedTherapist] = useState("All");
  const [workingStaff, setWorkingStaff] = useState(THERAPISTS); // default: all working
  const [showStaffPicker, setShowStaffPicker] = useState(false);
  const [toast, setToast] = useState(null);
  const [depositsForDate, setDepositsForDate] = useState([]);
  const [gcSyncLoading, setGcSyncLoading] = useState(false);
  const [depositSyncLoading, setDepositSyncLoading] = useState(false);
  const [reconcileLoading, setReconcileLoading] = useState(false);
  const [reconcileResult, setReconcileResult] = useState(null);

  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); };

  // <input type="number"> silently increments/decrements by 1 when the mouse wheel scrolls
  // over it while it's focused — easy to trigger by accident while scrolling a long form, and
  // it looks exactly like "I typed 30 but it saved as 29". Blur the field on wheel so scrolling
  // the page never changes a number input's value.
  useEffect(() => {
    const handler = (e) => {
      if (document.activeElement?.type === "number") document.activeElement.blur();
    };
    document.addEventListener("wheel", handler, { passive: true });
    return () => document.removeEventListener("wheel", handler);
  }, []);

  // Square同期・ギフトカード取得はネットワーク待ちがある非同期処理 — その待ち時間中にスタッフが
  // 手動で別の予約を追加/編集すると、fetch開始時点でクロージャに捕まえていた古い配列を使って
  // マージ・保存してしまい、待っている間に追加された分を上書きして消してしまう(stale closure)。
  // 常に最新の状態を読めるようrefに同期しておき、非同期処理の続き(awaitの後)はこちらを使う。
  const stateRef = useRef({});
  useEffect(() => {
    stateRef.current = { appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips, locked, workingStaff };
  });

  useEffect(() => {
    setReconcileResult(null);
    const saved = localStorage.getItem(`spa-sheet-${date}`);
    if (saved) {
      const d = JSON.parse(saved);
      setAppointments(d.appointments || []);
      setRetails(d.retails || []);
      setDeposits(d.deposits || []);
      setTicketPurchases(d.ticketPurchases || []);
      setStaffPurchases(d.staffPurchases || []);
      setRefunds(d.refunds || []);
      setForgottenTips(d.forgottenTips || []);
      setLocked(!!d.locked);
      // Working-staff selection is per-day (who actually came in that day) — without this,
      // leaving the day and coming back always reset it to "everyone", which pushed staff
      // who weren't actually working that day back into the schedule grid as extra columns.
      // Days saved before this field existed have no d.workingStaff — for those, derive it
      // from who actually has an appointment that day instead of falling back to "everyone".
      if (d.workingStaff && d.workingStaff.length > 0) {
        setWorkingStaff(d.workingStaff);
      } else {
        const derived = [...new Set((d.appointments || []).map(a => a.therapist).filter(t => t && THERAPISTS.includes(t)))];
        setWorkingStaff(derived.length > 0 ? derived : THERAPISTS);
      }
    } else {
      setAppointments([]); setRetails([]); setDeposits([]); setTicketPurchases([]); setStaffPurchases([]); setRefunds([]); setForgottenTips([]);
      setLocked(false);
      setWorkingStaff(THERAPISTS);
    }
  }, [date]);

  // Scan all localStorage entries for deposits/gift-card prepayments with appointmentDate
  // matching current date — covers both a partial deposit and a fully prepaid ("ギフト"
  // as advance payment) visit, since either can be booked for a future date that later changes.
  useEffect(() => {
    const found = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key?.startsWith("spa-sheet-")) continue;
      try {
        const d = JSON.parse(localStorage.getItem(key));
        const recordedDate = key.replace("spa-sheet-", "");
        (d.deposits || []).forEach(dep => {
          if (dep.appointmentDate === date && (dep.type === "deposit" || dep.type === "giftcard")) {
            found.push({ ...dep, recordedDate });
          }
        });
      } catch {}
    }
    setDepositsForDate(found.sort((a, b) => (a.appointmentTime || "").localeCompare(b.appointmentTime || "")));
  }, [date, deposits]);

  const save = useCallback((appts, rets, deps, tps, sps, refs, fts, ws) => {
    localStorage.setItem(`spa-sheet-${date}`, JSON.stringify({ appointments: appts, retails: rets, deposits: deps, ticketPurchases: tps || [], staffPurchases: sps || [], refunds: refs || [], forgottenTips: fts || [], locked, workingStaff: ws || workingStaff }));
  }, [date, locked, workingStaff]);

  // Toggling the lock writes immediately (doesn't wait for another edit) using current state.
  const setDayLocked = (newLocked) => {
    localStorage.setItem(`spa-sheet-${date}`, JSON.stringify({ appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips, locked: newLocked, workingStaff }));
    setLocked(newLocked);
  };

  const MANAGER_PIN = import.meta.env.VITE_MANAGER_PIN || "0000";
  const guardLocked = () => {
    if (locked) { showToast("🔒 この日は確定済みです。解除してから編集してください", "error"); return true; }
    return false;
  };
  const handleLockToggle = () => {
    if (!locked) {
      if (window.confirm(`${date} を確定してロックしますか？\nロック後は解除するまで編集・削除ができなくなります。`)) {
        setDayLocked(true);
        showToast("🔒 この日を確定しました");
      }
      return;
    }
    const pin = window.prompt("解除するにはPINコードを入力してください");
    if (pin === null) return;
    if (pin === MANAGER_PIN) {
      setDayLocked(false);
      showToast("🔓 ロックを解除しました");
    } else {
      showToast("PINが違います", "error");
    }
  };

  const fetchSquare = async () => {
    if (guardLocked()) return;
    setSquareLoading(true); setSquareStatus(null);
    try {
      // Calls the local dev-server proxy (vite.config.js → squareBookingsApi), which keeps
      // SQUARE_ACCESS_TOKEN server-side and resolves customer name / service name / therapist
      // for each booking. Only works while `npm run dev` is running.
      const res = await fetch(`/api/square-bookings?date=${date}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
      const bookings = data.bookings || [];

      if (!bookings.length) {
        setSquareStatus("この日の予約がSquareに見つかりませんでした。");
        showToast("予約なし", "info");
      } else {
        const newAppts = bookings.map((b, i) => ({
          ...EMPTY_APPOINTMENT,
          id: `sq-${b.squareId || Date.now()}-${i}`,
          clientName: b.clientName || "",
          therapist: b.therapist || "",
          startTime: b.startTime || "",
          duration: Number(b.duration) || 60,
          serviceName: b.serviceName || "",
          notes: b.notes || "",
          fromSquare: true,
        }));

        // Read the latest state via the ref, not the closure — a manual edit made while this
        // fetch was in flight would otherwise get silently overwritten/lost here.
        const cur = stateRef.current;
        // If the day got locked while this fetch was in flight, abort instead of writing —
        // otherwise this save (using a `locked: false` closure from before the lock happened)
        // would silently re-unlock the day and overwrite whatever was just locked in.
        if (cur.locked) {
          showToast("🔒 取得中にこの日がロックされたため中止しました", "error");
          return;
        }
        const merged = [...cur.appointments];
        newAppts.forEach(na => {
          // Same client + same start time isn't necessarily the same booking — a customer can
          // have two simultaneous bookings with two different therapists (e.g. a couple's
          // service, or a body+cav split). Matching therapist too avoids treating the second
          // one as a dup of the first and silently dropping it.
          const exists = merged.find(a =>
            (a.fromSquare && a.id === na.id) ||
            (a.clientName === na.clientName && a.startTime === na.startTime && a.therapist === na.therapist)
          );
          if (!exists) merged.push(na);
        });
        merged.sort((a, b) => (a.startTime || "").localeCompare(b.startTime || ""));
        setAppointments(merged);

        // Only show columns for staff who actually have a booking today.
        const bookingTherapists = [...new Set(newAppts.map(a => a.therapist).filter(t => t && THERAPISTS.includes(t)))];
        const nextWorkingStaff = bookingTherapists.length > 0 ? bookingTherapists : cur.workingStaff;
        if (bookingTherapists.length > 0) setWorkingStaff(bookingTherapists);
        save(merged, cur.retails, cur.deposits, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips, nextWorkingStaff);

        showToast(`✅ ${newAppts.length}件取得しました`);
      }
    } catch (e) {
      console.error("Square sync error:", e);
      setSquareStatus("Square接続エラー。手動入力してください。");
      showToast("接続エラー", "error");
    }
    setSquareLoading(false);
  };

  // Pull today's online e-gift-card purchases from Square (no client name, amount only)
  // and add them as "giftcard"-type deposits so they land in the day's sales summary.
  const syncOnlineGiftCards = async () => {
    if (guardLocked()) return;
    setGcSyncLoading(true);
    try {
      const res = await fetch(`/api/giftcard-activities?date=${date}`);
      const data = await res.json();
      if (!res.ok) {
        showToast(data.error || "取得エラー", "error");
        return;
      }
      const activities = data.activities || [];
      // Read the latest state via the ref, not the closure — a manual edit made while this
      // fetch was in flight would otherwise get silently overwritten/lost here.
      const cur = stateRef.current;
      // If the day got locked while this fetch was in flight, abort instead of writing —
      // otherwise this save (using a `locked: false` closure from before the lock happened)
      // would silently re-unlock the day and overwrite whatever was just locked in.
      if (cur.locked) {
        showToast("🔒 取得中にこの日がロックされたため中止しました", "error");
        return;
      }
      const existingIds = new Set(cur.deposits.map(d => d.id));
      const newDeposits = activities
        .filter(a => !existingIds.has(a.id))
        .map(a => ({
          id: a.id,
          type: "giftcard",
          amount: a.amount,
          clientName: "ギフトカード購入（お客様名不明）",
          paymentType: a.paymentType === "cash" ? "cash" : "card",
          tip: 0,
          tipPaymentType: "cash",
          appointmentDate: "",
          appointmentTime: "",
          notes: `Square自動取得${a.createdAt ? ` (${a.createdAt})` : ""}`,
        }));
      if (newDeposits.length === 0) {
        showToast(activities.length === 0 ? "この日のギフトカード購入はありません" : "すべて取得済みです", "info");
        return;
      }
      const next = [...cur.deposits, ...newDeposits];
      setDeposits(next);
      save(cur.appointments, cur.retails, next, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips);
      showToast(`✅ ギフトカード購入 ${newDeposits.length}件追加しました`);
    } catch (e) {
      console.error("Gift card sync error:", e);
      showToast("接続エラー", "error");
    } finally {
      setGcSyncLoading(false);
    }
  };

  // Pull today's deposit payments from Square (collected via a Payment Link with a custom
  // line item named "Deposit" / historically mistyped "Desit" — there's no dedicated Square
  // API for this the way gift cards have one, see api/square-deposits.js) and add them as
  // "deposit"-type deposits. Appointment date/time are left blank since Square doesn't say
  // which future visit the deposit is for — staff fills those in by editing the entry.
  const syncSquareDeposits = async () => {
    if (guardLocked()) return;
    setDepositSyncLoading(true);
    try {
      const res = await fetch(`/api/square-deposits?date=${date}`);
      const data = await res.json();
      if (!res.ok) {
        showToast(data.error || "取得エラー", "error");
        return;
      }
      const found = data.deposits || [];
      // Read the latest state via the ref, not the closure — a manual edit made while this
      // fetch was in flight would otherwise get silently overwritten/lost here.
      const cur = stateRef.current;
      // If the day got locked while this fetch was in flight, abort instead of writing —
      // otherwise this save (using a `locked: false` closure from before the lock happened)
      // would silently re-unlock the day and overwrite whatever was just locked in.
      if (cur.locked) {
        showToast("🔒 取得中にこの日がロックされたため中止しました", "error");
        return;
      }
      const existingIds = new Set(cur.deposits.map(d => d.id));
      const newDeposits = found
        .filter(d => !existingIds.has(d.id))
        .map(d => ({
          id: d.id,
          type: "deposit",
          amount: d.amount,
          clientName: d.clientName,
          paymentType: d.paymentType,
          tip: 0,
          tipPaymentType: "cash",
          appointmentDate: "",
          appointmentTime: "",
          notes: `Square自動取得${d.createdAt ? ` (${d.createdAt})` : ""}`,
        }));
      if (newDeposits.length === 0) {
        showToast(found.length === 0 ? "この日のデポジット支払いはありません" : "すべて取得済みです", "info");
        return;
      }
      const next = [...cur.deposits, ...newDeposits];
      setDeposits(next);
      save(cur.appointments, cur.retails, next, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips);
      showToast(`✅ デポジット ${newDeposits.length}件追加しました`);
    } catch (e) {
      console.error("Deposit sync error:", e);
      showToast("接続エラー", "error");
    } finally {
      setDepositSyncLoading(false);
    }
  };

  // Compare this sheet's manually-entered cash/card/tip totals against what Square's
  // Payments API actually recorded for the day, to catch entry typos same-day rather
  // than at month-end reporting time.
  const checkSquareReconciliation = async () => {
    setReconcileLoading(true);
    setReconcileResult(null);
    try {
      const res = await fetch(`/api/square-payments?date=${date}`);
      const data = await res.json();
      if (!res.ok) {
        showToast(data.error || "取得エラー", "error");
        return;
      }
      setReconcileResult({ ...data, checkedAt: Date.now() });
    } catch (e) {
      console.error("Square reconcile error:", e);
      showToast("接続エラー", "error");
    } finally {
      setReconcileLoading(false);
    }
  };

  // One-time helper for migrating data between environments (e.g. localhost -> production) —
  // this app has no backend, so each origin's localStorage is otherwise completely isolated
  // and there's no other way to carry data across them.
  const handleImportBackup = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        Object.entries(data).forEach(([d, v]) => localStorage.setItem(`spa-sheet-${d}`, v));
        showToast(`✅ ${Object.keys(data).length}日分のデータを復元しました`);
        window.location.reload();
      } catch (err) {
        showToast("読み込みに失敗しました: " + err.message, "error");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const calcCavStartTime = (startTime, duration, cavDuration = 15) => {
    if (!startTime) return "";
    const [h, m] = startTime.split(":").map(Number);
    const totalMins = h * 60 + m + (Number(duration) - Number(cavDuration));
    const ch = Math.floor(totalMins / 60);
    const cm = totalMins % 60;
    return `${String(ch).padStart(2, "0")}:${String(cm).padStart(2, "0")}`;
  };

  const saveAppt = (appt) => {
    if (guardLocked()) return;
    const cavSlotId = `cav-${appt.id}`;

    // A separate cav slot is needed when:
    // - A cav therapist is selected (meaning the main therapist can't do machine)
    // - There are cav amounts to record
    const needsCavSlot = !!appt.cavTherapist && (Number(appt.cavPrice) > 0 || Number(appt.cavTip) > 0);
    // Machine (cav) portion duration: use the minutes the staff actually entered
    // (needed e.g. for Weight Loss, where machine=40min / body=75min), falling back
    // to the old default of 15min for quick facial add-ons where it wasn't entered.
    const cavDuration = Number(appt.cavMins) || 15;
    const bodyDuration = appt.bodyMins ? Number(appt.bodyMins) : Number(appt.duration) - cavDuration;

    const cavSlot = needsCavSlot ? {
      id: cavSlotId,
      isCavSlot: true,
      parentId: appt.id,
      clientName: appt.clientName,
      therapist: appt.cavTherapist,
      bodyTherapist: appt.therapist,
      startTime: calcCavStartTime(appt.startTime, appt.duration, cavDuration),
      duration: cavDuration,
      isTicket: appt.isTicket,
      isSameDayTicket: appt.isSameDayTicket,
      isGiftCard: appt.isGiftCard,
      isPromo: appt.isPromo,
      ticketMenu: appt.ticketMenu,
      ticketTotal: appt.ticketTotal,
      ticketCurrent: appt.ticketCurrent,
      priceVersion: appt.priceVersion,
      price: Number(appt.cavPrice),   // machine service fee shows HERE
      tip: Number(appt.cavTip),       // machine tip shows HERE
      paymentType: appt.paymentType,
      tipPaymentType: appt.tipPaymentType,
      cavPrice: 0,
      cavTip: 0,
      cavTherapist: "",
      customerType: appt.customerType,
      notes: `⚡ 機械 for ${appt.clientName}（${appt.therapist}）`,
      fromSquare: false,
    } : null;

    // Remove old entries
    let next = appointments.filter(a => a.id !== appt.id && a.id !== cavSlotId);

    // Main appt: cav amounts stay on mainAppt for dual-license (they keep it all),
    // but are stripped when there's a separate cav slot
    const mainAppt = needsCavSlot
      ? { ...appt, cavPrice: 0, cavTip: 0, duration: bodyDuration }  // body only — machine goes to cav slot
      : appt;

    next = [...next, mainAppt];
    if (cavSlot) next = [...next, cavSlot];
    next.sort((a, b) => (a.startTime || "").localeCompare(b.startTime || ""));

    setAppointments(next); save(next, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); setEditingAppt(null); showToast("保存しました");
  };

  const deleteAppt = (id) => {
    if (guardLocked()) return;
    const cavSlotId = `cav-${id}`;
    const next = appointments.filter(a => a.id !== id && a.id !== cavSlotId);
    setAppointments(next); save(next, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); setEditingAppt(null);
  };

  // When editing an existing appointment, restore cavPrice/cavTip from the cav slot.
  // bodyMins/cavMins come from the saved durations (appt.duration is body-only once a cav
  // slot is split off; cavSlot.duration is the machine minutes actually entered) — but records
  // saved before that duration-split existed still have the OLD shape (appt.duration = full
  // total, cavSlot.duration hardcoded to 15), which would double-count into a bogus total. Detect
  // that by cross-checking against the course's real duration and blank the minutes out instead
  // of trusting stale numbers — the Weight Loss auto-fill effect will then re-default them.
  const openApptForEdit = (appt) => {
    if (!appt.isCavSlot && appt.cavTherapist) {
      const cavSlot = appointments.find(a => a.isCavSlot && a.parentId === appt.id);
      if (cavSlot) {
        const restoredCavPrice = Number(cavSlot.price || 0);
        const restoredCavTip = Number(cavSlot.tip || 0);
        const bodyPrice = Number(appt.price || 0);
        const svc = SQUARE_SERVICES.find(s => s.name === appt.serviceName);
        const savedBodyMins = Number(appt.duration || 0);
        const savedCavMins = Number(cavSlot.duration || 0);
        const trueTotal = svc ? svc.duration : savedBodyMins + savedCavMins;
        // For Weight Loss the machine is always exactly 40min — if a saved record has some other
        // cav duration (e.g. leftover minutes from a different course picked earlier in the same
        // form, saved before that bug was fixed), treat it as inconsistent too, even though the
        // two numbers happen to add up to the right total.
        const sumMatches = savedBodyMins + savedCavMins === trueTotal;
        const isConsistent = sumMatches && (!isWeightLossService(appt.serviceName) || savedCavMins === 40);
        const bodyMins = isConsistent ? savedBodyMins : 0;
        const cavMins = isConsistent ? savedCavMins : 0;
        setEditingAppt({
          ...appt,
          duration: trueTotal,
          cavPrice: restoredCavPrice,
          cavTip: restoredCavTip,
          bodyMins,
          cavMins,
          totalServiceInput: bodyPrice + restoredCavPrice,
          totalTipInput: Number(appt.tip || 0) + restoredCavTip,
        });
        return;
      }
    }
    setEditingAppt(appt);
  };
  const saveRetail = (r) => { if (guardLocked()) return; const next = retails.find(x => x.id === r.id) ? retails.map(x => x.id === r.id ? r : x) : [...retails, r]; setRetails(next); save(appointments, next, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); setEditingRetail(null); showToast("物販保存"); };
  const deleteRetail = (id) => { if (guardLocked()) return; const next = retails.filter(r => r.id !== id); setRetails(next); save(appointments, next, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); };
  const saveDeposit = (d) => { if (guardLocked()) return; const next = deposits.find(x => x.id === d.id) ? deposits.map(x => x.id === d.id ? d : x) : [...deposits, d]; setDeposits(next); save(appointments, retails, next, ticketPurchases, staffPurchases, refunds, forgottenTips); setEditingDeposit(null); showToast("保存しました"); };
  const deleteDeposit = (id) => { if (guardLocked()) return; const next = deposits.filter(d => d.id !== id); setDeposits(next); save(appointments, retails, next, ticketPurchases, staffPurchases, refunds, forgottenTips); };
  const saveTicketPurchase = (tp) => { if (guardLocked()) return; const next = ticketPurchases.find(x => x.id === tp.id) ? ticketPurchases.map(x => x.id === tp.id ? tp : x) : [...ticketPurchases, tp]; setTicketPurchases(next); save(appointments, retails, deposits, next, staffPurchases, refunds, forgottenTips); setEditingTicketPurchase(null); showToast("🎟️ チケット購入保存"); };
  const deleteTicketPurchase = (id) => { if (guardLocked()) return; const next = ticketPurchases.filter(tp => tp.id !== id); setTicketPurchases(next); save(appointments, retails, deposits, next, staffPurchases, refunds, forgottenTips); };
  const saveStaffPurchase = (sp) => { if (guardLocked()) return; const next = staffPurchases.find(x => x.id === sp.id) ? staffPurchases.map(x => x.id === sp.id ? sp : x) : [...staffPurchases, sp]; setStaffPurchases(next); save(appointments, retails, deposits, ticketPurchases, next, refunds, forgottenTips); setEditingStaffPurchase(null); showToast("👩‍💼 社販保存"); };
  const deleteStaffPurchase = (id) => { if (guardLocked()) return; const next = staffPurchases.filter(sp => sp.id !== id); setStaffPurchases(next); save(appointments, retails, deposits, ticketPurchases, next, refunds, forgottenTips); };
  const saveRefund = (rf) => { if (guardLocked()) return; const next = refunds.find(x => x.id === rf.id) ? refunds.map(x => x.id === rf.id ? rf : x) : [...refunds, rf]; setRefunds(next); save(appointments, retails, deposits, ticketPurchases, staffPurchases, next, forgottenTips); setEditingRefund(null); showToast("🔙 返金を記録しました"); };
  const deleteRefund = (id) => { if (guardLocked()) return; const next = refunds.filter(rf => rf.id !== id); setRefunds(next); save(appointments, retails, deposits, ticketPurchases, staffPurchases, next, forgottenTips); };
  const saveForgottenTip = (ft) => { if (guardLocked()) return; const next = forgottenTips.find(x => x.id === ft.id) ? forgottenTips.map(x => x.id === ft.id ? ft : x) : [...forgottenTips, ft]; setForgottenTips(next); save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, next); setEditingForgottenTip(null); showToast("🙏 打ち忘れ入力を記録しました"); };
  const deleteForgottenTip = (id) => { if (guardLocked()) return; const next = forgottenTips.filter(ft => ft.id !== id); setForgottenTips(next); save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, next); };

  // Summary — tickets excluded from today's revenue; cav slots excluded (counted in parent)
  const regularAppts = appointments.filter(a => !a.isTicket && !a.isCavSlot && !a.isGiftCard && !a.isPromo);
  const ticketAppts = appointments.filter(a => a.isTicket && !a.isCavSlot);
  const pureTicketAppts = ticketAppts.filter(a => !a.isSameDayTicket);
  const sameDayAppts = ticketAppts.filter(a => a.isSameDayTicket);
  const gcAppts = appointments.filter(a => a.isGiftCard && !a.isCavSlot);
  const promoAppts = appointments.filter(a => a.isPromo && !a.isCavSlot);
  // Addons marked "本日のお支払い" (countsAsRevenue===true) are real money received today;
  // "前回の未消化分" (countsAsRevenue===false) are makeup for an already-paid ticket session, so
  // they're excluded here — same rule as GC消化/PR無料 — but still counted in 💴給料集計 payroll.
  const revenueAddons = appointments.filter(a => !a.isCavSlot).flatMap(a => (a.addons || []).filter(ad => ad.countsAsRevenue === true));
  // The cav (machine) therapist's own portion of a split visit lives on its own isCavSlot row
  // (its price/tip fields hold the cav amounts, inheriting the same paymentType/tipPaymentType as
  // the body row) — must be folded into today's revenue/tip totals same as the body side.
  const cavSlotAppts = appointments.filter(a => a.isCavSlot && !a.isTicket && !a.isGiftCard && !a.isPromo);
  // Refunds for a past (possibly already-locked) day's visit are recorded today and deducted from
  // *today's* totals only — they never reach back to adjust that day's therapist payroll.
  const totalRefundService = refunds.reduce((s, rf) => s + Number(rf.serviceAmount || 0), 0);
  const totalRefundTip = refunds.reduce((s, rf) => s + Number(rf.tipAmount || 0), 0);
  const totalRefundCash = refunds.filter(rf => rf.paymentType === "cash").reduce((s, rf) => s + Number(rf.serviceAmount || 0), 0);
  const totalRefundCard = refunds.filter(rf => rf.paymentType === "card").reduce((s, rf) => s + Number(rf.serviceAmount || 0), 0);
  const totalRefundTipCash = refunds.filter(rf => rf.paymentType === "cash").reduce((s, rf) => s + Number(rf.tipAmount || 0), 0);
  const totalRefundTipCard = refunds.filter(rf => rf.paymentType === "card").reduce((s, rf) => s + Number(rf.tipAmount || 0), 0);
  // The opposite of a refund: a past visit's payment (usually a cash tip) that staff forgot to
  // ring in that day. Recorded today and ADDED to *today's* totals only — the register drawer is
  // physically counted today, so that's when this cash needs to show up in the system total too.
  const totalForgottenService = forgottenTips.reduce((s, ft) => s + Number(ft.serviceAmount || 0), 0);
  const totalForgottenTip = forgottenTips.reduce((s, ft) => s + Number(ft.tipAmount || 0), 0);
  const totalForgottenCash = forgottenTips.filter(ft => ft.paymentType === "cash").reduce((s, ft) => s + Number(ft.serviceAmount || 0), 0);
  const totalForgottenCard = forgottenTips.filter(ft => ft.paymentType === "card").reduce((s, ft) => s + Number(ft.serviceAmount || 0), 0);
  const totalForgottenTipCash = forgottenTips.filter(ft => ft.paymentType === "cash").reduce((s, ft) => s + Number(ft.tipAmount || 0), 0);
  const totalForgottenTipCard = forgottenTips.filter(ft => ft.paymentType === "card").reduce((s, ft) => s + Number(ft.tipAmount || 0), 0);

  // GC allocation: a giftCardUsed amount was already collected as revenue on the (possibly earlier)
  // day the gift card was purchased/loaded, so it must never also count as *today's* revenue —
  // same rule as GC消化/PR無料. Covers service first, then tip, then retail (in order).
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
  // Same allocation, but for a same-day ticket *purchase*'s package price/tip (giftCardUsed there
  // covers the package, not the per-visit staff-allocation price/tip fields).
  const gcAllocPackage = (a) => {
    const gc = Number(a.giftCardUsed || 0);
    const svc = Number(a.packagePrice || 0);
    const tip = Number(a.packageTip ?? a.tip ?? 0);
    const gcSvc = Math.min(gc, svc);
    const gcTip = Math.min(gc - gcSvc, tip);
    return { gcSvc, gcTip };
  };

  const totalRevenue = regularAppts.reduce((s, a) => s + Number(a.price || 0) - gcAlloc(a).gcSvc, 0)
    + sameDayAppts.reduce((s, a) => s + Number(a.packagePrice || 0) - gcAllocPackage(a).gcSvc, 0)
    + ticketAppts.reduce((s, a) => s + Number(a.extraPrice || 0), 0)
    + revenueAddons.reduce((s, ad) => s + Number(ad.price || 0), 0)
    + totalForgottenService
    - totalRefundService;
  const totalCavRevenue = regularAppts.reduce((s, a) => s + Number(a.cavPrice || 0), 0)
    + cavSlotAppts.reduce((s, a) => s + Number(a.price || 0), 0);
  const totalTips = regularAppts.reduce((s, a) => s + Number(a.tip || 0) - gcAlloc(a).gcTip, 0)
    + sameDayAppts.reduce((s, a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
    + ticketAppts.reduce((s, a) => s + Number(a.extraTip || 0), 0)
    + revenueAddons.reduce((s, ad) => s + Number(ad.tip || 0), 0)
    + totalForgottenTip
    - totalRefundTip;
  const totalCavTips = regularAppts.reduce((s, a) => s + Number(a.cavTip || 0), 0)
    + cavSlotAppts.reduce((s, a) => s + Number(a.tip || 0), 0);
  const totalGCUsed = regularAppts.reduce((s, a) => s + Number(a.giftCardUsed || 0), 0)
    + sameDayAppts.reduce((s, a) => s + Number(a.giftCardUsed || 0), 0);
  const totalCash = regularAppts.filter(a => !a.svcSplitPayment && a.paymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.price || 0) - gcAlloc(a).gcSvc), 0)
    + regularAppts.filter(a => a.svcSplitPayment).reduce((s, a) => s + Number(a.svcCashPortion || 0), 0)
    + sameDayAppts.filter(a => !a.packageSplitPayment && a.paymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.packagePrice || 0) - gcAllocPackage(a).gcSvc), 0)
    + sameDayAppts.filter(a => a.packageSplitPayment).reduce((s, a) => s + Number(a.packageCashPortion || 0), 0)
    + ticketAppts.filter(a => a.extraPricePaymentType === "cash").reduce((s, a) => s + Number(a.extraPrice || 0), 0)
    + revenueAddons.filter(ad => ad.paymentType === "cash").reduce((s, ad) => s + Number(ad.price || 0), 0)
    + cavSlotAppts.filter(a => a.paymentType === "cash").reduce((s, a) => s + Number(a.price || 0), 0)
    + totalForgottenCash
    - totalRefundCash;
  const totalCard = regularAppts.filter(a => !a.svcSplitPayment && a.paymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.price || 0) - gcAlloc(a).gcSvc), 0)
    + regularAppts.filter(a => a.svcSplitPayment).reduce((s, a) => s + Number(a.svcCardPortion || 0), 0)
    + sameDayAppts.filter(a => !a.packageSplitPayment && a.paymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.packagePrice || 0) - gcAllocPackage(a).gcSvc), 0)
    + sameDayAppts.filter(a => a.packageSplitPayment).reduce((s, a) => s + Number(a.packageCardPortion || 0), 0)
    + ticketAppts.filter(a => a.extraPricePaymentType === "card").reduce((s, a) => s + Number(a.extraPrice || 0), 0)
    + revenueAddons.filter(ad => ad.paymentType !== "cash").reduce((s, ad) => s + Number(ad.price || 0), 0)
    + cavSlotAppts.filter(a => a.paymentType !== "cash").reduce((s, a) => s + Number(a.price || 0), 0)
    + totalForgottenCard
    - totalRefundCard;
  const totalTipCash = regularAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.tip || 0) - gcAlloc(a).gcTip), 0)
    + regularAppts.filter(a => a.tipSplitPayment).reduce((s, a) => s + Number(a.tipCashPortion || 0), 0)
    + sameDayAppts.filter(a => a.tipPaymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip), 0)
    + ticketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s, a) => s + Number(a.extraTip || 0), 0)
    + revenueAddons.filter(ad => ad.tipPaymentType === "cash").reduce((s, ad) => s + Number(ad.tip || 0), 0)
    + cavSlotAppts.filter(a => a.tipPaymentType === "cash").reduce((s, a) => s + Number(a.tip || 0), 0)
    + deposits.filter(d => d.tipPaymentType === "cash" || !d.tipPaymentType).reduce((s, d) => s + Number(d.tip || 0), 0)
    + totalForgottenTipCash
    - totalRefundTipCash;
  const totalTipCard = regularAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.tip || 0) - gcAlloc(a).gcTip), 0)
    + regularAppts.filter(a => a.tipSplitPayment).reduce((s, a) => s + Number(a.tipCardPortion || 0), 0)
    + sameDayAppts.filter(a => a.tipPaymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip), 0)
    + ticketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s, a) => s + Number(a.extraTip || 0), 0)
    + revenueAddons.filter(ad => ad.tipPaymentType !== "cash").reduce((s, ad) => s + Number(ad.tip || 0), 0)
    + cavSlotAppts.filter(a => a.tipPaymentType === "card").reduce((s, a) => s + Number(a.tip || 0), 0)
    + deposits.filter(d => d.tipPaymentType === "card").reduce((s, d) => s + Number(d.tip || 0), 0)
    + totalForgottenTipCard
    - totalRefundTipCard;
  const totalRetail = retails.reduce((s, r) => s + Number(r.price || 0), 0);
  // 社販 (staff buying product for themselves) is still real product revenue for the salon.
  // A single 社販 record can cover several items bought in one go (getStaffPurchaseItems).
  const allStaffPurchaseItems = staffPurchases.flatMap(sp => getStaffPurchaseItems(sp));
  const spCash = allStaffPurchaseItems.filter(it => it.paymentType === "cash").reduce((s, it) => s + Number(it.amount || 0), 0);
  const spCard = allStaffPurchaseItems.filter(it => it.paymentType === "card").reduce((s, it) => s + Number(it.amount || 0), 0);
  const spTotal = allStaffPurchaseItems.reduce((s, it) => s + Number(it.amount || 0), 0);
  const totalDepositAmt = deposits.filter(d => d.type === "deposit").reduce((s, d) => s + Number(d.amount || 0), 0);
  const totalDepositApplied = regularAppts.reduce((s, a) => s + Number(a.depositApplied || 0), 0);
  const totalGiftCard = deposits.filter(d => d.type === "giftcard").reduce((s, d) => s + Number(d.amount || 0), 0);
  const totalCancellation = deposits.filter(d => d.type === "cancellation").reduce((s, d) => s + Number(d.amount || 0), 0);
  const totalCancellationCash = deposits.filter(d => d.type === "cancellation" && d.paymentType === "cash").reduce((s, d) => s + Number(d.amount || 0), 0);
  const totalCancellationCard = deposits.filter(d => d.type === "cancellation" && d.paymentType === "card").reduce((s, d) => s + Number(d.amount || 0), 0);
  // Deposit + gift card + cancellation combined by payment method — matches the Sales Report Excel's
  // "Treatment" column, which folds all three of these in (they're money received today, just not
  // tied to a specific service line item).
  const depositAllCash = deposits.filter(d => d.paymentType === "cash").reduce((s, d) => s + Number(d.amount || 0), 0);
  const depositAllCard = deposits.filter(d => d.paymentType === "card").reduce((s, d) => s + Number(d.amount || 0), 0);
  // Ticket new purchases (standalone section) — split-payment entries contribute cashPortion/cardPortion instead
  const tpCash = ticketPurchases.filter(tp => !tp.splitPayment && tp.paymentType === "cash").reduce((s, tp) => s + Number(tp.amount || 0), 0)
    + ticketPurchases.filter(tp => tp.splitPayment).reduce((s, tp) => s + Number(tp.cashPortion || 0), 0);
  const tpCard = ticketPurchases.filter(tp => !tp.splitPayment && tp.paymentType === "card").reduce((s, tp) => s + Number(tp.amount || 0), 0)
    + ticketPurchases.filter(tp => tp.splitPayment).reduce((s, tp) => s + Number(tp.cardPortion || 0), 0);
  const tpTotal = ticketPurchases.reduce((s, tp) => s + Number(tp.amount || 0), 0);
  const tpTipTotal = ticketPurchases.reduce((s, tp) => s + Number(tp.tip || 0), 0);
  const tpTipCash = ticketPurchases.filter(tp => tp.tipPaymentType === "cash" || !tp.tipPaymentType).reduce((s, tp) => s + Number(tp.tip || 0), 0);
  const tpTipCard = ticketPurchases.filter(tp => tp.tipPaymentType === "card").reduce((s, tp) => s + Number(tp.tip || 0), 0);
  // Inline purchases recorded inside appointment cards
  const allApptsList = appointments.filter(a => !a.isCavSlot);
  const inlineNewTicketAppts = allApptsList.filter(a => (a.purchaseTags||[]).includes("newTicket"));
  const inlineRetailAppts   = allApptsList.filter(a => (a.purchaseTags||[]).includes("retail"));
  const inlineGCAppts       = allApptsList.filter(a => (a.purchaseTags||[]).includes("giftCard"));
  const inlineNTCash  = inlineNewTicketAppts.filter(a=>!a.newTicketSplitPayment && a.newTicketPaymentType==="cash").reduce((s,a)=>s+Number(a.newTicketAmount||0),0)
    + inlineNewTicketAppts.filter(a=>a.newTicketSplitPayment).reduce((s,a)=>s+Number(a.newTicketCashPortion||0),0);
  const inlineNTCard  = inlineNewTicketAppts.filter(a=>!a.newTicketSplitPayment && a.newTicketPaymentType!=="cash").reduce((s,a)=>s+Number(a.newTicketAmount||0),0)
    + inlineNewTicketAppts.filter(a=>a.newTicketSplitPayment).reduce((s,a)=>s+Number(a.newTicketCardPortion||0),0);
  const inlineNTTipCash = inlineNewTicketAppts.filter(a=>a.newTicketTipPaymentType==="cash").reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  const inlineNTTipCard = inlineNewTicketAppts.filter(a=>a.newTicketTipPaymentType!=="cash").reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  const inlineNTTotal = inlineNewTicketAppts.reduce((s,a)=>s+Number(a.newTicketAmount||0),0);
  const inlineNTTipTotal = inlineNewTicketAppts.reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  // Comprehensive Total Tip — every tip source (main/sameday/extraTip/addon/cav/deposit already
  // folded into totalTipCash/totalTipCard above), plus standalone + inline new-ticket-purchase tips.
  const tipCashAllSources = totalTipCash + tpTipCash + inlineNTTipCash;
  const tipCardAllSources = totalTipCard + tpTipCard + inlineNTTipCard;
  const totalTipAllCC = tipCashAllSources + tipCardAllSources;
  // Multiple retail items can share one appointment; gc used against retail is distributed
  // across items proportionally to each item's share of the appointment's retail total.
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
  const inlineRetailCash = inlineRetailAppts.reduce((s,a)=>s+retailCashCardSplit(a).cash,0);
  const inlineRetailCard = inlineRetailAppts.reduce((s,a)=>s+retailCashCardSplit(a).card,0);
  const inlineRetailTotal = inlineRetailAppts.reduce((s,a)=>s+getRetailItems(a).reduce((s2,it)=>s2+Number(it.amount||0),0)-gcAlloc(a).gcRetail,0);
  const inlineGCCash = inlineGCAppts.filter(a=>a.giftCardPurchasePaymentType==="cash").reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);
  const inlineGCCard = inlineGCAppts.filter(a=>a.giftCardPurchasePaymentType!=="cash").reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);
  const inlineGCTotal = inlineGCAppts.reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);
  // Treatment (Cash/Card) — regular+sameday+addon+cav revenue, plus deposit/GC/cancellation, plus
  // any ticket purchased today (standalone ticket-purchase section or an inline チケット新規購入
  // add-on) — all real money received today that isn't a retail product.
  const cashTreatmentAll = totalCash + depositAllCash + tpCash + inlineNTCash;
  const cardTreatmentAll = totalCard + depositAllCard + tpCard + inlineNTCard;
  // Product (Cash/Card) — standalone retail sales plus inline 物販購入 recorded inside a visit.
  const cashProductStd = retails.filter(r => r.paymentType === "cash").reduce((s, r) => s + Number(r.price || 0), 0) + inlineRetailCash + spCash;
  const cardProductStd = retails.filter(r => r.paymentType === "card").reduce((s, r) => s + Number(r.price || 0), 0) + inlineRetailCard + spCard;
  const totalSalesAll = cashTreatmentAll + cashProductStd + cardTreatmentAll + cardProductStd;
  // All money actually received today by tender, tip included — this is what should match
  // Square's own Payments totals for the day (used by the Square照合 check below).
  const sheetCashTotal = r2(cashTreatmentAll + cashProductStd + tipCashAllSources);
  const sheetCardTotal = r2(cardTreatmentAll + cardProductStd + tipCardAllSources);
  const sheetCardTip = r2(tipCardAllSources);

  // 新規チケット販売（当日購入、または「当日の追加購入」→チケット新規購入タグ）— チケット消化とは別の集計。
  // 機械担当と2人で入って売れた場合は、件数・金額とも0.5ずつに割る。
  const newTicketByTherapist = {};
  THERAPISTS.forEach(t => { newTicketByTherapist[t] = { count: 0, amount: 0 }; });
  appointments.filter(a => !a.isCavSlot).forEach(a => {
    const events = [];
    if (a.isSameDayTicket && Number(a.packagePrice || 0) > 0) events.push(Number(a.packagePrice));
    if ((a.purchaseTags || []).includes("newTicket") && Number(a.newTicketAmount || 0) > 0) events.push(Number(a.newTicketAmount));
    events.forEach(amount => {
      const shared = a.therapist && a.cavTherapist && a.therapist !== a.cavTherapist;
      if (a.therapist && newTicketByTherapist[a.therapist]) {
        newTicketByTherapist[a.therapist].count += shared ? 0.5 : 1;
        newTicketByTherapist[a.therapist].amount += shared ? amount / 2 : amount;
      }
      if (shared && newTicketByTherapist[a.cavTherapist]) {
        newTicketByTherapist[a.cavTherapist].count += 0.5;
        newTicketByTherapist[a.cavTherapist].amount += amount / 2;
      }
    });
  });

  const summaryByTherapist = THERAPISTS.map(t => {
    const appts = regularAppts.filter(a => a.therapist === t);
    const cavAppts = regularAppts.filter(a => a.cavTherapist === t);
    const ticketAppts2 = ticketAppts.filter(a => a.therapist === t);
    const ticketCavAppts = ticketAppts.filter(a => a.cavTherapist === t);
    const gcAppts2 = gcAppts.filter(a => a.therapist === t);
    const revenue = appts.reduce((s, a) => s + Number(a.price || 0), 0);
    const tips = appts.reduce((s, a) => s + Number(a.tip || 0), 0);
    const tipCash = appts.filter(a => a.tipPaymentType === "cash").reduce((s, a) => s + Number(a.tip || 0), 0);
    const tipCard = appts.filter(a => a.tipPaymentType === "card").reduce((s, a) => s + Number(a.tip || 0), 0);
    const cavRevenue = [...appts.filter(a => a.cavTherapist === t), ...cavAppts].reduce((s, a) => s + Number(a.cavPrice || 0), 0);
    const ticketRevenue = ticketAppts2.reduce((s, a) => s + Number(a.price || 0), 0);
    const ticketCavRevenue = ticketCavAppts.reduce((s, a) => s + Number(a.cavPrice || 0), 0);
    const ticketTips = ticketAppts2.reduce((s, a) => s + Number(a.tip || 0), 0);
    const gcRevenue = gcAppts2.reduce((s, a) => s + Number(a.price || 0) + Number(a.cavPrice || 0), 0);
    const gcTips = gcAppts2.reduce((s, a) => s + Number(a.tip || 0) + Number(a.cavTip || 0), 0);
    const byType = {}; CUSTOMER_TYPES.forEach(ct => { byType[ct] = appointments.filter(a => a.therapist === t && a.customerType === ct).length; });
    // 人数: 機械担当が別にいる場合は、ボディ・機械それぞれ0.5人分として数える（1visitを二人で分ける）
    // 通常施術・チケット消化どちらも同じルールで数える
    const clients = appts.reduce((s, a) => s + (a.cavTherapist ? 0.5 : 1), 0) + cavAppts.length * 0.5
      + ticketAppts2.reduce((s, a) => s + (a.cavTherapist ? 0.5 : 1), 0) + ticketCavAppts.length * 0.5;
    const newTicket = newTicketByTherapist[t];
    return { therapist: t, revenue, tips, tipCash, tipCard, cavRevenue, ticketRevenue, ticketCavRevenue, ticketTips, gcRevenue, gcTips, clients, ticketClients: ticketAppts2.length, gcClients: gcAppts2.length, byType, newTicketCount: newTicket.count, newTicketAmount: newTicket.amount };
  });

  const visibleTherapists = selectedTherapist === "All"
    ? workingStaff
    : workingStaff.filter(t => t === selectedTherapist);

  return (
    <div style={{ fontFamily: "'Inter','Helvetica Neue',sans-serif", background: "#F7F4EE", minHeight: "100vh", color: "#1a1a1a" }}>
      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#0D4F4F,#1A6B5E)", padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ color: "#A8D5C8", fontSize: 11, letterSpacing: 3, textTransform: "uppercase", fontWeight: 600 }}>Daily Sheet</div>
          <div style={{ color: "#fff", fontSize: 22, fontWeight: 700 }}>🌺 Spa Schedule</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <input type="date" value={date} onChange={e => setDate(e.target.value)}
            style={{ padding: "8px 12px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontSize: 14 }} />
          <button onClick={fetchSquare} disabled={squareLoading}
            style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: squareLoading ? "#666" : "#E8A84A", color: "#fff", fontWeight: 700, cursor: squareLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {squareLoading ? "⏳ 取得中..." : "□ Square同期"}
          </button>
          <button onClick={syncOnlineGiftCards} disabled={gcSyncLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: gcSyncLoading ? "#666" : "#B45309", color: "#fff", fontWeight: 700, cursor: gcSyncLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {gcSyncLoading ? "⏳ 取得中..." : "🎁 ギフトカード購入取得"}
          </button>
          <button onClick={syncSquareDeposits} disabled={depositSyncLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: depositSyncLoading ? "#666" : "#00796B", color: "#fff", fontWeight: 700, cursor: depositSyncLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {depositSyncLoading ? "⏳ 取得中..." : "💰 デポジット自動取得"}
          </button>
          <button onClick={checkSquareReconciliation} disabled={reconcileLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: reconcileLoading ? "#666" : "#4A6572", color: "#fff", fontWeight: 700, cursor: reconcileLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {reconcileLoading ? "⏳ 照合中..." : "🔍 Square照合"}
          </button>
          <button onClick={handleLockToggle}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: locked ? "#C62828" : "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            {locked ? "🔒 確定済み（解除）" : "🔓 この日を確定する"}
          </button>
          <label style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            📥 データ復元
            <input type="file" accept="application/json" onChange={handleImportBackup} style={{ display: "none" }} />
          </label>
        </div>
      </div>
      {locked && (
        <div style={{ background: "#FFEBEE", padding: "8px 20px", fontSize: 13, color: "#C62828", fontWeight: 700, borderBottom: "2px solid #C62828" }}>
          🔒 この日は確定済みです。編集・削除するには「確定済み」ボタンからPINを入力して解除してください。
        </div>
      )}
      {squareStatus && <div style={{ background: "#FFF3CD", padding: "8px 20px", fontSize: 13, color: "#856404" }}>⚠️ {squareStatus}</div>}

      {reconcileResult && (() => {
        // Square doesn't record a separate tip amount for CASH payments (only card payments
        // go through a tip-prompt screen) — a cash sale is just one lump total in Square, so
        // there's nothing to compare the sheet's cash-tip figure against. Card tips, on the
        // other hand, are always captured separately and can be checked directly.
        const rows = [
          { label: "現金 合計（施術＋物販＋チップ）", sheet: sheetCashTotal, square: reconcileResult.cashTotal },
          { label: "カード 合計（施術＋物販＋チップ）", sheet: sheetCardTotal, square: reconcileResult.cardTotal },
          { label: "カード チップ", sheet: sheetCardTip, square: reconcileResult.cardTip },
        ];
        return (
          <div style={{ background: "#fff", margin: "10px 20px", borderRadius: 10, border: "1px solid #DDD", padding: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#0D4F4F" }}>🔍 Square照合結果</div>
              <button onClick={() => setReconcileResult(null)}
                style={{ border: "none", background: "none", color: "#999", cursor: "pointer", fontSize: 13 }}>✕ 閉じる</button>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                <thead>
                  <tr style={{ background: "#F7F4EE" }}>
                    <th style={{ textAlign: "left", padding: "6px 10px" }}>項目</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>シート入力</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>Square実績</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>差額</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map(r => {
                    const diff = r2(r.sheet - r.square);
                    const mismatch = Math.abs(diff) > 0.01;
                    return (
                      <tr key={r.label} style={{ borderTop: "1px solid #EEE" }}>
                        <td style={{ padding: "6px 10px" }}>{r.label}</td>
                        <td style={{ textAlign: "right", padding: "6px 10px" }}>{formatCurrency(r.sheet)}</td>
                        <td style={{ textAlign: "right", padding: "6px 10px" }}>{formatCurrency(r.square)}</td>
                        <td style={{ textAlign: "right", padding: "6px 10px", fontWeight: 700, color: mismatch ? "#C62828" : "#2E7D32" }}>
                          {mismatch ? `⚠️ ${diff > 0 ? "+" : ""}${formatCurrency(diff)}` : "✓ 一致"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: "#999", marginTop: 8 }}>
              ※ 現金チップはスクエア側で内訳が記録されないため照合対象外です（現金は合計金額のみ比較しています）
            </div>
          </div>
        );
      })()}

      {/* Tabs */}
      <div style={{ background: "#fff", borderBottom: "2px solid #E8E4DC", display: "flex", padding: "0 20px", overflowX: "auto" }}>
        {[["schedule","📅 スケジュール"],["summary","📊 集計"],["payroll","💴 給料集計"]].map(([tab, label]) => (
          <button key={tab} onClick={() => setActiveTab(tab)} style={{ padding: "12px 20px", border: "none", background: "none", cursor: "pointer", fontSize: 14, fontWeight: 600, color: activeTab === tab ? "#0D4F4F" : "#888", borderBottom: activeTab === tab ? "2px solid #0D4F4F" : "2px solid transparent", marginBottom: -2, whiteSpace: "nowrap" }}>
            {label}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", flexWrap: "wrap" }}>
          <select value={selectedTherapist} onChange={e => setSelectedTherapist(e.target.value)} style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid #DDD", fontSize: 13 }}>
            <option value="All">全員</option>
            {workingStaff.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          <button onClick={() => setShowStaffPicker(p => !p)}
            style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #0D4F4F", background: showStaffPicker ? "#0D4F4F" : "#fff", color: showStaffPicker ? "#fff" : "#0D4F4F", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>
            👥 出勤設定
          </button>
        </div>
      </div>

      {/* Working staff picker */}
      {showStaffPicker && (
        <div style={{ background: "#E8F5E9", padding: "10px 20px", borderBottom: "1px solid #C8E6C9", display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 12, fontWeight: 700, color: "#2E7D32" }}>今日の出勤スタッフ：</span>
          {THERAPISTS.map(t => (
            <button key={t} onClick={() => setWorkingStaff(prev => {
              const next = prev.includes(t) ? prev.filter(x => x !== t) : [...prev, t];
              save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips, next);
              return next;
            })} style={{
              padding: "4px 12px", borderRadius: 20, border: `2px solid ${workingStaff.includes(t) ? "#2E7D32" : "#CCC"}`,
              background: workingStaff.includes(t) ? "#2E7D32" : "#F5F5F5",
              color: workingStaff.includes(t) ? "#fff" : "#888",
              cursor: "pointer", fontWeight: 600, fontSize: 12
            }}>{t}</button>
          ))}
          <button onClick={() => setShowStaffPicker(false)}
            style={{ marginLeft: "auto", padding: "4px 12px", borderRadius: 8, border: "none", background: "#0D4F4F", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
            完了
          </button>
        </div>
      )}

      {/* Schedule */}
      {activeTab === "schedule" && (
        <div style={{ padding: "16px 12px", overflowX: "auto" }}>

          {/* 本日の来店人数 — cavSlot(同じお客様の機械担当分の複製行)を除いた実来店数 */}
          <div style={{ display: "inline-block", background: "#0D4F4F", color: "#fff", borderRadius: 20, padding: "6px 16px", fontSize: 13, fontWeight: 700, marginBottom: 12 }}>
            👥 本日の来店人数: {appointments.filter(a => !a.isCavSlot).length}名
          </div>

          {/* デポジット済み来店予定 */}
          {depositsForDate.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 14, border: "2px solid #1565C0" }}>
              <div style={{ fontWeight: 800, fontSize: 13, color: "#1565C0", marginBottom: 10 }}>
                💰 本日来店予定 — 領収済みのお客様 {depositsForDate.length}名
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {depositsForDate.map((dep, i) => (
                  <div key={dep.id || i} style={{
                    background: "#fff", borderRadius: 10, padding: "10px 14px",
                    border: "2px solid #1565C0", minWidth: 170,
                    boxShadow: "0 2px 6px rgba(21,101,192,0.12)"
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ fontWeight: 800, fontSize: 14, color: "#0D47A1" }}>{dep.clientName}</span>
                      <span style={{ fontSize: 9, fontWeight: 700, padding: "1px 6px", borderRadius: 8, background: dep.type === "giftcard" ? "#FFF3E0" : "#E3F2FD", color: dep.type === "giftcard" ? "#E65100" : "#1565C0" }}>
                        {dep.type === "giftcard" ? "🎁 全額ギフト" : "💰 デポジット"}
                      </span>
                    </div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#1565C0", marginTop: 3 }}>
                      💰 ${dep.amount}
                      <span style={{ fontSize: 11, fontWeight: 400, color: "#888", marginLeft: 6 }}>
                        {dep.paymentType === "cash" ? "現金" : dep.paymentType === "card" ? "カード" : "Check"}
                      </span>
                      {Number(dep.tip) > 0 && <span style={{ fontSize: 11, color: "#E65100", marginLeft: 6 }}>＋チップ${dep.tip}</span>}
                    </div>
                    {dep.appointmentTime && (
                      <div style={{ fontSize: 12, color: "#555", marginTop: 3 }}>🕐 {dep.appointmentTime}</div>
                    )}
                    {dep.notes && <div style={{ fontSize: 11, color: "#666", marginTop: 3 }}>{dep.notes}</div>}
                    <div style={{ fontSize: 10, color: "#AAA", marginTop: 5 }}>支払日：{dep.recordedDate}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 800 }}>
            <thead>
              <tr>
                <th style={{ width: 70, padding: "8px 4px", fontSize: 11, color: "#888", textAlign: "left", borderBottom: "2px solid #0D4F4F" }}>時間</th>
                {visibleTherapists.map(t => (
                  <th key={t} style={{ padding: "8px 6px", fontSize: 13, fontWeight: 700, color: "#0D4F4F", textAlign: "center", borderBottom: "2px solid #0D4F4F", minWidth: 140 }}>
                    {t}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {HOURS.map(hour => (
                <tr key={hour} style={{ borderBottom: "1px solid #EEE" }}>
                  <td style={{ padding: "6px 4px", fontSize: 11, color: "#999", verticalAlign: "top", whiteSpace: "nowrap" }}>{formatTime(hour)}</td>
                  {visibleTherapists.map(t => {
                    const appts = appointments.filter(a => {
                      if (a.therapist !== t) return false;
                      const [h] = (a.startTime || "").split(":").map(Number);
                      return h === hour;
                    });
                    const addonItems = appointments.flatMap(a => {
                      if (a.isCavSlot) return [];
                      const [h] = (a.startTime || "").split(":").map(Number);
                      if (h !== hour) return [];
                      return (a.addons || []).filter(ad => ad.therapist === t).map(ad => ({ parentAppt: a, addon: ad }));
                    });
                    const hasContent = appts.length > 0 || addonItems.length > 0;
                    return (
                      <td key={t} style={{ padding: "4px", verticalAlign: "top" }}
                        onClick={() => !hasContent && !locked && setEditingAppt({ ...EMPTY_APPOINTMENT, id: `m-${Date.now()}`, therapist: t, startTime: `${String(hour).padStart(2,"0")}:00` })}>
                        {appts.map(a => <ApptCard key={a.id} appt={a} allAppointments={appointments} onClick={() => {
                          if (locked) { showToast("🔒 この日は確定済みです。解除してから編集してください", "error"); return; }
                          if (a.isCavSlot) {
                            const parent = appointments.find(p => p.id === a.parentId);
                            if (parent) openApptForEdit(parent);
                          } else {
                            openApptForEdit(a);
                          }
                        }} />)}
                        {addonItems.map(({ parentAppt, addon }) => {
                          const aSvc = Number(addon.price||0);
                          const aTip = Number(addon.tip||0);
                          const aTotal = r2(aSvc + aTip);
                          const aColor = addon.countsAsRevenue === true ? REVENUE_COLOR : addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#005F4A";
                          const svcIcon = addon.paymentType === "card" ? "💳" : addon.paymentType === "cash" ? "💵" : "";
                          const tipIcon = addon.tipPaymentType === "card" ? "💳" : addon.tipPaymentType === "cash" ? "💵" : "";
                          return (
                            <div key={`${parentAppt.id}-${addon.id}`}
                              onClick={e => { e.stopPropagation(); openApptForEdit(parentAppt); }}
                              style={{ background: "#E0F2F1", border: "1.5px solid #00796B", borderRadius: 8, padding: "5px 7px", marginBottom: 3, cursor: "pointer" }}>
                              <div style={{ fontSize: 10, fontWeight: 800, color: "#00695C" }}>➕ オプション</div>
                              <div style={{ fontSize: 11, fontWeight: 700, color: "#004D40" }}>{parentAppt.clientName}</div>
                              <div style={{ fontSize: 10, color: "#00796B" }}>
                                {addon.serviceName || "オプション"}{addon.ticketCurrent ? ` ${addon.ticketCurrent}/${parentAppt.ticketTotal||3}` : ""}
                              </div>
                              {aTotal > 0 && (
                                <div style={{ fontSize: 10, fontWeight: 700, color: aColor }}>
                                  <div>{aSvc}{svcIcon}</div>
                                  {aTip > 0 && <div>{aTip}{tipIcon}</div>}
                                  <div style={{ fontWeight: 800 }}>{aTotal}</div>
                                </div>
                              )}
                            </div>
                          );
                        })}
                        {!hasContent && <div style={{ height: 36, border: "1px dashed #DDD", borderRadius: 6, cursor: "pointer", opacity: 0.4 }} />}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>

          {/* Retail */}
          <SectionBox title="🛍️ 物販" color="#6A1B9A" onAdd={() => setEditingRetail({ ...EMPTY_RETAIL, id: Date.now() })} disabled={locked}>
            {retails.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {retails.map(r => (
              <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ flex: 1, fontSize: 14 }}>{r.item || "（未入力）"}</span>
                <span style={{ fontWeight: 700, color: "#6A1B9A" }}>{formatCurrency(r.price)}</span>
                <PayBadge type={r.paymentType} />
                <span style={{ fontSize: 12, color: "#888" }}>
                  {(r.sellers || (r.soldBy ? [{ therapist: r.soldBy, amount: r.price }] : [])).filter(sel => sel.therapist).map(sel => `${sel.therapist}$${sel.amount||0}`).join(" / ")}
                </span>
                <button onClick={() => setEditingRetail(r)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                <button onClick={() => deleteRetail(r.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
              </div>
            ))}
          </SectionBox>

          {/* Deposits */}
          <SectionBox title="💰 デポジット・ギフトカード・キャンセル料" color="#1565C0" onAdd={() => setEditingDeposit({ ...EMPTY_DEPOSIT, id: Date.now() })} disabled={locked}>
            {deposits.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {deposits.map(d => (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, background: d.type === "deposit" ? "#E3F2FD" : d.type === "cancellation" ? "#FFEBEE" : "#FFF3E0", color: d.type === "deposit" ? "#1565C0" : d.type === "cancellation" ? "#C62828" : "#E65100", padding: "2px 8px", borderRadius: 10, fontWeight: 700 }}>
                  {d.type === "deposit" ? "デポジット" : d.type === "cancellation" ? "❌ キャンセル料" : "ギフトカード"}
                </span>
                <span style={{ flex: 1, fontSize: 14 }}>{d.clientName || "—"}</span>
                <span style={{ fontWeight: 700 }}>{formatCurrency(d.amount)}</span>
                <PayBadge type={d.paymentType} />
                <button onClick={() => setEditingDeposit(d)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                <button onClick={() => deleteDeposit(d.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
              </div>
            ))}
          </SectionBox>

          {/* Forgotten tips/payments — a past visit's payment (usually a cash tip) staff forgot to
              ring in that day. Added to *today's* totals (register is physically counted today)
              and credited to the therapist's payroll for today, without re-creating the old
              appointment (which would wrongly count as a new visit toward today's customer stats). */}
          <SectionBox title="🙏 打ち忘れ入力（デポジット/チップ）" color="#00695C" onAdd={() => setEditingForgottenTip({ ...EMPTY_FORGOTTEN_TIP, id: Date.now() })} disabled={locked}>
            {forgottenTips.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {forgottenTips.map(ft => {
              const svc = Number(ft.serviceAmount || 0);
              const tip = Number(ft.tipAmount || 0);
              return (
                <div key={ft.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{ft.clientName || "—"}</span>
                  {ft.therapist && <span style={{ fontSize: 11, color: "#00695C", background: "#E0F2F1", borderRadius: 8, padding: "2px 8px" }}>{ft.therapist}</span>}
                  {svc > 0 && <span style={{ fontWeight: 700, color: "#00695C" }}>施術{formatCurrency(svc)}</span>}
                  {tip > 0 && <span style={{ fontWeight: 700, color: "#00695C" }}>チップ{formatCurrency(tip)}</span>}
                  <PayBadge type={ft.paymentType} />
                  {ft.originalDate && <span style={{ fontSize: 11, color: "#888" }}>来店日 {ft.originalDate}</span>}
                  <button onClick={() => setEditingForgottenTip(ft)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                  <button onClick={() => deleteForgottenTip(ft.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                </div>
              );
            })}
            {forgottenTips.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#00695C", fontWeight: 700, textAlign: "right" }}>
                施術 {formatCurrency(totalForgottenService)}　チップ {formatCurrency(totalForgottenTip)}
              </div>
            )}
          </SectionBox>

          {/* Staff Purchases 社販 */}
          <SectionBox title="👩‍💼 社販（スタッフ購入・施術）" color="#37474F" onAdd={() => setEditingStaffPurchase({ ...EMPTY_STAFF_PURCHASE, id: Date.now() })} disabled={locked}>
            {staffPurchases.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {staffPurchases.map(sp => {
              const items = getStaffPurchaseItems(sp);
              const total = items.reduce((s, it) => s + Number(it.amount || 0), 0);
              return (
                <div key={sp.id} style={{ padding: "8px 0", borderBottom: "1px solid #F0F0F0" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span style={{ fontWeight: 700, fontSize: 14, color: "#37474F" }}>{sp.staffName || "—"}</span>
                    {items.length === 1 && items[0].productName && items[0].productName !== "__other__" && (
                      <span style={{ fontSize: 11, color: "#37474F", background: "#ECEFF1", borderRadius: 8, padding: "2px 8px" }}>{items[0].productName}</span>
                    )}
                    {items.length > 1 && (
                      <span style={{ fontSize: 11, color: "#37474F", background: "#ECEFF1", borderRadius: 8, padding: "2px 8px" }}>{items.length}点</span>
                    )}
                    <span style={{ fontWeight: 700, color: "#37474F" }}>{formatCurrency(total)}</span>
                    {items.length === 1 && <PayBadge type={items[0].paymentType} />}
                    {sp.notes && <span style={{ fontSize: 11, color: "#888" }}>{sp.notes}</span>}
                    <button onClick={() => setEditingStaffPurchase(sp)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                    <button onClick={() => deleteStaffPurchase(sp.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                  </div>
                  {items.length > 1 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 4, paddingLeft: 4 }}>
                      {items.map((it, idx) => (
                        <div key={idx} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "#555" }}>
                          <span style={{ flex: 1 }}>{it.productName || "（商品名なし）"}</span>
                          <span style={{ fontWeight: 700 }}>{formatCurrency(it.amount)}</span>
                          <PayBadge type={it.paymentType} />
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </SectionBox>

          {/* Ticket New Purchases */}
          <SectionBox title="🎟️ チケット新規購入（電話・当日）" color="#B71C1C" onAdd={() => setEditingTicketPurchase({ ...EMPTY_TICKET_PURCHASE, id: Date.now() })} disabled={locked}>
            {ticketPurchases.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {ticketPurchases.map(tp => {
              const tip = Number(tp.tip || 0);
              const amt = Number(tp.amount || 0);
              return (
                <div key={tp.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{tp.clientName || "—"}</span>
                  {tp.packageName && <span style={{ fontSize: 11, color: "#B71C1C", background: "#FFEBEE", borderRadius: 8, padding: "2px 8px" }}>{tp.packageName}</span>}
                  <span style={{ fontWeight: 700, color: "#B71C1C" }}>{formatCurrency(amt)}</span>
                  {tip > 0 && <span style={{ fontSize: 11, color: "#E65100" }}>チップ{formatCurrency(tip)}</span>}
                  {tp.splitPayment
                    ? <span style={{ fontSize: 11, color: "#B71C1C", background: "#FFEBEE", borderRadius: 8, padding: "2px 8px" }}>💵{formatCurrency(tp.cashPortion||0)}＋💳{formatCurrency(tp.cardPortion||0)}</span>
                    : <PayBadge type={tp.paymentType} />}
                  <button onClick={() => setEditingTicketPurchase(tp)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                  <button onClick={() => deleteTicketPurchase(tp.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                </div>
              );
            })}
            {ticketPurchases.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#B71C1C", fontWeight: 700, textAlign: "right" }}>
                合計 {formatCurrency(tpTotal)}　チップ {formatCurrency(tpTipTotal)}
              </div>
            )}
          </SectionBox>

          {/* Refunds — for a past visit's refund processed today (deducts from today's totals only) */}
          <SectionBox title="🔙 返金（リファンド）" color="#5D4037" onAdd={() => setEditingRefund({ ...EMPTY_REFUND, id: Date.now() })} disabled={locked}>
            {refunds.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {refunds.map(rf => {
              const svc = Number(rf.serviceAmount || 0);
              const tip = Number(rf.tipAmount || 0);
              return (
                <div key={rf.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{rf.clientName || "—"}</span>
                  {rf.originalDate && <span style={{ fontSize: 11, color: "#5D4037", background: "#EFEBE9", borderRadius: 8, padding: "2px 8px" }}>来店日: {rf.originalDate}</span>}
                  {svc > 0 && <span style={{ fontWeight: 700, color: "#C62828" }}>施術 -{formatCurrency(svc)}</span>}
                  {tip > 0 && <span style={{ fontWeight: 700, color: "#C62828" }}>チップ -{formatCurrency(tip)}</span>}
                  <PayBadge type={rf.paymentType} />
                  {rf.notes && <span style={{ fontSize: 11, color: "#888" }}>{rf.notes}</span>}
                  <button onClick={() => setEditingRefund(rf)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                  <button onClick={() => deleteRefund(rf.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                </div>
              );
            })}
            {refunds.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#C62828", fontWeight: 700, textAlign: "right" }}>
                施術 -{formatCurrency(totalRefundService)}　チップ -{formatCurrency(totalRefundTip)}
              </div>
            )}
          </SectionBox>
        </div>
      )}

      {/* Summary */}
      {activeTab === "summary" && (
        <div style={{ padding: 16 }}>

          {/* === SALES REPORT 形式 === */}
          <div style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 14, boxShadow: "0 2px 8px rgba(0,0,0,0.10)" }}>
            <div style={{ fontWeight: 800, fontSize: 15, color: "#0D4F4F", marginBottom: 12, borderBottom: "2px solid #0D4F4F", paddingBottom: 8 }}>
              📋 本日売上サマリー　<span style={{ fontSize: 12, color: "#888", fontWeight: 400 }}>Sales Report 形式</span>
            </div>

            {/* 来店人数 & 総売上 */}
            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              <div style={{ background: "#F0F7FF", borderRadius: 10, padding: "10px 18px", textAlign: "center", minWidth: 90 }}>
                <div style={{ fontSize: 11, color: "#888" }}>来店人数</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#0D4F4F" }}>{appointments.filter(a => !a.isCavSlot).length}<span style={{ fontSize: 13 }}>名</span></div>
              </div>
              <div style={{ background: "#FFF8E1", borderRadius: 10, padding: "10px 18px", textAlign: "center", flex: 1, minWidth: 120 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Total Sales（施術＋物販）</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#C62828" }}>{formatCurrency(totalSalesAll)}</div>
              </div>
              <div style={{ background: "#F3E5F5", borderRadius: 10, padding: "10px 18px", textAlign: "center", flex: 1, minWidth: 120 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Total Sales ＋ Tip</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#6A1B9A" }}>{formatCurrency(totalSalesAll + totalTipAllCC)}</div>
              </div>
            </div>

            {/* Cash セクション */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", background: "#2E7D32", padding: "5px 12px", borderRadius: 6, marginBottom: 8, display: "inline-block" }}>💵 CASH</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                {[
                  { label: "Treatment", sub: "デポジット・GC・キャンセル料・チケット購入込み", value: cashTreatmentAll, color: "#2E7D32" },
                  { label: "Product", value: cashProductStd, color: "#2E7D32" },
                  { label: "Total Cash", value: cashTreatmentAll + cashProductStd, color: "#1B5E20", bold: true },
                  { label: "Tip", value: tipCashAllSources, color: "#558B2F" },
                ].map(s => (
                  <div key={s.label} style={{ background: "#F1F8E9", borderRadius: 8, padding: "8px 10px", textAlign: "center", border: s.bold ? "1.5px solid #2E7D32" : "none" }}>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
                    {s.sub && <div style={{ fontSize: 8, color: "#999", marginBottom: 3 }}>({s.sub})</div>}
                    <div style={{ fontSize: 17, fontWeight: s.bold ? 800 : 700, color: s.color }}>{formatCurrency(s.value)}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Card セクション */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", background: "#1565C0", padding: "5px 12px", borderRadius: 6, marginBottom: 8, display: "inline-block" }}>💳 CARD</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                {[
                  { label: "Treatment", sub: "デポジット・GC・キャンセル料・チケット購入込み", value: cardTreatmentAll, color: "#1565C0" },
                  { label: "Product", value: cardProductStd, color: "#1565C0" },
                  { label: "Total Card", value: cardTreatmentAll + cardProductStd, color: "#0D47A1", bold: true },
                  { label: "Tip", value: tipCardAllSources, color: "#1976D2" },
                ].map(s => (
                  <div key={s.label} style={{ background: "#E8F0FE", borderRadius: 8, padding: "8px 10px", textAlign: "center", border: s.bold ? "1.5px solid #1565C0" : "none" }}>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
                    {s.sub && <div style={{ fontSize: 8, color: "#999", marginBottom: 3 }}>({s.sub})</div>}
                    <div style={{ fontSize: 17, fontWeight: s.bold ? 800 : 700, color: s.color }}>{formatCurrency(s.value)}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Total Tip */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ background: "#FFF3E0", borderRadius: 8, padding: "8px 10px", textAlign: "center", border: "1.5px solid #E65100" }}>
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Total Tip（Cash＋Card）</div>
                <div style={{ fontSize: 20, fontWeight: 800, color: "#E65100" }}>{formatCurrency(totalTipAllCC)}</div>
              </div>
            </div>

            {/* その他内訳（上のTreatmentに含まれています） */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(110px,1fr))", gap: 8, marginBottom: 12 }}>
              {[
                { label: "デポジット受取", note: "Treatmentに計上済み", value: totalDepositAmt, color: "#0277BD", bg: "#E1F5FE" },
                ...(totalDepositApplied > 0 ? [{ label: "💰 Payroll加算", value: totalDepositApplied, color: "#2E7D32", bg: "#C8E6C9", prefix: "+" }] : []),
                { label: "ギフトカード購入", note: "Treatmentに計上済み", value: totalGiftCard, color: "#00796B", bg: "#E0F2F1" },
                ...(totalGCUsed > 0 ? [{ label: "🎁 GC使用", value: totalGCUsed, color: "#00695C", bg: "#B2DFDB" }] : []),
                { label: "❌ キャンセル料", note: "Treatmentに計上済み", value: totalCancellation, color: "#C62828", bg: "#FFEBEE" },
                { label: "物販合計", value: totalRetail + inlineRetailTotal + spTotal, color: "#6A1B9A", bg: "#F3E5F5" },
              ].map(s => (
                <div key={s.label} style={{ background: s.bg, borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
                  {s.note && <div style={{ fontSize: 8, color: "#999", marginBottom: 3 }}>({s.note})</div>}
                  <div style={{ fontSize: 17, fontWeight: 700, color: s.color }}>{s.prefix || ""}{formatCurrency(s.value)}</div>
                </div>
              ))}
            </div>

            {/* Grand Total */}
            <div style={{ borderTop: "2px solid #0D4F4F", paddingTop: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#0D4F4F", marginBottom: 8 }}>📊 当日 GRAND TOTAL</div>
              {/* Treatment breakdown */}
              <div style={{ background: "#F9F9F9", borderRadius: 8, padding: "8px 12px", marginBottom: 8, fontSize: 11, color: "#555" }}>
                <span style={{ fontWeight: 700 }}>施術料合計（Deposit・GC・チケット購入含む）：</span>
                <span style={{ color: "#C62828", fontWeight: 800, fontSize: 14, marginLeft: 4 }}>{formatCurrency(totalCash + totalCard + totalDepositAmt + totalGiftCard + totalCancellation + tpTotal + inlineNTTotal + inlineGCTotal)}</span>
                <div style={{ marginTop: 4, color: "#999", fontSize: 10 }}>
                  Cash ${totalCash.toFixed(2)}　Card ${totalCard.toFixed(2)}
                  {totalDepositAmt > 0 && `　Deposit $${totalDepositAmt.toFixed(2)}`}
                  {totalGiftCard > 0 && `　GiftCard $${(totalGiftCard + inlineGCTotal).toFixed(2)}`}
                  {totalCancellation > 0 && `　キャンセル料 $${totalCancellation.toFixed(2)}`}
                  {(tpTotal + inlineNTTotal) > 0 && `　チケット購入 $${(tpTotal + inlineNTTotal).toFixed(2)}`}
                  {inlineRetailTotal > 0 && `　物販(inline) $${inlineRetailTotal.toFixed(2)}`}
                </div>
              </div>
              {/* Final totals grid */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6 }}>
                {[
                  { label: "Cash合計\n（全て）", value: totalCash + retails.filter(r=>r.paymentType==="cash").reduce((s,r)=>s+Number(r.price||0),0) + tpCash + inlineNTCash + inlineRetailCash + inlineGCCash + totalCancellationCash + spCash, color: "#2E7D32", bg: "#E8F5E9" },
                  { label: "Card合計\n（全て）", value: totalCard + retails.filter(r=>r.paymentType==="card").reduce((s,r)=>s+Number(r.price||0),0) + tpCard + inlineNTCard + inlineRetailCard + inlineGCCard + totalCancellationCard + spCard, color: "#1565C0", bg: "#E3F2FD" },
                  { label: "Tip合計", value: totalTipAllCC, color: "#E65100", bg: "#FFF3E0" },
                  { label: "🏆 TOTAL\n（全て込み）", value: totalCash + totalCard + totalDepositAmt + totalGiftCard + totalCancellation + totalRetail + totalTipAllCC + tpTotal + inlineNTTotal + inlineRetailTotal + inlineGCTotal + spTotal, color: "#fff", bg: "#0D4F4F", bold: true },
                ].map(s => (
                  <div key={s.label} style={{ background: s.bg, borderRadius: 8, padding: "8px 6px", textAlign: "center" }}>
                    <div style={{ fontSize: 9, color: s.color === "#fff" ? "#B2EBF2" : "#888", marginBottom: 3, whiteSpace: "pre-line", lineHeight: 1.3 }}>{s.label}</div>
                    <div style={{ fontSize: s.bold ? 18 : 16, fontWeight: 800, color: s.color }}>{formatCurrency(s.value)}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Same-day ticket purchase */}
          {sameDayAppts.length > 0 && (
            <div style={{ background: "#E8F5E9", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #2E7D32" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#2E7D32", marginBottom: 8 }}>🟢 当日購入チケット {sameDayAppts.length}件</div>
              {sameDayAppts.map(a => {
                const pkgTip = Number(a.packageTip ?? a.tip ?? 0);
                const pkgSvc = Number(a.packagePrice||0);
                const gc = Number(a.giftCardUsed||0);
                const gcSvc = Math.min(gc, pkgSvc);
                const gcTip = Math.min(gc - gcSvc, pkgTip);
                const extra = Number(a.extraTip||0);
                const extraSvc = Number(a.extraPrice||0);
                const total = pkgSvc + pkgTip - gcSvc - gcTip + extra + extraSvc;
                return (
                <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                    <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                    <span style={{ fontSize: 12, color: "#2E7D32" }}>{a.ticketMenu} {a.ticketCurrent > 0 ? `${a.ticketCurrent}/${a.ticketTotal}回目` : `購入のみ（${a.ticketTotal}回コース）`}</span>
                    <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                    <span style={{ fontSize: 11, color: "#2E7D32" }}>💰 {a.packageSplitPayment ? `💵$${a.packageCashPortion||0}＋💳$${a.packageCardPortion||0}` : `$${pkgSvc}${a.paymentType==="card"?"💳":"💵"}`} / チップ ${pkgTip}{a.tipPaymentType==="card"?"💳":"💵"}{extraSvc > 0 && <span style={{ color: "#F57F17" }}> +${extraSvc}{a.extraPricePaymentType==="card"?"💳":"💵"}</span>}{extra > 0 && <span style={{ color: "#F57F17" }}> +${extra}💝{a.extraTipPaymentType==="card"?"💳":"💵"}</span>}</span>
                    <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${pkgSvc + pkgTip + extra + extraSvc}</span>
                  </div>
                  {gc > 0 && (
                    <div style={{ marginTop: 4, fontSize: 10, color: NON_REVENUE_COLOR }}>
                      🔵 ギフトカード残高 ${gcSvc + gcTip} 使用（本日の売上には含まれません。実質の本日売上 ${total}）
                    </div>
                  )}
                </div>
                );
              })}
            </div>
          )}

          {/* Ticket (paid already) */}
          {pureTicketAppts.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #1565C0" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>🔵 領収済み（チケット消化）{pureTicketAppts.length}件</div>
              {pureTicketAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0) + Number(a.extraTip||0));
                const extraSvc = Number(a.extraPrice||0);
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#1565C0" }}>{a.ticketMenu} {a.ticketCurrent}/{a.ticketTotal}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4, flexWrap: "wrap", gap: 4 }}>
                      <span style={{ fontSize: 11, color: "#1565C0" }}>施術 ${svc} / チップ ${Number(a.tip||0)+Number(a.cavTip||0)}{extraSvc > 0 && <span style={{ color: "#F57F17" }}> +${extraSvc}{a.extraPricePaymentType==="card"?"💳":"💵"}</span>}{a.extraTip > 0 && <span style={{ color: "#F57F17" }}> +${a.extraTip}💝{a.extraTipPaymentType==="card"?"💳":"💵"}</span>}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${svc + tip + extraSvc}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* GC消化 (pre-purchased GC used today) */}
          {gcAppts.length > 0 && (
            <div style={{ background: "#FFFDE7", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #F59E0B" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#B45309", marginBottom: 8 }}>🎁 領収済み（GC消化）{gcAppts.length}件</div>
              {gcAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0));
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#B45309" }}>{a.serviceName || `${a.duration}分`}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                      <span style={{ fontSize: 11, color: "#B45309" }}>施術 ${svc} / チップ ${tip}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${svc + tip}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* PR無料 (comped treatments for influencers etc.) */}
          {promoAppts.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #1565C0" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>📸 PR無料（施術売上には計上されません）{promoAppts.length}件</div>
              {promoAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0));
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#1565C0" }}>{a.serviceName || `${a.duration}分`}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                      <span style={{ fontSize: 11, color: "#1565C0" }}>施術 ${svc} / チップ ${tip}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${svc + tip}</span>
                    </div>
                    {a.notes && <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>📝 {a.notes}</div>}
                  </div>
                );
              })}
            </div>
          )}

          {/* 顧客タイプ & 紹介元 */}
          <div style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 14, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
            <div style={{ fontWeight: 700, marginBottom: 10, color: "#0D4F4F" }}>顧客タイプ別</div>
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 12 }}>
              {CUSTOMER_TYPES.map(ct => {
                const colors = { RL:"#4CAF50", RT:"#2196F3", NL:"#FF9800", NT:"#9C27B0" };
                return (
                  <div key={ct} style={{ textAlign: "center", background: "#F9F9F9", borderRadius: 10, padding: "8px 16px", minWidth: 60 }}>
                    <div style={{ fontSize: 22, fontWeight: 800, color: colors[ct] }}>{appointments.filter(a => !a.isCavSlot && a.customerType === ct).length}</div>
                    <div style={{ fontSize: 12, color: "#888" }}>{ct}</div>
                  </div>
                );
              })}
            </div>
            {appointments.filter(a => !a.isCavSlot && (a.customerType === "NT" || a.customerType === "NL") && a.referralSource).length > 0 && (
              <>
                <div style={{ fontWeight: 700, marginBottom: 8, color: "#9C27B0", fontSize: 13 }}>📍 新規きっかけ</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {REFERRAL_SOURCES.map(src => {
                    const count = appointments.filter(a => !a.isCavSlot && a.referralSource === src).length;
                    if (!count) return null;
                    return (
                      <div key={src} style={{ background: "#F3E5F5", borderRadius: 8, padding: "6px 12px", textAlign: "center" }}>
                        <div style={{ fontSize: 18, fontWeight: 800, color: "#9C27B0" }}>{count}</div>
                        <div style={{ fontSize: 11, color: "#6A1B9A" }}>{src}</div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>

          {/* セラピスト別 */}
          <div style={{ background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)", overflowX: "auto" }}>
            <div style={{ fontWeight: 700, marginBottom: 12, color: "#0D4F4F" }}>セラピスト別</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: "2px solid #E0E0E0", background: "#F9F9F9" }}>
                  {["名前","人数",...CUSTOMER_TYPES,"チケット新規件数","チケット新規金額"].map(h => (
                    <th key={h} style={{ padding: "6px 8px", textAlign: h==="名前"?"left":"center", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {summaryByTherapist.filter(s => s.clients > 0 || s.ticketClients > 0).map(s => (
                  <tr key={s.therapist} style={{ borderBottom: "1px solid #F0F0F0" }}>
                    <td style={{ padding: "8px", fontWeight: 700, color: "#0D4F4F" }}>{s.therapist}</td>
                    <td style={{ textAlign: "center", padding: "8px" }}>{s.clients}</td>
                    {CUSTOMER_TYPES.map(ct => <td key={ct} style={{ textAlign: "center", padding: "8px", color: "#555" }}>{s.byType[ct]||0}</td>)}
                    <td style={{ textAlign: "center", padding: "8px", color: "#1565C0" }}>{s.newTicketCount || "—"}</td>
                    <td style={{ textAlign: "center", padding: "8px", color: "#1565C0", fontWeight: 700 }}>{s.newTicketAmount > 0 ? formatCurrency(s.newTicketAmount) : "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Payroll Tab */}
      {activeTab === "payroll" && <PayrollTab />}

      {/* Modals */}
      {editingAppt && <ApptModal appt={editingAppt} onSave={saveAppt} onDelete={() => deleteAppt(editingAppt.id)} onClose={() => setEditingAppt(null)}
        clientDeposits={(() => {
          const name = (editingAppt.clientName || "").toLowerCase().trim();
          if (!name) return [];
          const found = [];
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (!key?.startsWith("spa-sheet-")) continue;
            try {
              const d = JSON.parse(localStorage.getItem(key));
              const sheetDate = key.replace("spa-sheet-", "");
              (d.deposits || []).forEach(dep => {
                if ((dep.clientName || "").toLowerCase().trim() === name && (dep.type === "deposit" || dep.type === "giftcard")) {
                  found.push({ ...dep, sheetDate });
                }
              });
            } catch {}
          }
          return found.sort((a, b) => a.sheetDate.localeCompare(b.sheetDate));
        })()}
      />}
      {editingRetail && <RetailModal retail={editingRetail} onSave={saveRetail} onClose={() => setEditingRetail(null)} />}
      {editingDeposit && <DepositModal deposit={editingDeposit} onSave={saveDeposit} onDelete={() => { deleteDeposit(editingDeposit.id); setEditingDeposit(null); }} onClose={() => setEditingDeposit(null)} />}
      {editingTicketPurchase && <TicketPurchaseModal tp={editingTicketPurchase} onSave={saveTicketPurchase} onDelete={() => { deleteTicketPurchase(editingTicketPurchase.id); setEditingTicketPurchase(null); }} onClose={() => setEditingTicketPurchase(null)} />}
      {editingStaffPurchase && <StaffPurchaseModal sp={editingStaffPurchase} onSave={saveStaffPurchase} onDelete={() => { deleteStaffPurchase(editingStaffPurchase.id); setEditingStaffPurchase(null); }} onClose={() => setEditingStaffPurchase(null)} />}
      {editingRefund && <RefundModal rf={editingRefund} onSave={saveRefund} onDelete={() => { deleteRefund(editingRefund.id); setEditingRefund(null); }} onClose={() => setEditingRefund(null)} />}
      {editingForgottenTip && <ForgottenTipModal ft={editingForgottenTip} onSave={saveForgottenTip} onDelete={() => { deleteForgottenTip(editingForgottenTip.id); setEditingForgottenTip(null); }} onClose={() => setEditingForgottenTip(null)} />}

      {toast && (
        <div style={{ position: "fixed", bottom: 20, right: 20, padding: "12px 20px", borderRadius: 10, background: toast.type === "error" ? "#C62828" : toast.type === "info" ? "#1565C0" : "#0D4F4F", color: "#fff", fontWeight: 600, fontSize: 14, boxShadow: "0 4px 12px rgba(0,0,0,0.2)", zIndex: 9999 }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// Red = counts toward today's sales revenue (walk-in service, same-day ticket/GC purchase, extra tip).
// Blue = does NOT count toward today's revenue (redeeming a ticket/gift card paid for previously, applying a prior deposit).
const REVENUE_COLOR = "#C62828";
const NON_REVENUE_COLOR = "#1565C0";

function ApptCard({ appt, onClick, allAppointments }) {
  const typeColors = { RL: "#4CAF50", RT: "#2196F3", NL: "#FF9800", NT: "#9C27B0" };
  const isTicket = appt.isTicket;
  const isCavSlot = appt.isCavSlot;

  const isGiftCard = appt.isGiftCard;
  const isPromo = appt.isPromo;
  // A regular appointment can also be fully paid off an existing gift card balance
  // (giftCardUsed) rather than the isGiftCard type toggle — that money came in whenever the
  // gift card was originally purchased, not today, so it should read as non-revenue too.
  const totalOwed = Number(appt.price||0) + Number(appt.cavPrice||0) + Number(appt.tip||0) + Number(appt.cavTip||0);
  const gcFullyCovers = !isGiftCard && !isPromo && totalOwed > 0 && Number(appt.giftCardUsed||0) >= totalOwed;
  // Revenue today: regular walk-in visits, and same-day ticket purchases. Not revenue: redeeming
  // a previously-bought ticket or gift card (the money already came in on the day it was sold),
  // or a comped/PR treatment (no money is ever received for it).
  const isRevenueToday = !isGiftCard && !isPromo && !gcFullyCovers && (!isTicket || appt.isSameDayTicket);
  const amountColor = isRevenueToday ? REVENUE_COLOR : NON_REVENUE_COLOR;
  const borderColor = isCavSlot ? "#9C27B0" : isGiftCard ? "#F59E0B" : isPromo ? "#1565C0" : isTicket ? "#1565C0" : appt.paymentType === "cash" ? "#4CAF50" : "#2196F3";
  const bg = isCavSlot ? "#F9F0FF" : isGiftCard ? "#FFFDE7" : isPromo ? "#E3F2FD" : isTicket ? "#EEF5FF" : appt.paymentType === "cash" ? "#F1FBF3" : "#EEF5FF";

  if (isCavSlot) {
    const parent = (allAppointments || []).find(a => a.id === appt.parentId);
    const dep = Number(parent?.depositApplied || 0);
    let cavSvc = Number(appt.price||0);
    // Mirror the body card's deposit-inclusive minutes split (see the non-cav branch below and
    // PayrollTab) so this card shows the same deposit-adjusted share staff actually get paid.
    if (dep > 0 && parent && !isWeightLossService(parent.serviceName)) {
      const bodyReceived = Number(parent.price || 0);
      const bodyMins = Number(parent.duration || 0);
      const cavMins = Number(appt.duration || 0) || 15;
      const allMins = bodyMins + cavMins;
      if (allMins > 0) {
        const bodyShare = Math.round((bodyReceived + cavSvc + dep) * bodyMins / allMins * 100) / 100;
        cavSvc = Math.round((bodyReceived + cavSvc + dep - bodyShare) * 100) / 100;
      }
    }
    const cavTip = Number(appt.tip||0);
    const cavTotal = r2(cavSvc + cavTip);
    const svcIcon = appt.paymentType === "card" ? "💳" : appt.paymentType === "cash" ? "💵" : "";
    const tipIcon = appt.tipPaymentType === "card" ? "💳" : appt.tipPaymentType === "cash" ? "💵" : "";
    return (
      <div onClick={onClick} style={{ background: bg, border: `2px solid ${borderColor}`, borderRadius: 8, padding: "5px 8px", marginBottom: 3, cursor: "pointer", opacity: 0.9 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A" }}>⚡ 機械 {appt.isTicket ? "🎟️" : ""}</div>
        <div style={{ fontSize: 11, color: "#555" }}>
          {appt.clientName} <span style={{ color: "#888", fontSize: 10 }}>({appt.startTime}〜)</span>
        </div>
        {appt.bodyTherapist && (
          <div style={{ fontSize: 10, color: "#8E24AA" }}>with {appt.bodyTherapist}</div>
        )}
        {(cavSvc > 0 || cavTip > 0) && (
          <div style={{ fontSize: 11, color: amountColor, fontWeight: 700 }}>
            <div>{cavSvc}{svcIcon}</div>
            {cavTip > 0 && <div>{cavTip}{tipIcon}</div>}
            <div style={{ fontWeight: 800 }}>{cavTotal}</div>
          </div>
        )}
        {dep > 0 && !isWeightLossService(parent?.serviceName) && (
          <div style={{ color: "#2E7D32", fontSize: 9 }}>💰デポジット按分込み</div>
        )}
      </div>
    );
  }

  return (
    <div onClick={onClick} style={{ background: bg, border: `2px solid ${borderColor}`, borderRadius: 8, padding: "6px 8px", marginBottom: 3, cursor: "pointer" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <span style={{ fontSize: 12, fontWeight: 700, lineHeight: 1.3 }}>{appt.clientName || "（未設定）"}</span>
        <span style={{ fontSize: 10, fontWeight: 700, color: "#fff", background: typeColors[appt.customerType]||"#888", padding: "1px 5px", borderRadius: 8, marginLeft: 4, whiteSpace: "nowrap" }}>
          {appt.customerType}
        </span>
      </div>
      {(() => {
        const isSameDay = isTicket && appt.isSameDayTicket;
        const isRedemption = isTicket && !isSameDay;
        const dep = isSameDay ? Number(appt.packageDepositAmount||0) : Number(appt.depositApplied||0);
        const depDate = isSameDay ? appt.packageDepositDate : appt.depositPaidDate;

        // Red on this card is today's register-close figure — cash actually collected today —
        // NOT the payroll allocation. A prior-visit deposit already counted as revenue on the
        // day it was collected, so it must stay out of today's red total. We still need the
        // deposit-inclusive, minutes-split payroll share (same formula as PayrollTab) to find
        // the body therapist's fair cut of the deposit, then back the full deposit back out —
        // the machine (cav) card keeps its full payroll share untouched, matching how this
        // shop attributes deposits entirely against the body therapist's side.
        let bodySvc = Number(appt.price||0);
        if (dep > 0 && appt.cavTherapist && !isSameDay) {
          const cavSlot = (allAppointments || []).find(a => a.isCavSlot && a.parentId === appt.id);
          const cavReceived = Number(cavSlot?.price || 0);
          let bodyPayrollShare;
          if (isWeightLossService(appt.serviceName)) {
            bodyPayrollShare = r2(bodySvc + dep);
          } else {
            const bodyMins = Number(appt.duration || 0);
            const cavMins = Number(cavSlot?.duration || 0) || 15;
            const allMins = bodyMins + cavMins;
            bodyPayrollShare = allMins > 0 ? Math.round((bodySvc + cavReceived + dep) * bodyMins / allMins * 100) / 100 : r2(bodySvc + dep);
          }
          bodySvc = r2(bodyPayrollShare - dep);
        }
        const svc = r2(bodySvc + Number(appt.cavPrice||0));
        const tip = r2(Number(appt.tip||0) + Number(appt.cavTip||0));
        const extra = Number(appt.extraTip||0);
        const extraSvc = Number(appt.extraPrice||0);
        const paidIcon = isGiftCard ? "🎁" : isPromo ? "📸" : appt.paymentType === "card" ? "💳" : appt.paymentType === "cash" ? "💵" : "";
        const tipIcon = isGiftCard ? "🎁" : isPromo ? "📸" : appt.tipPaymentType === "card" ? "💳" : appt.tipPaymentType === "cash" ? "💵" : "";

        // Same-day ticket purchase: the real money is packagePrice/packageTip, not the staff-split price/tip.
        const pkgSvc = Number(appt.packagePrice||0);
        const pkgTip = Number(appt.packageTip ?? appt.tip ?? 0);

        const dispSvc = isSameDay ? pkgSvc : svc;
        const dispTip = isSameDay ? pkgTip : tip;
        const dispTotal = r2(dispSvc + dispTip);

        const svcText = isSameDay
          ? (appt.packageSplitPayment ? `💵$${appt.packageCashPortion||0}＋💳$${appt.packageCardPortion||0}` : `${dispSvc}${paidIcon}`)
          : !isRedemption && appt.svcSplitPayment ? `💵$${appt.svcCashPortion||0}＋💳$${appt.svcCardPortion||0}` : `${dispSvc}${paidIcon}`;
        const tipText = !isRedemption && !isSameDay && appt.tipSplitPayment
          ? `💵$${appt.tipCashPortion||0}＋💳$${appt.tipCardPortion||0}` : `${dispTip}${tipIcon}`;

        const courseName = isTicket ? (appt.serviceName || appt.ticketMenu) : (appt.serviceName || `${appt.duration}分`);
        const sessionSuffix = isRedemption && appt.ticketCurrent > 0 ? ` ${appt.ticketCurrent}/${appt.ticketTotal}`
          : isSameDay ? ` ×${appt.ticketTotal}回` : "";

        const gc = (!isGiftCard && !isPromo) ? Number(appt.giftCardUsed||0) : 0;

        return (
          <div style={{ fontSize: 11, marginTop: 2 }}>
            <div style={{ color: isGiftCard ? "#B45309" : isPromo ? "#1565C0" : "#222", fontWeight: 700, lineHeight: 1.4 }}>
              {isGiftCard && <span style={{ fontSize: 10, background: "#F59E0B", color: "#fff", padding: "1px 5px", borderRadius: 6, marginRight: 4 }}>🎁GC消化</span>}
              {isPromo && <span style={{ fontSize: 10, background: "#1565C0", color: "#fff", padding: "1px 5px", borderRadius: 6, marginRight: 4 }}>📸PR無料</span>}
              {courseName}{sessionSuffix}
              {(appt.extraServiceNames || []).map(n => ` + ${n}`).join("")}
            </div>
            {(dispSvc > 0 || dispTip > 0) && (
              <div style={{ color: amountColor, fontWeight: 700 }}>
                <div>{svcText}</div>
                {dispTip > 0 && <div>{tipText}</div>}
                <div style={{ fontWeight: 800 }}>
                  {dispTotal}
                  {extraSvc > 0 && <span style={{ color: REVENUE_COLOR }}> +Extra ${extraSvc}{appt.extraPricePaymentType==="card"?"💳":"💵"}</span>}
                  {extra > 0 && <span style={{ color: REVENUE_COLOR }}> +Extra tip{extra}💝{appt.extraTipPaymentType==="card"?"💳":"💵"}</span>}
                </div>
              </div>
            )}
            {/* Same-day ticket purchase where session 1 was also used today: show the
                per-session redemption reference (staff payroll split) alongside the
                package purchase amount above, so both totals are visible on the card. */}
            {isSameDay && appt.useToday !== false && (svc > 0 || tip > 0) && (
              <div style={{ color: NON_REVENUE_COLOR, fontWeight: 700 }}>
                <div>{svc}{paidIcon}</div>
                {tip > 0 && <div>{tip}{tipIcon}</div>}
                <div style={{ fontWeight: 800 }}>{r2(svc + tip)}</div>
              </div>
            )}
            {appt.cavTherapist && !isDualLicense(appt.therapist) && (
              <div style={{ color: "#6A1B9A", fontSize: 10 }}>with {appt.cavTherapist}</div>
            )}
            {dep > 0 && (
              <div style={{ color: "#2E7D32", fontSize: 10 }}>💰デポジット${dep} 支払済み（本日受取額には含まず）{depDate ? ` ${depDate}` : ""}</div>
            )}
            {gc > 0 && (
              <div style={{ color: "#2E7D32", fontSize: 10 }}>GC使用 ${gc}</div>
            )}
            {(appt.addons || []).length > 0 && (
              <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 3 }}>
                {(appt.addons || []).map(addon => {
                  const label = addon.serviceName || addon.name || "オプション";
                  const aSvc = Number(addon.price||0);
                  const aTip = Number(addon.tip||0);
                  const aTotal = r2(aSvc + aTip);
                  const chipColor = addon.countsAsRevenue === true ? REVENUE_COLOR : addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#00695C";
                  const svcIcon = addon.paymentType === "card" ? "💳" : addon.paymentType === "cash" ? "💵" : "";
                  const tipIcon = addon.tipPaymentType === "card" ? "💳" : addon.tipPaymentType === "cash" ? "💵" : "";
                  return (
                    <div key={addon.id} style={{ fontSize: 10, color: chipColor, fontWeight: 700 }}>
                      <div>➕ {label}{addon.ticketCurrent ? ` ${addon.ticketCurrent}/${appt.ticketTotal||3}` : ""}{addon.countsAsRevenue === null && " ⚠️"}</div>
                      {aTotal > 0 && (
                        <>
                          <div>{aSvc}{svcIcon}</div>
                          {aTip > 0 && <div>{aTip}{tipIcon}</div>}
                          <div style={{ fontWeight: 800 }}>{aTotal}</div>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })()}
      {(appt.purchaseTags || []).length > 0 && (
        <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 4 }}>
          {(appt.purchaseTags || []).map(tagId => {
            const tag = PURCHASE_TAGS.find(t => t.id === tagId);
            if (!tag) return null;
            const emptyChip = (
              <span key={tagId} style={{ fontSize: 9, background: tag.bg, color: tag.color, padding: "2px 6px", borderRadius: 8, fontWeight: 700, alignSelf: "flex-start" }}>{tag.label}</span>
            );

            if (tagId === "newTicket") {
              const ntSvc = Number(appt.newTicketAmount||0);
              const ntTip = Number(appt.newTicketTip||0);
              if (ntSvc + ntTip === 0) return emptyChip;
              const ntSvcIcon = appt.newTicketPaymentType==="card"?"💳":appt.newTicketPaymentType==="cash"?"💵":"";
              const ntTipIcon = appt.newTicketTipPaymentType==="card"?"💳":appt.newTicketTipPaymentType==="cash"?"💵":"";
              const ntSvcText = appt.newTicketSplitPayment ? `💵$${appt.newTicketCashPortion||0}＋💳$${appt.newTicketCardPortion||0}` : `${ntSvc}${ntSvcIcon}`;
              return (
                <div key={tagId} style={{ fontSize: 10, color: REVENUE_COLOR, fontWeight: 700 }}>
                  <div>{tag.label}{appt.newTicketPackageName && ` (${appt.newTicketPackageName})`}</div>
                  <div>{ntSvcText}</div>
                  {ntTip > 0 && <div>{ntTip}{ntTipIcon}</div>}
                  <div style={{ fontWeight: 800 }}>{r2(ntSvc + ntTip)}</div>
                </div>
              );
            }
            if (tagId === "retail") {
              const items = getRetailItems(appt);
              const total = items.reduce((s,it)=>s+Number(it.amount||0),0);
              if (total === 0) return emptyChip;
              const uniformType = items.every(it => it.paymentType === items[0].paymentType) ? items[0].paymentType : null;
              const icon = uniformType === "card" ? "💳" : uniformType === "cash" ? "💵" : "";
              return (
                <div key={tagId} style={{ fontSize: 10, color: REVENUE_COLOR, fontWeight: 700 }}>物販 ${total}{icon}</div>
              );
            }
            if (tagId === "giftCard") {
              const total = Number(appt.giftCardPurchaseAmount||0);
              if (total === 0) return emptyChip;
              const icon = appt.giftCardPurchasePaymentType==="card"?"💳":appt.giftCardPurchasePaymentType==="cash"?"💵":"";
              return (
                <div key={tagId} style={{ fontSize: 10, color: REVENUE_COLOR, fontWeight: 700 }}>{tag.label} ${total}{icon}</div>
              );
            }
            return emptyChip;
          })}
        </div>
      )}
    </div>
  );
}

const APPT_ERROR_LABELS = {
  clientName: "お客様名",
  serviceName: "コース名／施術メニュー",
  therapist: "セラピスト（ボディ）",
  startTime: "開始時間",
  customerType: "顧客タイプ（RL/RT/NL/NT）",
  referralSource: "どこで知りましたか（新規のきっかけ）",
  price: "施術料金",
  packagePrice: "パッケージ代金",
  ticketMenu: "施術メニュー・時間",
  ticketTotal: "コース回数（3回／5回）",
  priceVersionChosen: "料金バージョン（新料金／旧料金）",
  ticketCurrent: "今日は何回目か",
  paymentType: "施術 支払方法",
  tipPaymentType: "チップ 支払方法",
  packagePaymentType: "パッケージ代金 支払方法",
  packageTipPaymentType: "当日チップ 支払方法",
  extraPricePaymentType: "追加施術料 支払方法",
  extraTipPaymentType: "エクストラチップ 支払方法",
  newTicketPaymentType: "チケット新規購入 支払方法",
  newTicketTipPaymentType: "チケット新規購入チップ 支払方法",
  retailPurchasePaymentType: "物販購入 支払方法",
  giftCardPurchasePaymentType: "ギフトカード購入 支払方法",
  depositPaidDate: "デポジット 支払われた日",
  packageDepositDate: "デポジット 支払われた日",
};

function ApptModal({ appt, onSave, onDelete, onClose, clientDeposits = [] }) {
  const [form, setForm] = useState({
    tipPaymentType: "cash",
    // Legacy appointments saved before these validation flags existed already have a
    // real ticketMenu, so treat them as "chosen" rather than forcing a redundant re-click.
    ticketTotalChosen: !!appt.ticketMenu,
    priceVersionChosen: !!appt.ticketMenu,
    ...appt,
  });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [errors, setErrors] = useState([]);
  const isRegularWeightLoss = isWeightLossService(form.serviceName);

  const validate = (f) => {
    const errs = [];
    if (!f.clientName?.trim()) errs.push("clientName");
    if (!f.serviceName?.trim() && !f.ticketMenu) errs.push("serviceName");
    if (!f.therapist) errs.push("therapist");
    if (!f.startTime) errs.push("startTime");
    if (!f.customerType) errs.push("customerType");
    if ((f.customerType === "NL" || f.customerType === "NT") && !f.referralSource) errs.push("referralSource");
    if (f.isTicket) {
      if (!f.ticketMenu) errs.push("ticketMenu");
      if (!f.ticketTotalChosen) errs.push("ticketTotal");
      if (!f.priceVersionChosen) errs.push("priceVersionChosen");
      if (!f.isSameDayTicket && !f.ticketCurrent) errs.push("ticketCurrent");
    }
    if (f.isSameDayTicket) {
      if (f.useToday !== false && Number(f.packagePrice || 0) <= 0) errs.push("packagePrice");
      if (f.useToday !== false && !f.packageSplitPayment && Number(f.packagePrice || 0) > 0 && !f.paymentType) errs.push("packagePaymentType");
      if (f.useToday !== false && Number(f.packageTip || 0) > 0 && !f.tipPaymentType) errs.push("packageTipPaymentType");
      if (Number(f.packageDepositAmount || 0) > 0 && !f.packageDepositDate) errs.push("packageDepositDate");
    } else {
      const addonTotal = (f.addons || []).reduce((s, a) => s + Number(a.price || 0) + Number(a.tip || 0), 0);
      if (Number(f.price || 0) + Number(f.cavPrice || 0) + Number(f.tip || 0) + Number(f.cavTip || 0) + addonTotal <= 0) {
        errs.push("price");
      }
    }
    // チケット消化 (isTicket && !isSameDayTicket): price/tip here are payroll-reference only —
    // no real money changes hands (already paid when the ticket bundle was purchased), and there's
    // no payment-method control shown for it, so it must never require one. Same-day ticket
    // purchases ARE real money but are validated separately above via packagePaymentType.
    if (!f.isGiftCard && !f.isPromo && !f.isTicket) {
      const gc = Number(f.giftCardUsed || 0);
      const svc = Number(f.price || 0);
      const tip = Number(f.tip || 0);
      const gcSvc = Math.min(gc, svc);
      const gcTip = Math.min(Math.max(0, gc - gcSvc), tip);
      if (svc > 0 && gcSvc < svc && !f.svcSplitPayment && !f.paymentType) errs.push("paymentType");
      if (tip > 0 && gcTip < tip && !f.tipSplitPayment && !f.tipPaymentType) errs.push("tipPaymentType");
    }
    if (Number(f.extraPrice || 0) > 0 && !f.extraPricePaymentType) errs.push("extraPricePaymentType");
    if (Number(f.extraTip || 0) > 0 && !f.extraTipPaymentType) errs.push("extraTipPaymentType");
    if ((f.purchaseTags || []).includes("newTicket")) {
      if (!f.newTicketSplitPayment && Number(f.newTicketAmount || 0) > 0 && !f.newTicketPaymentType) errs.push("newTicketPaymentType");
      if (Number(f.newTicketTip || 0) > 0 && !f.newTicketTipPaymentType) errs.push("newTicketTipPaymentType");
    }
    if ((f.purchaseTags || []).includes("retail")) {
      if (Number(f.retailPurchaseAmount || 0) > 0 && !f.retailPurchasePaymentType) errs.push("retailPurchasePaymentType");
      if ((f.extraRetailItems || []).some(it => Number(it.amount || 0) > 0 && !it.paymentType)) errs.push("retailPurchasePaymentType");
    }
    if ((f.purchaseTags || []).includes("giftCard")) {
      if (Number(f.giftCardPurchaseAmount || 0) > 0 && !f.giftCardPurchasePaymentType) errs.push("giftCardPurchasePaymentType");
    }
    if (Number(f.depositApplied || 0) > 0 && !f.depositPaidDate) errs.push("depositPaidDate");
    return errs;
  };

  const handleSave = () => {
    const errs = validate(form);
    if (errs.length > 0) { setErrors(errs); return; }
    setErrors([]);
    onSave(form);
  };

  const ErrorBanner = () => errors.length === 0 ? null : (
    <div style={{ background: "#FFEBEE", border: "2px solid #C62828", borderRadius: 10, padding: "10px 14px", marginBottom: 12 }}>
      <div style={{ fontWeight: 800, fontSize: 13, color: "#C62828", marginBottom: 4 }}>⚠️ 保存できません — 未入力・未選択の項目があります</div>
      <div style={{ fontSize: 12, color: "#B71C1C" }}>
        {errors.map(e => APPT_ERROR_LABELS[e]).join("　／　")}
      </div>
    </div>
  );

  const therapistIsDual = isDualLicense(form.therapist);
  const menuHasCav = form.isTicket && form.ticketMenu ? !!PRICE_TABLE[form.priceVersion]?.[form.ticketMenu]?.cav : false;
  // Split only when body-only therapist + cav therapist selected
  const showSplitPrices = menuHasCav && isBodyOnly(form.therapist) && !!form.cavTherapist;

  // Auto-fill prices from ticket selection
  const applyTicketPrices = (menu, total, version, cavTherapist, therapist) => {
    const prices = PRICE_TABLE[version]?.[menu];
    if (!prices) return;
    const dual = isDualLicense(therapist ?? form.therapist);
    const hasCavTherapist = cavTherapist ?? form.cavTherapist;
    // Split prices when: body-only therapist has picked a cav therapist
    const willSplit = prices.cav && !dual && hasCavTherapist;
    setForm(f => ({
      ...f,
      ticketMenu: menu,
      ticketTotal: total,
      priceVersion: version,
      // If dual-license or cav therapist selected: split body+cav
      // Otherwise: total = body.service + cav.service combined into price
      price: willSplit ? prices.body.service : (prices.combined?.service ?? (prices.body.service + (prices.cav?.service || 0))),
      tip: willSplit ? prices.body.tip : (prices.combined?.tip ?? (prices.body.tip + (prices.cav?.tip || 0))),
      cavPrice: willSplit ? (prices.cav?.service || 0) : 0,
      cavTip: willSplit ? (prices.cav?.tip || 0) : 0,
    }));
  };

  // For the inline "📌 当日の追加購入 → 🎟️ チケット新規購入" block — a brand-new package sold
  // alongside today's visit. Auto-fills the package service/tip from TICKET_PACKAGE_PRICES
  // (still editable afterward, e.g. for occasional discounts).
  const applyNewTicketMenuPrice = (menu, total) => {
    const group = MENU_OPTIONS.find(m => menu.startsWith(m.prefix));
    const durLabel = menu.split("-")[1];
    const packageName = group ? `${group.group} ${durLabel}分 ×${total}回` : "";
    const pkg = TICKET_PACKAGE_PRICES[menu];
    setForm(f => ({
      ...f,
      newTicketMenu: menu,
      newTicketTotal: total,
      newTicketPackageName: packageName,
      ...(pkg ? { newTicketAmount: pkg.service, newTicketTip: pkg.tip } : {}),
    }));
  };

  const autoDetectAndApplyTicket = () => {
    if (!form.ticketMenu && form.serviceName) {
      const name = form.serviceName.toLowerCase();
      let prefix = null, dur = null;
      if (name.includes("improving posture")) prefix = "IP";
      else if (name.includes("weight loss")) prefix = "WL";
      else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "SF";
      else if (name.includes("channeling") || name.includes("procell")) prefix = "MC";
      if (prefix) {
        const m = form.serviceName.match(/(\d+)\s*min/i);
        const detectedDur = m ? Number(m[1]) : null;
        const validDurs = MENU_OPTIONS.find(o => o.prefix === prefix)?.durations || [];
        dur = validDurs.includes(detectedDur) ? detectedDur : validDurs[0];
        if (dur) {
          const menu = `${prefix}-${dur}-${form.ticketTotal || 3}`;
          applyTicketPrices(menu, form.ticketTotal || 3, form.priceVersion || "new", form.cavTherapist, form.therapist);
        }
      }
    }
  };

  // On open: if existing appointment has cavTherapist but cavPrice=0, auto-apply split prices
  useEffect(() => {
    if (appt.cavTherapist && Number(appt.cavPrice || 0) === 0 && isBodyOnly(appt.therapist)) {
      if (appt.ticketMenu) {
        applyTicketPrices(appt.ticketMenu, appt.ticketTotal, appt.priceVersion || "new", appt.cavTherapist, appt.therapist);
      } else if (appt.serviceName) {
        const name = appt.serviceName.toLowerCase();
        let prefix = null;
        if (name.includes("improving posture")) prefix = "IP";
        else if (name.includes("weight loss")) prefix = "WL";
        else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "SF";
        else if (name.includes("channeling") || name.includes("procell")) prefix = "MC";
        if (prefix) {
          const m = appt.serviceName.match(/(\d+)\s*min/i);
          const detectedDur = m ? Number(m[1]) : null;
          const validDurs = MENU_OPTIONS.find(o => o.prefix === prefix)?.durations || [];
          const dur = validDurs.includes(detectedDur) ? detectedDur : validDurs[0];
          if (dur) {
            const menu = `${prefix}-${dur}-${appt.ticketTotal || 3}`;
            applyTicketPrices(menu, appt.ticketTotal || 3, appt.priceVersion || "new", appt.cavTherapist, appt.therapist);
          }
        }
      }
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Machine (cav) minutes default to a fixed number regardless of the total course length —
  // 40min for Weight Loss, 15min for everything else (e.g. Improving Posture) — with the rest
  // going to the body therapist. Defaults in once a cav therapist is picked, but only if the
  // staff hasn't already typed their own minutes.
  useEffect(() => {
    if (!form.isTicket && form.cavTherapist && !form.bodyMins && !form.cavMins) {
      const cavMins = isRegularWeightLoss ? 40 : 15;
      const bodyMins = Math.max(0, Number(form.duration) - cavMins);
      setForm(f => ({ ...f, cavMins, bodyMins }));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.cavTherapist, form.duration, form.isTicket]);

  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#0D4F4F" }}>✏️ 予約編集</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer", color: "#888" }}>✕</button>
      </div>

      {clientDeposits.length > 0 && (
        <div style={{ background: "#FFF3E0", borderRadius: 10, padding: "10px 14px", marginBottom: 12, border: "2px solid #FF9800" }}>
          <div style={{ fontWeight: 800, fontSize: 13, color: "#E65100" }}>💰 デポジット／ギフト先払いあり！</div>
          {clientDeposits.map((d, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, fontSize: 12, color: "#BF360C", marginTop: 4 }}>
              <div>
                {d.type === "giftcard" ? "🎁 " : "💰 "}
                <strong>{d.sheetDate}</strong> — ${d.amount}{Number(d.tip) > 0 ? `（＋チップ$${d.tip}）` : ""} ({d.paymentType === "cash" ? "現金" : "カード"})
                {d.appointmentDate ? ` → 来店予定：${d.appointmentDate}${d.appointmentTime ? ` ${d.appointmentTime}` : ""}` : ""}
                {d.notes ? ` — ${d.notes}` : ""}
              </div>
              {d.type === "deposit" && (
                <button onClick={() => setForm(f => ({ ...f, depositApplied: Number(d.amount) || 0, depositPaidDate: d.sheetDate }))}
                  style={{ flexShrink: 0, padding: "4px 10px", borderRadius: 8, border: "none", background: "#E65100", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 11 }}>
                  この内容を使う
                </button>
              )}
            </div>
          ))}
        </div>
      )}

      <ErrorBanner />

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="顧客タイプ" error={errors.includes("customerType")}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {CUSTOMER_TYPES.map(ct => {
              const colors = { RL:"#4CAF50", RT:"#2196F3", NL:"#FF9800", NT:"#9C27B0" };
              const labels = { RL:"リピ・ローカル", RT:"リピ・トラベラー", NL:"ニュー・ローカル", NT:"ニュー・トラベラー" };
              return (
                <button key={ct} onClick={() => set("customerType", ct)}
                  style={{ flex: 1, minWidth: 90, padding: "7px 4px", borderRadius: 8, border: `2px solid ${form.customerType===ct?colors[ct]:(errors.includes("customerType")?"#C62828":"#DDD")}`, background: form.customerType===ct?colors[ct]:"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: form.customerType===ct?"#fff":"#888" }}>
                  {ct}<br /><span style={{ fontWeight: 400, fontSize: 9 }}>{labels[ct]}</span>
                </button>
              );
            })}
          </div>
        </Field>

        {/* How they found us — shown for NT and NL */}
        {(form.customerType === "NT" || form.customerType === "NL") && (
          <Field label={form.customerType === "NT" ? "📍 どこで知りましたか？（NT）" : "📍 どこで知りましたか？（NL）"} error={errors.includes("referralSource")}>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {REFERRAL_SOURCES.map(src => (
                <button key={src} onClick={() => set("referralSource", form.referralSource === src ? "" : src)}
                  style={{
                    padding: "6px 12px", borderRadius: 20, border: `2px solid ${form.referralSource === src ? "#9C27B0" : "#DDD"}`,
                    background: form.referralSource === src ? "#9C27B0" : "#fff",
                    cursor: "pointer", fontWeight: 600, fontSize: 12,
                    color: form.referralSource === src ? "#fff" : "#666"
                  }}>
                  {src}
                </button>
              ))}
            </div>
          </Field>
        )}

        <Field label="お客様名" error={errors.includes("clientName")}>
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={{ ...inputStyle, ...(errors.includes("clientName") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="お客様名" />
        </Field>
        <Field label="コース名" error={errors.includes("serviceName")}>
          <select value={SQUARE_SERVICES.find(s => s.name === form.serviceName) ? form.serviceName : ""} onChange={e => {
            const name = e.target.value;
            // Clear any body/cav minute split left over from a previously-selected course —
            // it belongs to that course's duration, not this one, and would otherwise block
            // the Weight Loss auto-default (fixed 40min machine) from kicking in.
            const svc = SQUARE_SERVICES.find(s => s.name === name);
            setForm(f => ({ ...f, serviceName: name, bodyMins: undefined, cavMins: undefined, ...(svc ? { duration: svc.duration } : {}) }));
          }} style={{ ...inputStyle, color: "#1a1a1a", fontWeight: 600, ...(errors.includes("serviceName") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }}>
            <option value="">コースを選択</option>
            {SQUARE_SERVICES.map(s => (
              <option key={s.name} value={s.name}>{s.name}</option>
            ))}
          </select>
          {/* Square同期の名前がリストにない場合は自由入力できる */}
          {!SQUARE_SERVICES.find(s => s.name === form.serviceName) && (
            <input type="text" value={form.serviceName || ""} placeholder="コース名を入力"
              onChange={e => set("serviceName", e.target.value)}
              style={{ ...inputStyle, marginTop: 4 }} />
          )}
          {form.serviceName && (
            <div style={{ fontSize: 11, color: "#0D4F4F", marginTop: 4, fontWeight: 600 }}>
              ✅ {form.serviceName} ({form.duration}分)
            </div>
          )}

          {/* Extra menu(s) done by the SAME therapist in the same visit — name only, no separate
              price/payment tracking (that combined amount goes straight into 施術合計・チップ合計
              below). Use オプション追加 instead when a DIFFERENT therapist is involved. */}
          {(form.extraServiceNames || []).map((name, idx) => (
            <div key={idx} style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 6 }}>
              <div style={{ flex: 1, fontSize: 12, fontWeight: 600, color: "#0D4F4F", background: "#EAF6F4", borderRadius: 8, padding: "6px 10px" }}>
                ✅ {name}
              </div>
              <button onClick={() => set("extraServiceNames", form.extraServiceNames.filter((_, i) => i !== idx))}
                style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA" }}>✕</button>
            </div>
          ))}
          <select value="" onChange={e => {
            const name = e.target.value;
            if (!name) return;
            set("extraServiceNames", [...(form.extraServiceNames || []), name]);
          }} style={{ ...inputStyle, marginTop: 6, fontSize: 12, color: "#888" }}>
            <option value="">＋ 同じセラピストのメニューを追加</option>
            {ADDON_PRESETS.map(p => <option key={p} value={p}>{p}</option>)}
            {SQUARE_SERVICES.map(s => (
              <option key={s.name} value={s.name}>{s.name}</option>
            ))}
          </select>
        </Field>

        {/* ── オプション追加（別担当者の追加施術・前回の未消化分キャビなど） ── */}
        <div>
          {/* Service picker — selecting immediately adds the row (same one-step pattern as
              the same-therapist メニュー追加 dropdown above) */}
          <div style={{ marginBottom: 8 }}>
            <select value="" onChange={e => {
              const name = e.target.value;
              if (!name) return;
              set("addons", [...(form.addons||[]), {
                id: `${Date.now()}`,
                serviceName: name,
                therapist: "",
                price: 0,
                tip: 0,
                paymentType: "",
                tipPaymentType: "",
                countsAsRevenue: null,
                ticketCurrent: null,
              }]);
            }} style={{ ...inputStyle, fontSize: 12, color: "#888" }}>
              <option value="">＋ 別のセラピストのメニューを追加</option>
              {ADDON_PRESETS.map(p => <option key={p} value={p}>{p}</option>)}
              {SQUARE_SERVICES.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
            </select>
          </div>

          {/* Added addons */}
          {(form.addons||[]).length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {(form.addons||[]).map((addon, idx) => {
                const upd = (patch) => set("addons", form.addons.map((a,i) => i===idx ? {...a, ...patch} : a));
                const svcAmt = Number(addon.price||0);
                const tipAmt = Number(addon.tip||0);
                return (
                  <div key={addon.id} style={{ background: "#fff", borderRadius: 8, padding: "10px 10px", border: "1.5px solid #D0E8E4" }}>
                    {/* Name row */}
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                      <input value={addon.serviceName||""} onChange={e => upd({ serviceName: e.target.value })}
                        style={{ ...inputStyle, flex: 1, fontSize: 12, fontWeight: 700, color: "#0D4F4F" }}
                        placeholder="サービス名を入力（例：前回キャビ未消化分 2/3回目）" />
                      <button onClick={() => set("addons", form.addons.filter((_,i) => i !== idx))}
                        style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA", marginLeft: 6, flexShrink: 0 }}>✕</button>
                    </div>

                    {/* Revenue vs. ticket-consumption toggle */}
                    <div style={{ marginBottom: 6 }}>
                      <div style={{ fontSize: 10, color: addon.countsAsRevenue === null ? "#C62828" : "#888", marginBottom: 3, fontWeight: addon.countsAsRevenue === null ? 700 : 400 }}>
                        この分の扱い{addon.countsAsRevenue === null && " ⚠️ 未選択"}
                      </div>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => upd({ countsAsRevenue: true, ticketCurrent: null })}
                          style={{ flex: 1, padding: "6px 4px", borderRadius: 8, border: `2px solid ${addon.countsAsRevenue === true ? REVENUE_COLOR : "#DDD"}`, background: addon.countsAsRevenue === true ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.countsAsRevenue === true ? REVENUE_COLOR : "#888" }}>
                          💰 都度払い（新規支払い）
                        </button>
                        <button onClick={() => upd({ countsAsRevenue: false })}
                          style={{ flex: 1, padding: "6px 4px", borderRadius: 8, border: `2px solid ${addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#DDD"}`, background: addon.countsAsRevenue === false ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#888" }}>
                          🎫 チケット消化（前回未消化分など）
                        </button>
                      </div>
                    </div>

                    {/* Session number — only when this addon is itself a ticket redemption */}
                    {addon.countsAsRevenue === false && form.isTicket && (
                      <div style={{ marginBottom: 6 }}>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>何回目の消化？</div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {Array.from({ length: form.ticketTotal||3 }, (_,i) => i+1).map(n => (
                            <button key={n} onClick={() => upd({ ticketCurrent: n })}
                              style={{ padding: "5px 10px", borderRadius: 8, border: `2px solid ${addon.ticketCurrent===n?"#0D4F4F":"#DDD"}`, background: addon.ticketCurrent===n?"#0D4F4F":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.ticketCurrent===n?"#fff":"#888" }}>
                              {n}/{form.ticketTotal||3}回目
                            </button>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Therapist */}
                    <div style={{ marginBottom: 6 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>担当セラピスト</div>
                      <select value={addon.therapist||""} onChange={e => upd({ therapist: e.target.value })}
                        style={{ ...inputStyle, fontSize: 12 }}>
                        <option value="">— 選択 —</option>
                        {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>

                    {/* Price + Tip inputs */}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 6 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術料 ($)</div>
                        <input type="number" value={addon.price||""} onChange={e => upd({ price: e.target.value })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>チップ ($)</div>
                        <input type="number" value={addon.tip||""} onChange={e => upd({ tip: e.target.value })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                    </div>

                    {/* Payment type — real money only when this is a new (revenue) payment, not a ticket redemption */}
                    {addon.countsAsRevenue !== false && (
                      <div style={{ marginBottom: tipAmt > 0 ? 6 : 0 }}>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法（施術）</div>
                        <PaymentToggle value={addon.paymentType} onChange={v => upd({ paymentType: v })} small />
                      </div>
                    )}

                    {/* Tip payment type */}
                    {tipAmt > 0 && addon.countsAsRevenue !== false && (
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法（チップ）</div>
                        <PaymentToggle value={addon.tipPaymentType} onChange={v => upd({ tipPaymentType: v })} small />
                      </div>
                    )}

                    {/* Sub-total */}
                    {(svcAmt + tipAmt) > 0 && (
                      <div style={{ marginTop: 6, textAlign: "right", fontSize: 11, color: "#0D4F4F", fontWeight: 700 }}>
                        施術 ${svcAmt}{tipAmt > 0 ? ` ＋ チップ $${tipAmt}` : ""} ＝ <strong>${svcAmt + tipAmt}</strong>
                      </div>
                    )}
                  </div>
                );
              })}

              {/* Grand addon total */}
              {(form.addons||[]).some(a => Number(a.price||0) + Number(a.tip||0) > 0) && (
                <div style={{ background: "#E0F2F1", borderRadius: 8, padding: "8px 12px", textAlign: "right" }}>
                  <span style={{ fontSize: 12, color: "#00695C" }}>オプション合計　</span>
                  <strong style={{ fontSize: 14, color: "#004D40" }}>
                    ${(form.addons||[]).reduce((s,a) => s + Number(a.price||0) + Number(a.tip||0), 0)}
                  </strong>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Ticket toggle */}
        <div style={{ display: "flex", gap: 4 }}>
          <button onClick={() => setForm(f => ({...f, isTicket: false, isSameDayTicket: false, isGiftCard: false, isPromo: false}))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${!form.isTicket && !form.isGiftCard && !form.isPromo ? "#C62828" : "#DDD"}`, background: !form.isTicket && !form.isGiftCard && !form.isPromo ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, color: !form.isTicket && !form.isGiftCard && !form.isPromo ? "#C62828" : "#888", fontSize: 11 }}>
            🔴 通常施術
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: false, isGiftCard: false, isPromo: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#DDD"}`, background: form.isTicket && !form.isSameDayTicket ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#888", fontSize: 11 }}>
            🔵 チケット消化
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: true, useToday: true, ticketCurrent: 1, isGiftCard: false, isPromo: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isSameDayTicket ? "#2E7D32" : "#DDD"}`, background: form.isSameDayTicket ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isSameDayTicket ? "#2E7D32" : "#888", fontSize: 11 }}>
            🟢 当日購入
          </button>
          <button onClick={() => setForm(f => ({...f, isPromo: true, isGiftCard: false, isTicket: false, isSameDayTicket: false}))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isPromo ? "#1565C0" : "#DDD"}`, background: form.isPromo ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isPromo ? "#1565C0" : "#888", fontSize: 11 }}>
            📸 PR無料
          </button>
        </div>
        {form.isGiftCard && (
          <div style={{ background: "#FFFDE7", border: "1.5px solid #F59E0B", borderRadius: 8, padding: "8px 10px", fontSize: 11, color: "#92400E" }}>
            ⚠️ この予約は過去に「GC消化」として登録されています。ギフトカードの一部だけ使った場合や、チップを現金・カードで受け取った場合は、下の「🔴 通常施術」を押してから「🎁 ギフトカード使用」で金額を入力し直してください。
          </div>
        )}

        {/* Ticket options */}
        {form.isTicket && (
          <div style={{ background: "#EEF4FF", borderRadius: 10, padding: 12 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#1565C0", marginBottom: 10 }}>🎟️ チケット設定</div>

            {/* Cav therapist — shown first; selecting triggers full auto-fill */}
            {isBodyOnly(form.therapist) && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ background: "#F3E5F5", borderRadius: 8, padding: 10, marginBottom: 10, border: `2px solid ${form.cavTherapist ? "#6A1B9A" : "#CE93D8"}` }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A", marginBottom: 6 }}>
                  ⚡ 機械担当セラピスト
                </div>
                <select value={form.cavTherapist || ""} onChange={e => {
                  const cav = e.target.value;
                  if (form.ticketMenu) {
                    // Menu already chosen — just re-apply with new cav
                    applyTicketPrices(form.ticketMenu, form.ticketTotal, form.priceVersion, cav, form.therapist);
                    setForm(f => ({ ...f, cavTherapist: cav }));
                  } else if (cav && form.serviceName) {
                    // Auto-detect menu from serviceName and apply split immediately
                    const name = form.serviceName.toLowerCase();
                    let prefix = null;
                    if (name.includes("improving posture")) prefix = "IP";
                    else if (name.includes("weight loss")) prefix = "WL";
                    else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "SF";
                    else if (name.includes("channeling") || name.includes("procell")) prefix = "MC";
                    if (prefix) {
                      const m = form.serviceName.match(/(\d+)\s*min/i);
                      const detectedDur = m ? Number(m[1]) : null;
                      const validDurs = MENU_OPTIONS.find(o => o.prefix === prefix)?.durations || [];
                      const dur = validDurs.includes(detectedDur) ? detectedDur : validDurs[0];
                      if (dur) {
                        const menu = `${prefix}-${dur}-${form.ticketTotal || 3}`;
                        applyTicketPrices(menu, form.ticketTotal || 3, form.priceVersion || "new", cav, form.therapist);
                        setForm(f => ({ ...f, cavTherapist: cav }));
                      } else {
                        set("cavTherapist", cav);
                      }
                    } else {
                      set("cavTherapist", cav);
                    }
                  } else {
                    set("cavTherapist", cav);
                  }
                }} style={{ ...inputStyle, borderColor: form.cavTherapist ? "#6A1B9A" : "#DDD" }}>
                  <option value="">なし（機械なし）</option>
                  {CAV_CAPABLE.filter(t => t !== form.therapist).map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                {form.cavTherapist ? (
                  <div style={{ fontSize: 11, color: "#6A1B9A", marginTop: 4, fontWeight: 600 }}>
                    ✅ {form.therapist}（ボディ）＋ {form.cavTherapist}（機械）<br />
                    → 保存すると <strong>{form.cavTherapist}</strong> のラインに施術料が自動で追加されます
                  </div>
                ) : (
                  <div style={{ fontSize: 11, color: "#999", marginTop: 4 }}>
                    機械担当がいる場合は選んでください → 自動で振り分けが入ります
                  </div>
                )}
              </div>
            )}
            {isDualLicense(form.therapist) && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ fontSize: 11, background: "#E8F5E9", color: "#2E7D32", padding: "6px 10px", borderRadius: 8, marginBottom: 10 }}>
                ✅ {form.therapist}はボディ＋機械両対応 → 合計金額で1本入力
              </div>
            )}

            {/* IP同額ショートカット — Facial/Glow系サービスでIP 90minチケット使用時 */}
            {form.serviceName && ["facial","glow","procell","channeling","hifu","sculpt"].some(k => form.serviceName.toLowerCase().includes(k)) && (
              <div style={{ background: "#FFF3E0", borderRadius: 8, padding: 8, marginBottom: 8, border: "1px solid #FFCC80" }}>
                <div style={{ fontSize: 11, color: "#E65100", fontWeight: 700, marginBottom: 6 }}>💡 Improving Posture 90分チケット同額オプション</div>
                <button onClick={() => {
                  const menu = `IP-90-${form.ticketTotal || 3}`;
                  applyTicketPrices(menu, form.ticketTotal || 3, form.priceVersion || "new", form.cavTherapist, form.therapist);
                }} style={{
                  width: "100%", padding: "8px", borderRadius: 8,
                  border: `2px solid ${form.ticketMenu?.startsWith("IP-90") ? "#E65100" : "#FFCC80"}`,
                  background: form.ticketMenu?.startsWith("IP-90") ? "#FFE0B2" : "#FFFDE7",
                  cursor: "pointer", fontWeight: 700, fontSize: 12,
                  color: form.ticketMenu?.startsWith("IP-90") ? "#E65100" : "#888"
                }}>
                  {form.ticketMenu?.startsWith("IP-90") ? "✅ IP 90分チケット使用中" : "🔄 Improving Posture 90分チケット使用（同額）"}
                </button>
              </div>
            )}

            {/* Treatment menu selector */}
            <Field label="施術メニュー" error={errors.includes("ticketMenu")}>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                {MENU_OPTIONS.map(({ group, prefix }) => {
                  const selected = form.ticketMenu?.startsWith(prefix);
                  return (
                    <button key={prefix} onClick={() => {
                      const dur = MENU_OPTIONS.find(m => m.prefix === prefix).durations[0];
                      const menu = `${prefix}-${dur}-${form.ticketTotal || 3}`;
                      applyTicketPrices(menu, form.ticketTotal || 3, form.priceVersion || "new", form.cavTherapist, form.therapist);
                    }} style={{
                      padding: "7px 10px", borderRadius: 8,
                      border: `2px solid ${selected ? "#1565C0" : (errors.includes("ticketMenu") ? "#C62828" : "#DDD")}`,
                      background: selected ? "#1565C0" : "#fff",
                      cursor: "pointer", fontWeight: 700, fontSize: 11,
                      color: selected ? "#fff" : "#888"
                    }}>{group}</button>
                  );
                })}
              </div>
            </Field>

            {/* Duration selector */}
            {form.ticketMenu && (() => {
              const prefix = form.ticketMenu.split("-")[0];
              const group = MENU_OPTIONS.find(m => m.prefix === prefix);
              const currentDur = Number(form.ticketMenu.split("-")[1]);
              return group ? (
                <Field label="時間">
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                    {group.durations.map(dur => (
                      <button key={dur} onClick={() => {
                        const menu = `${prefix}-${dur}-${form.ticketTotal || 3}`;
                        applyTicketPrices(menu, form.ticketTotal || 3, form.priceVersion || "new", form.cavTherapist, form.therapist);
                      }} style={{
                        padding: "7px 12px", borderRadius: 8,
                        border: `2px solid ${currentDur === dur ? "#0D4F4F" : "#DDD"}`,
                        background: currentDur === dur ? "#0D4F4F" : "#fff",
                        cursor: "pointer", fontWeight: 700, fontSize: 12,
                        color: currentDur === dur ? "#fff" : "#888"
                      }}>{dur}分</button>
                    ))}
                  </div>
                </Field>
              ) : null;
            })()}

            {/* Price version */}
            <Field label="料金バージョン" error={errors.includes("priceVersionChosen")}>
              <div style={{ display: "flex", gap: 8 }}>
                {[["new","🆕 新料金（2月以降）"],["old","📋 旧料金（2月以前）"]].map(([v,label]) => (
                  <button key={v} onClick={() => { setForm(f => ({ ...f, priceVersion: v, priceVersionChosen: true })); if(form.ticketMenu) applyTicketPrices(form.ticketMenu, form.ticketTotal, v, form.cavTherapist, form.therapist); }}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.priceVersion===v && form.priceVersionChosen?"#1565C0":(errors.includes("priceVersionChosen")?"#C62828":"#DDD")}`, background: form.priceVersion===v && form.priceVersionChosen?"#BBDEFB":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: form.priceVersion===v && form.priceVersionChosen?"#1565C0":"#888" }}>
                    {label}
                  </button>
                ))}
              </div>
            </Field>

            {/* Total sessions */}
            <Field label="コース回数" error={errors.includes("ticketTotal")}>
              <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                {[3,5].map(n => (
                  <button key={n} onClick={() => {
                    setForm(f => ({ ...f, ticketTotal: n, ticketTotalChosen: true }));
                    if(form.ticketMenu) {
                      const parts = form.ticketMenu.split("-");
                      applyTicketPrices(`${parts[0]}-${parts[1]}-${n}`, n, form.priceVersion||"new", form.cavTherapist, form.therapist);
                    }
                  }} style={{ flex: 1, padding: "8px", borderRadius: 8, border: `2px solid ${form.ticketTotal===n && form.ticketTotalChosen?"#1565C0":(errors.includes("ticketTotal")?"#C62828":"#DDD")}`, background: form.ticketTotal===n && form.ticketTotalChosen?"#1565C0":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketTotal===n && form.ticketTotalChosen?"#fff":"#888" }}>
                    {n}回コース
                  </button>
                ))}
              </div>
            </Field>

            {/* Current session — different UI for same-day purchase */}
            {form.isSameDayTicket ? (
              <Field label="今日使いますか？">
                <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                  <button onClick={() => setForm(f => ({...f, useToday: true, ticketCurrent: 1}))}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.useToday !== false ? "#2E7D32" : "#DDD"}`, background: form.useToday !== false ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, color: form.useToday !== false ? "#2E7D32" : "#888", fontSize: 12 }}>
                    ✅ 今日使う（1回目）
                  </button>
                  <button onClick={() => setForm(f => ({...f, useToday: false, ticketCurrent: 0, price: 0, tip: 0, cavPrice: 0, cavTip: 0}))}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.useToday === false ? "#888" : "#DDD"}`, background: form.useToday === false ? "#F5F5F5" : "#fff", cursor: "pointer", fontWeight: 700, color: form.useToday === false ? "#555" : "#888", fontSize: 12 }}>
                    ⭕ 次回から使う
                  </button>
                </div>
              </Field>
            ) : (
              <Field label="今日は何回目？" error={errors.includes("ticketCurrent")}>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                  {Array.from({ length: form.ticketTotal||3 }, (_,i) => i+1).map(n => (
                    <button key={n} onClick={() => set("ticketCurrent", n)}
                      style={{ padding: "6px 14px", borderRadius: 8, border: `2px solid ${form.ticketCurrent===n?"#0D4F4F":(errors.includes("ticketCurrent")?"#C62828":"#DDD")}`, background: form.ticketCurrent===n?"#0D4F4F":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketCurrent===n?"#fff":"#888" }}>
                      {n}/{form.ticketTotal||3}回目
                    </button>
                  ))}
                </div>
              </Field>
            )}


            {/* Prices — only when using today or チケット消化 */}
            {form.ticketMenu && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ marginTop: 8, background: "#fff", borderRadius: 8, padding: 10 }}>
                <div style={{ fontSize: 11, color: "#888", marginBottom: 6 }}>
                  {showSplitPrices ? `💡 ${form.therapist}（ボディ）＋ ${form.cavTherapist}（機械）振り分け` : `💡 ${form.therapist} 合計金額`}
                  {form.isSameDayTicket && <span style={{ color: "#2E7D32" }}> ※スタッフ給与参照用</span>}
                  {!form.isSameDayTicket && <span style={{ color: "#bbb" }}> ※変更可</span>}
                </div>
                {showSplitPrices ? (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label={`${form.therapist} 施術 ($)`} error={errors.includes("price")}><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label={`${form.therapist} チップ ($)`}><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                    <Field label={`${form.cavTherapist} 施術 ($)`} error={errors.includes("price")}><input type="number" value={form.cavPrice || ""} onChange={e => set("cavPrice", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label={`${form.cavTherapist} チップ ($)`}><input type="number" value={form.cavTip || ""} onChange={e => set("cavTip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label="施術 合計 ($)" error={errors.includes("price")}><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label="チップ 合計 ($)"><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                )}
              </div>
            )}

            {/* Extra service fee + tip — available for both チケット消化 and 当日購入. Needed
                because チケット消化's main 施術/チップ fields are a payroll-reference value only
                (the session was already paid for) and aren't counted as today's revenue — an
                add-on service actually paid for today (e.g. a same-therapist "追加マッサージ") has
                to go here instead, or it silently never shows up as real money received. */}
            {(!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ marginTop: 10, background: "#FFF8E1", borderRadius: 8, padding: 10, border: "1px solid #FFE082" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#F57F17", marginBottom: 8 }}>💝 エクストラ（任意）— 本日追加でお支払いがあった分</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="追加施術料 ($)">
                    <input type="number" value={form.extraPrice || ""} onChange={e => set("extraPrice", e.target.value)} style={inputStyle} placeholder="例: 27" />
                  </Field>
                  <Field label="追加チップ ($)">
                    <input type="number" value={form.extraTip || ""} onChange={e => set("extraTip", e.target.value)} style={inputStyle} placeholder="例: 20" />
                  </Field>
                </div>
                {(Number(form.extraPrice||0) > 0) && (
                  <div style={{ marginTop: 8 }}>
                    <Field label="追加施術料 支払方法" error={errors.includes("extraPricePaymentType")}><PaymentToggle value={form.extraPricePaymentType} onChange={v => set("extraPricePaymentType", v)} /></Field>
                  </div>
                )}
                {(form.extraTip > 0) && (
                  <div style={{ marginTop: 8 }}>
                    <Field label="チップ支払方法" error={errors.includes("extraTipPaymentType")}><PaymentToggle value={form.extraTipPaymentType} onChange={v => set("extraTipPaymentType", v)} /></Field>
                  </div>
                )}
                <div style={{ fontSize: 11, color: "#888", marginTop: 6 }}>物販は下の物販セクションへ</div>
              </div>
            )}

            {/* Ticket grand total display */}
            {form.ticketMenu && (() => {
              const svc = r2(Number(form.price||0) + Number(form.cavPrice||0));
              const tip = r2(Number(form.tip||0) + Number(form.cavTip||0));
              const extra = Number(form.extraTip||0);
              if (svc + tip === 0) return null;
              return (
                <div style={{ marginTop: 10, background: "#0D4F4F", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                  <div style={{ fontSize: 11, color: "#B2EBF2" }}>
                    施術 <strong style={{ color: "#fff" }}>${svc}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}</strong>
                    {extra > 0 && <span>　エクストラ <strong style={{ color: "#FFE082" }}>${extra}{form.extraTipPaymentType==="card"?"💳":"💵"}</strong></span>}
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ color: "#B2EBF2", fontSize: 11 }}>合計　</span>
                    <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip + extra}</span>
                  </div>
                </div>
              );
            })()}

            {/* Same-day purchase payment */}
            {form.isSameDayTicket && (
              <div style={{ marginTop: 10, background: "#E8F5E9", borderRadius: 8, padding: 10, border: "1px solid #A5D6A7" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#2E7D32", marginBottom: 8 }}>
                  💰 当日購入 — パッケージ代金（売上計上）
                  {form.useToday === false && <span style={{ color: "#888", fontWeight: 400 }}>（次回から使用）</span>}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="パッケージ代金 ($)" error={errors.includes("packagePrice")}>
                    <input type="number" value={form.packagePrice || ""} onChange={e => set("packagePrice", e.target.value)} style={{ ...inputStyle, ...(errors.includes("packagePrice") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="例: 432" disabled={form.packageSplitPayment} />
                  </Field>
                  <Field label={form.useToday === false ? "チップ ($)" : "当日チップ ($)"}>
                    <input type="number" value={form.packageTip || ""} onChange={e => set("packageTip", e.target.value)} style={inputStyle} placeholder="例: 91" />
                  </Field>
                </div>

                {/* Deposit note — packagePrice above is typed in by hand already net of any deposit; this only records that a deposit existed and when it was paid (allocation is already fixed by hand, no auto-calc) */}
                <div style={{ marginTop: 8, background: "#fff", borderRadius: 10, padding: 12, border: "1.5px solid #A5D6A7" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: "#2E7D32" }}>💰 デポジットあり</span>
                    <button onClick={() => set("packageDepositAmount", Number(form.packageDepositAmount) > 0 ? 0 : 20)}
                      style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.packageDepositAmount) > 0 ? "#2E7D32" : "#C8E6C9", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                      {Number(form.packageDepositAmount) > 0 ? "ON ✓" : "OFF"}
                    </button>
                  </div>
                  {Number(form.packageDepositAmount) > 0 && (
                    <div style={{ marginTop: 8, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>デポジット金額 ($)</div>
                        <input type="number" value={form.packageDepositAmount || ""}
                          onFocus={e => e.target.select()}
                          onChange={e => set("packageDepositAmount", e.target.value)}
                          style={{ ...inputStyle, borderColor: "#81C784" }} placeholder="例：20" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>支払われた日{errors.includes("packageDepositDate") && " ⚠️"}</div>
                        <input type="date" value={form.packageDepositDate || ""}
                          onChange={e => set("packageDepositDate", e.target.value)}
                          style={{ ...inputStyle, borderColor: errors.includes("packageDepositDate") ? "#C62828" : "#81C784", borderWidth: errors.includes("packageDepositDate") ? 2 : 1 }} />
                      </div>
                    </div>
                  )}
                </div>

                {/* Gift Card Usage — e.g. a previously-purchased balance covering part of this package */}
                <div style={{ background: "#E0F2F1", borderRadius: 10, padding: 12, marginTop: 8 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: "#00796B" }}>🎁 ギフトカード使用（既存残高）</span>
                    <button onClick={() => set("giftCardUsed", form.giftCardUsed === 0 ? (Number(form.packagePrice) || "") : 0)}
                      style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: form.giftCardUsed !== 0 ? "#00796B" : "#B2DFDB", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                      {form.giftCardUsed !== 0 ? "ON ✓" : "OFF"}
                    </button>
                  </div>
                  {(Number(form.giftCardUsed) > 0 || form.giftCardUsed === "") && (
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: "#00796B", marginBottom: 4 }}>使用金額 ($)</div>
                      <input type="number" value={form.giftCardUsed === "" ? "" : form.giftCardUsed}
                        onChange={e => set("giftCardUsed", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="例：89" />
                      {(() => {
                        const gc = Number(form.giftCardUsed || 0);
                        const svc = Number(form.packagePrice || 0);
                        const tip = Number(form.packageTip ?? 0);
                        const total = svc + tip;
                        if (!gc || !total) return null;
                        const remainder = Math.max(0, total - gc);
                        return (
                          <div style={{ marginTop: 6, fontSize: 12, color: "#00796B" }}>
                            <div style={{ color: "#555", marginBottom: 2 }}>
                              パッケージ${svc}{tip > 0 ? ` ＋ チップ$${tip}` : ""} ＝ <strong>${total}</strong>
                            </div>
                            🎁 GC(既存残高) <strong>${Math.min(gc, total)}</strong>(本日の売上には計上されません)
                            {remainder > 0
                              ? <> ＋ 残り <strong>${remainder}</strong>を現金・カードで受取</>
                              : <span style={{ color: "#2E7D32" }}> → 全額ギフトカード</span>}
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 10, background: "#fff", borderRadius: 8, padding: 10, border: "1.5px solid #A5D6A7" }}>
                  <div style={{ fontWeight: 700, fontSize: 12, color: "#2E7D32", marginBottom: 8 }}>💰 お会計</div>
                  <button onClick={() => set("packageSplitPayment", !form.packageSplitPayment)}
                    style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.packageSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.packageSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.packageSplitPayment ? "#2E7D32" : "#888" }}>
                    {form.packageSplitPayment ? "☑" : "☐"} 施術代金を現金・カードに分ける
                  </button>
                  {form.packageSplitPayment ? (
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術キャッシュ ($)</div>
                        <input type="number" value={form.packageCashPortion || ""} onChange={e => {
                          const raw = e.target.value;
                          setForm(f => ({ ...f, packageCashPortion: raw, packagePrice: (Number(raw) || 0) + Number(f.packageCardPortion||0) }));
                        }} style={inputStyle} placeholder="例：300" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術クレジット ($)</div>
                        <input type="number" value={form.packageCardPortion || ""} onChange={e => {
                          const raw = e.target.value;
                          setForm(f => ({ ...f, packageCardPortion: raw, packagePrice: Number(f.packageCashPortion||0) + (Number(raw) || 0) }));
                        }} style={inputStyle} placeholder="例：132" />
                      </div>
                    </div>
                  ) : (
                    <div style={{ marginBottom: 8 }}>
                      <Field label="施術 支払方法" error={errors.includes("packagePaymentType")}><PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} /></Field>
                    </div>
                  )}
                  <Field label="チップ 支払方法" error={errors.includes("packageTipPaymentType")}><PaymentToggle value={form.tipPaymentType} onChange={v => set("tipPaymentType", v)} /></Field>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Regular service */}
        {!form.isTicket && (
          <>
            {/* Cav therapist picker for regular service */}
            {!isCavCapable(form.therapist) && (
              <Field label="機械担当セラピスト（任意）">
                <select value={form.cavTherapist} onChange={e => {
                  const cav = e.target.value;
                  // Recalculate the split immediately using whatever total was already typed —
                  // otherwise picking/changing the machine therapist AFTER typing the total leaves
                  // the machine's price/tip stuck at 0 until the total is retyped.
                  const svcTotal = Number(form.totalServiceInput || (Number(form.price) + Number(form.cavPrice)) || 0);
                  const tipTotal = Number(form.totalTipInput || (Number(form.tip) + Number(form.cavTip)) || 0);
                  if (cav && isRegularWeightLoss && svcTotal > 0) {
                    const cavPrice = REGULAR_WL_CAV_PRICE;
                    const price = Math.round((svcTotal - cavPrice) * 100) / 100;
                    const truePriceTotal = svcTotal + Number(form.depositApplied || 0);
                    const pct = truePriceTotal > 0 ? tipTotal / truePriceTotal : 0;
                    const cavTip = Math.round(cavPrice * pct * 100) / 100;
                    const tip = Math.round((tipTotal - cavTip) * 100) / 100;
                    setForm(f => ({ ...f, cavTherapist: cav, price, cavPrice, tip, cavTip }));
                  } else if (cav && !isRegularWeightLoss && svcTotal > 0) {
                    // Everything else (e.g. Improving Posture) defaults to a fixed 15min machine
                    // slot, rest to body — same idea as Weight Loss's fixed 40min, just a different
                    // number. Still editable afterward in the minutes box.
                    const cavMins = Number(form.cavMins) || 15;
                    const bodyMins = Math.max(0, Number(form.duration) - cavMins);
                    const allMins = bodyMins + cavMins;
                    // Round the cav (machine) side first, then derive body as the exact remainder —
                    // rounding both sides independently can overshoot the original total by a cent.
                    const cavPrice = allMins > 0 ? Math.round(svcTotal * cavMins / allMins * 100) / 100 : 0;
                    const price = allMins > 0 ? Math.round((svcTotal - cavPrice) * 100) / 100 : svcTotal;
                    const cavTip = allMins > 0 ? Math.round(tipTotal * cavMins / allMins * 100) / 100 : 0;
                    const tip = allMins > 0 ? Math.round((tipTotal - cavTip) * 100) / 100 : tipTotal;
                    setForm(f => ({ ...f, cavTherapist: cav, bodyMins, cavMins, price, cavPrice, tip, cavTip }));
                  } else if (!cav) {
                    setForm(f => ({ ...f, cavTherapist: "", price: svcTotal, cavPrice: 0, tip: tipTotal, cavTip: 0 }));
                  } else {
                    set("cavTherapist", cav);
                  }
                }} style={inputStyle}>
                  <option value="">なし（機械なし）</option>
                  {CAV_CAPABLE.filter(t => t !== form.therapist).map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
            )}

            {/* Gift Card Usage — hidden for GC消化 (entire appointment IS the GC) and PR無料 */}
            {!form.isGiftCard && !form.isPromo && <div style={{ background: "#E0F2F1", borderRadius: 10, padding: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 700, fontSize: 13, color: "#00796B" }}>🎁 ギフトカード使用</span>
                <button onClick={() => set("giftCardUsed", form.giftCardUsed === 0 ? ((Number(form.price)||0) + (Number(form.cavPrice)||0) || "") : 0)}
                  style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: form.giftCardUsed !== 0 ? "#00796B" : "#B2DFDB", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                  {form.giftCardUsed !== 0 ? "ON ✓" : "OFF"}
                </button>
              </div>
              {(Number(form.giftCardUsed) > 0 || form.giftCardUsed === "") && (
                <div style={{ marginTop: 8 }}>
                  <div style={{ fontSize: 10, color: "#00796B", marginBottom: 4 }}>使用金額 ($)</div>
                  <input type="number" value={form.giftCardUsed === "" ? "" : form.giftCardUsed}
                    onChange={e => set("giftCardUsed", e.target.value)}
                    style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="例：150" />
                </div>
              )}
            </div>}

            {/* デポジット使用 */}
            {!form.isGiftCard && !form.isPromo && (
              <div style={{ background: "#E8F5E9", borderRadius: 10, padding: 12, border: "1.5px solid #A5D6A7" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 700, fontSize: 13, color: "#2E7D32" }}>💰 デポジット使用</span>
                  <button onClick={() => set("depositApplied", Number(form.depositApplied) > 0 ? 0 : 20)}
                    style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.depositApplied) > 0 ? "#2E7D32" : "#C8E6C9", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                    {Number(form.depositApplied) > 0 ? "ON ✓" : "OFF"}
                  </button>
                </div>
                {Number(form.depositApplied) > 0 && (
                  <div style={{ marginTop: 8, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <div>
                      <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>デポジット金額 ($)</div>
                      <input type="number" value={form.depositApplied || ""}
                        onFocus={e => e.target.select()}
                        onChange={e => set("depositApplied", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#81C784" }} placeholder="例：20" />
                    </div>
                    <div>
                      <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>支払われた日{errors.includes("depositPaidDate") && " ⚠️"}</div>
                      <input type="date" value={form.depositPaidDate || ""}
                        onChange={e => set("depositPaidDate", e.target.value)}
                        style={{ ...inputStyle, borderColor: errors.includes("depositPaidDate") ? "#C62828" : "#81C784", borderWidth: errors.includes("depositPaidDate") ? 2 : 1 }} />
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Total service + tip input */}
            <div style={{ background: "#F9F9F9", borderRadius: 10, padding: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "#333", marginBottom: 10 }}>
                💆 施術料金・チップ（合計入力）
                {Number(form.depositApplied) > 0 && <span style={{ fontSize: 11, color: "#2E7D32", fontWeight: 600, marginLeft: 8 }}>← もらった金額を入力</span>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <Field label="施術合計 ($)" error={errors.includes("price")}>
                  <input type="number" value={form.totalServiceInput || form.price || ""}
                    onFocus={e => e.target.select()}
                    onChange={e => {
                      const total = Number(e.target.value);
                      if (form.cavTherapist && isRegularWeightLoss) {
                        set("cavPrice", REGULAR_WL_CAV_PRICE);
                        set("price", Math.round((total - REGULAR_WL_CAV_PRICE) * 100) / 100);
                      } else {
                        const bodyMins = Number(form.bodyMins || form.duration);
                        const cavMins = Number(form.cavMins || 0);
                        const allMins = bodyMins + cavMins;
                        if (form.cavTherapist && cavMins > 0 && allMins > 0) {
                          // Round the machine side first, body gets the exact remainder (rounding
                          // both independently can overshoot the entered total by a cent).
                          const cavPrice = Math.round(total * cavMins / allMins * 100) / 100;
                          set("cavPrice", cavPrice);
                          set("price", Math.round((total - cavPrice) * 100) / 100);
                        } else {
                          set("price", total);
                          set("cavPrice", 0);
                        }
                      }
                      // Store the raw typed text (not the Number-converted value) so an
                      // in-progress decimal like "20." isn't stripped back to "20" by the
                      // next re-render before the digits after the dot are typed.
                      set("totalServiceInput", e.target.value);
                    }}
                    style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="例: 158" />
                </Field>
                <Field label="チップ合計 ($)">
                  <input type="number" value={form.totalTipInput || form.tip || ""}
                    onFocus={e => e.target.select()}
                    onChange={e => {
                      const total = Number(e.target.value);
                      if (form.cavTherapist && isRegularWeightLoss) {
                        // % is based on the TRUE course price (today's total + any deposit already
                        // paid), not just what's collected today — e.g. a $267 repeat-guest visit with
                        // a $20.9 deposit still bases the tip % on $267, not $246.1.
                        const svcTotal = Number(form.totalServiceInput || (Number(form.price) + Number(form.cavPrice)) || 0);
                        const truePriceTotal = svcTotal + Number(form.depositApplied || 0);
                        const pct = truePriceTotal > 0 ? total / truePriceTotal : 0;
                        const cavTip = Math.round(REGULAR_WL_CAV_PRICE * pct * 100) / 100;
                        set("cavTip", cavTip);
                        set("tip", Math.round((total - cavTip) * 100) / 100);
                      } else {
                        const bodyMins = Number(form.bodyMins || form.duration);
                        const cavMins = Number(form.cavMins || 0);
                        const allMins = bodyMins + cavMins;
                        if (form.cavTherapist && cavMins > 0 && allMins > 0) {
                          const cavTip = Math.round(total * cavMins / allMins * 100) / 100;
                          set("cavTip", cavTip);
                          set("tip", Math.round((total - cavTip) * 100) / 100);
                        } else {
                          set("tip", total);
                          set("cavTip", 0);
                        }
                      }
                      // Same fix as 施術合計: keep the raw typed text so a trailing "." mid-decimal
                      // isn't wiped out by the controlled re-render before it's fully typed.
                      set("totalTipInput", e.target.value);
                    }}
                    style={inputStyle} placeholder="例: 30" />
                </Field>
              </div>

              {/* Grand total — shown immediately below the amount/tip inputs so staff see the
                  combined total right away, without scrolling further down the form. */}
              {(() => {
                const svc = r2(Number(form.price||0) + Number(form.cavPrice||0));
                const tip = r2(Number(form.tip||0) + Number(form.cavTip||0));
                if (svc + tip === 0) return null;
                if (form.isGiftCard) {
                  return (
                    <div style={{ marginTop: 10, background: "#78350F", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontSize: 11, color: "#FDE68A" }}>
                        🎁 施術 <strong style={{ color: "#fff" }}>${svc}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}</strong>
                        {form.cavTherapist && !isDualLicense(form.therapist) && (
                          <span style={{ fontSize: 10, color: "#FDE68A", display: "block" }}>
                            {form.therapist} ${form.price||0}{Number(form.tip||0) > 0 ? ` +チップ$${form.tip}` : ""} / {form.cavTherapist} ${form.cavPrice||0}{Number(form.cavTip||0) > 0 ? ` +チップ$${form.cavTip}` : ""}
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#FDE68A", fontSize: 11 }}>GC消化合計　</span>
                        <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip}</span>
                      </div>
                    </div>
                  );
                }
                if (form.isPromo) {
                  return (
                    <div style={{ marginTop: 10, background: "#0D47A1", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontSize: 11, color: "#BBDEFB" }}>
                        📸 施術 <strong style={{ color: "#fff" }}>${svc}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}</strong>
                        {form.cavTherapist && !isDualLicense(form.therapist) && (
                          <span style={{ fontSize: 10, color: "#BBDEFB", display: "block" }}>
                            {form.therapist} ${form.price||0}{Number(form.tip||0) > 0 ? ` +チップ$${form.tip}` : ""} / {form.cavTherapist} ${form.cavPrice||0}{Number(form.cavTip||0) > 0 ? ` +チップ$${form.cavTip}` : ""}
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#BBDEFB", fontSize: 11 }}>PR無料合計　</span>
                        <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip}</span>
                      </div>
                    </div>
                  );
                }
                const gc = Number(form.giftCardUsed || 0);
                const gcSvc = Math.min(gc, svc);
                const gcTip = Math.min(Math.max(0, gc - gcSvc), tip);
                const receivedToday = r2((svc - gcSvc) + (tip - gcTip));
                const dep = Number(form.depositApplied || 0);
                // Payroll (deposit-inclusive) split per therapist — Weight Loss keeps the machine
                // fixed at $116 and gives the deposit entirely to the body therapist; everything
                // else splits the deposit-inclusive total by treatment MINUTES (body gets its
                // minute-share, machine gets the remainder), not by the price ratio, so the split
                // matches actual time spent regardless of how price/cavPrice happened to be entered.
                const hasCav = !!form.cavTherapist;
                const bodyMinsForSplit = Number(form.bodyMins) || Math.max(0, Number(form.duration||0) - (Number(form.cavMins)||15));
                const cavMinsForSplit = Number(form.cavMins) || 15;
                const allMinsForSplit = bodyMinsForSplit + cavMinsForSplit;
                let bodyPayroll, cavPayroll;
                if (dep > 0 && isRegularWeightLoss) {
                  bodyPayroll = Math.round((Number(form.price || 0) + dep) * 100) / 100;
                  cavPayroll = Number(form.cavPrice || 0);
                } else if (dep > 0 && hasCav && allMinsForSplit > 0) {
                  bodyPayroll = Math.round((svc + dep) * bodyMinsForSplit / allMinsForSplit * 100) / 100;
                  cavPayroll = Math.round(((svc + dep) - bodyPayroll) * 100) / 100;
                } else if (dep > 0 && svc > 0) {
                  bodyPayroll = Math.round((svc + dep) * Number(form.price || 0) / svc * 100) / 100;
                  cavPayroll = Math.round((svc + dep) * Number(form.cavPrice || 0) / svc * 100) / 100;
                } else {
                  bodyPayroll = Number(form.price || 0);
                  cavPayroll = Number(form.cavPrice || 0);
                }
                const bodySvcShown = dep > 0 ? bodyPayroll : Number(form.price || 0);
                const cavSvcShown = dep > 0 ? cavPayroll : Number(form.cavPrice || 0);
                const bodyTipShown = Number(form.tip || 0);
                const cavTipShown = Number(form.cavTip || 0);
                return (
                  <div style={{ marginTop: 10, background: "#0D4F4F", borderRadius: 8, padding: "10px 14px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontSize: 12, color: "#B2EBF2" }}>
                        施術 <strong style={{ color: "#fff" }}>${svc}{gcSvc >= svc && svc > 0 ? "🎁" : form.paymentType==="card"?"💳":"💵"}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}{gcTip >= tip && tip > 0 ? "🎁" : form.tipPaymentType==="card"?"💳":"💵"}</strong>
                        {gc > 0 && (
                          <span style={{ fontSize: 10, color: "#80CBC4", display: "block" }}>
                            🎁 ギフトカードで${gcSvc + gcTip}消化済み（本日の売上には含まれません）
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#B2EBF2", fontSize: 11 }}>{gc > 0 ? "本日の受取　" : "もらった金額　"}</span>
                        <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${gc > 0 ? receivedToday : r2(svc + tip)}</span>
                      </div>
                    </div>
                    {form.cavTherapist && !isDualLicense(form.therapist) && (
                      <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <div style={{ flex: 1, minWidth: 130, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                          <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.therapist}</div>
                          <div style={{ fontSize: 13, color: "#E0F7FA" }}>施術 ${bodySvcShown}　チップ ${bodyTipShown}</div>
                          <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>合計 ${r2(bodySvcShown + bodyTipShown)}</div>
                        </div>
                        <div style={{ flex: 1, minWidth: 130, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                          <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.cavTherapist}（機械）</div>
                          <div style={{ fontSize: 13, color: "#E0F7FA" }}>施術 ${cavSvcShown}　チップ ${cavTipShown}</div>
                          <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>合計 ${r2(cavSvcShown + cavTipShown)}</div>
                        </div>
                      </div>
                    )}
                    {/* Single therapist (no machine helper) with a deposit — the deposit-inclusive
                        payroll figure otherwise has nowhere to show, since the two-card layout
                        above only renders when there's a machine therapist to split against. */}
                    {!form.cavTherapist && dep > 0 && (
                      <div style={{ marginTop: 10, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.therapist}</div>
                        <div style={{ fontSize: 13, color: "#E0F7FA" }}>施術 ${bodyPayroll}　チップ ${bodyTipShown}</div>
                        <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>合計 ${r2(bodyPayroll + bodyTipShown)}</div>
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* Service payment method — right next to the total, before tip payment method */}
              {!form.isGiftCard && !form.isPromo && (() => {
                const gc = Number(form.giftCardUsed || 0);
                const svc = Number(form.price || 0);
                const gcSvc = Math.min(gc, svc);
                const svcCovered = gc > 0 && svc > 0 && gcSvc >= svc;
                if (svc === 0) return null;
                return (
                  <div style={{ marginTop: 10, background: "#F0F4FF", borderRadius: 10, padding: "10px 12px" }}>
                    <Field label={gc > 0 && !svcCovered ? "施術（差額）支払方法" : "施術 支払方法"} error={!svcCovered && errors.includes("paymentType")}>
                      {svcCovered ? (
                        <div style={{ fontSize: 12, color: "#00796B", fontWeight: 700, padding: "10px 0" }}>🎁 ギフトカード払い済み</div>
                      ) : (
                        <>
                          <button onClick={() => set("svcSplitPayment", !form.svcSplitPayment)}
                            style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.svcSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.svcSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.svcSplitPayment ? "#2E7D32" : "#888" }}>
                            {form.svcSplitPayment ? "☑" : "☐"} 現金・カードに分けて支払い
                          </button>
                          {form.svcSplitPayment ? (
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 現金 ($)</div>
                                <input type="number" value={form.svcCashPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, svcCashPortion: raw, svcCardPortion: Math.max(0, Number(f.price || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="例：50" />
                              </div>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 カード ($)</div>
                                <input type="number" value={form.svcCardPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, svcCardPortion: raw, svcCashPortion: Math.max(0, Number(f.price || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="例：50" />
                              </div>
                            </div>
                          ) : (
                            <PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} />
                          )}
                        </>
                      )}
                    </Field>
                  </div>
                );
              })()}

              {/* Tip payment method — right next to the tip amount so it's one unit, not two separate steps */}
              {Number(form.tip || 0) > 0 && (() => {
                const gc = Number(form.giftCardUsed || 0);
                const svc = Number(form.price || 0);
                const tip = Number(form.tip || 0);
                const gcSvc = Math.min(gc, svc);
                const gcTip = Math.min(Math.max(0, gc - gcSvc), tip);
                const tipCovered = gc > 0 && tip > 0 && gcTip >= tip;
                return (
                  <div style={{ marginTop: 8 }}>
                    <Field label="チップ 支払方法" error={!tipCovered && errors.includes("tipPaymentType")}>
                      {tipCovered ? (
                        <div style={{ fontSize: 12, color: "#00796B", fontWeight: 700, padding: "6px 0" }}>🎁 ギフトカード払い済み</div>
                      ) : (
                        <>
                          <button onClick={() => set("tipSplitPayment", !form.tipSplitPayment)}
                            style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.tipSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.tipSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.tipSplitPayment ? "#2E7D32" : "#888" }}>
                            {form.tipSplitPayment ? "☑" : "☐"} 現金・カードに分けて支払い
                          </button>
                          {form.tipSplitPayment ? (
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 現金 ($)</div>
                                <input type="number" value={form.tipCashPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, tipCashPortion: raw, tipCardPortion: Math.max(0, Number(f.tip || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="例：10" />
                              </div>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 カード ($)</div>
                                <input type="number" value={form.tipCardPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, tipCardPortion: raw, tipCashPortion: Math.max(0, Number(f.tip || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="例：20" />
                              </div>
                            </div>
                          ) : (
                            <PaymentToggle value={form.tipPaymentType} onChange={v => set("tipPaymentType", v)} />
                          )}
                        </>
                      )}
                    </Field>
                  </div>
                );
              })()}

              {/* Weight Loss — fixed cav allocation, no minutes needed */}
              {form.cavTherapist && isRegularWeightLoss && (Number(form.price) > 0 || Number(form.cavPrice) > 0) && (() => {
                const svcTotal = Number(form.totalServiceInput || (Number(form.price) + Number(form.cavPrice)) || 0);
                const truePriceTotal = svcTotal + Number(form.depositApplied || 0);
                const tipTotal = Number(form.totalTipInput || (Number(form.tip) + Number(form.cavTip)) || 0);
                const expectedCavPrice = REGULAR_WL_CAV_PRICE;
                const expectedCavTip = truePriceTotal > 0 ? Math.round(expectedCavPrice * tipTotal / truePriceTotal * 100) / 100 : 0;
                const priceStale = Math.abs(Number(form.cavPrice || 0) - expectedCavPrice) > 0.05;
                // Tip goes stale too when the deposit amount is toggled/changed AFTER the tip was
                // typed in — the % split isn't re-derived automatically, so cavTip can drift from
                // what it should be against the (now different) true course total.
                const tipStale = tipTotal > 0 && Math.abs(Number(form.cavTip || 0) - expectedCavTip) > 0.05;
                const isStale = priceStale || tipStale;
                if (!isStale) return null;
                return (
                  <div style={{ marginTop: 10, background: "#FFF3E0", borderRadius: 8, padding: 10, fontSize: 12, color: "#E65100" }}>
                    <button onClick={() => {
                      const cavPrice = expectedCavPrice;
                      const price = Math.round((svcTotal - cavPrice) * 100) / 100;
                      const pct = truePriceTotal > 0 ? tipTotal / truePriceTotal : 0;
                      const cavTip = Math.round(cavPrice * pct * 100) / 100;
                      const tip = Math.round((tipTotal - cavTip) * 100) / 100;
                      setForm(f => ({ ...f, price, cavPrice, tip, cavTip }));
                    }} style={{ padding: "5px 10px", borderRadius: 6, border: "none", background: "#E65100", color: "#fff", fontWeight: 700, fontSize: 11, cursor: "pointer" }}>
                      🔄 Weight Lossのルールで再計算する（{priceStale ? "現在$116になっていません" : "デポジット変更後にチップ％がずれています"}）
                    </button>
                  </div>
                );
              })()}

              {/* Regular (non-Weight-Loss) split — detect drift from the minutes ratio (e.g. a tip
                  typed directly into the machine therapist's field instead of via チップ合計) and
                  offer a one-click fix, mirroring the Weight Loss stale-check above. */}
              {form.cavTherapist && !isRegularWeightLoss && (Number(form.price) > 0 || Number(form.cavPrice) > 0 || Number(form.tip) > 0 || Number(form.cavTip) > 0) && (() => {
                const bodyMins = Number(form.bodyMins) || Math.max(0, Number(form.duration||0) - (Number(form.cavMins)||15));
                const cavMins = Number(form.cavMins) || 15;
                const allMins = bodyMins + cavMins;
                if (allMins <= 0) return null;
                const svcTotal = Number(form.totalServiceInput ?? (Number(form.price) + Number(form.cavPrice || 0)));
                const tipTotal = Number(form.totalTipInput ?? (Number(form.tip) + Number(form.cavTip || 0)));
                const expectedCavPrice = svcTotal > 0 ? Math.round(svcTotal * cavMins / allMins * 100) / 100 : 0;
                const expectedCavTip = tipTotal > 0 ? Math.round(tipTotal * cavMins / allMins * 100) / 100 : 0;
                const priceStale = svcTotal > 0 && Math.abs(Number(form.cavPrice || 0) - expectedCavPrice) > 0.05;
                const tipStale = tipTotal > 0 && Math.abs(Number(form.cavTip || 0) - expectedCavTip) > 0.05;
                if (!priceStale && !tipStale) return null;
                return (
                  <div style={{ marginTop: 10, background: "#FFF3E0", borderRadius: 8, padding: 10, fontSize: 12, color: "#E65100" }}>
                    <button onClick={() => {
                      const cavPrice = expectedCavPrice;
                      const price = Math.round((svcTotal - cavPrice) * 100) / 100;
                      const cavTip = expectedCavTip;
                      const tip = Math.round((tipTotal - cavTip) * 100) / 100;
                      setForm(f => ({ ...f, price, cavPrice, tip, cavTip }));
                    }} style={{ padding: "5px 10px", borderRadius: 6, border: "none", background: "#E65100", color: "#fff", fontWeight: 700, fontSize: 11, cursor: "pointer" }}>
                      🔄 分数比率（{bodyMins}分／{cavMins}分）で再計算する（{priceStale ? "施術料金" : "チップ"}がずれています）
                    </button>
                  </div>
                );
              })()}

              {/* Minutes split. Weight Loss always uses a fixed 40min machine / rest-to-body split
                  (auto-filled elsewhere, no manual entry needed) — this UI is only for everything
                  else, where minutes actually drive the $ split. */}
              {form.cavTherapist && !isRegularWeightLoss && (
                <div style={{ marginTop: 10, background: "#EEF4FF", borderRadius: 8, padding: 10 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>⏱️ 担当分数を入力 → 自動振り分け</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label={`${form.therapist} 担当（分）`}>
                      <input type="number" value={form.bodyMins || ""}
                        onChange={e => {
                          const bodyMins = Number(e.target.value);
                          const cavMins = Number(form.duration) - bodyMins;
                          const allMins = Number(form.duration);
                          if (isRegularWeightLoss) { setForm(f => ({ ...f, bodyMins, cavMins })); return; }
                          const svcTotal = Number(form.totalServiceInput ?? (Number(form.price) + Number(form.cavPrice || 0)));
                          const tipTotal = Number(form.totalTipInput ?? (Number(form.tip) + Number(form.cavTip || 0)));
                          // Round the machine side first, body gets the exact remainder — rounding
                          // both sides independently can overshoot the entered total by a cent.
                          const cavPrice = allMins > 0 ? Math.round(svcTotal * cavMins / allMins * 100) / 100 : form.cavPrice;
                          const price = allMins > 0 ? Math.round((svcTotal - cavPrice) * 100) / 100 : form.price;
                          const cavTip = allMins > 0 ? Math.round(tipTotal * cavMins / allMins * 100) / 100 : form.cavTip;
                          const tip = allMins > 0 ? Math.round((tipTotal - cavTip) * 100) / 100 : form.tip;
                          setForm(f => ({
                            ...f,
                            bodyMins,
                            cavMins,
                            price,
                            cavPrice,
                            tip,
                            cavTip,
                          }));
                        }}
                        style={inputStyle} placeholder={`例: ${form.duration - (isRegularWeightLoss ? 40 : 15)}`} />
                    </Field>
                    <Field label={`${form.cavTherapist} 担当（分）`}>
                      <input type="number" value={form.cavMins || ""}
                        onChange={e => {
                          const cavMins = Number(e.target.value);
                          const bodyMins = Number(form.duration) - cavMins;
                          const allMins = Number(form.duration);
                          if (isRegularWeightLoss) { setForm(f => ({ ...f, bodyMins, cavMins })); return; }
                          const svcTotal = Number(form.totalServiceInput ?? (Number(form.price) + Number(form.cavPrice || 0)));
                          const tipTotal = Number(form.totalTipInput ?? (Number(form.tip) + Number(form.cavTip || 0)));
                          // Round the machine side first, body gets the exact remainder — rounding
                          // both sides independently can overshoot the entered total by a cent.
                          const cavPrice = allMins > 0 ? Math.round(svcTotal * cavMins / allMins * 100) / 100 : form.cavPrice;
                          const price = allMins > 0 ? Math.round((svcTotal - cavPrice) * 100) / 100 : form.price;
                          const cavTip = allMins > 0 ? Math.round(tipTotal * cavMins / allMins * 100) / 100 : form.cavTip;
                          const tip = allMins > 0 ? Math.round((tipTotal - cavTip) * 100) / 100 : form.tip;
                          setForm(f => ({
                            ...f,
                            bodyMins,
                            cavMins,
                            price,
                            cavPrice,
                            tip,
                            cavTip,
                          }));
                        }}
                        style={inputStyle} placeholder={isRegularWeightLoss ? "例: 40" : "例: 15"} />
                    </Field>
                  </div>

                </div>
              )}

            </div>

            {/* GC消化 banner — replaces payment type selectors */}
            {form.isGiftCard && (
              <div style={{ background: "#FFFDE7", borderRadius: 10, padding: 12, border: "2px solid #F59E0B", marginTop: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 13, color: "#B45309" }}>🎁 GC消化 — 全額ギフトカードにて領収済</div>
                <div style={{ fontSize: 11, color: "#92400E", marginTop: 4 }}>事前購入のギフトカードを使用。本日の売上には計上されません。</div>
              </div>
            )}

            {/* PR無料 banner — replaces payment type selectors */}
            {form.isPromo && (
              <div style={{ background: "#E3F2FD", borderRadius: 10, padding: 12, border: "2px solid #1565C0", marginTop: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 13, color: "#1565C0" }}>📸 PR無料 — インスタグラマーなど無料施術</div>
                <div style={{ fontSize: 11, color: "#0D47A1", marginTop: 4 }}>お客様からは料金をいただいていません。本日の売上には計上されませんが、スタッフの施術実績（給料計算用）には入力した金額が反映されます。</div>
              </div>
            )}


          </>
        )}

        {/* ── 当日の追加購入 ── */}
        {(() => {
          const tags = form.purchaseTags || [];
          const hasTag = id => tags.includes(id);
          const toggleTag = id => set("purchaseTags", hasTag(id) ? tags.filter(t => t !== id) : [...tags, id]);
          const payBtns = (field, active, activeColor) => (
            <div style={{ display: "flex", gap: 5 }}>
              {[["cash","💵 Cash"],["card","💳 Card"]].map(([v,l]) => (
                <button key={v} onClick={() => set(field, v)}
                  style={{ flex:1, padding:"7px 4px", borderRadius:8, border:`2px solid ${form[field]===v?activeColor:"#DDD"}`, background:form[field]===v?"#FFF":"#fff", cursor:"pointer", fontWeight:form[field]===v?700:400, fontSize:11, color:form[field]===v?activeColor:"#AAA" }}>
                  {l}
                </button>
              ))}
            </div>
          );
          return (
            <div style={{ border: "2px solid #F06292", borderRadius: 12, overflow: "hidden" }}>
              <div style={{ background: "#FCE4EC", padding: "8px 12px", fontSize: 11, fontWeight: 700, color: "#880E4F" }}>
                📌 当日の追加購入（タップでON/OFF）
              </div>
              {/* Tag toggle buttons */}
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, padding: "10px 12px", background: "#FFF" }}>
                {PURCHASE_TAGS.map(tag => (
                  <button key={tag.id} onClick={() => toggleTag(tag.id)} style={{
                    padding: "7px 13px", borderRadius: 20,
                    border: `2px solid ${hasTag(tag.id) ? tag.color : "#DDD"}`,
                    background: hasTag(tag.id) ? tag.bg : "#F8F8F8",
                    color: hasTag(tag.id) ? tag.color : "#AAA",
                    fontWeight: hasTag(tag.id) ? 700 : 400, fontSize: 12, cursor: "pointer",
                  }}>
                    {tag.label}
                  </button>
                ))}
              </div>

              {/* 🎟️ チケット新規購入 fields */}
              {hasTag("newTicket") && (
                <div style={{ background: "#FFEBEE", padding: "10px 12px", borderTop: "1px solid #FFCDD2", display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#B71C1C" }}>🎟️ チケット新規購入 — 詳細</div>

                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>コース</div>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {MENU_OPTIONS.map(({ group, prefix }) => {
                        const selected = form.newTicketMenu?.startsWith(prefix);
                        return (
                          <button key={prefix} onClick={() => {
                            const dur = MENU_OPTIONS.find(m => m.prefix === prefix).durations[0];
                            applyNewTicketMenuPrice(`${prefix}-${dur}-${form.newTicketTotal || 3}`, form.newTicketTotal || 3);
                          }} style={{
                            padding: "6px 10px", borderRadius: 8,
                            border: `2px solid ${selected ? "#B71C1C" : "#DDD"}`,
                            background: selected ? "#B71C1C" : "#fff",
                            cursor: "pointer", fontWeight: 700, fontSize: 11,
                            color: selected ? "#fff" : "#888"
                          }}>{group}</button>
                        );
                      })}
                    </div>
                  </div>

                  {form.newTicketMenu && (() => {
                    const prefix = form.newTicketMenu.split("-")[0];
                    const group = MENU_OPTIONS.find(m => m.prefix === prefix);
                    const currentDur = Number(form.newTicketMenu.split("-")[1]);
                    return group ? (
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>時間</div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {group.durations.map(dur => (
                            <button key={dur} onClick={() => applyNewTicketMenuPrice(`${prefix}-${dur}-${form.newTicketTotal || 3}`, form.newTicketTotal || 3)}
                              style={{
                                padding: "6px 12px", borderRadius: 8,
                                border: `2px solid ${currentDur === dur ? "#B71C1C" : "#DDD"}`,
                                background: currentDur === dur ? "#B71C1C" : "#fff",
                                cursor: "pointer", fontWeight: 700, fontSize: 11,
                                color: currentDur === dur ? "#fff" : "#888"
                              }}>{dur}分</button>
                          ))}
                        </div>
                      </div>
                    ) : null;
                  })()}

                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>コース回数</div>
                    <div style={{ display: "flex", gap: 8 }}>
                      {[3,5].map(n => (
                        <button key={n} onClick={() => {
                          set("newTicketTotal", n);
                          if (form.newTicketMenu) {
                            const parts = form.newTicketMenu.split("-");
                            applyNewTicketMenuPrice(`${parts[0]}-${parts[1]}-${n}`, n);
                          }
                        }} style={{ flex: 1, padding: "7px", borderRadius: 8, border: `2px solid ${form.newTicketTotal===n?"#B71C1C":"#DDD"}`, background: form.newTicketTotal===n?"#B71C1C":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.newTicketTotal===n?"#fff":"#888" }}>
                          {n}回コース
                        </button>
                      ))}
                    </div>
                  </div>

                  <div style={{ display: "flex", gap: 8 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術料 ($)（自動入力・修正可）</div>
                      <input type="number" value={form.newTicketAmount || ""} onChange={e => set("newTicketAmount", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="例：719" disabled={form.newTicketSplitPayment} />
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>チップ ($)</div>
                      <input type="number" value={form.newTicketTip || ""} onChange={e => set("newTicketTip", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="例：46" />
                    </div>
                  </div>
                  <div>
                    <button onClick={() => set("newTicketSplitPayment", !form.newTicketSplitPayment)}
                      style={{ padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.newTicketSplitPayment ? "#B71C1C" : "#DDD"}`, background: form.newTicketSplitPayment ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.newTicketSplitPayment ? "#B71C1C" : "#888" }}>
                      {form.newTicketSplitPayment ? "☑" : "☐"} 現金・カードに分けて支払い
                    </button>
                  </div>
                  {form.newTicketSplitPayment ? (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術料の内訳（現金＋カード＝施術料の合計になるように）</div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 現金 ($)</div>
                          <input type="number" value={form.newTicketCashPortion || ""} onChange={e => {
                            const raw = e.target.value;
                            setForm(f => ({ ...f, newTicketCashPortion: raw, newTicketAmount: (Number(raw) || 0) + Number(f.newTicketCardPortion||0) }));
                          }} style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="例：500" />
                        </div>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 カード ($)</div>
                          <input type="number" value={form.newTicketCardPortion || ""} onChange={e => {
                            const raw = e.target.value;
                            setForm(f => ({ ...f, newTicketCardPortion: raw, newTicketAmount: Number(f.newTicketCashPortion||0) + (Number(raw) || 0) }));
                          }} style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="例：211" />
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>施術料の支払方法</div>
                      {payBtns("newTicketPaymentType", true, "#B71C1C")}
                    </div>
                  )}
                  {Number(form.newTicketTip) > 0 && (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>チップの支払方法</div>
                      {payBtns("newTicketTipPaymentType", true, "#E65100")}
                    </div>
                  )}
                  {(Number(form.newTicketAmount) > 0 || Number(form.newTicketTip) > 0) && (
                    <div style={{ textAlign: "right", fontSize: 12, fontWeight: 700, color: "#B71C1C" }}>
                      合計 ${(Number(form.newTicketAmount||0) + Number(form.newTicketTip||0)).toFixed(0)}
                    </div>
                  )}
                </div>
              )}

              {/* 🛍️ 物販購入 fields — supports multiple products in one visit (2nd/3rd item stored in extraRetailItems[]) */}
              {hasTag("retail") && (() => {
                const items = [
                  { productName: form.retailProductName, amount: form.retailPurchaseAmount, paymentType: form.retailPurchasePaymentType, sellers: form.retailSellers },
                  ...(form.extraRetailItems || [])
                ];
                const updateItem = (idx, patch) => {
                  if (idx === 0) {
                    if ("productName" in patch) set("retailProductName", patch.productName);
                    if ("amount" in patch) set("retailPurchaseAmount", patch.amount);
                    if ("paymentType" in patch) set("retailPurchasePaymentType", patch.paymentType);
                    if ("sellers" in patch) set("retailSellers", patch.sellers);
                  } else {
                    const extras = form.extraRetailItems || [];
                    set("extraRetailItems", extras.map((it, i) => i === idx - 1 ? { ...it, ...patch } : it));
                  }
                };
                const removeItem = (idx) => {
                  if (idx === 0) return;
                  const extras = form.extraRetailItems || [];
                  set("extraRetailItems", extras.filter((_, i) => i !== idx - 1));
                };
                return (
                  <div style={{ background: "#F3E5F5", padding: "10px 12px", borderTop: "1px solid #E1BEE7", display: "flex", flexDirection: "column", gap: 10 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A" }}>🛍️ 物販購入 — 詳細</div>
                    {items.map((item, idx) => (
                      <div key={idx} style={{ display: "flex", flexDirection: "column", gap: 8, ...(idx > 0 ? { borderTop: "1px dashed #CE93D8", paddingTop: 8 } : {}) }}>
                        {idx > 0 && (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: 10, fontWeight: 700, color: "#6A1B9A" }}>商品 {idx + 1}</span>
                            <button onClick={() => removeItem(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11, color: "#AAA" }}>✕ 削除</button>
                          </div>
                        )}
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>商品名</div>
                          <select value={RETAIL_PRODUCTS.find(p => p.name === item.productName) ? item.productName : (item.productName ? "__other__" : "")} onChange={e => {
                            const val = e.target.value;
                            const prod = RETAIL_PRODUCTS.find(p => p.name === val);
                            updateItem(idx, { productName: val === "__other__" ? "" : val, amount: prod?.price > 0 ? prod.price : item.amount });
                          }} style={{ ...inputStyle, borderColor: "#CE93D8" }}>
                            <option value="">— 商品を選択 —</option>
                            {RETAIL_PRODUCTS.map(p => (
                              <option key={p.name} value={p.name}>{p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
                            ))}
                            <option value="__other__">その他</option>
                          </select>
                          {(!RETAIL_PRODUCTS.find(p => p.name === item.productName)) && (
                            <input type="text" value={item.productName || ""} placeholder="商品名を入力" style={{ ...inputStyle, borderColor: "#CE93D8", marginTop: 4 }}
                              onChange={e => updateItem(idx, { productName: e.target.value })} />
                          )}
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>金額 ($)</div>
                          <input type="number" value={item.amount || ""} onChange={e => updateItem(idx, { amount: e.target.value })}
                            style={{ ...inputStyle, borderColor: "#CE93D8" }} placeholder="例：30" />
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法</div>
                          <div style={{ display: "flex", gap: 5 }}>
                            {[["cash","💵 Cash"],["card","💳 Card"]].map(([v,l]) => (
                              <button key={v} onClick={() => updateItem(idx, { paymentType: v })}
                                style={{ flex:1, padding:"7px 4px", borderRadius:8, border:`2px solid ${item.paymentType===v?"#6A1B9A":"#DDD"}`, background:"#fff", cursor:"pointer", fontWeight:item.paymentType===v?700:400, fontSize:11, color:item.paymentType===v?"#6A1B9A":"#AAA" }}>
                                {l}
                              </button>
                            ))}
                          </div>
                        </div>
                        {(() => {
                          // Seller amounts represent each person's share of the *tax-excluded* commission
                          // base (same convention as the standalone 物販 modal) — staff still work out
                          // their own 10%/4% commission from that base and type the dollar figure in.
                          const afterTaxTotal = Math.round(Number(item.amount || 0) * (1 - RETAIL_TAX_RATE) * 100) / 100;
                          // Defaults to just the body therapist getting the full (tax-excluded) amount —
                          // most retail sales are one person. Add rows (up to 3) for the rare cases where
                          // the machine therapist, or a third therapist from a multi-service visit, shares it.
                          const sellers = item.sellers && item.sellers.length > 0
                            ? item.sellers
                            : [{ therapist: form.therapist, amount: afterTaxTotal }];
                          const sellersTotal = Math.round(sellers.reduce((s, sel) => s + Number(sel.amount || 0), 0) * 100) / 100;
                          const updSeller = (sidx, patch) => updateItem(idx, { sellers: sellers.map((sel, i) => i === sidx ? { ...sel, ...patch } : sel) });
                          return (
                            <div>
                              <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>
                                誰と分けるか（最大3名まで・税抜金額${afterTaxTotal}を分けます）{afterTaxTotal > 0 && Math.abs(sellersTotal - afterTaxTotal) > 0.15 ? " ⚠️ 合計が税抜金額と大きくずれています" : ""}
                              </div>
                              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                                {afterTaxTotal > 0 && (
                                  <button onClick={() => {
                                    const each = Math.round((afterTaxTotal / sellers.length) * 100) / 100;
                                    updateItem(idx, { sellers: sellers.map(sel => ({ ...sel, amount: each })) });
                                  }} style={{ padding: "5px 10px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#fff", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 11, alignSelf: "flex-start" }}>
                                    ⚖️ 税抜金額を{sellers.length > 1 ? `${sellers.length}人で均等に分ける` : "自動入力"}
                                  </button>
                                )}
                                {sellers.map((sel, sidx) => (
                                  <div key={sidx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                    <select value={sel.therapist} onChange={e => updSeller(sidx, { therapist: e.target.value })} style={{ ...inputStyle, flex: 2, borderColor: "#CE93D8" }}>
                                      <option value="">— 選択 —</option>
                                      {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                                    </select>
                                    <input type="number" value={sel.amount || ""} onChange={e => updSeller(sidx, { amount: e.target.value })}
                                      style={{ ...inputStyle, flex: 1, borderColor: "#CE93D8" }} placeholder="金額" />
                                    {sellers.length > 1 && (
                                      <button onClick={() => updateItem(idx, { sellers: sellers.filter((_, i) => i !== sidx) })}
                                        style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA" }}>✕</button>
                                    )}
                                  </div>
                                ))}
                                {sellers.length < 3 && (
                                  <button onClick={() => updateItem(idx, { sellers: [...sellers, { therapist: "", amount: 0 }] })}
                                    style={{ padding: "5px 10px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#F3E5F5", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 11, alignSelf: "flex-start" }}>
                                    ＋ 販売者を追加
                                  </button>
                                )}
                              </div>
                            </div>
                          );
                        })()}
                      </div>
                    ))}
                    <button onClick={() => set("extraRetailItems", [...(form.extraRetailItems || []), { productName: "", amount: 0, paymentType: "", sellers: [] }])}
                      style={{ padding: "7px 12px", borderRadius: 8, border: "2px dashed #6A1B9A", background: "#fff", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                      ＋ 商品を追加（2つ目・3つ目...）
                    </button>
                  </div>
                );
              })()}

              {/* 🎁 ギフトカード購入 fields */}
              {hasTag("giftCard") && (
                <div style={{ background: "#E0F2F1", padding: "10px 12px", borderTop: "1px solid #B2DFDB", display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#00796B" }}>🎁 ギフトカード購入 — 詳細</div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>金額 ($)</div>
                    <input type="number" value={form.giftCardPurchaseAmount || ""} onChange={e => set("giftCardPurchaseAmount", e.target.value)}
                      style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="例：100" />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法</div>
                    {payBtns("giftCardPurchasePaymentType", true, "#00796B")}
                  </div>
                </div>
              )}
            </div>
          );
        })()}

        <Field label="メモ">
          <input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="備考など" />
        </Field>
      </div>

      <ErrorBanner />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        {appt.id && <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>}
      </div>
    </Modal>
  );
}

function RetailModal({ retail, onSave, onClose }) {
  const [form, setForm] = useState({
    ...retail,
    sellers: retail.sellers || (retail.soldBy ? [{ therapist: retail.soldBy, amount: retail.price || 0 }] : [{ therapist: "", amount: 0 }]),
  });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [paymentError, setPaymentError] = useState(false);
  const handleSave = () => {
    if (!form.paymentType) { setPaymentError(true); return; }
    onSave(form);
  };
  const sellers = form.sellers || [];
  const sellersTotal = Math.round(sellers.reduce((s, sel) => s + Number(sel.amount || 0), 0) * 100) / 100;
  // Seller amounts represent each person's share of the *tax-adjusted* commission-eligible base
  // (10% of this goes to staff as commission, 4% for Maki), not the raw sale price.
  const afterTaxTotal = Math.round(Number(form.price || 0) * (1 - RETAIL_TAX_RATE) * 100) / 100;
  const updSeller = (idx, patch) => set("sellers", sellers.map((sel, i) => i === idx ? { ...sel, ...patch } : sel));
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#6A1B9A" }}>🛍️ 物販</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="商品名">
          <select value={RETAIL_PRODUCTS.find(p => p.name === form.item) ? form.item : (form.item ? "__other__" : "")}
            onChange={e => {
              const val = e.target.value;
              const prod = RETAIL_PRODUCTS.find(p => p.name === val);
              setForm(f => ({ ...f, item: val === "__other__" ? "" : val, price: prod?.price > 0 ? prod.price : f.price }));
            }} style={inputStyle}>
            <option value="">— 商品を選択 —</option>
            {RETAIL_PRODUCTS.map(p => (
              <option key={p.name} value={p.name}>{p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
            ))}
            <option value="__other__">その他（カスタム入力）</option>
          </select>
          {(!RETAIL_PRODUCTS.find(p => p.name === form.item)) && (
            <input type="text" value={form.item || ""} placeholder="商品名を入力" style={{ ...inputStyle, marginTop: 4 }}
              onChange={e => set("item", e.target.value)} />
          )}
        </Field>
        <Field label="金額 ($)"><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={inputStyle} /></Field>
        <Field label={`販売者（最大3名まで・税抜金額$${afterTaxTotal}を分けます）${Number(form.price) > 0 && Math.abs(sellersTotal - afterTaxTotal) > 0.15 ? " ⚠️ 合計が税抜金額と大きくずれています" : ""}`}>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Number(form.price) > 0 && (
              <button onClick={() => {
                // Rounded to the nearest $0.10 — penny-level precision isn't needed for this split.
                const each = Math.round((afterTaxTotal / sellers.length) * 100) / 100;
                set("sellers", sellers.map(sel => ({ ...sel, amount: each })));
              }} style={{ padding: "6px 12px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#fff", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 12, alignSelf: "flex-start" }}>
                ⚖️ 税抜金額を{sellers.length > 1 ? `${sellers.length}人で均等に分ける` : "自動入力"}
              </button>
            )}
            {sellers.map((sel, idx) => (
              <div key={idx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <select value={sel.therapist} onChange={e => updSeller(idx, { therapist: e.target.value })} style={{ ...inputStyle, flex: 2 }}>
                  <option value="">— 選択 —</option>
                  {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                <input type="number" value={sel.amount || ""} onChange={e => updSeller(idx, { amount: e.target.value })}
                  style={{ ...inputStyle, flex: 1 }} placeholder="金額" />
                {sellers.length > 1 && (
                  <button onClick={() => set("sellers", sellers.filter((_, i) => i !== idx))}
                    style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA" }}>✕</button>
                )}
              </div>
            ))}
            {sellers.length < 3 && (
              <button onClick={() => set("sellers", [...sellers, { therapist: "", amount: 0 }])}
                style={{ padding: "6px 12px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#F3E5F5", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 12, alignSelf: "flex-start" }}>
                ＋ 販売者を追加
              </button>
            )}
          </div>
        </Field>
        <Field label="支払方法" error={paymentError && !form.paymentType}><PaymentToggle value={form.paymentType} onChange={v => { set("paymentType", v); setPaymentError(false); }} /></Field>
      </div>
      <button onClick={handleSave} style={{ width: "100%", marginTop: 16, padding: "12px", borderRadius: 10, border: "none", background: "#6A1B9A", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
    </Modal>
  );
}

function DepositModal({ deposit, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...deposit });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const isDeposit = form.type === "deposit";
  const isGiftCard = form.type === "giftcard";
  const isCancellation = form.type === "cancellation";
  const payTypes = [["cash","💵 現金"],["card","💳 カード"]];
  const [errors, setErrors] = useState([]);
  const handleSave = () => {
    const errs = [];
    if (!form.paymentType) errs.push("paymentType");
    if (Number(form.tip) > 0 && !form.tipPaymentType) errs.push("tipPaymentType");
    if (errs.length > 0) { setErrors(errs); return; }
    onSave(form);
  };
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#1565C0" }}>{isCancellation ? "❌ キャンセル料" : isGiftCard ? "🎁 ギフトカード" : "💰 デポジット"}</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="種類">
          <div style={{ display: "flex", gap: 8 }}>
            {[["deposit","💰 デポジット"],["giftcard","🎁 ギフトカード"],["cancellation","❌ キャンセル料"]].map(([val,label]) => (
              <button key={val} onClick={() => set("type", val)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${form.type===val?"#1565C0":"#DDD"}`, background: form.type===val?"#E3F2FD":"#fff", cursor: "pointer", fontWeight: 700, color: form.type===val?"#1565C0":"#888" }}>{label}</button>
            ))}
          </div>
        </Field>
        <Field label="お客様名"><input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="例：田中様" /></Field>
        {!isCancellation && (
          <Field label={`来店予定日時（任意${isGiftCard ? "・全額ギフトで先払いの場合など" : ""}）`}>
            <div style={{ display: "flex", gap: 8 }}>
              <input type="date" value={form.appointmentDate || ""} onChange={e => set("appointmentDate", e.target.value)} style={{ ...inputStyle, flex: 1 }} />
              <input type="time" value={form.appointmentTime || ""} onChange={e => set("appointmentTime", e.target.value)} style={{ ...inputStyle, width: 100 }} />
            </div>
            {form.appointmentDate && form.clientName && (
              <div style={{ fontSize: 11, color: "#1565C0", marginTop: 4 }}>
                {new Date(form.appointmentDate + "T00:00").toLocaleDateString("ja-JP", { month: "long", day: "numeric" })}
                {form.appointmentTime && ` ${form.appointmentTime}`}の{form.clientName}様の{isGiftCard ? "ギフト先払い" : "デポジット"}
                　※日付は後から変更できます（予定が変わったらここを編集）
              </div>
            )}
          </Field>
        )}
        <Field label={isCancellation ? "キャンセル料 ($)" : "金額 ($)"}><input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} /></Field>
        <div style={{ background: "#F0F4FF", borderRadius: 10, padding: "10px 12px" }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#1565C0", marginBottom: 8 }}>💳 支払い方法（施術・チップ別々に設定可）</div>
          <Field label={isCancellation ? "キャンセル料 支払方法" : isDeposit ? "デポジット 支払方法" : "金額 支払方法"} error={errors.includes("paymentType")}>
            <div style={{ display: "flex", gap: 6 }}>
              {payTypes.map(([val, label]) => (
                <button key={val} onClick={() => { set("paymentType", val); setErrors(errors.filter(e => e !== "paymentType")); }}
                  style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#1565C0":"#DDD"}`, background: form.paymentType===val?"#E3F2FD":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#1565C0":"#888" }}>
                  {label}
                </button>
              ))}
            </div>
          </Field>
          <div style={{ marginTop: 8 }}>
            <Field label="チップ ($)（任意）">
              <input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} placeholder="チップなしは空欄" />
            </Field>
          </div>
          {Number(form.tip) > 0 && (
            <div style={{ marginTop: 8 }}>
              <Field label="チップ 支払方法" error={errors.includes("tipPaymentType")}>
                <div style={{ display: "flex", gap: 6 }}>
                  {payTypes.map(([val, label]) => (
                    <button key={val} onClick={() => { set("tipPaymentType", val); setErrors(errors.filter(e => e !== "tipPaymentType")); }}
                      style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.tipPaymentType===val?"#E65100":"#DDD"}`, background: form.tipPaymentType===val?"#FFF3E0":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.tipPaymentType===val?"#E65100":"#888" }}>
                      {label}
                    </button>
                  ))}
                </div>
              </Field>
            </div>
          )}
        </div>
        <Field label="メモ"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function TicketPurchaseModal({ tp, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...tp });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const payTypes = [["cash","💵 Cash"],["card","💳 Card"]];
  const [errors, setErrors] = useState([]);
  const handleSave = () => {
    const errs = [];
    if (!form.splitPayment && !form.paymentType) errs.push("paymentType");
    if (Number(form.tip) > 0 && !form.tipPaymentType) errs.push("tipPaymentType");
    if (errs.length > 0) { setErrors(errs); return; }
    onSave(form);
  };

  // Auto-fills the package service/tip from TICKET_PACKAGE_PRICES (still editable afterward,
  // e.g. for occasional discounts).
  const applyMenuPrice = (menu, total) => {
    const group = MENU_OPTIONS.find(m => menu.startsWith(m.prefix));
    const durLabel = menu.split("-")[1];
    const packageName = group ? `${group.group} ${durLabel}分 ×${total}回` : form.packageName;
    const pkg = TICKET_PACKAGE_PRICES[menu];
    setForm(f => ({
      ...f,
      ticketMenu: menu,
      priceVersion: "new",
      ticketTotal: total,
      packageName,
      ...(pkg ? { amount: pkg.service, tip: pkg.tip } : {}),
    }));
  };

  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#B71C1C" }}>🎟️ チケット新規購入</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="お客様名">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="例：Mio" />
        </Field>

        <div style={{ background: "#FFF5F5", borderRadius: 10, padding: 12 }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#B71C1C", marginBottom: 8 }}>🎟️ コース選択</div>

          <Field label="コース">
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {MENU_OPTIONS.map(({ group, prefix }) => {
                const selected = form.ticketMenu?.startsWith(prefix);
                return (
                  <button key={prefix} onClick={() => {
                    const dur = MENU_OPTIONS.find(m => m.prefix === prefix).durations[0];
                    applyMenuPrice(`${prefix}-${dur}-${form.ticketTotal || 3}`, form.ticketTotal || 3);
                  }} style={{
                    padding: "7px 10px", borderRadius: 8,
                    border: `2px solid ${selected ? "#B71C1C" : "#DDD"}`,
                    background: selected ? "#B71C1C" : "#fff",
                    cursor: "pointer", fontWeight: 700, fontSize: 11,
                    color: selected ? "#fff" : "#888"
                  }}>{group}</button>
                );
              })}
            </div>
          </Field>

          {form.ticketMenu && (() => {
            const prefix = form.ticketMenu.split("-")[0];
            const group = MENU_OPTIONS.find(m => m.prefix === prefix);
            const currentDur = Number(form.ticketMenu.split("-")[1]);
            return group ? (
              <Field label="時間">
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                  {group.durations.map(dur => (
                    <button key={dur} onClick={() => applyMenuPrice(`${prefix}-${dur}-${form.ticketTotal || 3}`, form.ticketTotal || 3)}
                      style={{
                        padding: "7px 12px", borderRadius: 8,
                        border: `2px solid ${currentDur === dur ? "#B71C1C" : "#DDD"}`,
                        background: currentDur === dur ? "#B71C1C" : "#fff",
                        cursor: "pointer", fontWeight: 700, fontSize: 12,
                        color: currentDur === dur ? "#fff" : "#888"
                      }}>{dur}分</button>
                  ))}
                </div>
              </Field>
            ) : null;
          })()}

          <Field label="コース回数">
            <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
              {[3,5].map(n => (
                <button key={n} onClick={() => {
                  set("ticketTotal", n);
                  if (form.ticketMenu) {
                    const parts = form.ticketMenu.split("-");
                    applyMenuPrice(`${parts[0]}-${parts[1]}-${n}`, n);
                  }
                }} style={{ flex: 1, padding: "8px", borderRadius: 8, border: `2px solid ${form.ticketTotal===n?"#B71C1C":"#DDD"}`, background: form.ticketTotal===n?"#B71C1C":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketTotal===n?"#fff":"#888" }}>
                  {n}回コース
                </button>
              ))}
            </div>
          </Field>
        </div>

        <Field label="パッケージ名（自動入力・修正可）">
          <input value={form.packageName} onChange={e => set("packageName", e.target.value)} style={inputStyle} placeholder="例：Improving Posture 90分 ×3回" />
        </Field>
        <Field label="施術料 ($)">
          <input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} placeholder="例：719" disabled={form.splitPayment} />
        </Field>
        <div>
          <button onClick={() => set("splitPayment", !form.splitPayment)}
            style={{ padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.splitPayment ? "#B71C1C" : "#DDD"}`, background: form.splitPayment ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.splitPayment ? "#B71C1C" : "#888" }}>
            {form.splitPayment ? "☑" : "☐"} 現金・カードに分けて支払い
          </button>
        </div>
        {form.splitPayment ? (
          <Field label="内訳（現金＋カード＝施術料の合計になるように）">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <div>
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 現金 ($)</div>
                <input type="number" value={form.cashPortion || ""} onChange={e => {
                  const raw = e.target.value;
                  setForm(f => ({ ...f, cashPortion: raw, amount: (Number(raw) || 0) + Number(f.cardPortion||0) }));
                }} style={inputStyle} placeholder="例：500" />
              </div>
              <div>
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 カード ($)</div>
                <input type="number" value={form.cardPortion || ""} onChange={e => {
                  const raw = e.target.value;
                  setForm(f => ({ ...f, cardPortion: raw, amount: Number(f.cashPortion||0) + (Number(raw) || 0) }));
                }} style={inputStyle} placeholder="例：211" />
              </div>
            </div>
            <div style={{ marginTop: 6, fontSize: 12, color: "#B71C1C", fontWeight: 700 }}>
              合計 ${Number(form.cashPortion||0) + Number(form.cardPortion||0)}
            </div>
          </Field>
        ) : (
          <Field label="支払方法" error={errors.includes("paymentType")}>
            <div style={{ display: "flex", gap: 6 }}>
              {payTypes.map(([val, label]) => (
                <button key={val} onClick={() => { set("paymentType", val); setErrors(errors.filter(e => e !== "paymentType")); }}
                  style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#B71C1C":"#DDD"}`, background: form.paymentType===val?"#FFEBEE":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#B71C1C":"#888" }}>
                  {label}
                </button>
              ))}
            </div>
          </Field>
        )}
        <Field label="チップ ($)">
          <input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} placeholder="例：46" />
        </Field>
        {Number(form.tip) > 0 && (
          <Field label="チップ支払方法" error={errors.includes("tipPaymentType")}>
            <div style={{ display: "flex", gap: 6 }}>
              {payTypes.map(([val, label]) => (
                <button key={val} onClick={() => { set("tipPaymentType", val); setErrors(errors.filter(e => e !== "tipPaymentType")); }}
                  style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.tipPaymentType===val?"#E65100":"#DDD"}`, background: form.tipPaymentType===val?"#FFF3E0":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.tipPaymentType===val?"#E65100":"#888" }}>
                  {label}
                </button>
              ))}
            </div>
          </Field>
        )}
        <Field label="メモ">
          <input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="例：最終回使用後に購入" />
        </Field>
        {/* Total preview */}
        {(Number(form.amount) > 0 || Number(form.tip) > 0) && (
          <div style={{ background: "#B71C1C", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontSize: 11, color: "#FFCDD2" }}>
              施術 <strong style={{ color: "#fff" }}>${Number(form.amount||0).toFixed(0)}</strong>
              {Number(form.tip) > 0 && <span>　チップ <strong style={{ color: "#FFCC80" }}>${Number(form.tip||0).toFixed(0)}</strong></span>}
            </div>
            <div>
              <span style={{ color: "#FFCDD2", fontSize: 11 }}>合計　</span>
              <span style={{ color: "#fff", fontSize: 20, fontWeight: 800 }}>${(Number(form.amount||0) + Number(form.tip||0)).toFixed(0)}</span>
            </div>
          </div>
        )}
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#B71C1C", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function StaffPurchaseModal({ sp, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...sp, extraItems: sp.extraItems || [] });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [paymentError, setPaymentError] = useState(false);

  const items = [
    { productName: form.productName, amount: form.amount, paymentType: form.paymentType },
    ...(form.extraItems || []),
  ];
  const updateItem = (idx, patch) => {
    if (idx === 0) {
      if ("productName" in patch) set("productName", patch.productName);
      if ("amount" in patch) set("amount", patch.amount);
      if ("paymentType" in patch) set("paymentType", patch.paymentType);
    } else {
      const extras = form.extraItems || [];
      set("extraItems", extras.map((it, i) => i === idx - 1 ? { ...it, ...patch } : it));
    }
  };
  const addItem = () => set("extraItems", [...(form.extraItems || []), { productName: "", amount: 0, paymentType: "" }]);
  const removeItem = (idx) => {
    if (idx === 0) return;
    const extras = form.extraItems || [];
    set("extraItems", extras.filter((_, i) => i !== idx - 1));
  };
  const total = items.reduce((s, it) => s + Number(it.amount || 0), 0);

  const handleSave = () => {
    const missingPayment = items.some(it => Number(it.amount || 0) > 0 && !it.paymentType);
    if (missingPayment) { setPaymentError(true); return; }
    onSave(form);
  };

  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#37474F" }}>👩‍💼 社販</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="スタッフ名">
          <select value={form.staffName} onChange={e => set("staffName", e.target.value)} style={inputStyle}>
            <option value="">— 選択 —</option>
            {["Mami","Aya","Megumi","Hitomi","Maki","Yuka","Mai","Betsy"].map(t => (
              <option key={t} value={t}>{t}</option>
            ))}
          </select>
        </Field>
        {items.map((item, idx) => {
          const isOtherProduct = !!item.productName && !RETAIL_PRODUCTS.find(p => p.name === item.productName);
          return (
            <div key={idx} style={{ display: "flex", flexDirection: "column", gap: 8, ...(idx > 0 ? { background: "#F5F5F5", borderRadius: 8, padding: 10, border: "1px dashed #CCC" } : {}) }}>
              {idx > 0 && (
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 11, fontWeight: 700, color: "#37474F" }}>商品 {idx + 1}</span>
                  <button onClick={() => removeItem(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11, color: "#AAA" }}>✕ 削除</button>
                </div>
              )}
              <Field label="商品名 / 施術名">
                <select value={isOtherProduct ? "__other__" : (item.productName || "")} onChange={e => {
                  const val = e.target.value;
                  if (val === "__other__") { updateItem(idx, { productName: "" }); return; }
                  const prod = RETAIL_PRODUCTS.find(p => p.name === val);
                  updateItem(idx, { productName: val, amount: prod?.price > 0 ? prod.price : item.amount });
                }} style={inputStyle}>
                  <option value="">— 商品/施術を選択 —</option>
                  {RETAIL_PRODUCTS.map(p => (
                    <option key={p.name} value={p.name}>{p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
                  ))}
                  <option value="__other__">その他（施術名など直接入力）</option>
                </select>
                {isOtherProduct && (
                  <input type="text" value={item.productName || ""} placeholder="例：ProCell" style={{ ...inputStyle, marginTop: 4 }}
                    onChange={e => updateItem(idx, { productName: e.target.value })} />
                )}
              </Field>
              <Field label="金額 ($)"><input type="number" value={item.amount || ""} onChange={e => updateItem(idx, { amount: e.target.value })} style={inputStyle} /></Field>
              <Field label="支払方法" error={paymentError && Number(item.amount || 0) > 0 && !item.paymentType}>
                <div style={{ display: "flex", gap: 6 }}>
                  {[["cash","💵 現金"],["card","💳 カード"]].map(([val, label]) => (
                    <button key={val} onClick={() => { updateItem(idx, { paymentType: val }); setPaymentError(false); }}
                      style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${item.paymentType===val?"#37474F":"#DDD"}`, background: item.paymentType===val?"#ECEFF1":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: item.paymentType===val?"#37474F":"#888" }}>
                      {label}
                    </button>
                  ))}
                </div>
              </Field>
            </div>
          );
        })}
        <button onClick={addItem} style={{ background: "none", border: "none", color: "#37474F", fontSize: 13, textAlign: "left", padding: "4px 0", cursor: "pointer", textDecoration: "underline" }}>
          ＋ 商品を追加（2つ目・3つ目...）
        </button>
        {total > 0 && (
          <div style={{ textAlign: "right", fontSize: 13, fontWeight: 700, color: "#37474F" }}>合計 ${total}</div>
        )}
        <Field label="メモ"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="任意" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#37474F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function RefundModal({ rf, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...rf });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [paymentError, setPaymentError] = useState(false);
  const handleSave = () => {
    if ((Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && !form.paymentType) { setPaymentError(true); return; }
    onSave(form);
  };
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#5D4037" }}>🔙 返金（リファンド）</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 12, background: "#EFEBE9", borderRadius: 8, padding: "8px 10px" }}>
        過去の来店分の返金を、日をまたいだ「本日」の売上からマイナスするための記録です。担当セラピストの給料振り分けはここでは変更されません（必要な場合は元の日をロック解除して直接修正してください）。
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="お客様名">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="例：田中様" />
        </Field>
        <Field label="来店日（任意）">
          <input type="date" value={form.originalDate || ""} onChange={e => set("originalDate", e.target.value)} style={inputStyle} />
        </Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <Field label="施術返金額 ($)">
            <input type="number" value={form.serviceAmount || ""} onChange={e => set("serviceAmount", e.target.value)} style={inputStyle} placeholder="例：150" />
          </Field>
          <Field label="チップ返金額 ($)">
            <input type="number" value={form.tipAmount || ""} onChange={e => set("tipAmount", e.target.value)} style={inputStyle} placeholder="例：0" />
          </Field>
        </div>
        <Field label="返金方法" error={paymentError && !form.paymentType}>
          <div style={{ display: "flex", gap: 6 }}>
            {[["cash","💵 現金"],["card","💳 カード"]].map(([val, label]) => (
              <button key={val} onClick={() => { set("paymentType", val); setPaymentError(false); }}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#5D4037":"#DDD"}`, background: form.paymentType===val?"#EFEBE9":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#5D4037":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        {(Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && (
          <div style={{ background: "#5D4037", borderRadius: 8, padding: "10px 14px", textAlign: "right" }}>
            <span style={{ color: "#D7CCC8", fontSize: 11 }}>返金合計　</span>
            <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>-${Number(form.serviceAmount||0) + Number(form.tipAmount||0)}</span>
          </div>
        )}
        <Field label="メモ"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="返金理由など" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#5D4037", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        {rf.id && <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>}
      </div>
    </Modal>
  );
}

function ForgottenTipModal({ ft, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...ft });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [errors, setErrors] = useState([]);
  const handleSave = () => {
    const errs = [];
    if (!form.therapist) errs.push("therapist");
    if ((Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && !form.paymentType) errs.push("paymentType");
    if (errs.length > 0) { setErrors(errs); return; }
    onSave(form);
  };
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#00695C" }}>🙏 打ち忘れ入力</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 12, background: "#E0F2F1", borderRadius: 8, padding: "8px 10px" }}>
        過去の来店分でレジに打ち忘れた支払い（現金チップなど）を、日をまたいだ「本日」の売上に追加し、担当セラピストの本日分お給料にも計上するための記録です。来店日の予約は編集不要です（顧客タイプの件数には影響しません）。
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="お客様名">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="例：田中様" />
        </Field>
        <Field label="担当セラピスト" error={errors.includes("therapist")}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {THERAPISTS.map(t => (
              <button key={t} onClick={() => { set("therapist", t); setErrors(e => e.filter(x => x !== "therapist")); }}
                style={{ padding: "7px 12px", borderRadius: 8, border: `2px solid ${form.therapist===t?"#00695C":(errors.includes("therapist")?"#C62828":"#DDD")}`, background: form.therapist===t?"#00695C":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.therapist===t?"#fff":"#888" }}>
                {t}
              </button>
            ))}
          </div>
        </Field>
        <Field label="来店日（任意・実際に施術があった日）">
          <input type="date" value={form.originalDate || ""} onChange={e => set("originalDate", e.target.value)} style={inputStyle} />
        </Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <Field label="打ち忘れ施術料 ($)">
            <input type="number" value={form.serviceAmount || ""} onChange={e => set("serviceAmount", e.target.value)} style={inputStyle} placeholder="例：0" />
          </Field>
          <Field label="打ち忘れチップ ($)">
            <input type="number" value={form.tipAmount || ""} onChange={e => set("tipAmount", e.target.value)} style={inputStyle} placeholder="例：20" />
          </Field>
        </div>
        <Field label="支払方法" error={errors.includes("paymentType")}>
          <div style={{ display: "flex", gap: 6 }}>
            {[["cash","💵 現金"],["card","💳 カード"]].map(([val, label]) => (
              <button key={val} onClick={() => { set("paymentType", val); setErrors(e => e.filter(x => x !== "paymentType")); }}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#00695C":"#DDD"}`, background: form.paymentType===val?"#E0F2F1":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#00695C":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        {(Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && (
          <div style={{ background: "#00695C", borderRadius: 8, padding: "10px 14px", textAlign: "right" }}>
            <span style={{ color: "#B2DFDB", fontSize: 11 }}>本日追加合計　</span>
            <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>+${Number(form.serviceAmount||0) + Number(form.tipAmount||0)}</span>
          </div>
        )}
        <Field label="メモ"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="例：16日レジ打ち忘れ" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#00695C", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        {ft.id && <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>}
      </div>
    </Modal>
  );
}

function SectionBox({ title, color, onAdd, disabled, children }) {
  return (
    <div style={{ marginTop: 20, background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <span style={{ fontWeight: 700, color, fontSize: 15 }}>{title}</span>
        <button onClick={onAdd} disabled={disabled} style={{ padding: "6px 14px", borderRadius: 8, background: disabled ? "#CCC" : color, color: "#fff", border: "none", cursor: disabled ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 600 }}>+ 追加</button>
      </div>
      {children}
    </div>
  );
}

function Modal({ onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: 24, width: "100%", maxWidth: 500, maxHeight: "92vh", overflowY: "auto" }}>{children}</div>
    </div>
  );
}

function Field({ label, children, error }) {
  return (
    <div>
      <label style={{ fontSize: 12, fontWeight: 600, color: error ? "#C62828" : "#555", display: "block", marginBottom: 5 }}>
        {label}{error && <span style={{ marginLeft: 6, fontWeight: 800 }}>⚠️ 未入力・未選択</span>}
      </label>
      {children}
    </div>
  );
}

const inputStyle = { width: "100%", padding: "9px 12px", borderRadius: 8, border: "1.5px solid #DDD", fontSize: 14, boxSizing: "border-box", outline: "none" };
const iconBtn = { border: "none", background: "none", cursor: "pointer", fontSize: 16, padding: 2 };

// ============================================================
// SALES REPORT EXCEL EXPORT
// ============================================================
function exportSalesReportXlsx(monthStr) {
  const [y, m] = monthStr.split("-").map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const monthName = monthNames[m - 1];

  // Collect all daily data for the month
  const dailyData = [];
  // New-ticket sales events for the month — both a same-day package purchase and an inline
  // "🎟️チケット新規購入" add-on during a different-type visit count as one ticket sale. When two
  // therapists worked the visit (body + cav), the amount is split 50/50 between them for attribution.
  const ticketSaleEvents = [];
  // Local-customer visits (RL+NL — ticket buyers are mostly local) per staff member, and whether
  // that same visit converted into a ticket sale — lets the owner see "saw N locals, only sold M".
  const localVisitEvents = [];
  for (let d = 1; d <= lastDay; d++) {
    const dateStr = `${monthStr}-${String(d).padStart(2,"0")}`;
    const saved = localStorage.getItem(`spa-sheet-${dateStr}`);
    if (saved) {
      const data = JSON.parse(saved);
      const appts = (data.appointments || []).filter(a => !a.isCavSlot);
      const retails = data.retails || [];
      const refunds = data.refunds || [];
      const forgottenTips = data.forgottenTips || [];
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
        const gcSvc = Math.min(gc, svc);
        const gcTip = Math.min(gc - gcSvc, tip);
        return { gcSvc, gcTip };
      };
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
        + revenueAddons.filter(ad => ad.paymentType === "cash").reduce((s,ad) => s + Number(ad.price||0), 0)
        + depositCash
        + forgottenServiceCash
        - refundServiceCash;
      const cashProduct = retails.filter(r => r.paymentType === "cash").reduce((s,r) => s + Number(r.price||0), 0);
      const cashTip = revenueAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "cash").reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
        + revenueAppts.filter(a => a.tipSplitPayment).reduce((s,a) => s + Number(a.tipCashPortion||0), 0)
        + sameDayTicketAppts.filter(a => a.tipPaymentType === "cash").reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
        + pureTicketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s,a) => s + Number(a.extraTip||0), 0)
        + sameDayTicketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s,a) => s + Number(a.extraTip||0), 0)
        + revenueAddons.filter(ad => ad.tipPaymentType === "cash").reduce((s,ad) => s + Number(ad.tip||0), 0)
        + forgottenTipCash
        - refundTipCash;
      const cardTreatment = revenueAppts.filter(a => !a.svcSplitPayment && a.paymentType === "card").reduce((s,a) => s + Number(a.price||0) - gcAlloc(a).gcSvc, 0)
        + revenueAppts.filter(a => a.svcSplitPayment).reduce((s,a) => s + Number(a.svcCardPortion||0), 0)
        + sameDayTicketAppts.filter(a => !a.packageSplitPayment && a.paymentType === "card").reduce((s,a) => s + Number(a.packagePrice||0) - gcAllocPackage(a).gcSvc, 0)
        + sameDayTicketAppts.filter(a => a.packageSplitPayment).reduce((s,a) => s + Number(a.packageCardPortion||0), 0)
        + pureTicketAppts.filter(a => a.extraPricePaymentType === "card").reduce((s,a) => s + Number(a.extraPrice||0), 0)
        + sameDayTicketAppts.filter(a => a.extraPricePaymentType === "card").reduce((s,a) => s + Number(a.extraPrice||0), 0)
        + revenueAddons.filter(ad => ad.paymentType !== "cash").reduce((s,ad) => s + Number(ad.price||0), 0)
        + depositCard
        + forgottenServiceCard
        - refundServiceCard;
      const cardProduct = retails.filter(r => r.paymentType === "card").reduce((s,r) => s + Number(r.price||0), 0);
      const cardTip = revenueAppts.filter(a => !a.tipSplitPayment && a.tipPaymentType === "card").reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
        + revenueAppts.filter(a => a.tipSplitPayment).reduce((s,a) => s + Number(a.tipCardPortion||0), 0)
        + sameDayTicketAppts.filter(a => a.tipPaymentType === "card").reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
        + pureTicketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s,a) => s + Number(a.extraTip||0), 0)
        + sameDayTicketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s,a) => s + Number(a.extraTip||0), 0)
        + revenueAddons.filter(ad => ad.tipPaymentType !== "cash").reduce((s,ad) => s + Number(ad.tip||0), 0)
        + forgottenTipCard
        - refundTipCard;
      const totalTip = revenueAppts.reduce((s,a) => s + Number(a.tip||0) - gcAlloc(a).gcTip, 0)
        + sameDayTicketAppts.reduce((s,a) => s + Number(a.packageTip ?? a.tip ?? 0) - gcAllocPackage(a).gcTip, 0)
        + pureTicketAppts.reduce((s,a) => s + Number(a.extraTip||0), 0)
        + sameDayTicketAppts.reduce((s,a) => s + Number(a.extraTip||0), 0)
        + revenueAddons.reduce((s,ad) => s + Number(ad.tip||0), 0)
        + forgottenTipCash + forgottenTipCard
        - refundTipCash - refundTipCard;
      const totalSales = cashTreatment + cashProduct + cardTreatment + cardProduct;
      const totalCash = cashTreatment + cashProduct;
      const totalCard = cardTreatment + cardProduct;
      const clients = appts.filter(a => !a.isTicket).length + appts.filter(a => a.isTicket).length;

      dailyData.push({
        date: d,
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
      });
    } else {
      dailyData.push({ date: d, totalSales: 0, clients: "", cashTreatment:"", cashProduct:"", totalCash:"", cashTip:"", cardTreatment:"", cardProduct:"", totalCard:"", cardTip:"", totalTip:"", grandTotal: 0, rl: 0, rt: 0, nl: 0, nt: 0, referrals: REFERRAL_SOURCES.map(() => 0) });
    }
  }

  const wb = XLSX.utils.book_new();
  const wsData = [
    [`Dr.Body,Inc. Sales Report in ${monthName} ${y}`, "", "", "", "", "", "", "", "", "Please fill in the blue cells"],
    [],
    ["Date", "Total Sales", "客数", "Cash", "", "", "", "Card", "", "", "", "Total tip", "Total sales and Tip"],
    ["", "", "", "Treatment", "Product", "Total Cash", "Tip", "Treatment", "Product", "Total Card", "Tip", "", ""],
  ];

  dailyData.forEach(row => {
    wsData.push([
      row.date,
      row.totalSales || 0,
      row.clients || "",
      row.cashTreatment,
      row.cashProduct,
      row.totalCash,
      row.cashTip,
      row.cardTreatment,
      row.cardProduct,
      row.totalCard,
      row.cardTip,
      row.totalTip,
      row.grandTotal || 0,
    ]);
  });

  // Total row
  const dataStart = 5; // row index in wsData where data starts (1-based in Excel = row 5)
  const dataEnd = 5 + lastDay - 1;
  wsData.push([
    "Total",
    { f: `SUM(B${dataStart}:B${dataEnd})` },
    { f: `SUM(C${dataStart}:C${dataEnd})` },
    { f: `SUM(D${dataStart}:D${dataEnd})` },
    { f: `SUM(E${dataStart}:E${dataEnd})` },
    { f: `SUM(F${dataStart}:F${dataEnd})` },
    { f: `SUM(G${dataStart}:G${dataEnd})` },
    { f: `SUM(H${dataStart}:H${dataEnd})` },
    { f: `SUM(I${dataStart}:I${dataEnd})` },
    { f: `SUM(J${dataStart}:J${dataEnd})` },
    { f: `SUM(K${dataStart}:K${dataEnd})` },
    { f: `SUM(L${dataStart}:L${dataEnd})` },
    { f: `SUM(M${dataStart}:M${dataEnd})` },
  ]);

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws["!cols"] = [
    {wch:6},{wch:12},{wch:6},{wch:12},{wch:10},{wch:12},{wch:8},
    {wch:12},{wch:10},{wch:12},{wch:8},{wch:10},{wch:16}
  ];

  // Merge header cells
  ws["!merges"] = [
    { s:{r:2,c:3}, e:{r:2,c:6} },  // Cash
    { s:{r:2,c:7}, e:{r:2,c:10} }, // Card
  ];

  XLSX.utils.book_append_sheet(wb, ws, `${monthName} ${y}`);

  // Customer type (RL/RT/NL/NT) + new-customer referral source monthly tally
  const ctData = [
    [`Customer Type Summary — ${monthName} ${y}`],
    [],
    ["Date", "RL", "RT", "NL", "NT", "Total", ...REFERRAL_SOURCES],
  ];
  dailyData.forEach(row => {
    ctData.push([row.date, row.rl, row.rt, row.nl, row.nt, row.rl + row.rt + row.nl + row.nt, ...row.referrals]);
  });
  const ctDataStart = 4;
  const ctDataEnd = 4 + lastDay - 1;
  const ctCols = ["B","C","D","E","F","G","H","I","J","K"];
  ctData.push([
    "Total",
    ...ctCols.slice(0, 5 + REFERRAL_SOURCES.length).map(col => ({ f: `SUM(${col}${ctDataStart}:${col}${ctDataEnd})` })),
  ]);
  const ctWs = XLSX.utils.aoa_to_sheet(ctData);
  ctWs["!cols"] = [{wch:6},{wch:8},{wch:8},{wch:8},{wch:8},{wch:8}, ...REFERRAL_SOURCES.map(() => ({wch:14}))];
  XLSX.utils.book_append_sheet(wb, ctWs, "Customer Type");

  // Ticket sales summary — overall, new vs repeat, new-customer referral source, and per-staff
  const isNewCustomer = ct => ct === "NL" || ct === "NT";
  const countAmt = (list) => [list.length, Math.round(list.reduce((s, r) => s + r.amount, 0) * 100) / 100];
  const newSales = ticketSaleEvents.filter(r => isNewCustomer(r.customerType));
  const repeatSales = ticketSaleEvents.filter(r => !isNewCustomer(r.customerType));

  const tsData = [
    [`Ticket Sales Summary — ${monthName} ${y}`],
    [],
    ["", "件数 Count", "金額 Amount"],
    ["Total", ...countAmt(ticketSaleEvents)],
    ["New customers (NL+NT)", ...countAmt(newSales)],
    ["Repeat customers (RL+RT)", ...countAmt(repeatSales)],
    [],
    ["New customers by referral source", "件数 Count", "金額 Amount"],
  ];
  REFERRAL_SOURCES.forEach(src => {
    tsData.push([src, ...countAmt(newSales.filter(r => r.referralSource === src))]);
  });
  tsData.push([]);
  tsData.push(["By staff (ranked by New customers $ — who's converting new clients into ticket sales)", "件数 Count", "金額 Amount", "New customers $", "Repeat customers $", "Local visits (RL+NL)", "→ Converted to ticket", "Conversion %"]);
  // A visit worked by two therapists counts as 1 "件" for each, but the $ amount is split by share.
  const shareAmtFor = (t, list) => Math.round(list.reduce((s, r) => s + r.amount * r.therapists.find(th => th.name === t).share, 0) * 100) / 100;
  const byStaffRows = THERAPISTS.map(t => {
    const involved = ticketSaleEvents.filter(r => r.therapists.some(th => th.name === t));
    const newAmt = shareAmtFor(t, involved.filter(r => isNewCustomer(r.customerType)));
    const repeatAmt = shareAmtFor(t, involved.filter(r => !isNewCustomer(r.customerType)));
    const localVisits = localVisitEvents.filter(v => v.therapists.includes(t));
    const localConverted = localVisits.filter(v => v.hasTicketSale).length;
    const conversionPct = localVisits.length > 0 ? `${Math.round(localConverted / localVisits.length * 1000) / 10}%` : "";
    return [t, involved.length, shareAmtFor(t, involved), newAmt, repeatAmt, localVisits.length, localConverted, conversionPct];
  }).sort((a, b) => b[3] - a[3]);
  byStaffRows.forEach(row => tsData.push(row));
  const tsWs = XLSX.utils.aoa_to_sheet(tsData);
  tsWs["!cols"] = [{wch:26},{wch:12},{wch:14},{wch:14},{wch:16},{wch:14},{wch:16},{wch:12}];
  XLSX.utils.book_append_sheet(wb, tsWs, "Ticket Sales");

  XLSX.writeFile(wb, `SalesReport_${monthStr}.xlsx`);
}

// ============================================================
// PAYROLL TAB
// ============================================================
function PayrollTab() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");

  // Default: current period
  const defaultPeriod = today.getDate() <= 15 ? "first" : "second";
  const [period, setPeriod] = useState(defaultPeriod);
  const [month, setMonth] = useState(`${yyyy}-${mm}`);
  const [payrollData, setPayrollData] = useState(null);
  const [loading, setLoading] = useState(false);

  const calcPeriodDates = (monthStr, p) => {
    const [y, m] = monthStr.split("-").map(Number);
    if (p === "first") {
      return { start: `${monthStr}-01`, end: `${monthStr}-15` };
    } else {
      const lastDay = new Date(y, m, 0).getDate();
      return { start: `${monthStr}-16`, end: `${monthStr}-${String(lastDay).padStart(2, "0")}` };
    }
  };

  const loadPayroll = () => {
    setLoading(true);
    const { start, end } = calcPeriodDates(month, period);

    // Collect all dates in range from localStorage
    const startD = new Date(start);
    const endD = new Date(end);
    const allAppts = [];
    const allRetails = [];
    const allForgottenTips = [];

    for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
      const dateStr = d.toISOString().slice(0, 10);
      const saved = localStorage.getItem(`spa-sheet-${dateStr}`);
      if (saved) {
        const data = JSON.parse(saved);
        (data.appointments || []).forEach(a => {
          allAppts.push({ ...a, date: dateStr });
        });
        (data.retails || []).forEach(r => {
          allRetails.push({ ...r, date: dateStr });
        });
        (data.forgottenTips || []).forEach(ft => {
          allForgottenTips.push({ ...ft, date: dateStr });
        });
      }
    }

    // Aggregate per therapist
    const byTherapist = {};
    THERAPISTS.forEach(t => {
      byTherapist[t] = { therapist: t, rows: [], totalService: 0, totalTip: 0, totalServiceCard: 0, totalTipCard: 0, totalRetail: 0, totalRetailCard: 0 };
    });

    // Index by id for cav slot → parent lookup
    const apptById = {};
    allAppts.forEach(a => { if (a.id) apptById[a.id] = a; });

    // Deposit/gift-card usage changes where the money actually came from, so it's called out
    // in 備考 for every therapist — otherwise a payroll total can look "off" from the day's
    // revenue with no way to tell why.
    const moneyNote = (a) => [
      (Number(a.depositApplied || 0) + Number(a.packageDepositAmount || 0)) > 0
        ? `💰デポジット$${Number(a.depositApplied || 0) + Number(a.packageDepositAmount || 0)}分あり` : "",
      Number(a.giftCardUsed || 0) > 0 ? `🎁ギフトカード$${Number(a.giftCardUsed)}分使用` : "",
    ].filter(Boolean).join("　");

    allAppts.filter(a => !a.isCavSlot).forEach(a => {
      const t = a.therapist;
      if (!t || !byTherapist[t]) return;
      const depositAdd = Number(a.depositApplied || 0);
      const bodyReceived = Number(a.price || 0);
      let svc;
      if (depositAdd > 0 && isWeightLossService(a.serviceName) && a.cavTherapist) {
        // Weight Loss: machine payroll is a fixed $116 no matter what — any deposit on top
        // of what's collected today goes entirely to the body therapist, not split proportionally.
        svc = bodyReceived + depositAdd;
      } else if (depositAdd > 0 && a.cavTherapist) {
        // Split the deposit-inclusive total by treatment MINUTES (not the price ratio), so the
        // machine therapist's share reflects actual time spent, not how price/cavPrice were entered.
        const cavSlot = allAppts.find(s => s.isCavSlot && s.parentId === a.id);
        const cavReceived = cavSlot ? Number(cavSlot.price || 0) : 0;
        const totalReceived = bodyReceived + cavReceived;
        const bodyMins = Number(a.duration || 0);
        const cavMins = (cavSlot && Number(cavSlot.duration || 0)) || 15;
        const allMins = bodyMins + cavMins;
        svc = allMins > 0
          ? Math.round((totalReceived + depositAdd) * bodyMins / allMins * 100) / 100
          : totalReceived + depositAdd;
      } else if (depositAdd > 0) {
        svc = bodyReceived + depositAdd;
      } else {
        svc = bodyReceived;
      }
      // Extra service fee/tip (💝 given on top of a ticket redemption) always count toward the
      // therapist's payroll too.
      svc += Number(a.extraPrice || 0);
      const tip = Number(a.tip || 0) + Number(a.extraTip || 0);
      const isCard = a.paymentType === "card";
      const isTipCard = a.tipPaymentType === "card";
      byTherapist[t].rows.push({
        date: a.date,
        client: a.clientName,
        isTicket: a.isTicket,
        isGiftCard: a.isGiftCard,
        isPromo: a.isPromo,
        ticketInfo: a.isTicket ? `${a.ticketMenu} ${a.ticketCurrent}/${a.ticketTotal}` : "",
        partner: a.cavTherapist || "",
        duration: a.duration,
        service: svc,
        tip,
        paymentType: a.isGiftCard || a.isPromo ? "gc" : a.paymentType,
        tipPaymentType: a.isGiftCard || a.isPromo ? "gc" : a.tipPaymentType,
        notes: [a.notes || "", moneyNote(a), a.isPromo ? "📸PR無料" : ""].filter(Boolean).join("　"),
      });
      byTherapist[t].totalService += svc;
      byTherapist[t].totalTip += tip;
      if (isCard) byTherapist[t].totalServiceCard += svc;
      if (isTipCard) byTherapist[t].totalTipCard += tip;
    });

    // Retail from appointments → split among up to 3 sellers (defaults to just the body
    // therapist getting the full amount when no split was entered).
    allAppts.filter(a => !a.isCavSlot && a.purchaseTags?.includes("retail")).forEach(a => {
      getRetailItems(a).forEach(item => {
        const isCard = item.paymentType === "card";
        const sellers = (item.sellers && item.sellers.length > 0)
          ? item.sellers
          : [{ therapist: a.therapist, amount: Number(item.amount) }];
        sellers.filter(sel => sel.therapist && Number(sel.amount || 0) > 0).forEach(sel => {
          const t = sel.therapist;
          if (!t || !byTherapist[t]) return;
          const retail = Number(sel.amount);
          byTherapist[t].rows.push({
            date: a.date,
            client: a.clientName,
            isTicket: false,
            ticketInfo: "",
            duration: 0,
            service: 0,
            tip: 0,
            paymentType: item.paymentType,
            tipPaymentType: "",
            retail,
            retailProduct: item.productName || "",
            notes: `🛍️ 物販${item.productName ? ` (${item.productName})` : ""}`,
          });
          byTherapist[t].totalRetail += retail;
          if (isCard) byTherapist[t].totalRetailCard += retail;
        });
      });
    });

    // Retail — standalone (phone/walk-in, not tied to any appointment) → attributed via sellers,
    // each getting their own dollar share (commission rates differ per staff member, e.g. Maki's is
    // lower than everyone else's, so an even split isn't assumed — each amount is entered manually).
    allRetails.forEach(r => {
      const sellers = r.sellers || (r.soldBy ? [{ therapist: r.soldBy, amount: r.price }] : []);
      const isCard = r.paymentType === "card";
      sellers.filter(sel => sel.therapist && Number(sel.amount || 0) > 0).forEach(sel => {
        const t = sel.therapist;
        if (!byTherapist[t]) return;
        const retail = Number(sel.amount);
        byTherapist[t].rows.push({
          date: r.date,
          client: "",
          isTicket: false,
          ticketInfo: "",
          duration: 0,
          service: 0,
          tip: 0,
          paymentType: r.paymentType,
          tipPaymentType: "",
          retail,
          retailProduct: r.item || "",
          notes: `🛍️ 物販(電話・ウォークイン)${r.item ? ` (${r.item})` : ""}`,
        });
        byTherapist[t].totalRetail += retail;
        if (isCard) byTherapist[t].totalRetailCard += retail;
      });
    });

    // Forgotten tips/payments — a past visit's payment staff forgot to ring in that day, entered
    // later and credited to the therapist on the day it's *entered* (not the original visit date),
    // matching how the daily sheet folds it into that day's revenue too.
    allForgottenTips.forEach(ft => {
      const t = ft.therapist;
      if (!t || !byTherapist[t]) return;
      const svc = Number(ft.serviceAmount || 0);
      const tip = Number(ft.tipAmount || 0);
      const isCard = ft.paymentType === "card";
      byTherapist[t].rows.push({
        date: ft.date,
        client: ft.clientName,
        isTicket: false,
        ticketInfo: "",
        duration: 0,
        service: svc,
        tip,
        paymentType: ft.paymentType,
        tipPaymentType: ft.paymentType,
        notes: `🙏 打ち忘れ${ft.originalDate ? `(${ft.originalDate}来店分)` : ""}${ft.notes ? ` ${ft.notes}` : ""}`,
      });
      byTherapist[t].totalService += svc;
      byTherapist[t].totalTip += tip;
      if (isCard) byTherapist[t].totalServiceCard += svc;
      if (isCard) byTherapist[t].totalTipCard += tip;
    });

    // Add-on services → attributed therapist (or fall back to main therapist)
    allAppts.filter(a => !a.isCavSlot && (a.addons || []).length > 0).forEach(a => {
      (a.addons || []).forEach(addon => {
        const svc = Number(addon.price || 0);
        const tip = Number(addon.tip || 0);
        const addonTherapist = addon.therapist || "";
        const t = addonTherapist || a.therapist;
        if (!t || !byTherapist[t]) return;
        // Skip only if no explicit therapist AND no price (pure label with no data)
        if (!addonTherapist && svc + tip === 0) return;
        const isCard = addon.paymentType === "card";
        const isTipCard = addon.tipPaymentType === "card";
        byTherapist[t].rows.push({
          date: a.date,
          client: a.clientName,
          isTicket: false,
          isGiftCard: a.isGiftCard,
          ticketInfo: "",
          duration: 0,
          service: svc,
          tip,
          paymentType: a.isGiftCard ? "gc" : (addon.paymentType || "cash"),
          tipPaymentType: a.isGiftCard ? "gc" : (addon.tipPaymentType || "cash"),
          notes: `➕ ${addon.serviceName || addon.name || "オプション"}${addon.ticketCurrent ? ` ${addon.ticketCurrent}/${a.ticketTotal||3}` : ""}`,
        });
        byTherapist[t].totalService += svc;
        byTherapist[t].totalTip += tip;
        if (isCard && !a.isGiftCard) byTherapist[t].totalServiceCard += svc;
        if (isTipCard && !a.isGiftCard) byTherapist[t].totalTipCard += tip;
      });
    });

    // Cav slots count toward cav therapist
    allAppts.filter(a => a.isCavSlot).forEach(a => {
      const t = a.therapist;
      if (!t || !byTherapist[t]) return;
      const cavReceived = Number(a.price || 0);
      // Look up parent to check for depositApplied
      const parent = a.parentId ? apptById[a.parentId] : null;
      const depositAdd = Number(parent?.depositApplied || 0);
      let svc;
      if (depositAdd > 0 && parent && isWeightLossService(parent.serviceName) && parent.cavTherapist) {
        // Weight Loss: machine payroll always stays at the fixed $116 — the deposit goes to the
        // body therapist instead (see the matching branch above for the non-cav-slot rows).
        svc = cavReceived;
      } else if (depositAdd > 0 && parent) {
        // Split by treatment MINUTES (not price ratio) — machine gets the remainder after the
        // body therapist's minute-share, so the two always sum exactly to the deposit-inclusive total.
        const bodyReceived = Number(parent.price || 0);
        const totalReceived = bodyReceived + cavReceived;
        const bodyMins = Number(parent.duration || 0);
        const cavMins = Number(a.duration || 0) || 15;
        const allMins = bodyMins + cavMins;
        if (allMins > 0) {
          const bodyShare = Math.round((totalReceived + depositAdd) * bodyMins / allMins * 100) / 100;
          svc = Math.round(((totalReceived + depositAdd) - bodyShare) * 100) / 100;
        } else {
          svc = cavReceived;
        }
      } else {
        svc = cavReceived;
      }
      const tip = Number(a.tip || 0);
      byTherapist[t].rows.push({
        date: a.date,
        client: a.clientName,
        isTicket: a.isTicket,
        ticketInfo: a.isTicket ? `${a.ticketMenu} ${a.ticketCurrent}/${a.ticketTotal}` : "",
        partner: parent?.therapist || "",
        duration: 15,
        service: svc,
        tip,
        paymentType: a.paymentType,
        tipPaymentType: a.tipPaymentType,
        notes: ["⚡ 機械", parent ? moneyNote(parent) : ""].filter(Boolean).join("　"),
      });
      byTherapist[t].totalService += svc;
      byTherapist[t].totalTip += tip;
    });

    setPayrollData({ byTherapist, start, end });
    setLoading(false);
  };

  const downloadCSV = (therapist) => {
    const data = payrollData?.byTherapist[therapist];
    if (!data) return;
    const remarksFor = (r) => [
      r.isTicket && r.ticketInfo ? `🎟️ ${r.ticketInfo}` : "",
      r.partner ? `with ${r.partner}` : "",
      r.notes || "",
    ].filter(Boolean).join("　");
    const rows = [
      ["日付", "お客様", ...(therapist === "Maki" ? ["分数"] : []), "施術", "チップ", "合計", "物販", "備考"],
      ...data.rows.map(r => [
        r.date, r.client, ...(therapist === "Maki" ? [r.duration || ""] : []), r.service || "",
        r.tip || "", r2(r.service + r.tip), r.retail || "", remarksFor(r)
      ]),
      [],
      ["", "合計", ...(therapist === "Maki" ? [""] : []), data.totalService, data.totalTip, r2(data.totalService + data.totalTip), data.totalRetail, ""],
    ];
    const csv = rows.map(r => r.join(",")).join("\n");
    const bom = "\uFEFF";
    const blob = new Blob([bom + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${therapist}_${payrollData.start}_${payrollData.end}.csv`;
    a.click();
  };

  const downloadAllCSV = () => {
    if (!payrollData) return;
    const remarksForAll = (r) => [
      r.isTicket && r.ticketInfo ? `🎟️ ${r.ticketInfo}` : "",
      r.partner ? `with ${r.partner}` : "",
      r.notes || "",
    ].filter(Boolean).join("　");
    const allRows = [["担当者", "日付", "お客様", "分数（まきのみ）", "施術", "チップ", "合計", "物販", "備考"]];
    THERAPISTS.forEach(t => {
      const data = payrollData.byTherapist[t];
      if (!data || data.rows.length === 0) return;
      data.rows.forEach(r => {
        allRows.push([
          t, r.date, r.client, t === "Maki" ? (r.duration || "") : "",
          r.service || "", r.tip || "", r2(r.service + r.tip), r.retail || "", remarksForAll(r)
        ]);
      });
      allRows.push([t, "", "小計", "", data.totalService, data.totalTip, r2(data.totalService + data.totalTip), data.totalRetail, ""]);
      allRows.push([]);
    });
    const csv = allRows.map(r => r.join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `全員給料_${payrollData.start}_${payrollData.end}.csv`;
    a.click();
  };

  const { start, end } = calcPeriodDates(month, period);

  return (
    <div style={{ padding: 16 }}>
      {/* Sales Report Excel Export */}
      <div style={{ background: "linear-gradient(135deg,#E8F5E9,#C8E6C9)", borderRadius: 12, padding: 16, marginBottom: 16, border: "2px solid #4CAF50" }}>
        <div style={{ fontWeight: 700, color: "#2E7D32", marginBottom: 8, fontSize: 15 }}>📊 Sales Report Excel出力</div>
        <div style={{ fontSize: 12, color: "#555", marginBottom: 12 }}>月を選択してExcelをダウンロード（Dr.Body形式）</div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <input type="month" value={month} onChange={e => setMonth(e.target.value)}
            style={{ padding: "8px 12px", borderRadius: 8, border: "1.5px solid #4CAF50", fontSize: 14 }} />
          <button onClick={() => exportSalesReportXlsx(month)}
            style={{ padding: "10px 20px", borderRadius: 10, border: "none", background: "#2E7D32", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>
            ⬇️ Excel ダウンロード
          </button>
        </div>
      </div>
      <div style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
        <div style={{ fontWeight: 700, color: "#0D4F4F", marginBottom: 12, fontSize: 15 }}>💴 給料集計期間</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
          <Field label="月">
            <input type="month" value={month} onChange={e => setMonth(e.target.value)}
              style={{ padding: "8px 12px", borderRadius: 8, border: "1.5px solid #DDD", fontSize: 14 }} />
          </Field>
          <Field label="期間">
            <div style={{ display: "flex", gap: 8 }}>
              {[["first", `1〜15日`], ["second", `16〜末日`]].map(([val, label]) => (
                <button key={val} onClick={() => setPeriod(val)}
                  style={{ padding: "8px 16px", borderRadius: 8, border: `2px solid ${period === val ? "#0D4F4F" : "#DDD"}`, background: period === val ? "#0D4F4F" : "#fff", color: period === val ? "#fff" : "#888", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
                  {label}
                </button>
              ))}
            </div>
          </Field>
        </div>
        <div style={{ fontSize: 12, color: "#888", marginBottom: 12 }}>集計期間：{start} 〜 {end}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={loadPayroll} disabled={loading}
            style={{ padding: "10px 24px", borderRadius: 10, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>
            {loading ? "読み込み中..." : "📊 集計する"}
          </button>
          {payrollData && (
            <button onClick={downloadAllCSV}
              style={{ padding: "10px 20px", borderRadius: 10, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>
              ⬇️ 全員CSV
            </button>
          )}
        </div>
      </div>

      {/* Results */}
      {payrollData && THERAPISTS.map(t => {
        const data = payrollData.byTherapist[t];
        if (!data || data.rows.length === 0) return null;
        return (
          <div key={t} style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 14, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
              <div>
                <span style={{ fontWeight: 800, fontSize: 16, color: "#0D4F4F" }}>{t}</span>
                <span style={{ fontSize: 12, color: "#888", marginLeft: 10 }}>{data.rows.length}件</span>
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700, color: "#C62828", fontSize: 15 }}>施術 {formatCurrency(data.totalService)}</span>
                <span style={{ fontWeight: 700, color: "#E65100", fontSize: 15 }}>チップ {formatCurrency(data.totalTip)}</span>
                {data.totalRetail > 0 && <span style={{ fontWeight: 700, color: "#6A1B9A", fontSize: 15 }}>物販 {formatCurrency(data.totalRetail)}</span>}
                <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 15 }}>合計 {formatCurrency(data.totalService + data.totalTip + data.totalRetail)}</span>
                <button onClick={() => downloadCSV(t)}
                  style={{ padding: "5px 12px", borderRadius: 8, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                  ⬇️ CSV
                </button>
              </div>
            </div>

            {/* Rows */}
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ borderBottom: "2px solid #EEE", background: "#F9F9F9" }}>
                    {["日付", "お客様", ...(t === "Maki" ? ["分数"] : []), "施術", "チップ", "合計", "物販", "備考"].map(h => (
                      <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.rows.sort((a, b) => a.date.localeCompare(b.date)).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F5F5F5", background: r.isGiftCard ? "#FFFDE7" : r.isPromo ? "#E3F2FD" : r.isTicket ? "#F0F7FF" : "white" }}>
                      <td style={{ padding: "6px 8px", whiteSpace: "nowrap", color: "#888" }}>{r.date.slice(5)}</td>
                      <td style={{ padding: "6px 8px", fontWeight: 600 }}>{r.client}</td>
                      {t === "Maki" && <td style={{ padding: "6px 8px", textAlign: "center" }}>{r.duration}分</td>}
                      <td style={{ padding: "6px 8px", color: "#C62828", fontWeight: 700 }}>{formatCurrency(r.service)}</td>
                      <td style={{ padding: "6px 8px", color: "#E65100", fontWeight: 700 }}>{r.tip > 0 ? formatCurrency(r.tip) : "—"}</td>
                      <td style={{ padding: "6px 8px", color: "#0D4F4F", fontWeight: 800 }}>{formatCurrency(r2(r.service + r.tip))}</td>
                      <td style={{ padding: "6px 8px", color: "#6A1B9A", fontWeight: r.retail ? 700 : 400 }}>{r.retail ? formatCurrency(r.retail) : "—"}</td>
                      <td style={{ padding: "6px 8px", fontSize: 11, color: "#888", whiteSpace: "nowrap" }}>
                        {r.isTicket && r.ticketInfo && <span>🎟️ {r.ticketInfo}　</span>}
                        {r.partner && <span>with {r.partner}　</span>}
                        {r.notes && <span>{r.notes}</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: "2px solid #EEE", background: "#FFF8E1" }}>
                    <td colSpan={t === "Maki" ? 3 : 2} style={{ padding: "8px", fontWeight: 700, color: "#555" }}>合計</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#C62828" }}>{formatCurrency(data.totalService)}</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#E65100" }}>{formatCurrency(data.totalTip)}</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#0D4F4F" }}>{formatCurrency(r2(data.totalService + data.totalTip))}</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#6A1B9A" }}>{data.totalRetail > 0 ? formatCurrency(data.totalRetail) : "—"}</td>
                    <td />
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        );
      })}

      {payrollData && THERAPISTS.every(t => !payrollData.byTherapist[t]?.rows?.length) && (
        <div style={{ textAlign: "center", color: "#AAA", padding: 40, fontSize: 14 }}>
          この期間にデータがありません
        </div>
      )}
    </div>
  );
}
