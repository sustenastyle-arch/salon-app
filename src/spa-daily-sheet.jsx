import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { REFERRAL_SOURCES, getRetailItems, computeDayTotals } from "./lib/reportTotals.js";

const THERAPISTS = ["Mami", "Aya", "Megumi", "Hitomi", "Maki", "Yuka", "Mai", "Betsy"];
const CUSTOMER_TYPES = ["RL", "RT", "NL", "NT"];
// Stored value stays "紹介" (already-saved appointments are matched against it by exact equality
// in reporting) — this only translates it for display, same pattern as CUSTOMER_TYPE_LABELS below.
const REFERRAL_LABELS = { "紹介": "Referral" };
const HOURS = Array.from({ length: 11 }, (_, i) => i + 9);
// Weight Loss (regular, non-ticket): machine/cavitation allocation is a fixed $116 regardless
// of duration or total price — the rest goes to the body therapist.
const REGULAR_WL_CAV_PRICE = 116;
const isWeightLossService = (name) => (name || "").toLowerCase().includes("weight loss");
// Appointments synced from Square before the item-name cleanup in the booking sync (see
// vite.config.js/api/square-bookings.js) may still have serviceName saved with extra junk
// the catalog item's Name field was never meant to carry: a Japanese translation in
// parentheses, a bare Japanese translation/variation tacked on elsewhere (e.g. "-ハイフ全顔"),
// or even a call-to-action + phone number typed straight into the name (e.g. "HIFU(Full
// Face) -ハイフ全顔 Please Call us 808-922-5115 or Text us 808-971-1267 通常"). Strip all three
// for display only (the stored value is left alone; other code like isWeightLossService above
// just does a substring check on the English part, so it's unaffected either way).
const stripJpAnnotation = (name) => {
  if (!name) return "";
  let s = name.replace(/\s*\([^)]*[぀-ヿ一-鿿][^)]*\)/g, "");
  const ctaIdx = s.search(/(please\s+)?(call|text)\s+us|\d{3}[-.\s]\d{3}[-.\s]\d{4}/i);
  if (ctaIdx !== -1) s = s.slice(0, ctaIdx);
  return s.split(/\s+/).filter(t => t && !/[぀-ヿ一-鿿]/.test(t)).join(" ").trim();
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
  { id: "newTicket", label: "🎟️ New Ticket Purchase", color: "#B71C1C", bg: "#FFEBEE" },
  { id: "giftCard",  label: "🎁 Gift Card Purchase",   color: "#00796B", bg: "#E0F2F1" },
  { id: "retail",    label: "🛍️ Retail Purchase",      color: "#6A1B9A", bg: "#F3E5F5" },
];

const ADDON_PRESETS = [
  "Previous unused cav time",
  "Cav add-on +10min",
  "Cav add-on +20min",
  "Cav add-on +30min",
  "Cav add-on +40min",
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
// Stored value stays "電気バリブラシ" (already-saved retail sales are matched against it by exact
// name equality) — this only translates it for display, same pattern as REFERRAL_LABELS above.
const RETAIL_PRODUCT_LABELS = { "電気バリブラシ": "Electric Facial Brush" };

const formatCurrency = (n) => `$${Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 0 })}`;
const r2 = (n) => Math.round(n * 100) / 100;
const formatTime = (hour) => { const h = hour % 12 || 12; return `${h}:00 ${hour < 12 ? "AM" : "PM"}`; };

// All day data now lives in a shared cloud store (see api/_lib/dayDataStore.js) instead of
// localStorage, so the same data shows up on every computer instead of being stuck in
// whichever browser it was typed into. The site password (already shipped to this bundle
// for PasswordGate) doubles as this API's shared secret — see checkAuth in dayDataStore.js.
const API_PASSWORD = import.meta.env.VITE_APP_PASSWORD || "";
async function apiFetch(path, options = {}) {
  const res = await fetch(path, {
    ...options,
    headers: { "Content-Type": "application/json", "X-App-Password": API_PASSWORD, ...(options.headers || {}) },
  });
  if (!res.ok) throw new Error(`API error ${res.status}`);
  return res.json();
}

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
        {pt === "cash" ? "💵 Cash" : "💳 Card"}
      </button>
    ))}
  </div>
);

const PayBadge = ({ type }) => (
  <span style={{ fontSize: 11, background: type === "cash" ? "#E8F5E9" : "#E3F2FD", color: type === "cash" ? "#2E7D32" : "#1565C0", padding: "2px 8px", borderRadius: 10, whiteSpace: "nowrap" }}>
    {type === "cash" ? "Cash" : "Card"}
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
  const [clientDepositHistory, setClientDepositHistory] = useState([]);
  const [gcSyncLoading, setGcSyncLoading] = useState(false);
  const [depositSyncLoading, setDepositSyncLoading] = useState(false);
  const [reconcileLoading, setReconcileLoading] = useState(false);
  const [reconcileResult, setReconcileResult] = useState(null);
  const [dayLoading, setDayLoading] = useState(true);
  const [autoBackupDates, setAutoBackupDates] = useState(null);
  const [autoBackupListLoading, setAutoBackupListLoading] = useState(false);

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
    let cancelled = false;
    setReconcileResult(null);
    setDayLoading(true);
    (async () => {
      try {
        const { data: d } = await apiFetch(`/api/day-data?date=${date}`);
        if (cancelled) return;
        if (d) {
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
      } catch (e) {
        if (!cancelled) {
          console.error("Day load error:", e);
          showToast("Failed to load data. Please check your internet connection and reload", "error");
        }
      } finally {
        if (!cancelled) setDayLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [date]);

  // Scan every saved day for deposits/gift-card prepayments with appointmentDate matching
  // the current date — covers both a partial deposit and a fully prepaid ("ギフト" as advance
  // payment) visit, since either can be booked for a future date that later changes. Runs
  // server-side now (api/deposits.js) instead of enumerating every localStorage key.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { deposits: found } = await apiFetch(`/api/deposits?mode=date&date=${date}`);
        if (!cancelled) setDepositsForDate(found);
      } catch (e) {
        if (!cancelled) console.error("Deposit scan error:", e);
      }
    })();
    return () => { cancelled = true; };
  }, [date, deposits]);

  // A client's deposit/gift-card history, shown when their appointment modal opens — was an
  // inline localStorage-enumerating IIFE before; now a network call, so it has to live in
  // state+effect instead of a synchronous JSX expression.
  useEffect(() => {
    const name = (editingAppt?.clientName || "").toLowerCase().trim();
    if (!name) { setClientDepositHistory([]); return; }
    let cancelled = false;
    (async () => {
      try {
        const { deposits: found } = await apiFetch(`/api/deposits?mode=client&name=${encodeURIComponent(name)}`);
        if (!cancelled) setClientDepositHistory(found);
      } catch (e) {
        if (!cancelled) console.error("Client deposit history error:", e);
      }
    })();
    return () => { cancelled = true; };
  }, [editingAppt?.clientName]);

  const save = useCallback(async (appts, rets, deps, tps, sps, refs, fts, ws) => {
    try {
      await apiFetch("/api/day-data", {
        method: "POST",
        body: JSON.stringify({
          date,
          data: { appointments: appts, retails: rets, deposits: deps, ticketPurchases: tps || [], staffPurchases: sps || [], refunds: refs || [], forgottenTips: fts || [], locked, workingStaff: ws || workingStaff },
        }),
      });
      return true;
    } catch (e) {
      console.error("Save error:", e);
      showToast("Failed to save. Please check your internet connection and try again", "error");
      return false;
    }
  }, [date, locked, workingStaff]);

  // Toggling the lock writes immediately (doesn't wait for another edit) using current state.
  // Only flips the UI state after the cloud write actually succeeds, so "locked" shown on
  // screen can't diverge from what the server recorded.
  const setDayLocked = async (newLocked) => {
    try {
      await apiFetch("/api/day-data", {
        method: "POST",
        body: JSON.stringify({
          date,
          data: { appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips, locked: newLocked, workingStaff },
        }),
      });
      setLocked(newLocked);
      return true;
    } catch (e) {
      console.error("Lock toggle error:", e);
      showToast("Failed to save. Please check your internet connection and try again", "error");
      return false;
    }
  };

  const MANAGER_PIN = import.meta.env.VITE_MANAGER_PIN || "0000";
  const guardLocked = () => {
    if (locked) { showToast("🔒 This day is finalized. Unlock it before editing", "error"); return true; }
    return false;
  };
  const handleLockToggle = async () => {
    if (!locked) {
      if (window.confirm(`Finalize and lock ${date}?\nOnce locked, no edits or deletions can be made until it's unlocked.`)) {
        if (await setDayLocked(true)) showToast("🔒 This day has been finalized");
      }
      return;
    }
    const pin = window.prompt("Enter the PIN to unlock");
    if (pin === null) return;
    if (pin === MANAGER_PIN) {
      if (await setDayLocked(false)) showToast("🔓 Unlocked");
    } else {
      showToast("Incorrect PIN", "error");
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
        setSquareStatus("No bookings found in Square for this day.");
        showToast("No bookings", "info");
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
          showToast("🔒 Cancelled — this day was locked while fetching", "error");
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
        await save(merged, cur.retails, cur.deposits, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips, nextWorkingStaff);

        showToast(`✅ Fetched ${newAppts.length} booking(s)`);
      }
    } catch (e) {
      console.error("Square sync error:", e);
      setSquareStatus("Square connection error. Please enter manually.");
      showToast("Connection error", "error");
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
        showToast(data.error || "Fetch error", "error");
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
        showToast("🔒 Cancelled — this day was locked while fetching", "error");
        return;
      }
      const existingIds = new Set(cur.deposits.map(d => d.id));
      const newDeposits = activities
        .filter(a => !existingIds.has(a.id))
        .map(a => ({
          id: a.id,
          type: "giftcard",
          amount: a.amount,
          clientName: "Gift Card Purchase (customer name unknown)",
          paymentType: a.paymentType === "cash" ? "cash" : "card",
          tip: 0,
          tipPaymentType: "cash",
          appointmentDate: "",
          appointmentTime: "",
          notes: `Auto-fetched from Square${a.createdAt ? ` (${a.createdAt})` : ""}`,
        }));
      if (newDeposits.length === 0) {
        showToast(activities.length === 0 ? "No gift card purchases this day" : "Already up to date", "info");
        return;
      }
      const next = [...cur.deposits, ...newDeposits];
      setDeposits(next);
      await save(cur.appointments, cur.retails, next, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips);
      showToast(`✅ Added ${newDeposits.length} gift card purchase(s)`);
    } catch (e) {
      console.error("Gift card sync error:", e);
      showToast("Connection error", "error");
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
        showToast(data.error || "Fetch error", "error");
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
        showToast("🔒 Cancelled — this day was locked while fetching", "error");
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
          notes: `Auto-fetched from Square${d.createdAt ? ` (${d.createdAt})` : ""}`,
        }));
      if (newDeposits.length === 0) {
        showToast(found.length === 0 ? "No deposit payments this day" : "Already up to date", "info");
        return;
      }
      const next = [...cur.deposits, ...newDeposits];
      setDeposits(next);
      await save(cur.appointments, cur.retails, next, cur.ticketPurchases, cur.staffPurchases, cur.refunds, cur.forgottenTips);
      showToast(`✅ Added ${newDeposits.length} deposit(s)`);
    } catch (e) {
      console.error("Deposit sync error:", e);
      showToast("Connection error", "error");
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
        showToast(data.error || "Fetch error", "error");
        return;
      }
      setReconcileResult({ ...data, checkedAt: Date.now() });
    } catch (e) {
      console.error("Square reconcile error:", e);
      showToast("Connection error", "error");
    } finally {
      setReconcileLoading(false);
    }
  };

  // Downloads every saved day (from the cloud store, not this browser) as one JSON file — the
  // export counterpart to handleImportBackup below, for an offline copy or a manual transfer.
  const handleExportBackup = async () => {
    try {
      const { days } = await apiFetch("/api/export-all");
      const blob = new Blob([JSON.stringify(days)], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `spa-sheet-backup-${new Date().toISOString().slice(0, 10)}.json`;
      a.click();
      URL.revokeObjectURL(a.href);
      showToast(`✅ Backed up ${Object.keys(days).length} day(s) of data`);
    } catch (e) {
      console.error("Export error:", e);
      showToast("Backup failed. Please check your internet connection and try again", "error");
    }
  };

  // Restores a previously exported JSON file into the shared cloud store (bulk upsert via
  // /api/import-all — accepts both this format and the older per-browser localStorage export
  // format, since setDays() on the server normalizes either raw JSON strings or parsed objects).
  const handleImportBackup = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const data = JSON.parse(reader.result);
        const { count } = await apiFetch("/api/import-all", { method: "POST", body: JSON.stringify(data) });
        showToast(`✅ Restored ${count} day(s) of data`);
        window.location.reload();
      } catch (err) {
        showToast("Failed to load: " + err.message, "error");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  // A snapshot of every saved day is taken automatically once a day (see api/auto-backup.js,
  // triggered by the Vercel Cron in vercel.json) — this lets a day get recovered even if nobody
  // remembered to click the manual 📤 Backup Data button.
  const openAutoBackupList = async () => {
    setAutoBackupListLoading(true);
    try {
      const { dates } = await apiFetch("/api/auto-backup-list");
      setAutoBackupDates(dates);
    } catch (e) {
      console.error("Auto-backup list error:", e);
      showToast("Failed to load the backup list. Please check your internet connection and try again", "error");
    } finally {
      setAutoBackupListLoading(false);
    }
  };

  const restoreFromAutoBackup = async (date) => {
    if (!window.confirm(`This will overwrite the current cloud data with the automatic backup from ${date}. Continue?`)) return;
    try {
      const { count } = await apiFetch("/api/auto-backup-restore", { method: "POST", body: JSON.stringify({ date }) });
      showToast(`✅ Restored ${count} day(s) of data from the ${date} backup`);
      setAutoBackupDates(null);
      window.location.reload();
    } catch (e) {
      console.error("Auto-backup restore error:", e);
      showToast("Restore failed. Please check your internet connection and try again", "error");
    }
  };

  const calcCavStartTime = (startTime, duration, cavDuration = 15) => {
    if (!startTime) return "";
    const [h, m] = startTime.split(":").map(Number);
    const totalMins = h * 60 + m + (Number(duration) - Number(cavDuration));
    const ch = Math.floor(totalMins / 60);
    const cm = totalMins % 60;
    return `${String(ch).padStart(2, "0")}:${String(cm).padStart(2, "0")}`;
  };

  const saveAppt = async (appt) => {
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
      notes: `⚡ Machine for ${appt.clientName} (${appt.therapist})`,
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

    setAppointments(next);
    if (await save(next, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips)) { setEditingAppt(null); showToast("Saved"); }
  };

  const deleteAppt = async (id) => {
    if (guardLocked()) return;
    const cavSlotId = `cav-${id}`;
    const next = appointments.filter(a => a.id !== id && a.id !== cavSlotId);
    setAppointments(next); await save(next, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); setEditingAppt(null);
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
  const saveRetail = async (r) => { if (guardLocked()) return; const next = retails.find(x => x.id === r.id) ? retails.map(x => x.id === r.id ? r : x) : [...retails, r]; setRetails(next); if (await save(appointments, next, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips)) { setEditingRetail(null); showToast("Retail sale saved"); } };
  const deleteRetail = async (id) => { if (guardLocked()) return; const next = retails.filter(r => r.id !== id); setRetails(next); await save(appointments, next, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips); };
  const saveDeposit = async (d) => { if (guardLocked()) return; const next = deposits.find(x => x.id === d.id) ? deposits.map(x => x.id === d.id ? d : x) : [...deposits, d]; setDeposits(next); if (await save(appointments, retails, next, ticketPurchases, staffPurchases, refunds, forgottenTips)) { setEditingDeposit(null); showToast("Saved"); } };
  const deleteDeposit = async (id) => { if (guardLocked()) return; const next = deposits.filter(d => d.id !== id); setDeposits(next); await save(appointments, retails, next, ticketPurchases, staffPurchases, refunds, forgottenTips); };
  const saveTicketPurchase = async (tp) => { if (guardLocked()) return; const next = ticketPurchases.find(x => x.id === tp.id) ? ticketPurchases.map(x => x.id === tp.id ? tp : x) : [...ticketPurchases, tp]; setTicketPurchases(next); if (await save(appointments, retails, deposits, next, staffPurchases, refunds, forgottenTips)) { setEditingTicketPurchase(null); showToast("🎟️ Ticket purchase saved"); } };
  const deleteTicketPurchase = async (id) => { if (guardLocked()) return; const next = ticketPurchases.filter(tp => tp.id !== id); setTicketPurchases(next); await save(appointments, retails, deposits, next, staffPurchases, refunds, forgottenTips); };
  const saveStaffPurchase = async (sp) => { if (guardLocked()) return; const next = staffPurchases.find(x => x.id === sp.id) ? staffPurchases.map(x => x.id === sp.id ? sp : x) : [...staffPurchases, sp]; setStaffPurchases(next); if (await save(appointments, retails, deposits, ticketPurchases, next, refunds, forgottenTips)) { setEditingStaffPurchase(null); showToast("👩‍💼 Staff purchase saved"); } };
  const deleteStaffPurchase = async (id) => { if (guardLocked()) return; const next = staffPurchases.filter(sp => sp.id !== id); setStaffPurchases(next); await save(appointments, retails, deposits, ticketPurchases, next, refunds, forgottenTips); };
  const saveRefund = async (rf) => { if (guardLocked()) return; const next = refunds.find(x => x.id === rf.id) ? refunds.map(x => x.id === rf.id ? rf : x) : [...refunds, rf]; setRefunds(next); if (await save(appointments, retails, deposits, ticketPurchases, staffPurchases, next, forgottenTips)) { setEditingRefund(null); showToast("🔙 Refund recorded"); } };
  const deleteRefund = async (id) => { if (guardLocked()) return; const next = refunds.filter(rf => rf.id !== id); setRefunds(next); await save(appointments, retails, deposits, ticketPurchases, staffPurchases, next, forgottenTips); };
  const saveForgottenTip = async (ft) => { if (guardLocked()) return; const next = forgottenTips.find(x => x.id === ft.id) ? forgottenTips.map(x => x.id === ft.id ? ft : x) : [...forgottenTips, ft]; setForgottenTips(next); if (await save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, next)) { setEditingForgottenTip(null); showToast("🙏 Forgotten entry recorded"); } };
  const deleteForgottenTip = async (id) => { if (guardLocked()) return; const next = forgottenTips.filter(ft => ft.id !== id); setForgottenTips(next); await save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, next); };

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
  const sheetCashTip = r2(tipCashAllSources);
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
          <DatePicker value={date} onChange={setDate} allowClear={false}
            style={{ padding: "8px 12px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontSize: 14 }} />
          <button onClick={fetchSquare} disabled={squareLoading}
            style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: squareLoading ? "#666" : "#E8A84A", color: "#fff", fontWeight: 700, cursor: squareLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {squareLoading ? "⏳ Fetching..." : "□ Square Sync"}
          </button>
          <button onClick={syncOnlineGiftCards} disabled={gcSyncLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: gcSyncLoading ? "#666" : "#B45309", color: "#fff", fontWeight: 700, cursor: gcSyncLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {gcSyncLoading ? "⏳ Fetching..." : "🎁 Fetch Gift Card Purchases"}
          </button>
          <button onClick={syncSquareDeposits} disabled={depositSyncLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: depositSyncLoading ? "#666" : "#00796B", color: "#fff", fontWeight: 700, cursor: depositSyncLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {depositSyncLoading ? "⏳ Fetching..." : "💰 Auto-Fetch Deposits"}
          </button>
          <button onClick={checkSquareReconciliation} disabled={reconcileLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: reconcileLoading ? "#666" : "#4A6572", color: "#fff", fontWeight: 700, cursor: reconcileLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {reconcileLoading ? "⏳ Checking..." : "🔍 Square Reconcile"}
          </button>
          <button onClick={handleLockToggle}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: locked ? "#C62828" : "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            {locked ? "🔒 Finalized (Unlock)" : "🔓 Finalize This Day"}
          </button>
          <button onClick={handleExportBackup}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            📤 Backup Data
          </button>
          <label style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            📥 Restore Data
            <input type="file" accept="application/json" onChange={handleImportBackup} style={{ display: "none" }} />
          </label>
          <button onClick={openAutoBackupList} disabled={autoBackupListLoading}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "rgba(255,255,255,0.15)", color: "#fff", fontWeight: 700, cursor: autoBackupListLoading ? "not-allowed" : "pointer", fontSize: 13 }}>
            {autoBackupListLoading ? "⏳ Loading..." : "🕐 Restore from Auto-Backup"}
          </button>
        </div>
      </div>
      {dayLoading && (
        <div style={{ background: "#E3F2FD", padding: "8px 20px", fontSize: 13, color: "#1565C0", fontWeight: 700 }}>
          ⏳ Loading data...
        </div>
      )}
      {locked && (
        <div style={{ background: "#FFEBEE", padding: "8px 20px", fontSize: 13, color: "#C62828", fontWeight: 700, borderBottom: "2px solid #C62828" }}>
          🔒 This day is finalized. To edit or delete, unlock it with the "Finalized" button and enter the PIN.
        </div>
      )}
      {squareStatus && <div style={{ background: "#FFF3CD", padding: "8px 20px", fontSize: 13, color: "#856404" }}>⚠️ {squareStatus}</div>}

      {reconcileResult && (() => {
        // Cash payments never go through Square's tip-prompt screen, so tip_money is always 0
        // for them — but /api/square-payments falls back to the order's itemized "Tip" line
        // (same as it already did for card) when staff rang the tip up as a manual line item
        // inside a cash order, so cash tips are comparable too as long as that's how they were
        // entered in Square. A cash tip paid with no "Tip" line item at all still won't show up
        // here (there's nothing in Square to recover it from).
        const rows = [
          { label: "Cash Total (Treatment + Retail + Tip)", sheet: sheetCashTotal, square: reconcileResult.cashTotal },
          { label: "Card Total (Treatment + Retail + Tip)", sheet: sheetCardTotal, square: reconcileResult.cardTotal },
          { label: "Cash Tip", sheet: sheetCashTip, square: reconcileResult.cashTip },
          { label: "Card Tip", sheet: sheetCardTip, square: reconcileResult.cardTip },
        ];
        return (
          <div style={{ background: "#fff", margin: "10px 20px", borderRadius: 10, border: "1px solid #DDD", padding: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#0D4F4F" }}>🔍 Square Reconciliation Result</div>
              <button onClick={() => setReconcileResult(null)}
                style={{ border: "none", background: "none", color: "#999", cursor: "pointer", fontSize: 13 }}>✕ Close</button>
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
                <thead>
                  <tr style={{ background: "#F7F4EE" }}>
                    <th style={{ textAlign: "left", padding: "6px 10px" }}>Item</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>Sheet Entry</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>Square Actual</th>
                    <th style={{ textAlign: "right", padding: "6px 10px" }}>Difference</th>
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
                          {mismatch ? `⚠️ ${diff > 0 ? "+" : ""}${formatCurrency(diff)}` : "✓ Match"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: "#999", marginTop: 8 }}>
              ※ Cash tips aren't broken out separately on Square's side, so they're excluded from this check (only the cash total is compared).
            </div>
          </div>
        );
      })()}

      {/* Tabs */}
      <div style={{ background: "#fff", borderBottom: "2px solid #E8E4DC", display: "flex", padding: "0 20px", overflowX: "auto" }}>
        {[["schedule","📅 Schedule"],["summary","📊 Summary"],["payroll","💴 Payroll"]].map(([tab, label]) => (
          <button key={tab} onClick={() => setActiveTab(tab)} style={{ padding: "12px 20px", border: "none", background: "none", cursor: "pointer", fontSize: 14, fontWeight: 600, color: activeTab === tab ? "#0D4F4F" : "#888", borderBottom: activeTab === tab ? "2px solid #0D4F4F" : "2px solid transparent", marginBottom: -2, whiteSpace: "nowrap" }}>
            {label}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", flexWrap: "wrap" }}>
          <select value={selectedTherapist} onChange={e => setSelectedTherapist(e.target.value)} style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid #DDD", fontSize: 13 }}>
            <option value="All">Everyone</option>
            {workingStaff.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          <button onClick={() => setShowStaffPicker(p => !p)}
            style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #0D4F4F", background: showStaffPicker ? "#0D4F4F" : "#fff", color: showStaffPicker ? "#fff" : "#0D4F4F", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>
            👥 Working Staff
          </button>
        </div>
      </div>

      {/* Working staff picker */}
      {showStaffPicker && (
        <div style={{ background: "#E8F5E9", padding: "10px 20px", borderBottom: "1px solid #C8E6C9", display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 12, fontWeight: 700, color: "#2E7D32" }}>Today's working staff:</span>
          {THERAPISTS.map(t => (
            <button key={t} onClick={() => {
              const next = workingStaff.includes(t) ? workingStaff.filter(x => x !== t) : [...workingStaff, t];
              setWorkingStaff(next);
              save(appointments, retails, deposits, ticketPurchases, staffPurchases, refunds, forgottenTips, next);
            }} style={{
              padding: "4px 12px", borderRadius: 20, border: `2px solid ${workingStaff.includes(t) ? "#2E7D32" : "#CCC"}`,
              background: workingStaff.includes(t) ? "#2E7D32" : "#F5F5F5",
              color: workingStaff.includes(t) ? "#fff" : "#888",
              cursor: "pointer", fontWeight: 600, fontSize: 12
            }}>{t}</button>
          ))}
          <button onClick={() => setShowStaffPicker(false)}
            style={{ marginLeft: "auto", padding: "4px 12px", borderRadius: 8, border: "none", background: "#0D4F4F", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
            Done
          </button>
        </div>
      )}

      {/* Schedule */}
      {activeTab === "schedule" && (
        <div style={{ padding: "16px 12px", overflowX: "auto" }}>

          {/* Today's visit count — excludes cavSlot rows (duplicate row for the machine therapist's share of the same client) */}
          <div style={{ display: "inline-block", background: "#0D4F4F", color: "#fff", borderRadius: 20, padding: "6px 16px", fontSize: 13, fontWeight: 700, marginBottom: 12 }}>
            👥 Today's Visits: {appointments.filter(a => !a.isCavSlot).length}
          </div>

          {/* Clients arriving today who already have a deposit on file */}
          {depositsForDate.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 14, border: "2px solid #1565C0" }}>
              <div style={{ fontWeight: 800, fontSize: 13, color: "#1565C0", marginBottom: 10 }}>
                💰 Arriving Today — {depositsForDate.length} client(s) already paid
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
                        {dep.type === "giftcard" ? "🎁 Fully Prepaid" : "💰 Deposit"}
                      </span>
                    </div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#1565C0", marginTop: 3 }}>
                      💰 ${dep.amount}
                      <span style={{ fontSize: 11, fontWeight: 400, color: "#888", marginLeft: 6 }}>
                        {dep.paymentType === "cash" ? "Cash" : dep.paymentType === "card" ? "Card" : "Check"}
                      </span>
                      {Number(dep.tip) > 0 && <span style={{ fontSize: 11, color: "#E65100", marginLeft: 6 }}>+ Tip ${dep.tip}</span>}
                    </div>
                    {dep.appointmentTime && (
                      <div style={{ fontSize: 12, color: "#555", marginTop: 3 }}>🕐 {dep.appointmentTime}</div>
                    )}
                    {dep.notes && <div style={{ fontSize: 11, color: "#666", marginTop: 3 }}>{dep.notes}</div>}
                    <div style={{ fontSize: 10, color: "#AAA", marginTop: 5 }}>Paid on: {dep.recordedDate}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 800 }}>
            <thead>
              <tr>
                <th style={{ width: 70, padding: "8px 4px", fontSize: 11, color: "#888", textAlign: "left", borderBottom: "2px solid #0D4F4F" }}>Time</th>
                {visibleTherapists.map(t => (
                  <th key={t} style={{ padding: "8px 6px", fontSize: 13, fontWeight: 700, color: "#0D4F4F", textAlign: "center", borderBottom: "2px solid #0D4F4F", minWidth: 190 }}>
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
                          if (locked) { showToast("🔒 This day is finalized. Unlock it before editing", "error"); return; }
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
                              <div style={{ fontSize: 10, fontWeight: 800, color: "#00695C" }}>➕ Add-on</div>
                              <div style={{ fontSize: 11, fontWeight: 700, color: "#004D40" }}>{parentAppt.clientName}</div>
                              <div style={{ fontSize: 10, color: "#00796B" }}>
                                {addon.serviceName || "Add-on"}{addon.ticketCurrent ? ` ${addon.ticketCurrent}/${parentAppt.ticketTotal||3}` : ""}
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
          <SectionBox title="🛍️ Retail" color="#6A1B9A" onAdd={() => setEditingRetail({ ...EMPTY_RETAIL, id: Date.now() })} disabled={locked}>
            {retails.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
            {retails.map(r => (
              <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ flex: 1, fontSize: 14 }}>{r.item || "(not entered)"}</span>
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
          <SectionBox title="💰 Deposits / Gift Cards / Cancellation Fees" color="#1565C0" onAdd={() => setEditingDeposit({ ...EMPTY_DEPOSIT, id: Date.now() })} disabled={locked}>
            {deposits.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
            {deposits.map(d => (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, background: d.type === "deposit" ? "#E3F2FD" : d.type === "cancellation" ? "#FFEBEE" : "#FFF3E0", color: d.type === "deposit" ? "#1565C0" : d.type === "cancellation" ? "#C62828" : "#E65100", padding: "2px 8px", borderRadius: 10, fontWeight: 700 }}>
                  {d.type === "deposit" ? "Deposit" : d.type === "cancellation" ? "❌ Cancellation Fee" : "Gift Card"}
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
          <SectionBox title="🙏 Forgotten Entry (Deposit/Tip)" color="#00695C" onAdd={() => setEditingForgottenTip({ ...EMPTY_FORGOTTEN_TIP, id: Date.now() })} disabled={locked}>
            {forgottenTips.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
            {forgottenTips.map(ft => {
              const svc = Number(ft.serviceAmount || 0);
              const tip = Number(ft.tipAmount || 0);
              return (
                <div key={ft.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{ft.clientName || "—"}</span>
                  {ft.therapist && <span style={{ fontSize: 11, color: "#00695C", background: "#E0F2F1", borderRadius: 8, padding: "2px 8px" }}>{ft.therapist}</span>}
                  {svc > 0 && <span style={{ fontWeight: 700, color: "#00695C" }}>Treatment {formatCurrency(svc)}</span>}
                  {tip > 0 && <span style={{ fontWeight: 700, color: "#00695C" }}>Tip {formatCurrency(tip)}</span>}
                  <PayBadge type={ft.paymentType} />
                  {ft.originalDate && <span style={{ fontSize: 11, color: "#888" }}>Visit date {ft.originalDate}</span>}
                  <button onClick={() => setEditingForgottenTip(ft)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                  <button onClick={() => deleteForgottenTip(ft.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                </div>
              );
            })}
            {forgottenTips.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#00695C", fontWeight: 700, textAlign: "right" }}>
                Treatment {formatCurrency(totalForgottenService)}　Tip {formatCurrency(totalForgottenTip)}
              </div>
            )}
          </SectionBox>

          {/* Staff Purchases 社販 */}
          <SectionBox title="👩‍💼 Staff Purchases (Employee Purchase/Treatment)" color="#37474F" onAdd={() => setEditingStaffPurchase({ ...EMPTY_STAFF_PURCHASE, id: Date.now() })} disabled={locked}>
            {staffPurchases.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
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
                      <span style={{ fontSize: 11, color: "#37474F", background: "#ECEFF1", borderRadius: 8, padding: "2px 8px" }}>{items.length} items</span>
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
                          <span style={{ flex: 1 }}>{it.productName || "(no product name)"}</span>
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
          <SectionBox title="🎟️ New Ticket Purchase (Phone/Walk-in)" color="#B71C1C" onAdd={() => setEditingTicketPurchase({ ...EMPTY_TICKET_PURCHASE, id: Date.now() })} disabled={locked}>
            {ticketPurchases.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
            {ticketPurchases.map(tp => {
              const tip = Number(tp.tip || 0);
              const amt = Number(tp.amount || 0);
              return (
                <div key={tp.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{tp.clientName || "—"}</span>
                  {tp.packageName && <span style={{ fontSize: 11, color: "#B71C1C", background: "#FFEBEE", borderRadius: 8, padding: "2px 8px" }}>{tp.packageName}</span>}
                  <span style={{ fontWeight: 700, color: "#B71C1C" }}>{formatCurrency(amt)}</span>
                  {tip > 0 && <span style={{ fontSize: 11, color: "#E65100" }}>Tip {formatCurrency(tip)}</span>}
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
                Total {formatCurrency(tpTotal)}　Tip {formatCurrency(tpTipTotal)}
              </div>
            )}
          </SectionBox>

          {/* Refunds — for a past visit's refund processed today (deducts from today's totals only) */}
          <SectionBox title="🔙 Refund" color="#5D4037" onAdd={() => setEditingRefund({ ...EMPTY_REFUND, id: Date.now() })} disabled={locked}>
            {refunds.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>None</p>}
            {refunds.map(rf => {
              const svc = Number(rf.serviceAmount || 0);
              const tip = Number(rf.tipAmount || 0);
              return (
                <div key={rf.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                  <span style={{ flex: 1, fontWeight: 700, fontSize: 14 }}>{rf.clientName || "—"}</span>
                  {rf.originalDate && <span style={{ fontSize: 11, color: "#5D4037", background: "#EFEBE9", borderRadius: 8, padding: "2px 8px" }}>Visit date: {rf.originalDate}</span>}
                  {svc > 0 && <span style={{ fontWeight: 700, color: "#C62828" }}>Treatment -{formatCurrency(svc)}</span>}
                  {tip > 0 && <span style={{ fontWeight: 700, color: "#C62828" }}>Tip -{formatCurrency(tip)}</span>}
                  <PayBadge type={rf.paymentType} />
                  {rf.notes && <span style={{ fontSize: 11, color: "#888" }}>{rf.notes}</span>}
                  <button onClick={() => setEditingRefund(rf)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>✏️</button>
                  <button onClick={() => deleteRefund(rf.id)} disabled={locked} style={{...iconBtn, opacity: locked ? 0.35 : 1, cursor: locked ? "not-allowed" : "pointer"}}>🗑️</button>
                </div>
              );
            })}
            {refunds.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#C62828", fontWeight: 700, textAlign: "right" }}>
                Treatment -{formatCurrency(totalRefundService)}　Tip -{formatCurrency(totalRefundTip)}
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
              📋 Today's Sales Summary　<span style={{ fontSize: 12, color: "#888", fontWeight: 400 }}>Sales Report format</span>
            </div>

            {/* Visit count & total sales */}
            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              <div style={{ background: "#F0F7FF", borderRadius: 10, padding: "10px 18px", textAlign: "center", minWidth: 90 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Visits</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#0D4F4F" }}>{appointments.filter(a => !a.isCavSlot).length}</div>
              </div>
              <div style={{ background: "#FFF8E1", borderRadius: 10, padding: "10px 18px", textAlign: "center", flex: 1, minWidth: 120 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Total Sales (Treatment + Retail)</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#C62828" }}>{formatCurrency(totalSalesAll)}</div>
              </div>
              <div style={{ background: "#F3E5F5", borderRadius: 10, padding: "10px 18px", textAlign: "center", flex: 1, minWidth: 120 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Total Sales + Tip</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#6A1B9A" }}>{formatCurrency(totalSalesAll + totalTipAllCC)}</div>
              </div>
            </div>

            {/* Cash section */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", background: "#2E7D32", padding: "5px 12px", borderRadius: 6, marginBottom: 8, display: "inline-block" }}>💵 CASH</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                {[
                  { label: "Treatment", sub: "Includes deposits/GC/cancellation fees/ticket purchases", value: cashTreatmentAll, color: "#2E7D32" },
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
                  { label: "Treatment", sub: "Includes deposits/GC/cancellation fees/ticket purchases", value: cardTreatmentAll, color: "#1565C0" },
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
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Total Tip (Cash + Card)</div>
                <div style={{ fontSize: 20, fontWeight: 800, color: "#E65100" }}>{formatCurrency(totalTipAllCC)}</div>
              </div>
            </div>

            {/* Other breakdown (already included in Treatment above) */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(110px,1fr))", gap: 8, marginBottom: 12 }}>
              {[
                { label: "Deposits Received", note: "Already counted in Treatment", value: totalDepositAmt, color: "#0277BD", bg: "#E1F5FE" },
                ...(totalDepositApplied > 0 ? [{ label: "💰 Added to Payroll", value: totalDepositApplied, color: "#2E7D32", bg: "#C8E6C9", prefix: "+" }] : []),
                { label: "Gift Card Purchases", note: "Already counted in Treatment", value: totalGiftCard, color: "#00796B", bg: "#E0F2F1" },
                ...(totalGCUsed > 0 ? [{ label: "🎁 GC Used", value: totalGCUsed, color: "#00695C", bg: "#B2DFDB" }] : []),
                { label: "❌ Cancellation Fees", note: "Already counted in Treatment", value: totalCancellation, color: "#C62828", bg: "#FFEBEE" },
                { label: "Retail Total", value: totalRetail + inlineRetailTotal + spTotal, color: "#6A1B9A", bg: "#F3E5F5" },
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
              <div style={{ fontSize: 12, fontWeight: 700, color: "#0D4F4F", marginBottom: 8 }}>📊 Today's GRAND TOTAL</div>
              {/* Treatment breakdown */}
              <div style={{ background: "#F9F9F9", borderRadius: 8, padding: "8px 12px", marginBottom: 8, fontSize: 11, color: "#555" }}>
                <span style={{ fontWeight: 700 }}>Treatment total (incl. Deposit/GC/ticket purchases):</span>
                <span style={{ color: "#C62828", fontWeight: 800, fontSize: 14, marginLeft: 4 }}>{formatCurrency(totalCash + totalCard + totalDepositAmt + totalGiftCard + totalCancellation + tpTotal + inlineNTTotal + inlineGCTotal)}</span>
                <div style={{ marginTop: 4, color: "#999", fontSize: 10 }}>
                  Cash ${totalCash.toFixed(2)}　Card ${totalCard.toFixed(2)}
                  {totalDepositAmt > 0 && `　Deposit $${totalDepositAmt.toFixed(2)}`}
                  {totalGiftCard > 0 && `　GiftCard $${(totalGiftCard + inlineGCTotal).toFixed(2)}`}
                  {totalCancellation > 0 && `　Cancellation Fee $${totalCancellation.toFixed(2)}`}
                  {(tpTotal + inlineNTTotal) > 0 && `　Ticket Purchase $${(tpTotal + inlineNTTotal).toFixed(2)}`}
                  {inlineRetailTotal > 0 && `　Retail (inline) $${inlineRetailTotal.toFixed(2)}`}
                </div>
              </div>
              {/* Final totals grid */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6 }}>
                {[
                  { label: "Cash Total\n(all)", value: totalCash + retails.filter(r=>r.paymentType==="cash").reduce((s,r)=>s+Number(r.price||0),0) + tpCash + inlineNTCash + inlineRetailCash + inlineGCCash + totalCancellationCash + spCash, color: "#2E7D32", bg: "#E8F5E9" },
                  { label: "Card Total\n(all)", value: totalCard + retails.filter(r=>r.paymentType==="card").reduce((s,r)=>s+Number(r.price||0),0) + tpCard + inlineNTCard + inlineRetailCard + inlineGCCard + totalCancellationCard + spCard, color: "#1565C0", bg: "#E3F2FD" },
                  { label: "Tip Total", value: totalTipAllCC, color: "#E65100", bg: "#FFF3E0" },
                  { label: "🏆 TOTAL\n(everything)", value: totalCash + totalCard + totalDepositAmt + totalGiftCard + totalCancellation + totalRetail + totalTipAllCC + tpTotal + inlineNTTotal + inlineRetailTotal + inlineGCTotal + spTotal, color: "#fff", bg: "#0D4F4F", bold: true },
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
              <div style={{ fontSize: 12, fontWeight: 700, color: "#2E7D32", marginBottom: 8 }}>🟢 Same-Day Ticket Purchases: {sameDayAppts.length}</div>
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
                    <span style={{ fontSize: 12, color: "#2E7D32" }}>{a.ticketMenu} {a.ticketCurrent > 0 ? `session ${a.ticketCurrent}/${a.ticketTotal}` : `purchase only (${a.ticketTotal}-session course)`}</span>
                    <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` + ${a.cavTherapist} (machine)` : ""}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                    <span style={{ fontSize: 11, color: "#2E7D32" }}>💰 {a.packageSplitPayment ? `💵$${a.packageCashPortion||0}＋💳$${a.packageCardPortion||0}` : `$${pkgSvc}${a.paymentType==="card"?"💳":"💵"}`} / Tip ${pkgTip}{a.tipPaymentType==="card"?"💳":"💵"}{extraSvc > 0 && <span style={{ color: "#F57F17" }}> +${extraSvc}{a.extraPricePaymentType==="card"?"💳":"💵"}</span>}{extra > 0 && <span style={{ color: "#F57F17" }}> +${extra}💝{a.extraTipPaymentType==="card"?"💳":"💵"}</span>}</span>
                    <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>Total ${pkgSvc + pkgTip + extra + extraSvc}</span>
                  </div>
                  {gc > 0 && (
                    <div style={{ marginTop: 4, fontSize: 10, color: NON_REVENUE_COLOR }}>
                      🔵 Used ${gcSvc + gcTip} of gift card balance (not counted in today's sales — today's actual sales: ${total})
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
              <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>🔵 Already Paid (Ticket Redemption): {pureTicketAppts.length}</div>
              {pureTicketAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0) + Number(a.extraTip||0));
                const extraSvc = Number(a.extraPrice||0);
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#1565C0" }}>{a.ticketMenu} {a.ticketCurrent}/{a.ticketTotal}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` + ${a.cavTherapist} (machine)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4, flexWrap: "wrap", gap: 4 }}>
                      <span style={{ fontSize: 11, color: "#1565C0" }}>Treatment ${svc} / Tip ${Number(a.tip||0)+Number(a.cavTip||0)}{extraSvc > 0 && <span style={{ color: "#F57F17" }}> +${extraSvc}{a.extraPricePaymentType==="card"?"💳":"💵"}</span>}{a.extraTip > 0 && <span style={{ color: "#F57F17" }}> +${a.extraTip}💝{a.extraTipPaymentType==="card"?"💳":"💵"}</span>}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>Total ${svc + tip + extraSvc}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* GC消化 (pre-purchased GC used today) */}
          {gcAppts.length > 0 && (
            <div style={{ background: "#FFFDE7", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #F59E0B" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#B45309", marginBottom: 8 }}>🎁 Already Paid (Gift Card Redemption): {gcAppts.length}</div>
              {gcAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0));
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#B45309" }}>{a.serviceName || `${a.duration}min`}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` + ${a.cavTherapist} (machine)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                      <span style={{ fontSize: 11, color: "#B45309" }}>Treatment ${svc} / Tip ${tip}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>Total ${svc + tip}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* PR無料 (comped treatments for influencers etc.) */}
          {promoAppts.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 12, borderLeft: "4px solid #1565C0" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>📸 Complimentary PR (not counted in treatment sales): {promoAppts.length}</div>
              {promoAppts.map(a => {
                const svc = r2(Number(a.price||0) + Number(a.cavPrice||0));
                const tip = r2(Number(a.tip||0) + Number(a.cavTip||0));
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#1565C0" }}>{a.serviceName || `${a.duration}min`}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` + ${a.cavTherapist} (machine)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                      <span style={{ fontSize: 11, color: "#1565C0" }}>Treatment ${svc} / Tip ${tip}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>Total ${svc + tip}</span>
                    </div>
                    {a.notes && <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>📝 {a.notes}</div>}
                  </div>
                );
              })}
            </div>
          )}

          {/* Customer type & referral source */}
          <div style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 14, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
            <div style={{ fontWeight: 700, marginBottom: 10, color: "#0D4F4F" }}>By Customer Type</div>
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
                <div style={{ fontWeight: 700, marginBottom: 8, color: "#9C27B0", fontSize: 13 }}>📍 New Customer Source</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {REFERRAL_SOURCES.map(src => {
                    const count = appointments.filter(a => !a.isCavSlot && a.referralSource === src).length;
                    if (!count) return null;
                    return (
                      <div key={src} style={{ background: "#F3E5F5", borderRadius: 8, padding: "6px 12px", textAlign: "center" }}>
                        <div style={{ fontSize: 18, fontWeight: 800, color: "#9C27B0" }}>{count}</div>
                        <div style={{ fontSize: 11, color: "#6A1B9A" }}>{REFERRAL_LABELS[src] || src}</div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>

          {/* By Therapist */}
          <div style={{ background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)", overflowX: "auto" }}>
            <div style={{ fontWeight: 700, marginBottom: 12, color: "#0D4F4F" }}>By Therapist</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: "2px solid #E0E0E0", background: "#F9F9F9" }}>
                  {["Name","Clients",...CUSTOMER_TYPES,"New Tickets Sold","New Ticket Amount"].map(h => (
                    <th key={h} style={{ padding: "6px 8px", textAlign: h==="Name"?"left":"center", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
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
        clientDeposits={clientDepositHistory}
      />}
      {editingRetail && <RetailModal retail={editingRetail} onSave={saveRetail} onClose={() => setEditingRetail(null)} />}
      {editingDeposit && <DepositModal deposit={editingDeposit} onSave={saveDeposit} onDelete={() => { deleteDeposit(editingDeposit.id); setEditingDeposit(null); }} onClose={() => setEditingDeposit(null)} />}
      {editingTicketPurchase && <TicketPurchaseModal tp={editingTicketPurchase} onSave={saveTicketPurchase} onDelete={() => { deleteTicketPurchase(editingTicketPurchase.id); setEditingTicketPurchase(null); }} onClose={() => setEditingTicketPurchase(null)} />}
      {editingStaffPurchase && <StaffPurchaseModal sp={editingStaffPurchase} onSave={saveStaffPurchase} onDelete={() => { deleteStaffPurchase(editingStaffPurchase.id); setEditingStaffPurchase(null); }} onClose={() => setEditingStaffPurchase(null)} />}
      {editingRefund && <RefundModal rf={editingRefund} onSave={saveRefund} onDelete={() => { deleteRefund(editingRefund.id); setEditingRefund(null); }} onClose={() => setEditingRefund(null)} />}
      {editingForgottenTip && <ForgottenTipModal ft={editingForgottenTip} onSave={saveForgottenTip} onDelete={() => { deleteForgottenTip(editingForgottenTip.id); setEditingForgottenTip(null); }} onClose={() => setEditingForgottenTip(null)} />}
      {autoBackupDates && <AutoBackupModal dates={autoBackupDates} onRestore={restoreFromAutoBackup} onClose={() => setAutoBackupDates(null)} />}

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
        <div style={{ fontSize: 17, fontWeight: 700, color: "#6A1B9A" }}>⚡ Machine {appt.isTicket ? "🎟️" : ""}</div>
        <div style={{ fontSize: 17, color: "#555" }}>
          {appt.clientName} <span style={{ color: "#888", fontSize: 13 }}>({appt.startTime}〜)</span>
        </div>
        {appt.bodyTherapist && (
          <div style={{ fontSize: 13, color: "#8E24AA" }}>with {appt.bodyTherapist}</div>
        )}
        {(cavSvc > 0 || cavTip > 0) && (
          <div style={{ fontSize: 17, color: amountColor, fontWeight: 700 }}>
            <div>{cavSvc}{svcIcon}</div>
            {cavTip > 0 && <div>{cavTip}{tipIcon}</div>}
            <div style={{ fontWeight: 800 }}>{cavTotal}</div>
          </div>
        )}
        {dep > 0 && !isWeightLossService(parent?.serviceName) && (
          <div style={{ color: "#2E7D32", fontSize: 12 }}>💰Deposit split included</div>
        )}
        {(() => {
          // A retail sale on the parent visit can be split with this machine (cav) therapist
          // (see the "retail" purchaseTag branch below) — the payroll split already accounts
          // for it, but without this it never showed up anywhere on the cav therapist's own
          // card, so it was easy to forget to double-check later.
          if (!parent) return null;
          const items = getRetailItems(parent);
          let myShare = 0;
          items.forEach(it => (it.sellers || []).forEach(sel => {
            if (sel.therapist === appt.therapist) myShare += Number(sel.amount || 0);
          }));
          if (myShare <= 0) return null;
          // Black, not red — the full sale total already shows in red on the body therapist's
          // own card; this is a payroll-split detail, not a second sale to add on top.
          return (
            <div style={{ color: "#333", fontSize: 12, fontWeight: 700 }}>
              🛍️ Retail ${r2(myShare)} (with {appt.bodyTherapist})
            </div>
          );
        })()}
      </div>
    );
  }

  return (
    <div onClick={onClick} style={{ background: bg, border: `2px solid ${borderColor}`, borderRadius: 8, padding: "6px 8px", marginBottom: 3, cursor: "pointer" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <span style={{ fontSize: 18, fontWeight: 700, lineHeight: 1.3 }}>{appt.clientName || "(not set)"}</span>
        <span style={{ fontSize: 13, fontWeight: 700, color: "#fff", background: typeColors[appt.customerType]||"#888", padding: "1px 5px", borderRadius: 8, marginLeft: 4, whiteSpace: "nowrap" }}>
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

        const courseName = isTicket ? (stripJpAnnotation(appt.serviceName) || appt.ticketMenu) : (stripJpAnnotation(appt.serviceName) || `${appt.duration}min`);
        const sessionSuffix = isRedemption && appt.ticketCurrent > 0 ? ` ${appt.ticketCurrent}/${appt.ticketTotal}`
          : isSameDay ? ` x${appt.ticketTotal}` : "";

        const gc = (!isGiftCard && !isPromo) ? Number(appt.giftCardUsed||0) : 0;

        return (
          <div style={{ fontSize: 17, marginTop: 2 }}>
            <div style={{ color: isGiftCard ? "#B45309" : isPromo ? "#1565C0" : "#222", fontWeight: 700, lineHeight: 1.4 }}>
              {isGiftCard && <span style={{ fontSize: 13, background: "#F59E0B", color: "#fff", padding: "1px 5px", borderRadius: 6, marginRight: 4 }}>🎁GC Redemption</span>}
              {isPromo && <span style={{ fontSize: 13, background: "#1565C0", color: "#fff", padding: "1px 5px", borderRadius: 6, marginRight: 4 }}>📸Complimentary PR</span>}
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
              <div style={{ color: "#6A1B9A", fontSize: 13 }}>with {appt.cavTherapist}</div>
            )}
            {dep > 0 && (
              <div style={{ color: "#2E7D32", fontSize: 13 }}>💰Deposit ${dep} paid (not included in today's total received){depDate ? ` ${depDate}` : ""}</div>
            )}
            {gc > 0 && (
              <div style={{ color: "#2E7D32", fontSize: 13 }}>GC used ${gc}</div>
            )}
            {(appt.addons || []).length > 0 && (
              <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 3 }}>
                {(appt.addons || []).map(addon => {
                  const label = addon.serviceName || addon.name || "Add-on";
                  const aSvc = Number(addon.price||0);
                  const aTip = Number(addon.tip||0);
                  const aTotal = r2(aSvc + aTip);
                  const chipColor = addon.countsAsRevenue === true ? REVENUE_COLOR : addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#00695C";
                  const svcIcon = addon.paymentType === "card" ? "💳" : addon.paymentType === "cash" ? "💵" : "";
                  const tipIcon = addon.tipPaymentType === "card" ? "💳" : addon.tipPaymentType === "cash" ? "💵" : "";
                  return (
                    <div key={addon.id} style={{ fontSize: 13, color: chipColor, fontWeight: 700 }}>
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
              <span key={tagId} style={{ fontSize: 12, background: tag.bg, color: tag.color, padding: "2px 6px", borderRadius: 8, fontWeight: 700, alignSelf: "flex-start" }}>{tag.label}</span>
            );

            if (tagId === "newTicket") {
              const ntSvc = Number(appt.newTicketAmount||0);
              const ntTip = Number(appt.newTicketTip||0);
              if (ntSvc + ntTip === 0) return emptyChip;
              const ntSvcIcon = appt.newTicketPaymentType==="card"?"💳":appt.newTicketPaymentType==="cash"?"💵":"";
              const ntTipIcon = appt.newTicketTipPaymentType==="card"?"💳":appt.newTicketTipPaymentType==="cash"?"💵":"";
              const ntSvcText = appt.newTicketSplitPayment ? `💵$${appt.newTicketCashPortion||0}＋💳$${appt.newTicketCardPortion||0}` : `${ntSvc}${ntSvcIcon}`;
              return (
                <div key={tagId} style={{ fontSize: 13, color: REVENUE_COLOR, fontWeight: 700 }}>
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
              const names = items.map(it => it.productName).filter(Boolean).join(" / ");
              // Red is reserved for the one number that should match Square/the register total —
              // the full amount actually collected. The seller split is a payroll-allocation
              // detail, not extra revenue on top, so it's rendered in black: a staff member
              // scanning cards and mentally adding up red numbers should never double-count a
              // single sale as if the split were a second sale. An item with no explicit seller
              // split defaults its full amount to this card's own therapist (same fallback
              // PayrollTab uses).
              const sellerTotals = {};
              items.forEach(it => {
                const sellers = (it.sellers && it.sellers.length > 0) ? it.sellers : [{ therapist: appt.therapist, amount: Number(it.amount || 0) }];
                sellers.forEach(sel => {
                  if (!sel.therapist) return;
                  sellerTotals[sel.therapist] = (sellerTotals[sel.therapist] || 0) + Number(sel.amount || 0);
                });
              });
              const sellerNames = Object.keys(sellerTotals);
              return (
                <div key={tagId}>
                  <div style={{ fontSize: 13, color: REVENUE_COLOR, fontWeight: 700 }}>Retail ${total}{icon}</div>
                  {names && <div style={{ fontWeight: 600, color: "#333", fontSize: 12 }}>{names}</div>}
                  {sellerNames.length > 1 && (
                    <div style={{ fontSize: 11, fontWeight: 600, color: "#333" }}>
                      {sellerNames.map(n => `${n} $${r2(sellerTotals[n])}`).join(" / ")}
                    </div>
                  )}
                </div>
              );
            }
            if (tagId === "giftCard") {
              const total = Number(appt.giftCardPurchaseAmount||0);
              if (total === 0) return emptyChip;
              const icon = appt.giftCardPurchasePaymentType==="card"?"💳":appt.giftCardPurchasePaymentType==="cash"?"💵":"";
              return (
                <div key={tagId} style={{ fontSize: 13, color: REVENUE_COLOR, fontWeight: 700 }}>{tag.label} ${total}{icon}</div>
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
  clientName: "Client Name",
  serviceName: "Course / Treatment Menu",
  therapist: "Therapist (Body)",
  startTime: "Start Time",
  customerType: "Customer Type (RL/RT/NL/NT)",
  referralSource: "How did they hear about us (new customer)",
  price: "Treatment Price",
  packagePrice: "Package Price",
  ticketMenu: "Treatment Menu / Duration",
  ticketTotal: "Sessions in course (3 / 5)",
  priceVersionChosen: "Price Version (new / old)",
  ticketCurrent: "Which session number today",
  paymentType: "Treatment Payment Method",
  tipPaymentType: "Tip Payment Method",
  packagePaymentType: "Package Price Payment Method",
  packageTipPaymentType: "Same-Day Tip Payment Method",
  extraPricePaymentType: "Extra Treatment Fee Payment Method",
  extraTipPaymentType: "Extra Tip Payment Method",
  newTicketPaymentType: "New Ticket Purchase Payment Method",
  newTicketTipPaymentType: "New Ticket Purchase Tip Payment Method",
  retailPurchasePaymentType: "Retail Purchase Payment Method",
  giftCardPurchasePaymentType: "Gift Card Purchase Payment Method",
  depositPaidDate: "Deposit Paid Date",
  packageDepositDate: "Deposit Paid Date",
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
      <div style={{ fontWeight: 800, fontSize: 13, color: "#C62828", marginBottom: 4 }}>⚠️ Can't save — some required fields are missing</div>
      <div style={{ fontSize: 12, color: "#B71C1C" }}>
        {errors.map(e => APPT_ERROR_LABELS[e]).join(" / ")}
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
    const packageName = group ? `${group.group} ${durLabel}min x${total}` : "";
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#0D4F4F" }}>✏️ Edit Appointment</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer", color: "#888" }}>✕</button>
      </div>

      {clientDeposits.length > 0 && (
        <div style={{ background: "#FFF3E0", borderRadius: 10, padding: "10px 14px", marginBottom: 12, border: "2px solid #FF9800" }}>
          <div style={{ fontWeight: 800, fontSize: 13, color: "#E65100" }}>💰 Has a deposit / gift prepayment!</div>
          {clientDeposits.map((d, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, fontSize: 12, color: "#BF360C", marginTop: 4 }}>
              <div>
                {d.type === "giftcard" ? "🎁 " : "💰 "}
                <strong>{d.sheetDate}</strong> — ${d.amount}{Number(d.tip) > 0 ? ` (+ Tip $${d.tip})` : ""} ({d.paymentType === "cash" ? "Cash" : "Card"})
                {d.appointmentDate ? ` → Scheduled visit: ${d.appointmentDate}${d.appointmentTime ? ` ${d.appointmentTime}` : ""}` : ""}
                {d.notes ? ` — ${d.notes}` : ""}
              </div>
              {d.type === "deposit" && (
                <button onClick={() => setForm(f => ({ ...f, depositApplied: Number(d.amount) || 0, depositPaidDate: d.sheetDate }))}
                  style={{ flexShrink: 0, padding: "4px 10px", borderRadius: 8, border: "none", background: "#E65100", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 11 }}>
                  Use This
                </button>
              )}
            </div>
          ))}
        </div>
      )}

      <ErrorBanner />

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Customer Type" error={errors.includes("customerType")}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {CUSTOMER_TYPES.map(ct => {
              const colors = { RL:"#4CAF50", RT:"#2196F3", NL:"#FF9800", NT:"#9C27B0" };
              const labels = { RL:"Returning Local", RT:"Returning Traveler", NL:"New Local", NT:"New Traveler" };
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
          <Field label={form.customerType === "NT" ? "📍 How did they hear about us? (NT)" : "📍 How did they hear about us? (NL)"} error={errors.includes("referralSource")}>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {REFERRAL_SOURCES.map(src => (
                <button key={src} onClick={() => set("referralSource", form.referralSource === src ? "" : src)}
                  style={{
                    padding: "6px 12px", borderRadius: 20, border: `2px solid ${form.referralSource === src ? "#9C27B0" : "#DDD"}`,
                    background: form.referralSource === src ? "#9C27B0" : "#fff",
                    cursor: "pointer", fontWeight: 600, fontSize: 12,
                    color: form.referralSource === src ? "#fff" : "#666"
                  }}>
                  {REFERRAL_LABELS[src] || src}
                </button>
              ))}
            </div>
          </Field>
        )}

        <Field label="Client Name" error={errors.includes("clientName")}>
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={{ ...inputStyle, ...(errors.includes("clientName") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="Client Name" />
        </Field>
        <Field label="Course Name" error={errors.includes("serviceName")}>
          <select value={SQUARE_SERVICES.find(s => s.name === form.serviceName) ? form.serviceName : ""} onChange={e => {
            const name = e.target.value;
            // Clear any body/cav minute split left over from a previously-selected course —
            // it belongs to that course's duration, not this one, and would otherwise block
            // the Weight Loss auto-default (fixed 40min machine) from kicking in.
            const svc = SQUARE_SERVICES.find(s => s.name === name);
            setForm(f => ({ ...f, serviceName: name, bodyMins: undefined, cavMins: undefined, ...(svc ? { duration: svc.duration } : {}) }));
          }} style={{ ...inputStyle, color: "#1a1a1a", fontWeight: 600, ...(errors.includes("serviceName") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }}>
            <option value="">Select a course</option>
            {SQUARE_SERVICES.map(s => (
              <option key={s.name} value={s.name}>{s.name}</option>
            ))}
          </select>
          {/* Free-text entry when the Square-synced name isn't in the list */}
          {!SQUARE_SERVICES.find(s => s.name === form.serviceName) && (
            <input type="text" value={form.serviceName || ""} placeholder="Enter course name"
              onChange={e => set("serviceName", e.target.value)}
              style={{ ...inputStyle, marginTop: 4 }} />
          )}
          {form.serviceName && (
            <div style={{ fontSize: 11, color: "#0D4F4F", marginTop: 4, fontWeight: 600 }}>
              ✅ {form.serviceName} ({form.duration}min)
            </div>
          )}

          {/* Extra menu(s) done by the SAME therapist in the same visit — name only, no separate
              price/payment tracking (that combined amount goes straight into the treatment/tip
              totals below). Use the Add-on section instead when a DIFFERENT therapist is involved. */}
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
            <option value="">+ Add a menu item for the same therapist</option>
            {ADDON_PRESETS.map(p => <option key={p} value={p}>{p}</option>)}
            {SQUARE_SERVICES.map(s => (
              <option key={s.name} value={s.name}>{s.name}</option>
            ))}
          </select>
        </Field>

        {/* ── Add-on (a different therapist's extra treatment, previous unused cav time, etc.) ── */}
        <div>
          {/* Service picker — selecting immediately adds the row (same one-step pattern as
              the same-therapist menu-add dropdown above) */}
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
              <option value="">+ Add a menu item for a different therapist</option>
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
                        placeholder="Enter service name (e.g. Previous unused cav time, session 2/3)" />
                      <button onClick={() => set("addons", form.addons.filter((_,i) => i !== idx))}
                        style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA", marginLeft: 6, flexShrink: 0 }}>✕</button>
                    </div>

                    {/* Revenue vs. ticket-consumption toggle */}
                    <div style={{ marginBottom: 6 }}>
                      <div style={{ fontSize: 10, color: addon.countsAsRevenue === null ? "#C62828" : "#888", marginBottom: 3, fontWeight: addon.countsAsRevenue === null ? 700 : 400 }}>
                        How to treat this{addon.countsAsRevenue === null && " ⚠️ Not selected"}
                      </div>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => upd({ countsAsRevenue: true, ticketCurrent: null })}
                          style={{ flex: 1, padding: "6px 4px", borderRadius: 8, border: `2px solid ${addon.countsAsRevenue === true ? REVENUE_COLOR : "#DDD"}`, background: addon.countsAsRevenue === true ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.countsAsRevenue === true ? REVENUE_COLOR : "#888" }}>
                          💰 Pay-as-you-go (new payment)
                        </button>
                        <button onClick={() => upd({ countsAsRevenue: false })}
                          style={{ flex: 1, padding: "6px 4px", borderRadius: 8, border: `2px solid ${addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#DDD"}`, background: addon.countsAsRevenue === false ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.countsAsRevenue === false ? NON_REVENUE_COLOR : "#888" }}>
                          🎫 Ticket Redemption (previous unused session, etc.)
                        </button>
                      </div>
                    </div>

                    {/* Session number — only when this addon is itself a ticket redemption */}
                    {addon.countsAsRevenue === false && form.isTicket && (
                      <div style={{ marginBottom: 6 }}>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Which session is this?</div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {Array.from({ length: form.ticketTotal||3 }, (_,i) => i+1).map(n => (
                            <button key={n} onClick={() => upd({ ticketCurrent: n })}
                              style={{ padding: "5px 10px", borderRadius: 8, border: `2px solid ${addon.ticketCurrent===n?"#0D4F4F":"#DDD"}`, background: addon.ticketCurrent===n?"#0D4F4F":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: addon.ticketCurrent===n?"#fff":"#888" }}>
                              {n}/{form.ticketTotal||3}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}

                    {/* Therapist */}
                    <div style={{ marginBottom: 6 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Therapist</div>
                      <select value={addon.therapist||""} onChange={e => upd({ therapist: e.target.value })}
                        style={{ ...inputStyle, fontSize: 12 }}>
                        <option value="">— Select —</option>
                        {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </div>

                    {/* Price + Tip inputs */}
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 6 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment Price ($)</div>
                        <input type="number" value={addon.price||""} onChange={e => upd({ price: e.target.value })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Tip ($)</div>
                        <input type="number" value={addon.tip||""} onChange={e => upd({ tip: e.target.value })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                    </div>

                    {/* Payment type — real money only when this is a new (revenue) payment, not a ticket redemption */}
                    {addon.countsAsRevenue !== false && (
                      <div style={{ marginBottom: tipAmt > 0 ? 6 : 0 }}>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Payment Method (Treatment)</div>
                        <PaymentToggle value={addon.paymentType} onChange={v => upd({ paymentType: v })} small />
                      </div>
                    )}

                    {/* Tip payment type */}
                    {tipAmt > 0 && addon.countsAsRevenue !== false && (
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Payment Method (Tip)</div>
                        <PaymentToggle value={addon.tipPaymentType} onChange={v => upd({ tipPaymentType: v })} small />
                      </div>
                    )}

                    {/* Sub-total */}
                    {(svcAmt + tipAmt) > 0 && (
                      <div style={{ marginTop: 6, textAlign: "right", fontSize: 11, color: "#0D4F4F", fontWeight: 700 }}>
                        Treatment ${svcAmt}{tipAmt > 0 ? ` + Tip $${tipAmt}` : ""} = <strong>${svcAmt + tipAmt}</strong>
                      </div>
                    )}
                  </div>
                );
              })}

              {/* Grand addon total */}
              {(form.addons||[]).some(a => Number(a.price||0) + Number(a.tip||0) > 0) && (
                <div style={{ background: "#E0F2F1", borderRadius: 8, padding: "8px 12px", textAlign: "right" }}>
                  <span style={{ fontSize: 12, color: "#00695C" }}>Add-on Total　</span>
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
            🔴 Regular Treatment
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: false, isGiftCard: false, isPromo: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#DDD"}`, background: form.isTicket && !form.isSameDayTicket ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#888", fontSize: 11 }}>
            🔵 Ticket Redemption
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: true, useToday: true, ticketCurrent: 1, isGiftCard: false, isPromo: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isSameDayTicket ? "#2E7D32" : "#DDD"}`, background: form.isSameDayTicket ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isSameDayTicket ? "#2E7D32" : "#888", fontSize: 11 }}>
            🟢 Same-Day Purchase
          </button>
          <button onClick={() => setForm(f => ({...f, isPromo: true, isGiftCard: false, isTicket: false, isSameDayTicket: false}))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isPromo ? "#1565C0" : "#DDD"}`, background: form.isPromo ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isPromo ? "#1565C0" : "#888", fontSize: 11 }}>
            📸 Complimentary PR
          </button>
        </div>
        {form.isGiftCard && (
          <div style={{ background: "#FFFDE7", border: "1.5px solid #F59E0B", borderRadius: 8, padding: "8px 10px", fontSize: 11, color: "#92400E" }}>
            ⚠️ This appointment was previously recorded as "GC Redemption". If only part of the gift card was used, or if a tip was received in cash/card, click "🔴 Regular Treatment" below and re-enter the amounts under "🎁 Gift Card Used".
          </div>
        )}

        {/* Ticket options */}
        {form.isTicket && (
          <div style={{ background: "#EEF4FF", borderRadius: 10, padding: 12 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#1565C0", marginBottom: 10 }}>🎟️ Ticket Settings</div>

            {/* Cav therapist — shown first; selecting triggers full auto-fill */}
            {isBodyOnly(form.therapist) && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ background: "#F3E5F5", borderRadius: 8, padding: 10, marginBottom: 10, border: `2px solid ${form.cavTherapist ? "#6A1B9A" : "#CE93D8"}` }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A", marginBottom: 6 }}>
                  ⚡ Machine Therapist
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
                  <option value="">None (no machine)</option>
                  {CAV_CAPABLE.filter(t => t !== form.therapist).map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                {form.cavTherapist ? (
                  <div style={{ fontSize: 11, color: "#6A1B9A", marginTop: 4, fontWeight: 600 }}>
                    ✅ {form.therapist} (Body) + {form.cavTherapist} (Machine)<br />
                    → Saving will automatically add the treatment fee to <strong>{form.cavTherapist}</strong>'s line
                  </div>
                ) : (
                  <div style={{ fontSize: 11, color: "#999", marginTop: 4 }}>
                    Select if a machine therapist is involved → the split will be applied automatically
                  </div>
                )}
              </div>
            )}
            {isDualLicense(form.therapist) && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ fontSize: 11, background: "#E8F5E9", color: "#2E7D32", padding: "6px 10px", borderRadius: 8, marginBottom: 10 }}>
                ✅ {form.therapist} is dual-licensed (Body + Machine) → enter the total as a single amount
              </div>
            )}

            {/* IP-equivalent shortcut — for Facial/Glow-type services when using an IP 90min ticket */}
            {form.serviceName && ["facial","glow","procell","channeling","hifu","sculpt"].some(k => form.serviceName.toLowerCase().includes(k)) && (
              <div style={{ background: "#FFF3E0", borderRadius: 8, padding: 8, marginBottom: 8, border: "1px solid #FFCC80" }}>
                <div style={{ fontSize: 11, color: "#E65100", fontWeight: 700, marginBottom: 6 }}>💡 Improving Posture 90min ticket equivalent option</div>
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
                  {form.ticketMenu?.startsWith("IP-90") ? "✅ Using IP 90min ticket" : "🔄 Use Improving Posture 90min ticket (equivalent)"}
                </button>
              </div>
            )}

            {/* Treatment menu selector */}
            <Field label="Treatment Menu" error={errors.includes("ticketMenu")}>
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
                <Field label="Duration">
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
                      }}>{dur}min</button>
                    ))}
                  </div>
                </Field>
              ) : null;
            })()}

            {/* Price version */}
            <Field label="Price Version" error={errors.includes("priceVersionChosen")}>
              <div style={{ display: "flex", gap: 8 }}>
                {[["new","🆕 New Price (Feb onward)"],["old","📋 Old Price (before Feb)"]].map(([v,label]) => (
                  <button key={v} onClick={() => { setForm(f => ({ ...f, priceVersion: v, priceVersionChosen: true })); if(form.ticketMenu) applyTicketPrices(form.ticketMenu, form.ticketTotal, v, form.cavTherapist, form.therapist); }}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.priceVersion===v && form.priceVersionChosen?"#1565C0":(errors.includes("priceVersionChosen")?"#C62828":"#DDD")}`, background: form.priceVersion===v && form.priceVersionChosen?"#BBDEFB":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: form.priceVersion===v && form.priceVersionChosen?"#1565C0":"#888" }}>
                    {label}
                  </button>
                ))}
              </div>
            </Field>

            {/* Total sessions */}
            <Field label="Sessions in Course" error={errors.includes("ticketTotal")}>
              <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                {[3,5].map(n => (
                  <button key={n} onClick={() => {
                    setForm(f => ({ ...f, ticketTotal: n, ticketTotalChosen: true }));
                    if(form.ticketMenu) {
                      const parts = form.ticketMenu.split("-");
                      applyTicketPrices(`${parts[0]}-${parts[1]}-${n}`, n, form.priceVersion||"new", form.cavTherapist, form.therapist);
                    }
                  }} style={{ flex: 1, padding: "8px", borderRadius: 8, border: `2px solid ${form.ticketTotal===n && form.ticketTotalChosen?"#1565C0":(errors.includes("ticketTotal")?"#C62828":"#DDD")}`, background: form.ticketTotal===n && form.ticketTotalChosen?"#1565C0":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketTotal===n && form.ticketTotalChosen?"#fff":"#888" }}>
                    {n}-session course
                  </button>
                ))}
              </div>
            </Field>

            {/* Current session — different UI for same-day purchase */}
            {form.isSameDayTicket ? (
              <Field label="Use it today?">
                <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                  <button onClick={() => setForm(f => ({...f, useToday: true, ticketCurrent: 1}))}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.useToday !== false ? "#2E7D32" : "#DDD"}`, background: form.useToday !== false ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, color: form.useToday !== false ? "#2E7D32" : "#888", fontSize: 12 }}>
                    ✅ Use today (session 1)
                  </button>
                  <button onClick={() => setForm(f => ({...f, useToday: false, ticketCurrent: 0, price: 0, tip: 0, cavPrice: 0, cavTip: 0}))}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.useToday === false ? "#888" : "#DDD"}`, background: form.useToday === false ? "#F5F5F5" : "#fff", cursor: "pointer", fontWeight: 700, color: form.useToday === false ? "#555" : "#888", fontSize: 12 }}>
                    ⭕ Use starting next visit
                  </button>
                </div>
              </Field>
            ) : (
              <Field label="Which session is today?" error={errors.includes("ticketCurrent")}>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                  {Array.from({ length: form.ticketTotal||3 }, (_,i) => i+1).map(n => (
                    <button key={n} onClick={() => set("ticketCurrent", n)}
                      style={{ padding: "6px 14px", borderRadius: 8, border: `2px solid ${form.ticketCurrent===n?"#0D4F4F":(errors.includes("ticketCurrent")?"#C62828":"#DDD")}`, background: form.ticketCurrent===n?"#0D4F4F":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketCurrent===n?"#fff":"#888" }}>
                      {n}/{form.ticketTotal||3}
                    </button>
                  ))}
                </div>
              </Field>
            )}


            {/* Prices — only when using today or redeeming a ticket */}
            {form.ticketMenu && (!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ marginTop: 8, background: "#fff", borderRadius: 8, padding: 10 }}>
                <div style={{ fontSize: 11, color: "#888", marginBottom: 6 }}>
                  {showSplitPrices ? `💡 ${form.therapist} (Body) + ${form.cavTherapist} (Machine) split` : `💡 ${form.therapist} Total Amount`}
                  {form.isSameDayTicket && <span style={{ color: "#2E7D32" }}> ※For staff payroll reference</span>}
                  {!form.isSameDayTicket && <span style={{ color: "#bbb" }}> ※Editable</span>}
                </div>
                {showSplitPrices ? (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label={`${form.therapist} Treatment ($)`} error={errors.includes("price")}><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label={`${form.therapist} Tip ($)`}><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                    <Field label={`${form.cavTherapist} Treatment ($)`} error={errors.includes("price")}><input type="number" value={form.cavPrice || ""} onChange={e => set("cavPrice", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label={`${form.cavTherapist} Tip ($)`}><input type="number" value={form.cavTip || ""} onChange={e => set("cavTip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label="Treatment Total ($)" error={errors.includes("price")}><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} /></Field>
                    <Field label="Tip Total ($)"><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                )}
              </div>
            )}

            {/* Extra service fee + tip — available for both ticket redemption and same-day
                purchase. Needed because redemption's main treatment/tip fields are a payroll-
                reference value only (the session was already paid for) and aren't counted as
                today's revenue — an add-on service actually paid for today (e.g. a same-therapist
                "additional massage") has to go here instead, or it silently never shows up as
                real money received. */}
            {(!form.isSameDayTicket || form.useToday !== false) && (
              <div style={{ marginTop: 10, background: "#FFF8E1", borderRadius: 8, padding: 10, border: "1px solid #FFE082" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#F57F17", marginBottom: 8 }}>💝 Extra (optional) — additional payment received today</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="Extra Treatment Fee ($)">
                    <input type="number" value={form.extraPrice || ""} onChange={e => set("extraPrice", e.target.value)} style={inputStyle} placeholder="e.g. 27" />
                  </Field>
                  <Field label="Extra Tip ($)">
                    <input type="number" value={form.extraTip || ""} onChange={e => set("extraTip", e.target.value)} style={inputStyle} placeholder="e.g. 20" />
                  </Field>
                </div>
                {(Number(form.extraPrice||0) > 0) && (
                  <div style={{ marginTop: 8 }}>
                    <Field label="Extra Treatment Fee Payment Method" error={errors.includes("extraPricePaymentType")}><PaymentToggle value={form.extraPricePaymentType} onChange={v => set("extraPricePaymentType", v)} /></Field>
                  </div>
                )}
                {(form.extraTip > 0) && (
                  <div style={{ marginTop: 8 }}>
                    <Field label="Tip Payment Method" error={errors.includes("extraTipPaymentType")}><PaymentToggle value={form.extraTipPaymentType} onChange={v => set("extraTipPaymentType", v)} /></Field>
                  </div>
                )}
                <div style={{ fontSize: 11, color: "#888", marginTop: 6 }}>For retail purchases, use the Retail section below</div>
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
                    Treatment <strong style={{ color: "#fff" }}>${svc}</strong>　Tip <strong style={{ color: "#fff" }}>${tip}</strong>
                    {extra > 0 && <span>　Extra <strong style={{ color: "#FFE082" }}>${extra}{form.extraTipPaymentType==="card"?"💳":"💵"}</strong></span>}
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ color: "#B2EBF2", fontSize: 11 }}>Total　</span>
                    <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip + extra}</span>
                  </div>
                </div>
              );
            })()}

            {/* Same-day purchase payment */}
            {form.isSameDayTicket && (
              <div style={{ marginTop: 10, background: "#E8F5E9", borderRadius: 8, padding: 10, border: "1px solid #A5D6A7" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#2E7D32", marginBottom: 8 }}>
                  💰 Same-Day Purchase — Package Price (counted as revenue)
                  {form.useToday === false && <span style={{ color: "#888", fontWeight: 400 }}>(to be used starting next visit)</span>}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="Package Price ($)" error={errors.includes("packagePrice")}>
                    <input type="number" value={form.packagePrice || ""} onChange={e => set("packagePrice", e.target.value)} style={{ ...inputStyle, ...(errors.includes("packagePrice") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="e.g. 432" disabled={form.packageSplitPayment} />
                  </Field>
                  <Field label={form.useToday === false ? "Tip ($)" : "Same-Day Tip ($)"}>
                    <input type="number" value={form.packageTip || ""} onChange={e => set("packageTip", e.target.value)} style={inputStyle} placeholder="e.g. 91" />
                  </Field>
                </div>

                {/* Deposit note — packagePrice above is typed in by hand already net of any deposit; this only records that a deposit existed and when it was paid (allocation is already fixed by hand, no auto-calc) */}
                <div style={{ marginTop: 8, background: "#fff", borderRadius: 10, padding: 12, border: "1.5px solid #A5D6A7" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: "#2E7D32" }}>💰 Has a Deposit</span>
                    <button onClick={() => set("packageDepositAmount", Number(form.packageDepositAmount) > 0 ? 0 : 20)}
                      style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.packageDepositAmount) > 0 ? "#2E7D32" : "#C8E6C9", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                      {Number(form.packageDepositAmount) > 0 ? "ON ✓" : "OFF"}
                    </button>
                  </div>
                  {Number(form.packageDepositAmount) > 0 && (
                    <div style={{ marginTop: 8, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>Deposit Amount ($)</div>
                        <input type="number" value={form.packageDepositAmount || ""}
                          onFocus={e => e.target.select()}
                          onChange={e => set("packageDepositAmount", e.target.value)}
                          style={{ ...inputStyle, borderColor: "#81C784" }} placeholder="e.g. 20" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>Date Paid{errors.includes("packageDepositDate") && " ⚠️"}</div>
                        <DatePicker value={form.packageDepositDate || ""}
                          onChange={v => set("packageDepositDate", v)}
                          style={{ ...inputStyle, borderColor: errors.includes("packageDepositDate") ? "#C62828" : "#81C784", borderWidth: errors.includes("packageDepositDate") ? 2 : 1 }} />
                      </div>
                    </div>
                  )}
                </div>

                {/* Gift Card Usage — e.g. a previously-purchased balance covering part of this package */}
                <div style={{ background: "#E0F2F1", borderRadius: 10, padding: 12, marginTop: 8 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, fontSize: 13, color: "#00796B" }}>🎁 Gift Card Used (existing balance)</span>
                    <button onClick={() => set("giftCardUsed", form.giftCardUsed === 0 ? (Number(form.packagePrice) || "") : 0)}
                      style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: form.giftCardUsed !== 0 ? "#00796B" : "#B2DFDB", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                      {form.giftCardUsed !== 0 ? "ON ✓" : "OFF"}
                    </button>
                  </div>
                  {(Number(form.giftCardUsed) > 0 || form.giftCardUsed === "") && (
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: "#00796B", marginBottom: 4 }}>Amount Used ($)</div>
                      <input type="number" value={form.giftCardUsed === "" ? "" : form.giftCardUsed}
                        onChange={e => set("giftCardUsed", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="e.g. 89" />
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
                              Package ${svc}{tip > 0 ? ` + Tip $${tip}` : ""} = <strong>${total}</strong>
                            </div>
                            🎁 GC (existing balance) <strong>${Math.min(gc, total)}</strong> (not counted in today's sales)
                            {remainder > 0
                              ? <> + remaining <strong>${remainder}</strong> received in cash/card</>
                              : <span style={{ color: "#2E7D32" }}> → Fully covered by gift card</span>}
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 10, background: "#fff", borderRadius: 8, padding: 10, border: "1.5px solid #A5D6A7" }}>
                  <div style={{ fontWeight: 700, fontSize: 12, color: "#2E7D32", marginBottom: 8 }}>💰 Payment</div>
                  <button onClick={() => set("packageSplitPayment", !form.packageSplitPayment)}
                    style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.packageSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.packageSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.packageSplitPayment ? "#2E7D32" : "#888" }}>
                    {form.packageSplitPayment ? "☑" : "☐"} Split treatment payment between cash and card
                  </button>
                  {form.packageSplitPayment ? (
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment Cash ($)</div>
                        <input type="number" value={form.packageCashPortion || ""} onChange={e => {
                          const raw = e.target.value;
                          setForm(f => ({ ...f, packageCashPortion: raw, packagePrice: (Number(raw) || 0) + Number(f.packageCardPortion||0) }));
                        }} style={inputStyle} placeholder="e.g. 300" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment Card ($)</div>
                        <input type="number" value={form.packageCardPortion || ""} onChange={e => {
                          const raw = e.target.value;
                          setForm(f => ({ ...f, packageCardPortion: raw, packagePrice: Number(f.packageCashPortion||0) + (Number(raw) || 0) }));
                        }} style={inputStyle} placeholder="e.g. 132" />
                      </div>
                    </div>
                  ) : (
                    <div style={{ marginBottom: 8 }}>
                      <Field label="Treatment Payment Method" error={errors.includes("packagePaymentType")}><PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} /></Field>
                    </div>
                  )}
                  <Field label="Tip Payment Method" error={errors.includes("packageTipPaymentType")}><PaymentToggle value={form.tipPaymentType} onChange={v => set("tipPaymentType", v)} /></Field>
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
              <Field label="Machine Therapist (optional)">
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
                  <option value="">None (no machine)</option>
                  {CAV_CAPABLE.filter(t => t !== form.therapist).map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
            )}

            {/* Gift Card Usage — hidden for GC redemption (entire appointment IS the GC) and complimentary PR */}
            {!form.isGiftCard && !form.isPromo && <div style={{ background: "#E0F2F1", borderRadius: 10, padding: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 700, fontSize: 13, color: "#00796B" }}>🎁 Gift Card Used</span>
                <button onClick={() => set("giftCardUsed", form.giftCardUsed === 0 ? ((Number(form.price)||0) + (Number(form.cavPrice)||0) || "") : 0)}
                  style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: form.giftCardUsed !== 0 ? "#00796B" : "#B2DFDB", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                  {form.giftCardUsed !== 0 ? "ON ✓" : "OFF"}
                </button>
              </div>
              {(Number(form.giftCardUsed) > 0 || form.giftCardUsed === "") && (
                <div style={{ marginTop: 8 }}>
                  <div style={{ fontSize: 10, color: "#00796B", marginBottom: 4 }}>Amount Used ($)</div>
                  <input type="number" value={form.giftCardUsed === "" ? "" : form.giftCardUsed}
                    onChange={e => set("giftCardUsed", e.target.value)}
                    style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="e.g. 150" />
                </div>
              )}
            </div>}

            {/* Deposit Used */}
            {!form.isGiftCard && !form.isPromo && (
              <div style={{ background: "#E8F5E9", borderRadius: 10, padding: 12, border: "1.5px solid #A5D6A7" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 700, fontSize: 13, color: "#2E7D32" }}>💰 Deposit Used</span>
                  <button onClick={() => set("depositApplied", Number(form.depositApplied) > 0 ? 0 : 20)}
                    style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.depositApplied) > 0 ? "#2E7D32" : "#C8E6C9", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                    {Number(form.depositApplied) > 0 ? "ON ✓" : "OFF"}
                  </button>
                </div>
                {Number(form.depositApplied) > 0 && (
                  <div style={{ marginTop: 8, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <div>
                      <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>Deposit Amount ($)</div>
                      <input type="number" value={form.depositApplied || ""}
                        onFocus={e => e.target.select()}
                        onChange={e => set("depositApplied", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#81C784" }} placeholder="e.g. 20" />
                    </div>
                    <div>
                      <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>Date Paid{errors.includes("depositPaidDate") && " ⚠️"}</div>
                      <DatePicker value={form.depositPaidDate || ""}
                        onChange={v => set("depositPaidDate", v)}
                        style={{ ...inputStyle, borderColor: errors.includes("depositPaidDate") ? "#C62828" : "#81C784", borderWidth: errors.includes("depositPaidDate") ? 2 : 1 }} />
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Total service + tip input */}
            <div style={{ background: "#F9F9F9", borderRadius: 10, padding: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "#333", marginBottom: 10 }}>
                💆 Treatment Price / Tip (enter total)
                {Number(form.depositApplied) > 0 && <span style={{ fontSize: 11, color: "#2E7D32", fontWeight: 600, marginLeft: 8 }}>← Enter the amount actually received</span>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <Field label="Treatment Total ($)" error={errors.includes("price")}>
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
                    style={{ ...inputStyle, ...(errors.includes("price") ? { borderColor: "#C62828", borderWidth: 2 } : {}) }} placeholder="e.g. 158" />
                </Field>
                <Field label="Tip Total ($)">
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
                      // Same fix as Treatment Total: keep the raw typed text so a trailing "." mid-decimal
                      // isn't wiped out by the controlled re-render before it's fully typed.
                      set("totalTipInput", e.target.value);
                    }}
                    style={inputStyle} placeholder="e.g. 30" />
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
                        🎁 Treatment <strong style={{ color: "#fff" }}>${svc}</strong>　Tip <strong style={{ color: "#fff" }}>${tip}</strong>
                        {form.cavTherapist && !isDualLicense(form.therapist) && (
                          <span style={{ fontSize: 10, color: "#FDE68A", display: "block" }}>
                            {form.therapist} ${form.price||0}{Number(form.tip||0) > 0 ? ` +Tip $${form.tip}` : ""} / {form.cavTherapist} ${form.cavPrice||0}{Number(form.cavTip||0) > 0 ? ` +Tip $${form.cavTip}` : ""}
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#FDE68A", fontSize: 11 }}>GC Redemption Total　</span>
                        <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip}</span>
                      </div>
                    </div>
                  );
                }
                if (form.isPromo) {
                  return (
                    <div style={{ marginTop: 10, background: "#0D47A1", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontSize: 11, color: "#BBDEFB" }}>
                        📸 Treatment <strong style={{ color: "#fff" }}>${svc}</strong>　Tip <strong style={{ color: "#fff" }}>${tip}</strong>
                        {form.cavTherapist && !isDualLicense(form.therapist) && (
                          <span style={{ fontSize: 10, color: "#BBDEFB", display: "block" }}>
                            {form.therapist} ${form.price||0}{Number(form.tip||0) > 0 ? ` +Tip $${form.tip}` : ""} / {form.cavTherapist} ${form.cavPrice||0}{Number(form.cavTip||0) > 0 ? ` +Tip $${form.cavTip}` : ""}
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#BBDEFB", fontSize: 11 }}>Complimentary PR Total　</span>
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
                        Treatment <strong style={{ color: "#fff" }}>${svc}{gcSvc >= svc && svc > 0 ? "🎁" : form.paymentType==="card"?"💳":"💵"}</strong>　Tip <strong style={{ color: "#fff" }}>${tip}{gcTip >= tip && tip > 0 ? "🎁" : form.tipPaymentType==="card"?"💳":"💵"}</strong>
                        {gc > 0 && (
                          <span style={{ fontSize: 10, color: "#80CBC4", display: "block" }}>
                            🎁 ${gcSvc + gcTip} covered by gift card (not counted in today's sales)
                          </span>
                        )}
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <span style={{ color: "#B2EBF2", fontSize: 11 }}>{gc > 0 ? "Received Today　" : "Amount Received　"}</span>
                        <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${gc > 0 ? receivedToday : r2(svc + tip)}</span>
                      </div>
                    </div>
                    {form.cavTherapist && !isDualLicense(form.therapist) && (
                      <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <div style={{ flex: 1, minWidth: 130, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                          <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.therapist}</div>
                          <div style={{ fontSize: 13, color: "#E0F7FA" }}>Treatment ${bodySvcShown}　Tip ${bodyTipShown}</div>
                          <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>Total ${r2(bodySvcShown + bodyTipShown)}</div>
                        </div>
                        <div style={{ flex: 1, minWidth: 130, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                          <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.cavTherapist} (Machine)</div>
                          <div style={{ fontSize: 13, color: "#E0F7FA" }}>Treatment ${cavSvcShown}　Tip ${cavTipShown}</div>
                          <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>Total ${r2(cavSvcShown + cavTipShown)}</div>
                        </div>
                      </div>
                    )}
                    {/* Single therapist (no machine helper) with a deposit — the deposit-inclusive
                        payroll figure otherwise has nowhere to show, since the two-card layout
                        above only renders when there's a machine therapist to split against. */}
                    {!form.cavTherapist && dep > 0 && (
                      <div style={{ marginTop: 10, background: "rgba(255,255,255,0.14)", borderRadius: 6, padding: "6px 10px" }}>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>{form.therapist}</div>
                        <div style={{ fontSize: 13, color: "#E0F7FA" }}>Treatment ${bodyPayroll}　Tip ${bodyTipShown}</div>
                        <div style={{ fontSize: 13, color: "#fff", fontWeight: 700 }}>Total ${r2(bodyPayroll + bodyTipShown)}</div>
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
                    <Field label={gc > 0 && !svcCovered ? "Treatment (remaining balance) Payment Method" : "Treatment Payment Method"} error={!svcCovered && errors.includes("paymentType")}>
                      {svcCovered ? (
                        <div style={{ fontSize: 12, color: "#00796B", fontWeight: 700, padding: "10px 0" }}>🎁 Paid with gift card</div>
                      ) : (
                        <>
                          <button onClick={() => set("svcSplitPayment", !form.svcSplitPayment)}
                            style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.svcSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.svcSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.svcSplitPayment ? "#2E7D32" : "#888" }}>
                            {form.svcSplitPayment ? "☑" : "☐"} Split payment between cash and card
                          </button>
                          {form.svcSplitPayment ? (
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 Cash ($)</div>
                                <input type="number" value={form.svcCashPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, svcCashPortion: raw, svcCardPortion: Math.max(0, Number(f.price || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="e.g. 50" />
                              </div>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 Card ($)</div>
                                <input type="number" value={form.svcCardPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, svcCardPortion: raw, svcCashPortion: Math.max(0, Number(f.price || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="e.g. 50" />
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
                    <Field label="Tip Payment Method" error={!tipCovered && errors.includes("tipPaymentType")}>
                      {tipCovered ? (
                        <div style={{ fontSize: 12, color: "#00796B", fontWeight: 700, padding: "6px 0" }}>🎁 Paid with gift card</div>
                      ) : (
                        <>
                          <button onClick={() => set("tipSplitPayment", !form.tipSplitPayment)}
                            style={{ marginBottom: 8, padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.tipSplitPayment ? "#2E7D32" : "#DDD"}`, background: form.tipSplitPayment ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.tipSplitPayment ? "#2E7D32" : "#888" }}>
                            {form.tipSplitPayment ? "☑" : "☐"} Split payment between cash and card
                          </button>
                          {form.tipSplitPayment ? (
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 Cash ($)</div>
                                <input type="number" value={form.tipCashPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, tipCashPortion: raw, tipCardPortion: Math.max(0, Number(f.tip || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="e.g. 10" />
                              </div>
                              <div>
                                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 Card ($)</div>
                                <input type="number" value={form.tipCardPortion || ""} onChange={e => {
                                  const raw = e.target.value;
                                  setForm(f => ({ ...f, tipCardPortion: raw, tipCashPortion: Math.max(0, Number(f.tip || 0) - (Number(raw) || 0)) }));
                                }} style={inputStyle} placeholder="e.g. 20" />
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
                      🔄 Recalculate using Weight Loss rules ({priceStale ? "not currently $116" : "tip % is off after the deposit changed"})
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
                      🔄 Recalculate using the minute ratio ({bodyMins}min / {cavMins}min) ({priceStale ? "treatment price" : "tip"} is off)
                    </button>
                  </div>
                );
              })()}

              {/* Minutes split. Weight Loss always uses a fixed 40min machine / rest-to-body split
                  (auto-filled elsewhere, no manual entry needed) — this UI is only for everything
                  else, where minutes actually drive the $ split. */}
              {form.cavTherapist && !isRegularWeightLoss && (
                <div style={{ marginTop: 10, background: "#EEF4FF", borderRadius: 8, padding: 10 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>⏱️ Enter each therapist's minutes → auto-split</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label={`${form.therapist} (min)`}>
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
                        style={inputStyle} placeholder={`e.g. ${form.duration - (isRegularWeightLoss ? 40 : 15)}`} />
                    </Field>
                    <Field label={`${form.cavTherapist} (min)`}>
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
                        style={inputStyle} placeholder={isRegularWeightLoss ? "e.g. 40" : "e.g. 15"} />
                    </Field>
                  </div>

                </div>
              )}

            </div>

            {/* Gift Card Redemption banner — replaces payment type selectors */}
            {form.isGiftCard && (
              <div style={{ background: "#FFFDE7", borderRadius: 10, padding: 12, border: "2px solid #F59E0B", marginTop: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 13, color: "#B45309" }}>🎁 GC Redemption — Paid in full with gift card</div>
                <div style={{ fontSize: 11, color: "#92400E", marginTop: 4 }}>Uses a previously-purchased gift card. Not counted in today's sales.</div>
              </div>
            )}

            {/* Complimentary PR banner — replaces payment type selectors */}
            {form.isPromo && (
              <div style={{ background: "#E3F2FD", borderRadius: 10, padding: 12, border: "2px solid #1565C0", marginTop: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 13, color: "#1565C0" }}>📸 Complimentary PR — free treatment for influencers, etc.</div>
                <div style={{ fontSize: 11, color: "#0D47A1", marginTop: 4 }}>No payment was received from the client. Not counted in today's sales, but the entered amount is still reflected in the therapist's treatment record for payroll purposes.</div>
              </div>
            )}


          </>
        )}

        {/* ── Same-Day Additional Purchase ── */}
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
                📌 Same-Day Additional Purchase (tap to toggle)
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

              {/* 🎟️ New Ticket Purchase fields */}
              {hasTag("newTicket") && (
                <div style={{ background: "#FFEBEE", padding: "10px 12px", borderTop: "1px solid #FFCDD2", display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#B71C1C" }}>🎟️ New Ticket Purchase — Details</div>

                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Course</div>
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
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Duration</div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {group.durations.map(dur => (
                            <button key={dur} onClick={() => applyNewTicketMenuPrice(`${prefix}-${dur}-${form.newTicketTotal || 3}`, form.newTicketTotal || 3)}
                              style={{
                                padding: "6px 12px", borderRadius: 8,
                                border: `2px solid ${currentDur === dur ? "#B71C1C" : "#DDD"}`,
                                background: currentDur === dur ? "#B71C1C" : "#fff",
                                cursor: "pointer", fontWeight: 700, fontSize: 11,
                                color: currentDur === dur ? "#fff" : "#888"
                              }}>{dur}min</button>
                          ))}
                        </div>
                      </div>
                    ) : null;
                  })()}

                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Number of Sessions</div>
                    <div style={{ display: "flex", gap: 8 }}>
                      {[3,5].map(n => (
                        <button key={n} onClick={() => {
                          set("newTicketTotal", n);
                          if (form.newTicketMenu) {
                            const parts = form.newTicketMenu.split("-");
                            applyNewTicketMenuPrice(`${parts[0]}-${parts[1]}-${n}`, n);
                          }
                        }} style={{ flex: 1, padding: "7px", borderRadius: 8, border: `2px solid ${form.newTicketTotal===n?"#B71C1C":"#DDD"}`, background: form.newTicketTotal===n?"#B71C1C":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.newTicketTotal===n?"#fff":"#888" }}>
                          {n}-session package
                        </button>
                      ))}
                    </div>
                  </div>

                  <div style={{ display: "flex", gap: 8 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment Price ($) (auto-filled, editable)</div>
                      <input type="number" value={form.newTicketAmount || ""} onChange={e => set("newTicketAmount", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="e.g. 719" disabled={form.newTicketSplitPayment} />
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Tip ($)</div>
                      <input type="number" value={form.newTicketTip || ""} onChange={e => set("newTicketTip", e.target.value)}
                        style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="e.g. 46" />
                    </div>
                  </div>
                  <div>
                    <button onClick={() => set("newTicketSplitPayment", !form.newTicketSplitPayment)}
                      style={{ padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.newTicketSplitPayment ? "#B71C1C" : "#DDD"}`, background: form.newTicketSplitPayment ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.newTicketSplitPayment ? "#B71C1C" : "#888" }}>
                      {form.newTicketSplitPayment ? "☑" : "☐"} Split payment between cash and card
                    </button>
                  </div>
                  {form.newTicketSplitPayment ? (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment price breakdown (cash + card = total treatment price)</div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 Cash ($)</div>
                          <input type="number" value={form.newTicketCashPortion || ""} onChange={e => {
                            const raw = e.target.value;
                            setForm(f => ({ ...f, newTicketCashPortion: raw, newTicketAmount: (Number(raw) || 0) + Number(f.newTicketCardPortion||0) }));
                          }} style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="e.g. 500" />
                        </div>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 Card ($)</div>
                          <input type="number" value={form.newTicketCardPortion || ""} onChange={e => {
                            const raw = e.target.value;
                            setForm(f => ({ ...f, newTicketCardPortion: raw, newTicketAmount: Number(f.newTicketCashPortion||0) + (Number(raw) || 0) }));
                          }} style={{ ...inputStyle, borderColor: "#EF9A9A" }} placeholder="e.g. 211" />
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Treatment price payment method</div>
                      {payBtns("newTicketPaymentType", true, "#B71C1C")}
                    </div>
                  )}
                  {Number(form.newTicketTip) > 0 && (
                    <div>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Tip payment method</div>
                      {payBtns("newTicketTipPaymentType", true, "#E65100")}
                    </div>
                  )}
                  {(Number(form.newTicketAmount) > 0 || Number(form.newTicketTip) > 0) && (
                    <div style={{ textAlign: "right", fontSize: 12, fontWeight: 700, color: "#B71C1C" }}>
                      Total ${(Number(form.newTicketAmount||0) + Number(form.newTicketTip||0)).toFixed(0)}
                    </div>
                  )}
                </div>
              )}

              {/* 🛍️ Retail Purchase fields — supports multiple products in one visit (2nd/3rd item stored in extraRetailItems[]) */}
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
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A" }}>🛍️ Retail Purchase — Details</div>
                    {items.map((item, idx) => (
                      <div key={idx} style={{ display: "flex", flexDirection: "column", gap: 8, ...(idx > 0 ? { borderTop: "1px dashed #CE93D8", paddingTop: 8 } : {}) }}>
                        {idx > 0 && (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: 10, fontWeight: 700, color: "#6A1B9A" }}>Item {idx + 1}</span>
                            <button onClick={() => removeItem(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11, color: "#AAA" }}>✕ Remove</button>
                          </div>
                        )}
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Product Name</div>
                          <select value={RETAIL_PRODUCTS.find(p => p.name === item.productName) ? item.productName : (item.productName ? "__other__" : "")} onChange={e => {
                            const val = e.target.value;
                            const prod = RETAIL_PRODUCTS.find(p => p.name === val);
                            updateItem(idx, { productName: val === "__other__" ? "" : val, amount: prod?.price > 0 ? prod.price : item.amount });
                          }} style={{ ...inputStyle, borderColor: "#CE93D8" }}>
                            <option value="">— Select a product —</option>
                            {RETAIL_PRODUCTS.map(p => (
                              <option key={p.name} value={p.name}>{RETAIL_PRODUCT_LABELS[p.name] || p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
                            ))}
                            <option value="__other__">Other</option>
                          </select>
                          {(!RETAIL_PRODUCTS.find(p => p.name === item.productName)) && (
                            <input type="text" value={item.productName || ""} placeholder="Enter product name" style={{ ...inputStyle, borderColor: "#CE93D8", marginTop: 4 }}
                              onChange={e => updateItem(idx, { productName: e.target.value })} />
                          )}
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Amount ($)</div>
                          <input type="number" value={item.amount || ""} onChange={e => updateItem(idx, { amount: e.target.value })}
                            style={{ ...inputStyle, borderColor: "#CE93D8" }} placeholder="e.g. 30" />
                        </div>
                        <div>
                          <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Payment Method</div>
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
                          // base (same convention as the standalone Retail modal) — staff still work out
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
                                Split between (up to 3 people, splitting the tax-excluded amount of ${afterTaxTotal}){afterTaxTotal > 0 && Math.abs(sellersTotal - afterTaxTotal) > 0.15 ? " ⚠️ The total is significantly off from the tax-excluded amount" : ""}
                              </div>
                              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                                {afterTaxTotal > 0 && (
                                  <button onClick={() => {
                                    const each = Math.round((afterTaxTotal / sellers.length) * 100) / 100;
                                    updateItem(idx, { sellers: sellers.map(sel => ({ ...sel, amount: each })) });
                                  }} style={{ padding: "5px 10px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#fff", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 11, alignSelf: "flex-start" }}>
                                    ⚖️ {sellers.length > 1 ? `Split the tax-excluded amount evenly ${sellers.length} ways` : "Auto-fill the tax-excluded amount"}
                                  </button>
                                )}
                                {sellers.map((sel, sidx) => (
                                  <div key={sidx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                    <select value={sel.therapist} onChange={e => updSeller(sidx, { therapist: e.target.value })} style={{ ...inputStyle, flex: 2, borderColor: "#CE93D8" }}>
                                      <option value="">— Select —</option>
                                      {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                                    </select>
                                    <input type="number" value={sel.amount || ""} onChange={e => updSeller(sidx, { amount: e.target.value })}
                                      style={{ ...inputStyle, flex: 1, borderColor: "#CE93D8" }} placeholder="Amount" />
                                    {sellers.length > 1 && (
                                      <button onClick={() => updateItem(idx, { sellers: sellers.filter((_, i) => i !== sidx) })}
                                        style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA" }}>✕</button>
                                    )}
                                  </div>
                                ))}
                                {sellers.length < 3 && (
                                  <button onClick={() => updateItem(idx, { sellers: [...sellers, { therapist: "", amount: 0 }] })}
                                    style={{ padding: "5px 10px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#F3E5F5", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 11, alignSelf: "flex-start" }}>
                                    ＋ Add a seller
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
                      ＋ Add product (2nd, 3rd, ...)
                    </button>
                  </div>
                );
              })()}

              {/* 🎁 Gift Card Purchase fields */}
              {hasTag("giftCard") && (
                <div style={{ background: "#E0F2F1", padding: "10px 12px", borderTop: "1px solid #B2DFDB", display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#00796B" }}>🎁 Gift Card Purchase — Details</div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Amount ($)</div>
                    <input type="number" value={form.giftCardPurchaseAmount || ""} onChange={e => set("giftCardPurchaseAmount", e.target.value)}
                      style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="e.g. 100" />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>Payment Method</div>
                    {payBtns("giftCardPurchasePaymentType", true, "#00796B")}
                  </div>
                </div>
              )}
            </div>
          );
        })()}

        <Field label="Memo">
          <input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="Additional notes" />
        </Field>
      </div>

      <ErrorBanner />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#6A1B9A" }}>🛍️ Retail</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Product Name">
          <select value={RETAIL_PRODUCTS.find(p => p.name === form.item) ? form.item : (form.item ? "__other__" : "")}
            onChange={e => {
              const val = e.target.value;
              const prod = RETAIL_PRODUCTS.find(p => p.name === val);
              setForm(f => ({ ...f, item: val === "__other__" ? "" : val, price: prod?.price > 0 ? prod.price : f.price }));
            }} style={inputStyle}>
            <option value="">— Select a product —</option>
            {RETAIL_PRODUCTS.map(p => (
              <option key={p.name} value={p.name}>{RETAIL_PRODUCT_LABELS[p.name] || p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
            ))}
            <option value="__other__">Other (custom entry)</option>
          </select>
          {(!RETAIL_PRODUCTS.find(p => p.name === form.item)) && (
            <input type="text" value={form.item || ""} placeholder="Enter product name" style={{ ...inputStyle, marginTop: 4 }}
              onChange={e => set("item", e.target.value)} />
          )}
        </Field>
        <Field label="Amount ($)"><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={inputStyle} /></Field>
        <Field label={`Split between (up to 3 people, splitting the tax-excluded amount of $${afterTaxTotal})${Number(form.price) > 0 && Math.abs(sellersTotal - afterTaxTotal) > 0.15 ? " ⚠️ The total is significantly off from the tax-excluded amount" : ""}`}>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Number(form.price) > 0 && (
              <button onClick={() => {
                // Rounded to the nearest $0.10 — penny-level precision isn't needed for this split.
                const each = Math.round((afterTaxTotal / sellers.length) * 100) / 100;
                set("sellers", sellers.map(sel => ({ ...sel, amount: each })));
              }} style={{ padding: "6px 12px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#fff", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 12, alignSelf: "flex-start" }}>
                ⚖️ {sellers.length > 1 ? `Split the tax-excluded amount evenly ${sellers.length} ways` : "Auto-fill the tax-excluded amount"}
              </button>
            )}
            {sellers.map((sel, idx) => (
              <div key={idx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <select value={sel.therapist} onChange={e => updSeller(idx, { therapist: e.target.value })} style={{ ...inputStyle, flex: 2 }}>
                  <option value="">— Select —</option>
                  {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
                <input type="number" value={sel.amount || ""} onChange={e => updSeller(idx, { amount: e.target.value })}
                  style={{ ...inputStyle, flex: 1 }} placeholder="Amount" />
                {sellers.length > 1 && (
                  <button onClick={() => set("sellers", sellers.filter((_, i) => i !== idx))}
                    style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA" }}>✕</button>
                )}
              </div>
            ))}
            {sellers.length < 3 && (
              <button onClick={() => set("sellers", [...sellers, { therapist: "", amount: 0 }])}
                style={{ padding: "6px 12px", borderRadius: 8, border: "2px solid #6A1B9A", background: "#F3E5F5", color: "#6A1B9A", fontWeight: 700, cursor: "pointer", fontSize: 12, alignSelf: "flex-start" }}>
                ＋ Add a seller
              </button>
            )}
          </div>
        </Field>
        <Field label="Payment Method" error={paymentError && !form.paymentType}><PaymentToggle value={form.paymentType} onChange={v => { set("paymentType", v); setPaymentError(false); }} /></Field>
      </div>
      <button onClick={handleSave} style={{ width: "100%", marginTop: 16, padding: "12px", borderRadius: 10, border: "none", background: "#6A1B9A", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
    </Modal>
  );
}

function DepositModal({ deposit, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...deposit });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const isDeposit = form.type === "deposit";
  const isGiftCard = form.type === "giftcard";
  const isCancellation = form.type === "cancellation";
  const payTypes = [["cash","💵 Cash"],["card","💳 Card"]];
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#1565C0" }}>{isCancellation ? "❌ Cancellation Fee" : isGiftCard ? "🎁 Gift Card" : "💰 Deposit"}</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Type">
          <div style={{ display: "flex", gap: 8 }}>
            {[["deposit","💰 Deposit"],["giftcard","🎁 Gift Card"],["cancellation","❌ Cancellation Fee"]].map(([val,label]) => (
              <button key={val} onClick={() => set("type", val)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${form.type===val?"#1565C0":"#DDD"}`, background: form.type===val?"#E3F2FD":"#fff", cursor: "pointer", fontWeight: 700, color: form.type===val?"#1565C0":"#888" }}>{label}</button>
            ))}
          </div>
        </Field>
        <Field label="Client Name"><input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="e.g. Tanaka" /></Field>
        {!isCancellation && (
          <Field label={`Scheduled Visit Date/Time (optional${isGiftCard ? " — e.g. if fully prepaid by gift card" : ""})`}>
            <div style={{ display: "flex", gap: 8 }}>
              <DatePicker value={form.appointmentDate || ""} onChange={v => set("appointmentDate", v)} style={{ ...inputStyle, flex: 1 }} />
              <input type="time" value={form.appointmentTime || ""} onChange={e => set("appointmentTime", e.target.value)} style={{ ...inputStyle, width: 100 }} />
            </div>
            {form.appointmentDate && form.clientName && (
              <div style={{ fontSize: 11, color: "#1565C0", marginTop: 4 }}>
                {isGiftCard ? "Gift prepayment" : "Deposit"} for {form.clientName} on {new Date(form.appointmentDate + "T00:00").toLocaleDateString("en-US", { month: "long", day: "numeric" })}
                {form.appointmentTime && ` ${form.appointmentTime}`}
                　※The date can be changed later (edit here if the plan changes)
              </div>
            )}
          </Field>
        )}
        <Field label={isCancellation ? "Cancellation Fee ($)" : "Amount ($)"}><input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} /></Field>
        <div style={{ background: "#F0F4FF", borderRadius: 10, padding: "10px 12px" }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#1565C0", marginBottom: 8 }}>💳 Payment Method (treatment and tip can be set separately)</div>
          <Field label={isCancellation ? "Cancellation Fee Payment Method" : isDeposit ? "Deposit Payment Method" : "Amount Payment Method"} error={errors.includes("paymentType")}>
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
            <Field label="Tip ($) (optional)">
              <input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} placeholder="Leave blank if no tip" />
            </Field>
          </div>
          {Number(form.tip) > 0 && (
            <div style={{ marginTop: 8 }}>
              <Field label="Tip Payment Method" error={errors.includes("tipPaymentType")}>
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
        <Field label="Memo"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
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
    const packageName = group ? `${group.group} ${durLabel}min x${total}` : form.packageName;
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#B71C1C" }}>🎟️ New Ticket Purchase</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Client Name">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="e.g. Mio" />
        </Field>

        <div style={{ background: "#FFF5F5", borderRadius: 10, padding: 12 }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#B71C1C", marginBottom: 8 }}>🎟️ Course Selection</div>

          <Field label="Course">
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
              <Field label="Duration">
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                  {group.durations.map(dur => (
                    <button key={dur} onClick={() => applyMenuPrice(`${prefix}-${dur}-${form.ticketTotal || 3}`, form.ticketTotal || 3)}
                      style={{
                        padding: "7px 12px", borderRadius: 8,
                        border: `2px solid ${currentDur === dur ? "#B71C1C" : "#DDD"}`,
                        background: currentDur === dur ? "#B71C1C" : "#fff",
                        cursor: "pointer", fontWeight: 700, fontSize: 12,
                        color: currentDur === dur ? "#fff" : "#888"
                      }}>{dur}min</button>
                  ))}
                </div>
              </Field>
            ) : null;
          })()}

          <Field label="Number of Sessions">
            <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
              {[3,5].map(n => (
                <button key={n} onClick={() => {
                  set("ticketTotal", n);
                  if (form.ticketMenu) {
                    const parts = form.ticketMenu.split("-");
                    applyMenuPrice(`${parts[0]}-${parts[1]}-${n}`, n);
                  }
                }} style={{ flex: 1, padding: "8px", borderRadius: 8, border: `2px solid ${form.ticketTotal===n?"#B71C1C":"#DDD"}`, background: form.ticketTotal===n?"#B71C1C":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketTotal===n?"#fff":"#888" }}>
                  {n}-session package
                </button>
              ))}
            </div>
          </Field>
        </div>

        <Field label="Package Name (auto-filled, editable)">
          <input value={form.packageName} onChange={e => set("packageName", e.target.value)} style={inputStyle} placeholder="e.g. Improving Posture 90min x3" />
        </Field>
        <Field label="Treatment Price ($)">
          <input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} placeholder="e.g. 719" disabled={form.splitPayment} />
        </Field>
        <div>
          <button onClick={() => set("splitPayment", !form.splitPayment)}
            style={{ padding: "6px 12px", borderRadius: 8, border: `2px solid ${form.splitPayment ? "#B71C1C" : "#DDD"}`, background: form.splitPayment ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.splitPayment ? "#B71C1C" : "#888" }}>
            {form.splitPayment ? "☑" : "☐"} Split payment between cash and card
          </button>
        </div>
        {form.splitPayment ? (
          <Field label="Breakdown (cash + card = total treatment price)">
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <div>
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💵 Cash ($)</div>
                <input type="number" value={form.cashPortion || ""} onChange={e => {
                  const raw = e.target.value;
                  setForm(f => ({ ...f, cashPortion: raw, amount: (Number(raw) || 0) + Number(f.cardPortion||0) }));
                }} style={inputStyle} placeholder="e.g. 500" />
              </div>
              <div>
                <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>💳 Card ($)</div>
                <input type="number" value={form.cardPortion || ""} onChange={e => {
                  const raw = e.target.value;
                  setForm(f => ({ ...f, cardPortion: raw, amount: Number(f.cashPortion||0) + (Number(raw) || 0) }));
                }} style={inputStyle} placeholder="e.g. 211" />
              </div>
            </div>
            <div style={{ marginTop: 6, fontSize: 12, color: "#B71C1C", fontWeight: 700 }}>
              Total ${Number(form.cashPortion||0) + Number(form.cardPortion||0)}
            </div>
          </Field>
        ) : (
          <Field label="Payment Method" error={errors.includes("paymentType")}>
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
        <Field label="Tip ($)">
          <input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} placeholder="e.g. 46" />
        </Field>
        {Number(form.tip) > 0 && (
          <Field label="Tip Payment Method" error={errors.includes("tipPaymentType")}>
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
        <Field label="Memo">
          <input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="e.g. Purchased after final session used" />
        </Field>
        {/* Total preview */}
        {(Number(form.amount) > 0 || Number(form.tip) > 0) && (
          <div style={{ background: "#B71C1C", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ fontSize: 11, color: "#FFCDD2" }}>
              Treatment <strong style={{ color: "#fff" }}>${Number(form.amount||0).toFixed(0)}</strong>
              {Number(form.tip) > 0 && <span>　Tip <strong style={{ color: "#FFCC80" }}>${Number(form.tip||0).toFixed(0)}</strong></span>}
            </div>
            <div>
              <span style={{ color: "#FFCDD2", fontSize: 11 }}>Total　</span>
              <span style={{ color: "#fff", fontSize: 20, fontWeight: 800 }}>${(Number(form.amount||0) + Number(form.tip||0)).toFixed(0)}</span>
            </div>
          </div>
        )}
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#B71C1C", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#37474F" }}>👩‍💼 Staff Purchase</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Staff Name">
          <select value={form.staffName} onChange={e => set("staffName", e.target.value)} style={inputStyle}>
            <option value="">— Select —</option>
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
                  <span style={{ fontSize: 11, fontWeight: 700, color: "#37474F" }}>Item {idx + 1}</span>
                  <button onClick={() => removeItem(idx)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 11, color: "#AAA" }}>✕ Remove</button>
                </div>
              )}
              <Field label="Product / Treatment Name">
                <select value={isOtherProduct ? "__other__" : (item.productName || "")} onChange={e => {
                  const val = e.target.value;
                  if (val === "__other__") { updateItem(idx, { productName: "" }); return; }
                  const prod = RETAIL_PRODUCTS.find(p => p.name === val);
                  updateItem(idx, { productName: val, amount: prod?.price > 0 ? prod.price : item.amount });
                }} style={inputStyle}>
                  <option value="">— Select a product/treatment —</option>
                  {RETAIL_PRODUCTS.map(p => (
                    <option key={p.name} value={p.name}>{RETAIL_PRODUCT_LABELS[p.name] || p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
                  ))}
                  <option value="__other__">Other (enter treatment name, etc. directly)</option>
                </select>
                {isOtherProduct && (
                  <input type="text" value={item.productName || ""} placeholder="e.g. ProCell" style={{ ...inputStyle, marginTop: 4 }}
                    onChange={e => updateItem(idx, { productName: e.target.value })} />
                )}
              </Field>
              <Field label="Amount ($)"><input type="number" value={item.amount || ""} onChange={e => updateItem(idx, { amount: e.target.value })} style={inputStyle} /></Field>
              <Field label="Payment Method" error={paymentError && Number(item.amount || 0) > 0 && !item.paymentType}>
                <div style={{ display: "flex", gap: 6 }}>
                  {[["cash","💵 Cash"],["card","💳 Card"]].map(([val, label]) => (
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
          ＋ Add product (2nd, 3rd, ...)
        </button>
        {total > 0 && (
          <div style={{ textAlign: "right", fontSize: 13, fontWeight: 700, color: "#37474F" }}>Total ${total}</div>
        )}
        <Field label="Memo"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="Optional" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#37474F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#5D4037" }}>🔙 Refund</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 12, background: "#EFEBE9", borderRadius: 8, padding: "8px 10px" }}>
        A record used to deduct a refund for a past visit from "today's" sales, across day boundaries. This does not change the therapist's payroll split — if that needs adjusting, unlock the original day and edit it directly.
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Client Name">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="e.g. Tanaka" />
        </Field>
        <Field label="Visit Date (optional)">
          <DatePicker value={form.originalDate || ""} onChange={v => set("originalDate", v)} style={inputStyle} />
        </Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <Field label="Treatment Refund Amount ($)">
            <input type="number" value={form.serviceAmount || ""} onChange={e => set("serviceAmount", e.target.value)} style={inputStyle} placeholder="e.g. 150" />
          </Field>
          <Field label="Tip Refund Amount ($)">
            <input type="number" value={form.tipAmount || ""} onChange={e => set("tipAmount", e.target.value)} style={inputStyle} placeholder="e.g. 0" />
          </Field>
        </div>
        <Field label="Refund Method" error={paymentError && !form.paymentType}>
          <div style={{ display: "flex", gap: 6 }}>
            {[["cash","💵 Cash"],["card","💳 Card"]].map(([val, label]) => (
              <button key={val} onClick={() => { set("paymentType", val); setPaymentError(false); }}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#5D4037":"#DDD"}`, background: form.paymentType===val?"#EFEBE9":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#5D4037":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        {(Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && (
          <div style={{ background: "#5D4037", borderRadius: 8, padding: "10px 14px", textAlign: "right" }}>
            <span style={{ color: "#D7CCC8", fontSize: 11 }}>Total Refund　</span>
            <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>-${Number(form.serviceAmount||0) + Number(form.tipAmount||0)}</span>
          </div>
        )}
        <Field label="Memo"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="Reason for refund, etc." /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#5D4037", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
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
        <h2 style={{ margin: 0, fontSize: 18, color: "#00695C" }}>🙏 Forgotten Entry</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 12, background: "#E0F2F1", borderRadius: 8, padding: "8px 10px" }}>
        A record for a payment (usually a cash tip) staff forgot to ring in for a past visit — adds it to "today's" sales across day boundaries, and to the therapist's payroll for today. No need to edit the appointment on the original visit date (it doesn't affect customer-type counts).
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="Client Name">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="e.g. Tanaka" />
        </Field>
        <Field label="Therapist" error={errors.includes("therapist")}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {THERAPISTS.map(t => (
              <button key={t} onClick={() => { set("therapist", t); setErrors(e => e.filter(x => x !== "therapist")); }}
                style={{ padding: "7px 12px", borderRadius: 8, border: `2px solid ${form.therapist===t?"#00695C":(errors.includes("therapist")?"#C62828":"#DDD")}`, background: form.therapist===t?"#00695C":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.therapist===t?"#fff":"#888" }}>
                {t}
              </button>
            ))}
          </div>
        </Field>
        <Field label="Visit Date (optional — the date the treatment actually happened)">
          <DatePicker value={form.originalDate || ""} onChange={v => set("originalDate", v)} style={inputStyle} />
        </Field>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <Field label="Forgotten Treatment Amount ($)">
            <input type="number" value={form.serviceAmount || ""} onChange={e => set("serviceAmount", e.target.value)} style={inputStyle} placeholder="e.g. 0" />
          </Field>
          <Field label="Forgotten Tip ($)">
            <input type="number" value={form.tipAmount || ""} onChange={e => set("tipAmount", e.target.value)} style={inputStyle} placeholder="e.g. 20" />
          </Field>
        </div>
        <Field label="Payment Method" error={errors.includes("paymentType")}>
          <div style={{ display: "flex", gap: 6 }}>
            {[["cash","💵 Cash"],["card","💳 Card"]].map(([val, label]) => (
              <button key={val} onClick={() => { set("paymentType", val); setErrors(e => e.filter(x => x !== "paymentType")); }}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#00695C":"#DDD"}`, background: form.paymentType===val?"#E0F2F1":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#00695C":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        {(Number(form.serviceAmount) > 0 || Number(form.tipAmount) > 0) && (
          <div style={{ background: "#00695C", borderRadius: 8, padding: "10px 14px", textAlign: "right" }}>
            <span style={{ color: "#B2DFDB", fontSize: 11 }}>Total Added Today　</span>
            <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>+${Number(form.serviceAmount||0) + Number(form.tipAmount||0)}</span>
          </div>
        )}
        <Field label="Memo"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="e.g. Forgot to ring in on the 16th" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={handleSave} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#00695C", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 Save</button>
        {ft.id && <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>}
      </div>
    </Modal>
  );
}

function AutoBackupModal({ dates, onRestore, onClose }) {
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#0D4F4F" }}>🕐 Restore from Auto-Backup</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ fontSize: 12, color: "#888", marginBottom: 12 }}>
        A snapshot of all data is saved automatically once a day. Pick a date below to restore the cloud store to that day's snapshot — this overwrites whatever's currently saved.
      </div>
      {dates.length === 0 ? (
        <div style={{ textAlign: "center", color: "#AAA", padding: 24, fontSize: 13 }}>No automatic backups yet</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: "50vh", overflowY: "auto" }}>
          {dates.map(date => (
            <div key={date} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: "#F5F5F5", borderRadius: 8 }}>
              <span style={{ fontWeight: 600, fontSize: 13 }}>{date}</span>
              <button onClick={() => onRestore(date)}
                style={{ padding: "6px 14px", borderRadius: 8, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                Restore
              </button>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function SectionBox({ title, color, onAdd, disabled, children }) {
  return (
    <div style={{ marginTop: 20, background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <span style={{ fontWeight: 700, color, fontSize: 15 }}>{title}</span>
        <button onClick={onAdd} disabled={disabled} style={{ padding: "6px 14px", borderRadius: 8, background: disabled ? "#CCC" : color, color: "#fff", border: "none", cursor: disabled ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 600 }}>+ Add</button>
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

// Native <input type="month"> renders its picker/label using the browser's UI language, not the
// page's <html lang>, so on a browser set to Japanese it shows month names like "7月" regardless
// of the app's own translation — this custom picker guarantees English month names everywhere.
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
function MonthPicker({ value, onChange, style }) {
  const [y, m] = value.split("-").map(Number);
  return (
    <div style={{ display: "flex", gap: 6 }}>
      <select value={m} onChange={e => onChange(`${y}-${String(e.target.value).padStart(2, "0")}`)} style={style}>
        {MONTH_NAMES.map((name, i) => <option key={i} value={i + 1}>{name}</option>)}
      </select>
      <input type="number" value={y} onChange={e => onChange(`${e.target.value}-${String(m).padStart(2, "0")}`)}
        style={{ ...style, width: 80 }} />
    </div>
  );
}

const WEEKDAY_NAMES = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

// Same reasoning as MonthPicker above: native <input type="date">'s calendar popup (era labels
// like "令和", "今日"/Today, "削除"/Clear) renders in the browser's UI language, not the page's —
// this is a fully custom calendar-grid popup instead, so there's no native popup to leak the
// browser's language through, while keeping the "click a day" interaction staff are used to
// (a plain month/day/year dropdown trio works but is slower to use than picking off a calendar).
function DatePicker({ value, onChange, style, allowClear = true }) {
  const [open, setOpen] = useState(false);
  const today = new Date();
  const [y, m, d] = value ? value.split("-").map(Number) : [null, null, null];
  const [viewYear, setViewYear] = useState(y || today.getFullYear());
  const [viewMonth, setViewMonth] = useState(m || today.getMonth() + 1); // 1-12

  const openPicker = () => {
    setViewYear(y || today.getFullYear());
    setViewMonth(m || today.getMonth() + 1);
    setOpen(true);
  };

  const daysInMonth = new Date(viewYear, viewMonth, 0).getDate();
  const firstWeekday = new Date(viewYear, viewMonth - 1, 1).getDay();
  const cells = [...Array(firstWeekday).fill(null), ...Array.from({ length: daysInMonth }, (_, i) => i + 1)];

  const changeMonth = (delta) => {
    let nm = viewMonth + delta, ny = viewYear;
    if (nm < 1) { nm = 12; ny -= 1; } else if (nm > 12) { nm = 1; ny += 1; }
    setViewMonth(nm); setViewYear(ny);
  };

  const selectDay = (day) => {
    onChange(`${viewYear}-${String(viewMonth).padStart(2, "0")}-${String(day).padStart(2, "0")}`);
    setOpen(false);
  };

  const displayText = value ? `${MONTH_NAMES[m - 1].slice(0, 3)} ${d}, ${y}` : "Select a date";

  return (
    <div style={{ position: "relative", display: "inline-block" }}>
      <button type="button" onClick={() => open ? setOpen(false) : openPicker()} style={{ ...style, cursor: "pointer", whiteSpace: "nowrap" }}>
        📅 {displayText}
      </button>
      {open && (
        <>
          <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 1000 }} />
          <div style={{ position: "absolute", top: "100%", left: 0, marginTop: 4, background: "#fff", borderRadius: 10, boxShadow: "0 4px 16px rgba(0,0,0,0.25)", padding: 12, zIndex: 1001, width: 230 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <button type="button" onClick={() => changeMonth(-1)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#0D4F4F", fontWeight: 700, padding: "0 8px" }}>‹</button>
              <span style={{ fontWeight: 700, fontSize: 13, color: "#0D4F4F" }}>{MONTH_NAMES[viewMonth - 1]} {viewYear}</span>
              <button type="button" onClick={() => changeMonth(1)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#0D4F4F", fontWeight: 700, padding: "0 8px" }}>›</button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 2, marginBottom: 4 }}>
              {WEEKDAY_NAMES.map(w => <div key={w} style={{ textAlign: "center", fontSize: 10, color: "#AAA", fontWeight: 700 }}>{w}</div>)}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 2 }}>
              {cells.map((day, i) => {
                if (day == null) return <div key={i} />;
                const isSelected = value && day === d && viewMonth === m && viewYear === y;
                const isToday = day === today.getDate() && viewMonth === today.getMonth() + 1 && viewYear === today.getFullYear();
                return (
                  <button type="button" key={i} onClick={() => selectDay(day)} style={{
                    border: isToday && !isSelected ? "1.5px solid #0D4F4F" : "none", borderRadius: 6, padding: "6px 0",
                    cursor: "pointer", fontSize: 12, background: isSelected ? "#0D4F4F" : "transparent",
                    color: isSelected ? "#fff" : "#333", fontWeight: isSelected || isToday ? 700 : 400,
                  }}>{day}</button>
                );
              })}
            </div>
            {value && allowClear && (
              <button type="button" onClick={() => { onChange(""); setOpen(false); }}
                style={{ marginTop: 8, width: "100%", padding: "6px", borderRadius: 6, border: "1px solid #DDD", background: "#fff", color: "#888", fontSize: 11, cursor: "pointer" }}>
                Clear
              </button>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function Field({ label, children, error }) {
  return (
    <div>
      <label style={{ fontSize: 12, fontWeight: 600, color: error ? "#C62828" : "#555", display: "block", marginBottom: 5 }}>
        {label}{error && <span style={{ marginLeft: 6, fontWeight: 800 }}>⚠️ Required</span>}
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
async function exportSalesReportXlsx(monthStr) {
  const [y, m] = monthStr.split("-").map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const monthName = monthNames[m - 1];

  // One round trip for every day in the month instead of a per-day localStorage read.
  const startDate = `${monthStr}-01`;
  const endDate = `${monthStr}-${String(lastDay).padStart(2, "0")}`;
  const { days } = await apiFetch(`/api/day-data-range?start=${startDate}&end=${endDate}`);

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
    const data = days[dateStr];
    if (data) {
      const appts = (data.appointments || []).filter(a => !a.isCavSlot);
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

      dailyData.push({ date: d, ...computeDayTotals(data) });
    } else {
      dailyData.push({ date: d, totalSales: 0, clients: "", cashTreatment:"", cashProduct:"", totalCash:"", cashTip:"", cardTreatment:"", cardProduct:"", totalCard:"", cardTip:"", totalTip:"", grandTotal: 0, rl: 0, rt: 0, nl: 0, nt: 0, referrals: REFERRAL_SOURCES.map(() => 0) });
    }
  }

  const wb = XLSX.utils.book_new();
  const wsData = [
    [`Dr.Body,Inc. Sales Report in ${monthName} ${y}`, "", "", "", "", "", "", "", "", "Please fill in the blue cells"],
    [],
    ["Date", "Total Sales", "Clients", "Cash", "", "", "", "Card", "", "", "", "Total tip", "Total sales and Tip"],
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
    ["Date", "RL", "RT", "NL", "NT", "Total", ...REFERRAL_SOURCES.map(src => REFERRAL_LABELS[src] || src)],
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
    ["", "Count", "Amount"],
    ["Total", ...countAmt(ticketSaleEvents)],
    ["New customers (NL+NT)", ...countAmt(newSales)],
    ["Repeat customers (RL+RT)", ...countAmt(repeatSales)],
    [],
    ["New customers by referral source", "Count", "Amount"],
  ];
  REFERRAL_SOURCES.forEach(src => {
    tsData.push([REFERRAL_LABELS[src] || src, ...countAmt(newSales.filter(r => r.referralSource === src))]);
  });
  tsData.push([]);
  tsData.push(["By staff (ranked by New customers $ — who's converting new clients into ticket sales)", "Count", "Amount", "New customers $", "Repeat customers $", "Local visits (RL+NL)", "→ Converted to ticket", "Conversion %"]);
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
  const [loadError, setLoadError] = useState(null);
  const [xlsxLoading, setXlsxLoading] = useState(false);

  const calcPeriodDates = (monthStr, p) => {
    const [y, m] = monthStr.split("-").map(Number);
    if (p === "first") {
      return { start: `${monthStr}-01`, end: `${monthStr}-15` };
    } else {
      const lastDay = new Date(y, m, 0).getDate();
      return { start: `${monthStr}-16`, end: `${monthStr}-${String(lastDay).padStart(2, "0")}` };
    }
  };

  const loadPayroll = async () => {
    setLoading(true);
    setLoadError(null);
    const { start, end } = calcPeriodDates(month, period);

    try {
      // One round trip for every day in the period instead of a per-day localStorage read.
      const { days } = await apiFetch(`/api/day-data-range?start=${start}&end=${end}`);
      const startD = new Date(start);
      const endD = new Date(end);
      const allAppts = [];
      const allRetails = [];
      const allForgottenTips = [];

      for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
        const dateStr = d.toISOString().slice(0, 10);
        const data = days[dateStr];
        if (data) {
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
        ? `💰Includes $${Number(a.depositApplied || 0) + Number(a.packageDepositAmount || 0)} deposit` : "",
      Number(a.giftCardUsed || 0) > 0 ? `🎁Used $${Number(a.giftCardUsed)} gift card` : "",
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
        notes: [a.notes || "", moneyNote(a), a.isPromo ? "📸Complimentary PR" : ""].filter(Boolean).join("　"),
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
            notes: `🛍️ Retail${item.productName ? ` (${item.productName})` : ""}`,
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
          notes: `🛍️ Retail (phone/walk-in)${r.item ? ` (${r.item})` : ""}`,
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
        notes: `🙏 Forgotten Entry${ft.originalDate ? ` (visit on ${ft.originalDate})` : ""}${ft.notes ? ` ${ft.notes}` : ""}`,
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
          notes: `➕ ${addon.serviceName || addon.name || "Add-on"}${addon.ticketCurrent ? ` ${addon.ticketCurrent}/${a.ticketTotal||3}` : ""}`,
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
        notes: ["⚡ Machine", parent ? moneyNote(parent) : ""].filter(Boolean).join("　"),
      });
      byTherapist[t].totalService += svc;
      byTherapist[t].totalTip += tip;
    });

    setPayrollData({ byTherapist, start, end });
    } catch (e) {
      console.error("Payroll load error:", e);
      setLoadError("Failed to load. Please check your internet connection and try again");
    } finally {
      setLoading(false);
    }
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
      ["Date", "Client", ...(therapist === "Maki" ? ["Minutes"] : []), "Treatment", "Tip", "Total", "Retail", "Remarks"],
      ...data.rows.map(r => [
        r.date, r.client, ...(therapist === "Maki" ? [r.duration || ""] : []), r.service || "",
        r.tip || "", r2(r.service + r.tip), r.retail || "", remarksFor(r)
      ]),
      [],
      ["", "Total", ...(therapist === "Maki" ? [""] : []), data.totalService, data.totalTip, r2(data.totalService + data.totalTip), data.totalRetail, ""],
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
    const allRows = [["Therapist", "Date", "Client", "Minutes (Maki only)", "Treatment", "Tip", "Total", "Retail", "Remarks"]];
    THERAPISTS.forEach(t => {
      const data = payrollData.byTherapist[t];
      if (!data || data.rows.length === 0) return;
      data.rows.forEach(r => {
        allRows.push([
          t, r.date, r.client, t === "Maki" ? (r.duration || "") : "",
          r.service || "", r.tip || "", r2(r.service + r.tip), r.retail || "", remarksForAll(r)
        ]);
      });
      allRows.push([t, "", "Subtotal", "", data.totalService, data.totalTip, r2(data.totalService + data.totalTip), data.totalRetail, ""]);
      allRows.push([]);
    });
    const csv = allRows.map(r => r.join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `AllStaffPayroll_${payrollData.start}_${payrollData.end}.csv`;
    a.click();
  };

  const { start, end } = calcPeriodDates(month, period);

  return (
    <div style={{ padding: 16 }}>
      {/* Sales Report Excel Export */}
      <div style={{ background: "linear-gradient(135deg,#E8F5E9,#C8E6C9)", borderRadius: 12, padding: 16, marginBottom: 16, border: "2px solid #4CAF50" }}>
        <div style={{ fontWeight: 700, color: "#2E7D32", marginBottom: 8, fontSize: 15 }}>📊 Sales Report Excel Export</div>
        <div style={{ fontSize: 12, color: "#555", marginBottom: 12 }}>Select a month and download the Excel report (Dr.Body format)</div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <MonthPicker value={month} onChange={setMonth}
            style={{ padding: "8px 12px", borderRadius: 8, border: "1.5px solid #4CAF50", fontSize: 14 }} />
          <button onClick={async () => {
              setXlsxLoading(true);
              try { await exportSalesReportXlsx(month); }
              catch (e) { console.error("Excel export error:", e); window.alert("Excel export failed. Please check your internet connection and try again"); }
              finally { setXlsxLoading(false); }
            }} disabled={xlsxLoading}
            style={{ padding: "10px 20px", borderRadius: 10, border: "none", background: xlsxLoading ? "#666" : "#2E7D32", color: "#fff", fontWeight: 700, cursor: xlsxLoading ? "not-allowed" : "pointer", fontSize: 14 }}>
            {xlsxLoading ? "⏳ Exporting..." : "⬇️ Download Excel"}
          </button>
        </div>
      </div>
      <div style={{ background: "#fff", borderRadius: 12, padding: 16, marginBottom: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
        <div style={{ fontWeight: 700, color: "#0D4F4F", marginBottom: 12, fontSize: 15 }}>💴 Payroll Period</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 12 }}>
          <Field label="Month">
            <MonthPicker value={month} onChange={setMonth}
              style={{ padding: "8px 12px", borderRadius: 8, border: "1.5px solid #DDD", fontSize: 14 }} />
          </Field>
          <Field label="Period">
            <div style={{ display: "flex", gap: 8 }}>
              {[["first", `1st-15th`], ["second", `16th-end of month`]].map(([val, label]) => (
                <button key={val} onClick={() => setPeriod(val)}
                  style={{ padding: "8px 16px", borderRadius: 8, border: `2px solid ${period === val ? "#0D4F4F" : "#DDD"}`, background: period === val ? "#0D4F4F" : "#fff", color: period === val ? "#fff" : "#888", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
                  {label}
                </button>
              ))}
            </div>
          </Field>
        </div>
        <div style={{ fontSize: 12, color: "#888", marginBottom: 12 }}>Period: {start} – {end}</div>
        {loadError && <div style={{ color: "#C62828", fontSize: 13, marginBottom: 12 }}>{loadError}</div>}
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={loadPayroll} disabled={loading}
            style={{ padding: "10px 24px", borderRadius: 10, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>
            {loading ? "Loading..." : "📊 Calculate"}
          </button>
          {payrollData && (
            <button onClick={downloadAllCSV}
              style={{ padding: "10px 20px", borderRadius: 10, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 14 }}>
              ⬇️ All Staff CSV
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
                <span style={{ fontSize: 12, color: "#888", marginLeft: 10 }}>{data.rows.length} visits</span>
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700, color: "#C62828", fontSize: 15 }}>Treatment {formatCurrency(data.totalService)}</span>
                <span style={{ fontWeight: 700, color: "#E65100", fontSize: 15 }}>Tip {formatCurrency(data.totalTip)}</span>
                {data.totalRetail > 0 && <span style={{ fontWeight: 700, color: "#6A1B9A", fontSize: 15 }}>Retail {formatCurrency(data.totalRetail)}</span>}
                <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 15 }}>Total {formatCurrency(data.totalService + data.totalTip + data.totalRetail)}</span>
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
                    {["Date", "Client", ...(t === "Maki" ? ["Minutes"] : []), "Treatment", "Tip", "Total", "Retail", "Remarks"].map(h => (
                      <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.rows.sort((a, b) => a.date.localeCompare(b.date)).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F5F5F5", background: r.isGiftCard ? "#FFFDE7" : r.isPromo ? "#E3F2FD" : r.isTicket ? "#F0F7FF" : "white" }}>
                      <td style={{ padding: "6px 8px", whiteSpace: "nowrap", color: "#888" }}>{r.date.slice(5)}</td>
                      <td style={{ padding: "6px 8px", fontWeight: 600 }}>{r.client}</td>
                      {t === "Maki" && <td style={{ padding: "6px 8px", textAlign: "center" }}>{r.duration}min</td>}
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
                    <td colSpan={t === "Maki" ? 3 : 2} style={{ padding: "8px", fontWeight: 700, color: "#555" }}>Total</td>
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
          No data for this period
        </div>
      )}
    </div>
  );
}
