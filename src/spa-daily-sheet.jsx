import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const THERAPISTS = ["Mami", "Aya", "Megumi", "Hitomi", "Maki", "Yuka", "Mai", "Betsy"];
const CUSTOMER_TYPES = ["RL", "RT", "NL", "NT"];
const HOURS = Array.from({ length: 11 }, (_, i) => i + 9);

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

const MENU_OPTIONS = [
  { group: "Improving Posture", prefix: "IP", durations: [60, 75, 90, 120] },
  { group: "Weight Loss", prefix: "WL", durations: [75, 90, 115] },
  { group: "Glow Facial", prefix: "FA", durations: [60] },
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

const WL_CAV_FIXED = 116;
const isWLService = (name) => (name || '').toLowerCase().includes('weight loss');

const EMPTY_APPOINTMENT = {
  id: null,
  clientName: "",
  therapist: "",
  startTime: "",
  duration: 60,
  isTicket: false,
  ticketMenu: "",       // e.g. "IP-90-3"
  ticketTotal: 3,       // 3 or 5
  ticketCurrent: 1,     // which session (1/3, 2/3...)
  priceVersion: "new",  // new | old
  cavTherapist: "",
  price: 0,
  paymentType: "cash",
  tip: 0,
  tipPaymentType: "cash",
  cavPrice: 0,
  cavTip: 0,
  customerType: "RL",
  serviceName: "",  // コース名（Squareから自動入力）
  notes: "",
  purchaseTags: [],
  // Inline purchase records (recorded on same appointment)
  newTicketAmount: 0, newTicketTip: 0, newTicketPaymentType: "card", newTicketTipPaymentType: "cash",
  retailPurchaseAmount: 0, retailPurchasePaymentType: "cash", retailProductName: "",
  giftCardUsed: 0,
  giftCardPurchaseAmount: 0, giftCardPurchasePaymentType: "card",
  depositApplied: 0,
  isGiftCard: false,
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
  { name: "Epicutis Glow Facial 60min", duration: 60 },
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
  { name: "Cavitation 40min", duration: 40 },
  { name: "Electric Brush 15min", duration: 15 },
  { name: "Hot Stone 30min", duration: 30 },
  { name: "Shaving 30min", duration: 30 },
  { name: "Head Massage 15min", duration: 15 },
  { name: "Radio Frequency 15min", duration: 15 },
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
const EMPTY_DEPOSIT = { id: Date.now(), type: "deposit", amount: 0, clientName: "", paymentType: "card", appointmentDate: "", appointmentTime: "", tip: 0, tipPaymentType: "cash", notes: "" };
const EMPTY_TICKET_PURCHASE = { id: Date.now(), clientName: "", packageName: "", amount: 0, paymentType: "card", tip: 0, tipPaymentType: "cash", notes: "" };
const EMPTY_STAFF_PURCHASE = { id: Date.now(), staffName: "", productName: "", amount: 0, paymentType: "cash", notes: "" };

const PURCHASE_TAGS = [
  { id: "ticketEnd",  label: "🏁 チケット最終回",   color: "#5D4037", bg: "#EFEBE9" },
  { id: "giftCard",  label: "🎁 ギフトカード購入",  color: "#00796B", bg: "#E0F2F1" },
  { id: "retail",    label: "🛍️ 物販購入",         color: "#6A1B9A", bg: "#F3E5F5" },
];

const ADDON_PRESETS = [
  "キャビ追加 +10分",
  "キャビ追加 +20分",
  "キャビ追加 +30分",
  "マッサージ追加",
  "ひしゃ追加",
];

const RETAIL_PRODUCTS = [
  { name: "Sheet Mask", price: 10 },
  { name: "Epicutis Sheet Mask", price: 30 },
  { name: "Epicutis Sample Set", price: 42 },
  { name: "Koso Shot", price: 10 },
  { name: "Koso Drink", price: 99 },
  { name: "Belly Detox Tea", price: 25 },
  { name: "Decaf Chocolate Truffle Tea", price: 25 },
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
  const [editingStaffPurchase, setEditingStaffPurchase] = useState(null);
  const [editingAppt, setEditingAppt] = useState(null);
  const [editingRetail, setEditingRetail] = useState(null);
  const [editingDeposit, setEditingDeposit] = useState(null);
  const [editingTicketPurchase, setEditingTicketPurchase] = useState(null);
  const [squareLoading, setSquareLoading] = useState(false);
  const [squareStatus, setSquareStatus] = useState(null);
  const [activeTab, setActiveTab] = useState("schedule");
  const [selectedTherapist, setSelectedTherapist] = useState("All");
  const [workingStaff, setWorkingStaff] = useState(THERAPISTS); // default: all working
  const [showStaffPicker, setShowStaffPicker] = useState(false);
  const [toast, setToast] = useState(null);
  const [depositsForDate, setDepositsForDate] = useState([]);

  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); };

  // All Square data by date
  const ALL_SQUARE_DATA = {
    "2026-06-20": [
      { id: "sq-njozv4l63l9po2", clientName: "Tony Quach",        therapist: "Mami",  startTime: "09:15", duration: 90,  serviceName: "Improving Posture 90min" },
      { id: "sq-56o0myzdf4ut6b", clientName: "Natsume Minegishi", therapist: "Yuka",  startTime: "09:30", duration: 60,  serviceName: "Sculpted Face Lift 60min" },
      { id: "sq-did02ytnrwdlgj", clientName: "Jamielynn Estrada", therapist: "Mami",  startTime: "11:00", duration: 60,  serviceName: "HIFU Full Face 60min" },
      { id: "sq-pb3qfamgawhhjp", clientName: "Marilyn Remiglo",   therapist: "Yuka",  startTime: "11:00", duration: 60,  serviceName: "Epicutis Glow Facial 60min" },
      { id: "sq-sbruu6ass34oir", clientName: "Nui",               therapist: "Yuka",  startTime: "12:30", duration: 60,  serviceName: "HIFU Full Face 60min", notes: "Couple gift card あり" },
      { id: "sq-qzqxkcv7047ml3", clientName: "Jannie Lai",        therapist: "Mami",  startTime: "13:00", duration: 150, serviceName: "Improving Posture 90min + HIFU Full Face 60min" },
    ],
    "2026-06-21": [
      { id: "p35cktixyc5n4y", clientName: "Brian Hasegawa",  therapist: "Mami",  startTime: "10:00", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "zg8iv61nvltg53", clientName: "Lisa Man",         therapist: "Hitomi",startTime: "10:00", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "2p4jkq26j781vc", clientName: "Esther",          therapist: "Maki",  startTime: "10:30", duration: 75, serviceName: "Improving Posture 75min" },
      { id: "l8hakzoxx1jyna", clientName: "Anita",           therapist: "Betsy", startTime: "10:30", duration: 75, serviceName: "Improving Posture 75min" },
      { id: "6jjrlzm6ofvba0", clientName: "Lani Arakaki",    therapist: "Hitomi",startTime: "12:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "nepeme2kn1qa7p", clientName: "Willo Shimamura", therapist: "Betsy", startTime: "12:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "ygk9eocp80snk4", clientName: "David Arakaki",   therapist: "Maki",  startTime: "12:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "ja225auquxd3s0", clientName: "Trina Le",        therapist: "Mami",  startTime: "14:00", duration: 60, serviceName: "Epicutis Glow Facial 60min" },
      { id: "m32x1d808ofp5s", clientName: "Mio Lau",         therapist: "Maki",  startTime: "14:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "zmph4tmlbendps", clientName: "Miki Garcia",     therapist: "Mami",  startTime: "15:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "4tjqb1isq99s56", clientName: "Elizabeth Park",  therapist: "Maki",  startTime: "16:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "6odddpottup0zs", clientName: "Patrick Lau",     therapist: "Hitomi",startTime: "14:30", duration: 90, serviceName: "Improving Posture 90min" },
      { id: "nk5rhk6b8jf1u6", clientName: "CeCe Bernstine", therapist: "Hitomi",startTime: "17:00", duration: 90, serviceName: "Deep Tissue / Lomi Lomi 90min", notes: "Booked for Jabari Bernstine" },
    ],
  };

  useEffect(() => {
    const saved = localStorage.getItem(`spa-sheet-${date}`);
    const squareRef = ALL_SQUARE_DATA[date] || [];
    if (saved) {
      const d = JSON.parse(saved);
      // Always patch serviceName from Square reference data
      const patched = (d.appointments || []).map(a => {
        const ref = squareRef.find(r => r.id === a.id);
        if (ref && !a.serviceName) return { ...a, serviceName: ref.serviceName };
        return a;
      });
      setAppointments(patched);
      setRetails(d.retails || []);
      setDeposits(d.deposits || []);
      setTicketPurchases(d.ticketPurchases || []);
      setStaffPurchases(d.staffPurchases || []);
      // Set working staff from Square data for this date
      if (squareRef.length > 0) {
        const staff = [...new Set(squareRef.map(a => a.therapist).filter(Boolean))];
        if (staff.length > 0) setWorkingStaff(staff);
      }
    } else if (squareRef.length > 0) {
      // No saved data — auto-load Square data for this date
      const appts = squareRef.map(a => ({
        ...EMPTY_APPOINTMENT, ...a, fromSquare: true,
      }));
      setAppointments(appts);
      setRetails([]);
      setDeposits([]);
      setTicketPurchases([]);
      setStaffPurchases([]);
      // Auto-set working staff from Square data
      const staff = [...new Set(appts.map(a => a.therapist).filter(Boolean))];
      if (staff.length > 0) setWorkingStaff(staff);
    } else {
      setAppointments([]); setRetails([]); setDeposits([]); setTicketPurchases([]); setStaffPurchases([]);
    }
  }, [date]);

  // Scan all localStorage entries for deposits with appointmentDate matching current date
  useEffect(() => {
    const found = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key?.startsWith("spa-sheet-")) continue;
      try {
        const d = JSON.parse(localStorage.getItem(key));
        const recordedDate = key.replace("spa-sheet-", "");
        (d.deposits || []).forEach(dep => {
          if (dep.appointmentDate === date && dep.type === "deposit") {
            found.push({ ...dep, recordedDate });
          }
        });
      } catch {}
    }
    setDepositsForDate(found.sort((a, b) => (a.appointmentTime || "").localeCompare(b.appointmentTime || "")));
  }, [date, deposits]);

  const save = useCallback((appts, rets, deps, tps, sps) => {
    localStorage.setItem(`spa-sheet-${date}`, JSON.stringify({ appointments: appts, retails: rets, deposits: deps, ticketPurchases: tps || [], staffPurchases: sps || [] }));
  }, [date]);

  const fetchSquare = async () => {
    setSquareLoading(true); setSquareStatus(null);
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 2000,
          system: `You have access to Square via MCP tool.
Location ID: J84JMVZ7QNE10

Steps:
1. Call Square bookings list with location_id="J84JMVZ7QNE10", start_at_min and start_at_max covering ${date} in Hawaii time (UTC-10, so ${date}T10:00:00Z to next day T09:59:59Z)
2. For each booking, get customer name via customers get endpoint
3. Map team_member_id to therapist name using this map:
   TME2bzx3XoNFV-jw=Mami, TMWQwOib7pJ7Sw22=Yuka, TMez6BJMpuw8PpVZ=Megumi,
   TMZUOeE7bTyj6JrW=Betsy, TMhfddQBlrnZkR9n=Aya, TMu1v7kb0cgbl5mn=Mai,
   TMwBMik2I8Xpyog2=Maki, TMW8eccMtM4pdMiM=Hitomi
4. Convert start_at from UTC to Hawaii time (subtract 10 hours), format as HH:MM

Return ONLY valid JSON array (no markdown, no explanation):
[{"squareId":"...","clientName":"First Last","therapist":"Mami","startTime":"09:15","duration":90,"notes":""}]

If no bookings return [].`,
          messages: [{ role: "user", content: `Fetch Square bookings for ${date} Hawaii time and return JSON only.` }],
          mcp_servers: [{ type: "url", url: "https://mcp.squareup.com/mcp", name: "square-mcp" }],
        }),
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();

      // Extract JSON from response - look through all content blocks
      let bookings = [];
      const allText = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
      const jsonMatch = allText.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        bookings = JSON.parse(jsonMatch[0]);
      }

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
          notes: b.notes || "",
          fromSquare: true,
        }));

        const merged = [...appointments];
        newAppts.forEach(na => {
          const exists = merged.find(a =>
            (a.fromSquare && a.id === na.id) ||
            (a.clientName === na.clientName && a.startTime === na.startTime)
          );
          if (!exists) merged.push(na);
        });
        merged.sort((a, b) => (a.startTime || "").localeCompare(b.startTime || ""));
        setAppointments(merged);
        save(merged, retails, deposits, ticketPurchases, staffPurchases);
        showToast(`✅ ${newAppts.length}件取得しました`);
      }
    } catch (e) {
      console.error("Square sync error:", e);
      setSquareStatus("Square接続エラー。手動入力してください。");
      showToast("接続エラー", "error");
    }
    setSquareLoading(false);
  };

  const loadTodayFromSquare = () => {
    const filtered = ALL_SQUARE_DATA[date] || [];
    if (filtered.length === 0) {
      showToast(`${date}のデータがありません`, "info");
      return;
    }

    // Build complete appointment objects with ALL fields
    const newAppts = filtered.map(a => ({
      id: a.id,
      clientName: a.clientName,
      therapist: a.therapist,
      startTime: a.startTime,
      duration: a.duration,
      serviceName: a.serviceName,
      notes: a.notes || "",
      isTicket: false,
      isCavSlot: false,
      ticketMenu: "",
      ticketTotal: 3,
      ticketCurrent: 1,
      priceVersion: "new",
      cavTherapist: "",
      price: 0,
      paymentType: "cash",
      tip: 0,
      tipPaymentType: "cash",
      cavPrice: 0,
      cavTip: 0,
      customerType: "RL",
      fromSquare: true,
      bodyMins: 0,
      cavMins: 0,
      referralSource: "",
    }));

    // ALWAYS overwrite — clear old Square data for this date, keep manual entries
    const manualAppts = appointments.filter(a => !a.fromSquare && !a.isCavSlot);
    const combined = [...manualAppts, ...newAppts];
    combined.sort((a, b) => (a.startTime || "").localeCompare(b.startTime || ""));

    // Also clear localStorage for this date and rewrite fresh
    localStorage.removeItem(`spa-sheet-${date}`);
    setAppointments(combined);
    save(combined, retails, deposits, ticketPurchases, staffPurchases);

    const staffInBookings = [...new Set(newAppts.map(a => a.therapist).filter(Boolean))];
    if (staffInBookings.length > 0) setWorkingStaff(staffInBookings);

    showToast(`✅ ${newAppts.length}件ロードしました！`);
    // DEBUG: confirm serviceName
    console.log("LOADED APPTS:", JSON.stringify(newAppts.map(a => ({name: a.clientName, serviceName: a.serviceName}))));
  };
  const calcCavStartTime = (startTime, duration) => {
    if (!startTime) return "";
    const [h, m] = startTime.split(":").map(Number);
    const totalMins = h * 60 + m + (Number(duration) - 15);
    const ch = Math.floor(totalMins / 60);
    const cm = totalMins % 60;
    return `${String(ch).padStart(2, "0")}:${String(cm).padStart(2, "0")}`;
  };

  const saveAppt = (appt) => {
    const cavSlotId = `cav-${appt.id}`;

    // A separate cav slot is needed when:
    // - A cav therapist is selected (meaning the main therapist can't do machine)
    // - There are cav amounts to record
    const needsCavSlot = !!appt.cavTherapist && (Number(appt.cavPrice) > 0 || Number(appt.cavTip) > 0);

    const cavSlot = needsCavSlot ? {
      id: cavSlotId,
      isCavSlot: true,
      parentId: appt.id,
      clientName: appt.clientName,
      therapist: appt.cavTherapist,
      startTime: calcCavStartTime(appt.startTime, appt.duration),
      duration: 15,
      isTicket: appt.isTicket,
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
      ? { ...appt, cavPrice: 0, cavTip: 0 }  // body only — machine goes to cav slot
      : appt;

    next = [...next, mainAppt];
    if (cavSlot) next = [...next, cavSlot];
    next.sort((a, b) => (a.startTime || "").localeCompare(b.startTime || ""));

    setAppointments(next); save(next, retails, deposits, ticketPurchases, staffPurchases); setEditingAppt(null); showToast("保存しました");
  };

  const deleteAppt = (id) => {
    const cavSlotId = `cav-${id}`;
    const next = appointments.filter(a => a.id !== id && a.id !== cavSlotId);
    setAppointments(next); save(next, retails, deposits, ticketPurchases, staffPurchases); setEditingAppt(null);
  };

  // When editing an existing appointment, restore cavPrice/cavTip from the cav slot
  // and infer bodyMins/cavMins from the price ratio so splits recalculate correctly
  const openApptForEdit = (appt) => {
    if (!appt.isCavSlot && appt.cavTherapist) {
      const cavSlot = appointments.find(a => a.isCavSlot && a.parentId === appt.id);
      if (cavSlot) {
        const restoredCavPrice = Number(cavSlot.price || 0);
        const restoredCavTip = Number(cavSlot.tip || 0);
        const bodyPrice = Number(appt.price || 0);
        const total = bodyPrice + restoredCavPrice;
        const dur = Number(appt.duration || 90);
        const bodyMins = total > 0 ? Math.round(bodyPrice / total * dur) : 0;
        const cavMins = bodyMins > 0 ? dur - bodyMins : 0;
        setEditingAppt({
          ...appt,
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
  const saveRetail = (r) => { const next = retails.find(x => x.id === r.id) ? retails.map(x => x.id === r.id ? r : x) : [...retails, r]; setRetails(next); save(appointments, next, deposits, ticketPurchases, staffPurchases); setEditingRetail(null); showToast("物販保存"); };
  const deleteRetail = (id) => { const next = retails.filter(r => r.id !== id); setRetails(next); save(appointments, next, deposits, ticketPurchases, staffPurchases); };
  const saveDeposit = (d) => { const next = deposits.find(x => x.id === d.id) ? deposits.map(x => x.id === d.id ? d : x) : [...deposits, d]; setDeposits(next); save(appointments, retails, next, ticketPurchases, staffPurchases); setEditingDeposit(null); showToast("保存しました"); };
  const deleteDeposit = (id) => { const next = deposits.filter(d => d.id !== id); setDeposits(next); save(appointments, retails, next, ticketPurchases, staffPurchases); };
  const saveTicketPurchase = (tp) => { const next = ticketPurchases.find(x => x.id === tp.id) ? ticketPurchases.map(x => x.id === tp.id ? tp : x) : [...ticketPurchases, tp]; setTicketPurchases(next); save(appointments, retails, deposits, next, staffPurchases); setEditingTicketPurchase(null); showToast("🎟️ チケット購入保存"); };
  const deleteTicketPurchase = (id) => { const next = ticketPurchases.filter(tp => tp.id !== id); setTicketPurchases(next); save(appointments, retails, deposits, next, staffPurchases); };
  const saveStaffPurchase = (sp) => { const next = staffPurchases.find(x => x.id === sp.id) ? staffPurchases.map(x => x.id === sp.id ? sp : x) : [...staffPurchases, sp]; setStaffPurchases(next); save(appointments, retails, deposits, ticketPurchases, next); setEditingStaffPurchase(null); showToast("👩‍💼 社販保存"); };
  const deleteStaffPurchase = (id) => { const next = staffPurchases.filter(sp => sp.id !== id); setStaffPurchases(next); save(appointments, retails, deposits, ticketPurchases, next); };

  // Summary — tickets excluded from today's revenue; cav slots excluded (counted in parent)
  const regularAppts = appointments.filter(a => !a.isTicket && !a.isCavSlot && !a.isGiftCard);
  const ticketAppts = appointments.filter(a => a.isTicket && !a.isCavSlot);
  const pureTicketAppts = ticketAppts.filter(a => !a.isSameDayTicket);
  const sameDayAppts = ticketAppts.filter(a => a.isSameDayTicket);
  const gcAppts = appointments.filter(a => a.isGiftCard && !a.isCavSlot);

  const totalRevenue = regularAppts.reduce((s, a) => s + Number(a.price || 0), 0)
    + sameDayAppts.reduce((s, a) => s + Number(a.packagePrice || 0), 0);
  const totalCavRevenue = regularAppts.reduce((s, a) => s + Number(a.cavPrice || 0), 0);
  const totalTips = regularAppts.reduce((s, a) => s + Number(a.tip || 0), 0)
    + sameDayAppts.reduce((s, a) => s + Number(a.packageTip ?? a.tip ?? 0), 0)
    + pureTicketAppts.reduce((s, a) => s + Number(a.extraTip || 0), 0);
  const totalCavTips = regularAppts.reduce((s, a) => s + Number(a.cavTip || 0), 0);
  // GC allocation: covers service first, then tip, then retail (in order)
  const gcAlloc = (a) => {
    const gc = Number(a.giftCardUsed || 0);
    const svc = Number(a.price || 0);
    const tip = Number(a.tip || 0);
    const retail = (a.purchaseTags?.includes("retail")) ? Number(a.retailPurchaseAmount || 0) : 0;
    const gcSvc = Math.min(gc, svc);
    const gcTip = Math.min(gc - gcSvc, tip);
    const gcRetail = Math.min(gc - gcSvc - gcTip, retail);
    return { gcSvc, gcTip, gcRetail };
  };
  const totalGCUsed = regularAppts.reduce((s, a) => s + Number(a.giftCardUsed || 0), 0);
  const totalCash = regularAppts.filter(a => a.paymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.price || 0) - gcAlloc(a).gcSvc), 0)
    + sameDayAppts.filter(a => a.paymentType === "cash").reduce((s, a) => s + Number(a.packagePrice || 0), 0);
  const totalCard = regularAppts.filter(a => a.paymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.price || 0) - gcAlloc(a).gcSvc), 0)
    + sameDayAppts.filter(a => a.paymentType === "card").reduce((s, a) => s + Number(a.packagePrice || 0), 0);
  const totalTipCash = regularAppts.filter(a => a.tipPaymentType === "cash").reduce((s, a) => s + Math.max(0, Number(a.tip || 0) - gcAlloc(a).gcTip), 0)
    + sameDayAppts.filter(a => a.tipPaymentType === "cash").reduce((s, a) => s + Number(a.packageTip ?? a.tip ?? 0), 0)
    + pureTicketAppts.filter(a => a.extraTipPaymentType === "cash").reduce((s, a) => s + Number(a.extraTip || 0), 0)
    + deposits.filter(d => d.tipPaymentType === "cash" || !d.tipPaymentType).reduce((s, d) => s + Number(d.tip || 0), 0);
  const totalTipCard = regularAppts.filter(a => a.tipPaymentType === "card").reduce((s, a) => s + Math.max(0, Number(a.tip || 0) - gcAlloc(a).gcTip), 0)
    + sameDayAppts.filter(a => a.tipPaymentType === "card").reduce((s, a) => s + Number(a.packageTip ?? a.tip ?? 0), 0)
    + pureTicketAppts.filter(a => a.extraTipPaymentType === "card").reduce((s, a) => s + Number(a.extraTip || 0), 0)
    + deposits.filter(d => d.tipPaymentType === "card").reduce((s, d) => s + Number(d.tip || 0), 0);
  const totalRetail = retails.reduce((s, r) => s + Number(r.price || 0), 0);
  const totalDepositAmt = deposits.filter(d => d.type === "deposit").reduce((s, d) => s + Number(d.amount || 0), 0);
  const totalDepositApplied = regularAppts.reduce((s, a) => s + Number(a.depositApplied || 0), 0);
  const totalGiftCard = deposits.filter(d => d.type === "giftcard").reduce((s, d) => s + Number(d.amount || 0), 0);
  // Ticket new purchases (standalone section)
  const tpCash = ticketPurchases.filter(tp => tp.paymentType === "cash").reduce((s, tp) => s + Number(tp.amount || 0), 0);
  const tpCard = ticketPurchases.filter(tp => tp.paymentType === "card").reduce((s, tp) => s + Number(tp.amount || 0), 0);
  const tpCheck = ticketPurchases.filter(tp => tp.paymentType === "check").reduce((s, tp) => s + Number(tp.amount || 0), 0);
  const tpTotal = ticketPurchases.reduce((s, tp) => s + Number(tp.amount || 0), 0);
  const tpTipCash = ticketPurchases.filter(tp => tp.tipPaymentType === "cash").reduce((s, tp) => s + Number(tp.tip || 0), 0);
  const tpTipCard = ticketPurchases.filter(tp => tp.tipPaymentType === "card").reduce((s, tp) => s + Number(tp.tip || 0), 0);
  const tpTipCheck = ticketPurchases.filter(tp => tp.tipPaymentType === "check").reduce((s, tp) => s + Number(tp.tip || 0), 0);
  const tpTipTotal = ticketPurchases.reduce((s, tp) => s + Number(tp.tip || 0), 0);
  // Inline purchases recorded inside appointment cards
  const allApptsList = appointments.filter(a => !a.isCavSlot);
  const inlineNewTicketAppts = allApptsList.filter(a => (a.purchaseTags||[]).includes("newTicket"));
  const inlineRetailAppts   = allApptsList.filter(a => (a.purchaseTags||[]).includes("retail"));
  const inlineGCAppts       = allApptsList.filter(a => (a.purchaseTags||[]).includes("giftCard"));
  const inlineNTCash  = inlineNewTicketAppts.filter(a=>a.newTicketPaymentType==="cash").reduce((s,a)=>s+Number(a.newTicketAmount||0),0);
  const inlineNTCard  = inlineNewTicketAppts.filter(a=>a.newTicketPaymentType!=="cash").reduce((s,a)=>s+Number(a.newTicketAmount||0),0);
  const inlineNTTipCash = inlineNewTicketAppts.filter(a=>a.newTicketTipPaymentType==="cash").reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  const inlineNTTipCard = inlineNewTicketAppts.filter(a=>a.newTicketTipPaymentType!=="cash").reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  const inlineNTTotal = inlineNewTicketAppts.reduce((s,a)=>s+Number(a.newTicketAmount||0),0);
  const inlineNTTipTotal = inlineNewTicketAppts.reduce((s,a)=>s+Number(a.newTicketTip||0),0);
  const inlineRetailCash = inlineRetailAppts.filter(a=>a.retailPurchasePaymentType==="cash").reduce((s,a)=>s+Math.max(0,Number(a.retailPurchaseAmount||0)-gcAlloc(a).gcRetail),0);
  const inlineRetailCard = inlineRetailAppts.filter(a=>a.retailPurchasePaymentType!=="cash").reduce((s,a)=>s+Math.max(0,Number(a.retailPurchaseAmount||0)-gcAlloc(a).gcRetail),0);
  const inlineRetailTotal = inlineRetailAppts.reduce((s,a)=>s+Number(a.retailPurchaseAmount||0),0);
  const inlineGCCash = inlineGCAppts.filter(a=>a.giftCardPurchasePaymentType==="cash").reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);
  const inlineGCCard = inlineGCAppts.filter(a=>a.giftCardPurchasePaymentType!=="cash").reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);
  const inlineGCTotal = inlineGCAppts.reduce((s,a)=>s+Number(a.giftCardPurchaseAmount||0),0);

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
    return { therapist: t, revenue, tips, tipCash, tipCard, cavRevenue, ticketRevenue, ticketCavRevenue, ticketTips, gcRevenue, gcTips, clients: appts.length, ticketClients: ticketAppts2.length, gcClients: gcAppts2.length, byType };
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
          <button onClick={loadTodayFromSquare}
            style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "#5C6BC0", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            📥 今日のデータ
          </button>
        </div>
      </div>
      {squareStatus && <div style={{ background: "#FFF3CD", padding: "8px 20px", fontSize: 13, color: "#856404" }}>⚠️ {squareStatus}</div>}

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
            <button key={t} onClick={() => setWorkingStaff(prev =>
              prev.includes(t) ? prev.filter(x => x !== t) : [...prev, t]
            )} style={{
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

          {/* デポジット済み来店予定 */}
          {depositsForDate.length > 0 && (
            <div style={{ background: "#E3F2FD", borderRadius: 12, padding: 12, marginBottom: 14, border: "2px solid #1565C0" }}>
              <div style={{ fontWeight: 800, fontSize: 13, color: "#1565C0", marginBottom: 10 }}>
                💰 本日来店予定 — デポジット済みのお客様 {depositsForDate.length}名
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {depositsForDate.map((dep, i) => (
                  <div key={dep.id || i} style={{
                    background: "#fff", borderRadius: 10, padding: "10px 14px",
                    border: "2px solid #1565C0", minWidth: 170,
                    boxShadow: "0 2px 6px rgba(21,101,192,0.12)"
                  }}>
                    <div style={{ fontWeight: 800, fontSize: 14, color: "#0D47A1" }}>{dep.clientName}</div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#1565C0", marginTop: 3 }}>
                      💰 ${dep.amount}
                      <span style={{ fontSize: 11, fontWeight: 400, color: "#888", marginLeft: 6 }}>
                        {dep.paymentType === "cash" ? "現金" : dep.paymentType === "card" ? "カード" : "Check"}
                      </span>
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
                    <button onClick={() => setEditingAppt({ ...EMPTY_APPOINTMENT, id: `m-${Date.now()}`, therapist: t, startTime: "09:00" })}
                      style={{ display: "block", margin: "4px auto 0", fontSize: 10, padding: "2px 8px", borderRadius: 10, border: "1px dashed #0D4F4F", background: "none", cursor: "pointer", color: "#0D4F4F" }}>
                      + 追加
                    </button>
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
                        onClick={() => !hasContent && setEditingAppt({ ...EMPTY_APPOINTMENT, id: `m-${Date.now()}`, therapist: t, startTime: `${String(hour).padStart(2,"0")}:00` })}>
                        {appts.map(a => <ApptCard key={a.id} appt={a} onClick={() => {
                          if (a.isCavSlot) {
                            const parent = appointments.find(p => p.id === a.parentId);
                            if (parent) openApptForEdit(parent);
                          } else {
                            openApptForEdit(a);
                          }
                        }} />)}
                        {addonItems.map(({ parentAppt, addon }) => (
                          <div key={`${parentAppt.id}-${addon.id}`}
                            onClick={e => { e.stopPropagation(); openApptForEdit(parentAppt); }}
                            style={{ background: "#E0F2F1", border: "1.5px solid #00796B", borderRadius: 8, padding: "5px 7px", marginBottom: 3, cursor: "pointer" }}>
                            <div style={{ fontSize: 10, fontWeight: 800, color: "#00695C" }}>➕ オプション</div>
                            <div style={{ fontSize: 11, fontWeight: 700, color: "#004D40" }}>{parentAppt.clientName}</div>
                            <div style={{ fontSize: 10, color: "#00796B" }}>{addon.serviceName || "オプション"}</div>
                            {(Number(addon.price||0) + Number(addon.tip||0)) > 0 && (
                              <div style={{ fontSize: 10, fontWeight: 700, color: "#005F4A" }}>
                                施術${addon.price||0}{Number(addon.tip||0) > 0 ? ` チップ$${addon.tip}` : ""}
                              </div>
                            )}
                          </div>
                        ))}
                        {!hasContent && <div style={{ height: 36, border: "1px dashed #DDD", borderRadius: 6, cursor: "pointer", opacity: 0.4 }} />}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>

          {/* Retail */}
          <SectionBox title="🛍️ 物販" color="#6A1B9A" onAdd={() => setEditingRetail({ ...EMPTY_RETAIL, id: Date.now() })}>
            {retails.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {retails.map(r => (
              <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ flex: 1, fontSize: 14 }}>{r.item || "（未入力）"}</span>
                <span style={{ fontWeight: 700, color: "#6A1B9A" }}>{formatCurrency(r.price)}</span>
                <PayBadge type={r.paymentType} />
                <span style={{ fontSize: 12, color: "#888" }}>{r.soldBy}</span>
                <button onClick={() => setEditingRetail(r)} style={iconBtn}>✏️</button>
                <button onClick={() => deleteRetail(r.id)} style={iconBtn}>🗑️</button>
              </div>
            ))}
          </SectionBox>

          {/* Deposits */}
          <SectionBox title="💰 デポジット・ギフトカード" color="#1565C0" onAdd={() => setEditingDeposit({ ...EMPTY_DEPOSIT, id: Date.now() })}>
            {deposits.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {deposits.map(d => (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, background: d.type === "deposit" ? "#E3F2FD" : "#FFF3E0", color: d.type === "deposit" ? "#1565C0" : "#E65100", padding: "2px 8px", borderRadius: 10, fontWeight: 700 }}>
                  {d.type === "deposit" ? "デポジット" : "ギフトカード"}
                </span>
                <span style={{ flex: 1, fontSize: 14 }}>{d.clientName || "—"}</span>
                <span style={{ fontWeight: 700 }}>{formatCurrency(d.amount)}</span>
                <PayBadge type={d.paymentType} />
                <button onClick={() => setEditingDeposit(d)} style={iconBtn}>✏️</button>
                <button onClick={() => deleteDeposit(d.id)} style={iconBtn}>🗑️</button>
              </div>
            ))}
          </SectionBox>

          {/* Staff Purchases 社販 */}
          <SectionBox title="👩‍💼 社販（スタッフ購入）" color="#37474F" onAdd={() => setEditingStaffPurchase({ ...EMPTY_STAFF_PURCHASE, id: Date.now() })}>
            {staffPurchases.length === 0 && <p style={{ color: "#AAA", fontSize: 13 }}>なし</p>}
            {staffPurchases.map(sp => (
              <div key={sp.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: "1px solid #F0F0F0", flexWrap: "wrap" }}>
                <span style={{ fontWeight: 700, fontSize: 14, color: "#37474F" }}>{sp.staffName || "—"}</span>
                {sp.productName && sp.productName !== "__other__" && <span style={{ fontSize: 11, color: "#37474F", background: "#ECEFF1", borderRadius: 8, padding: "2px 8px" }}>{sp.productName}</span>}
                <span style={{ fontWeight: 700, color: "#37474F" }}>{formatCurrency(sp.amount)}</span>
                <PayBadge type={sp.paymentType} />
                {sp.notes && <span style={{ fontSize: 11, color: "#888" }}>{sp.notes}</span>}
                <button onClick={() => setEditingStaffPurchase(sp)} style={iconBtn}>✏️</button>
                <button onClick={() => deleteStaffPurchase(sp.id)} style={iconBtn}>🗑️</button>
              </div>
            ))}
          </SectionBox>

          {/* Ticket New Purchases */}
          <SectionBox title="🎟️ チケット新規購入（電話・当日）" color="#B71C1C" onAdd={() => setEditingTicketPurchase({ ...EMPTY_TICKET_PURCHASE, id: Date.now() })}>
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
                  <PayBadge type={tp.paymentType} />
                  <button onClick={() => setEditingTicketPurchase(tp)} style={iconBtn}>✏️</button>
                  <button onClick={() => deleteTicketPurchase(tp.id)} style={iconBtn}>🗑️</button>
                </div>
              );
            })}
            {ticketPurchases.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 11, color: "#B71C1C", fontWeight: 700, textAlign: "right" }}>
                合計 {formatCurrency(tpTotal)}　チップ {formatCurrency(tpTipTotal)}
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
                <div style={{ fontSize: 26, fontWeight: 800, color: "#C62828" }}>{formatCurrency(totalRevenue + totalCavRevenue + totalRetail)}</div>
              </div>
              <div style={{ background: "#F3E5F5", borderRadius: 10, padding: "10px 18px", textAlign: "center", flex: 1, minWidth: 120 }}>
                <div style={{ fontSize: 11, color: "#888" }}>Total Sales ＋ Tip</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#6A1B9A" }}>{formatCurrency(totalRevenue + totalCavRevenue + totalRetail + totalTips + totalCavTips)}</div>
              </div>
            </div>

            {/* Cash セクション */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#2E7D32", background: "#E8F5E9", padding: "4px 10px", borderRadius: 6, marginBottom: 8, display: "inline-block" }}>💵 CASH</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                {[
                  { label: "Treatment", value: totalCash, color: "#2E7D32" },
                  { label: "Product", value: retails.filter(r=>r.paymentType==="cash").reduce((s,r)=>s+Number(r.price||0),0), color: "#2E7D32" },
                  { label: "Total Cash", value: totalCash + retails.filter(r=>r.paymentType==="cash").reduce((s,r)=>s+Number(r.price||0),0), color: "#1B5E20", bold: true },
                  { label: "Tip", value: totalTipCash, color: "#558B2F" },
                ].map(s => (
                  <div key={s.label} style={{ background: "#F1F8E9", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
                    <div style={{ fontSize: 17, fontWeight: s.bold ? 800 : 700, color: s.color }}>{formatCurrency(s.value)}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Card セクション */}
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", background: "#E3F2FD", padding: "4px 10px", borderRadius: 6, marginBottom: 8, display: "inline-block" }}>💳 CARD</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                {[
                  { label: "Treatment", value: totalCard, color: "#1565C0" },
                  { label: "Product", value: retails.filter(r=>r.paymentType==="card").reduce((s,r)=>s+Number(r.price||0),0), color: "#1565C0" },
                  { label: "Total Card", value: totalCard + retails.filter(r=>r.paymentType==="card").reduce((s,r)=>s+Number(r.price||0),0), color: "#0D47A1", bold: true },
                  { label: "Tip", value: totalTipCard, color: "#1976D2" },
                ].map(s => (
                  <div key={s.label} style={{ background: "#E8F0FE", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
                    <div style={{ fontSize: 17, fontWeight: s.bold ? 800 : 700, color: s.color }}>{formatCurrency(s.value)}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Tip & その他 */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(110px,1fr))", gap: 8, marginBottom: 12 }}>
              {[
                { label: "Total Tip", value: totalTips + totalCavTips, color: "#E65100", bg: "#FFF3E0" },
                { label: "物販合計", value: totalRetail, color: "#6A1B9A", bg: "#F3E5F5" },
                { label: "デポジット受取", value: totalDepositAmt, color: "#0277BD", bg: "#E1F5FE" },
                ...(totalDepositApplied > 0 ? [{ label: "💰 Payroll加算", value: totalDepositApplied, color: "#2E7D32", bg: "#C8E6C9", prefix: "+" }] : []),
                { label: "ギフトカード購入", value: totalGiftCard, color: "#00796B", bg: "#E0F2F1" },
                ...(totalGCUsed > 0 ? [{ label: "🎁 GC使用", value: totalGCUsed, color: "#00695C", bg: "#B2DFDB" }] : []),
              ].map(s => (
                <div key={s.label} style={{ background: s.bg, borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                  <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>{s.label}</div>
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
                <span style={{ color: "#C62828", fontWeight: 800, fontSize: 14, marginLeft: 4 }}>{formatCurrency(totalCash + totalCard + totalDepositAmt + totalGiftCard + tpTotal + inlineNTTotal + inlineGCTotal)}</span>
                <div style={{ marginTop: 4, color: "#999", fontSize: 10 }}>
                  Cash ${totalCash.toFixed(2)}　Card ${totalCard.toFixed(2)}
                  {totalDepositAmt > 0 && `　Deposit $${totalDepositAmt.toFixed(2)}`}
                  {totalGiftCard > 0 && `　GiftCard $${(totalGiftCard + inlineGCTotal).toFixed(2)}`}
                  {(tpTotal + inlineNTTotal) > 0 && `　チケット購入 $${(tpTotal + inlineNTTotal).toFixed(2)}${tpCheck > 0 ? `（Check $${tpCheck.toFixed(2)}）` : ""}`}
                  {inlineRetailTotal > 0 && `　物販(inline) $${inlineRetailTotal.toFixed(2)}`}
                </div>
              </div>
              {/* Final totals grid */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6 }}>
                {[
                  { label: "Cash合計\n（全て）", value: totalCash + retails.filter(r=>r.paymentType==="cash").reduce((s,r)=>s+Number(r.price||0),0) + tpCash + inlineNTCash + inlineRetailCash + inlineGCCash, color: "#2E7D32", bg: "#E8F5E9" },
                  { label: "Card合計\n（全て）", value: totalCard + retails.filter(r=>r.paymentType==="card").reduce((s,r)=>s+Number(r.price||0),0) + tpCard + tpCheck + inlineNTCard + inlineRetailCard + inlineGCCard, color: "#1565C0", bg: "#E3F2FD" },
                  { label: "Tip合計", value: totalTips + totalCavTips + tpTipTotal + inlineNTTipTotal, color: "#E65100", bg: "#FFF3E0" },
                  { label: "🏆 TOTAL\n（全て込み）", value: totalCash + totalCard + totalDepositAmt + totalGiftCard + totalRetail + totalTips + totalCavTips + tpTotal + tpTipTotal + inlineNTTotal + inlineNTTipTotal + inlineRetailTotal + inlineGCTotal, color: "#fff", bg: "#0D4F4F", bold: true },
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
                const total = Number(a.packagePrice||0) + pkgTip;
                return (
                <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                    <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                    <span style={{ fontSize: 12, color: "#2E7D32" }}>{a.ticketMenu} {a.ticketCurrent > 0 ? `${a.ticketCurrent}/${a.ticketTotal}回目` : `購入のみ（${a.ticketTotal}回コース）`}</span>
                    <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4 }}>
                    <span style={{ fontSize: 11, color: "#2E7D32" }}>💰 ${Number(a.packagePrice||0)} / チップ ${pkgTip}{a.paymentType==="card"?"💳":"💵"}</span>
                    <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${total}</span>
                  </div>
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
                const svc = Number(a.price||0) + Number(a.cavPrice||0);
                const tip = Number(a.tip||0) + Number(a.cavTip||0) + Number(a.extraTip||0);
                return (
                  <div key={a.id} onClick={() => openApptForEdit(a)} style={{ background: "#fff", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6 }}>
                      <span style={{ fontWeight: 700 }}>{a.clientName}</span>
                      <span style={{ fontSize: 12, color: "#1565C0" }}>{a.ticketMenu} {a.ticketCurrent}/{a.ticketTotal}</span>
                      <span style={{ fontSize: 12, color: "#888" }}>{a.therapist}{a.cavTherapist ? ` ＋ ${a.cavTherapist}(機械)` : ""}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 4, flexWrap: "wrap", gap: 4 }}>
                      <span style={{ fontSize: 11, color: "#1565C0" }}>施術 ${svc - Number(a.extraTip||0) === svc ? svc : (Number(a.price||0)+Number(a.cavPrice||0))} / チップ ${Number(a.tip||0)+Number(a.cavTip||0)}{a.extraTip > 0 && <span style={{ color: "#F57F17" }}> +${a.extraTip}💝{a.extraTipPaymentType==="card"?"💳":"💵"}</span>}</span>
                      <span style={{ fontWeight: 800, color: "#0D4F4F", fontSize: 14 }}>合計 ${svc + tip}</span>
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
                const svc = Number(a.price||0) + Number(a.cavPrice||0);
                const tip = Number(a.tip||0) + Number(a.cavTip||0);
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
                  {["Google / Website", "Google Map", "Instagram", "Yelp", "紹介"].map(src => {
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
                  {["名前","人数","施術","チップ","現金tip","カードtip","RL","RT","NL","NT","チケット"].map(h => (
                    <th key={h} style={{ padding: "6px 8px", textAlign: h==="名前"?"left":"center", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {summaryByTherapist.filter(s => s.clients > 0 || s.ticketClients > 0).map(s => (
                  <tr key={s.therapist} style={{ borderBottom: "1px solid #F0F0F0" }}>
                    <td style={{ padding: "8px", fontWeight: 700, color: "#0D4F4F" }}>{s.therapist}</td>
                    <td style={{ textAlign: "center", padding: "8px" }}>{s.clients}</td>
                    <td style={{ textAlign: "center", padding: "8px", color: "#C62828", fontWeight: 700 }}>{formatCurrency(s.revenue)}</td>
                    <td style={{ textAlign: "center", padding: "8px", color: "#E65100" }}>{formatCurrency(s.tips)}</td>
                    <td style={{ textAlign: "center", padding: "8px", color: "#2E7D32" }}>{formatCurrency(s.tipCash)}</td>
                    <td style={{ textAlign: "center", padding: "8px", color: "#1565C0" }}>{formatCurrency(s.tipCard)}</td>
                    {CUSTOMER_TYPES.map(ct => <td key={ct} style={{ textAlign: "center", padding: "8px", color: "#555" }}>{s.byType[ct]||0}</td>)}
                    <td style={{ textAlign: "center", padding: "8px", color: "#1565C0" }}>{s.ticketClients}</td>
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
                if ((dep.clientName || "").toLowerCase().trim() === name && dep.type === "deposit") {
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

      {toast && (
        <div style={{ position: "fixed", bottom: 20, right: 20, padding: "12px 20px", borderRadius: 10, background: toast.type === "error" ? "#C62828" : toast.type === "info" ? "#1565C0" : "#0D4F4F", color: "#fff", fontWeight: 600, fontSize: 14, boxShadow: "0 4px 12px rgba(0,0,0,0.2)", zIndex: 9999 }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

function ApptCard({ appt, onClick }) {
  const typeColors = { RL: "#4CAF50", RT: "#2196F3", NL: "#FF9800", NT: "#9C27B0" };
  const isTicket = appt.isTicket;
  const isCavSlot = appt.isCavSlot;

  const isGiftCard = appt.isGiftCard;
  const borderColor = isCavSlot ? "#9C27B0" : isGiftCard ? "#F59E0B" : isTicket ? "#1565C0" : appt.paymentType === "cash" ? "#4CAF50" : "#2196F3";
  const bg = isCavSlot ? "#F9F0FF" : isGiftCard ? "#FFFDE7" : isTicket ? "#EEF5FF" : appt.paymentType === "cash" ? "#F1FBF3" : "#EEF5FF";

  if (isCavSlot) {
    const cavTotal = Number(appt.price||0) + Number(appt.tip||0);
    return (
      <div onClick={onClick} style={{ background: bg, border: `2px solid ${borderColor}`, borderRadius: 8, padding: "5px 8px", marginBottom: 3, cursor: "pointer", opacity: 0.9 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A" }}>⚡ 機械 {appt.isTicket ? "🎟️" : ""}</div>
        <div style={{ fontSize: 11, color: "#555" }}>
          {appt.clientName} <span style={{ color: "#888", fontSize: 10 }}>({appt.startTime}〜)</span>
        </div>
        <div style={{ fontSize: 11, color: "#6A1B9A", fontWeight: 700 }}>
          施術${appt.price}{appt.tip > 0 ? ` チップ$${appt.tip}` : ""}
          {cavTotal > 0 && <span style={{ color: "#0D4F4F", fontWeight: 800 }}> ＝${cavTotal}</span>}
        </div>
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
      {isTicket ? (
        <div style={{ fontSize: 11, marginTop: 2 }}>
          <div style={{ color: "#1565C0", fontWeight: 700 }}>
            🎟️ {appt.serviceName || appt.ticketMenu} {appt.ticketCurrent > 0 ? `${appt.ticketCurrent}/${appt.ticketTotal}` : `（${appt.ticketTotal}回購入）`}
            {appt.cavTherapist && !isDualLicense(appt.therapist) && (
              <span style={{ color: "#6A1B9A", fontWeight: 400, fontSize: 10 }}> +{appt.cavTherapist}</span>
            )}
          </div>
          {(appt.price > 0 || appt.tip > 0) && (() => {
            const svc = Number(appt.price||0) + Number(appt.cavPrice||0);
            const tip = Number(appt.tip||0) + Number(appt.cavTip||0);
            const extra = Number(appt.extraTip||0);
            return (
              <div>
                <span style={{ color: "#C62828" }}>施術${svc}</span>
                <span style={{ color: "#E65100" }}> チップ${tip}</span>
                <span style={{ color: "#0D4F4F", fontWeight: 800 }}> ＝${svc + tip}</span>
                {extra > 0 && <span style={{ color: "#F57F17", fontWeight: 700 }}> +${extra}💝{appt.extraTipPaymentType==="card"?"💳":"💵"}</span>}
              </div>
            );
          })()}
        </div>
      ) : (
        <div style={{ fontSize: 11, marginTop: 2 }}>
          <div style={{ color: isGiftCard ? "#B45309" : "#222", fontWeight: 700, lineHeight: 1.4 }}>
            {isGiftCard && <span style={{ fontSize: 10, background: "#F59E0B", color: "#fff", padding: "1px 5px", borderRadius: 6, marginRight: 4 }}>🎁GC消化</span>}
            {appt.serviceName || `${appt.duration}分`}
          </div>
          {(appt.price > 0 || appt.tip > 0) && (() => {
            const svc = Number(appt.price||0) + Number(appt.cavPrice||0);
            const tip = Number(appt.tip||0) + Number(appt.cavTip||0);
            const total = svc + tip;
            return (
              <div>
                <span style={{ color: "#C62828", fontWeight: 700 }}>施術${svc}{isGiftCard ? "🎁" : appt.paymentType === "card" ? "💳" : "💵"}</span>
                {!isGiftCard && Number(appt.giftCardUsed) > 0 && <span style={{ color: "#00796B", fontSize: 10 }}> (GC🎁${appt.giftCardUsed})</span>}
                {tip > 0 && <span style={{ color: "#E65100" }}> チップ${tip}{isGiftCard ? "🎁" : appt.tipPaymentType==="card"?"💳":"💵"}</span>}
                {total > 0 && <span style={{ color: "#0D4F4F", fontWeight: 800 }}> ＝${total}</span>}
                {Number(appt.depositApplied) > 0 && (
                  <span style={{ fontSize: 10, background: "#C8E6C9", color: "#1B5E20", padding: "1px 5px", borderRadius: 6, fontWeight: 700, marginLeft: 3 }}>
                    💰payroll${svc + Number(appt.depositApplied)}
                  </span>
                )}
              </div>
            );
          })()}
          {appt.cavTherapist && !isDualLicense(appt.therapist) && (
            <div style={{ color: "#6A1B9A", fontSize: 10 }}>（with {appt.cavTherapist} 機械）</div>
          )}
          {(appt.addons || []).length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginTop: 3 }}>
              {(appt.addons || []).map(addon => {
                const label = addon.serviceName || addon.name || "オプション";
                const total = Number(addon.price||0) + Number(addon.tip||0);
                return (
                  <span key={addon.id} style={{ fontSize: 9, background: "#E0F2F1", color: "#00695C", padding: "1px 6px", borderRadius: 8, fontWeight: 700 }}>
                    ➕ {label}{total > 0 ? ` $${total}` : ""}
                  </span>
                );
              })}
            </div>
          )}
        </div>
      )}
      {(appt.purchaseTags || []).length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginTop: 4 }}>
          {(appt.purchaseTags || []).map(tagId => {
            const tag = PURCHASE_TAGS.find(t => t.id === tagId);
            if (!tag) return null;
            let amt = null;
            if (tagId === "newTicket") {
              const total = Number(appt.newTicketAmount||0) + Number(appt.newTicketTip||0);
              if (total > 0) amt = `$${total}`;
            } else if (tagId === "retail") {
              const total = Number(appt.retailPurchaseAmount||0);
              if (total > 0) amt = `$${total}`;
            } else if (tagId === "giftCard") {
              const total = Number(appt.giftCardPurchaseAmount||0);
              if (total > 0) amt = `$${total}`;
            }
            return (
              <span key={tagId} style={{ fontSize: 9, background: tag.bg, color: tag.color, padding: "2px 6px", borderRadius: 8, fontWeight: 700 }}>
                {tag.label}{amt ? ` ${amt}` : ""}
              </span>
            );
          })}
        </div>
      )}
    </div>
  );
}

function ApptModal({ appt, onSave, onDelete, onClose, clientDeposits = [] }) {
  const [form, setForm] = useState({ tipPaymentType: "cash", ...appt });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const [addonPick, setAddonPick] = useState("");

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

  const autoDetectAndApplyTicket = () => {
    if (!form.ticketMenu && form.serviceName) {
      const name = form.serviceName.toLowerCase();
      let prefix = null, dur = null;
      if (name.includes("improving posture")) prefix = "IP";
      else if (name.includes("weight loss")) prefix = "WL";
      else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "FA";
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
        else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "FA";
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

  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#0D4F4F" }}>✏️ 予約編集</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer", color: "#888" }}>✕</button>
      </div>

      {clientDeposits.length > 0 && (
        <div style={{ background: "#FFF3E0", borderRadius: 10, padding: "10px 14px", marginBottom: 12, border: "2px solid #FF9800" }}>
          <div style={{ fontWeight: 800, fontSize: 13, color: "#E65100" }}>💰 デポジットあり！</div>
          {clientDeposits.map((d, i) => (
            <div key={i} style={{ fontSize: 12, color: "#BF360C", marginTop: 4 }}>
              <strong>{d.sheetDate}</strong> — ${d.amount} ({d.paymentType === "cash" ? "現金" : "カード"})
              {d.appointmentDate ? ` → 予約：${d.appointmentDate}${d.appointmentTime ? ` ${d.appointmentTime}` : ""}` : ""}
              {d.notes ? ` — ${d.notes}` : ""}
            </div>
          ))}
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="お客様名">
          <input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="お客様名" />
        </Field>
        <Field label="コース名">
          <select value={form.serviceName || ""} onChange={e => {
            const name = e.target.value;
            const svc = SQUARE_SERVICES.find(s => s.name === name);
            setForm(f => ({ ...f, serviceName: name, ...(svc ? { duration: svc.duration } : {}) }));
          }} style={{ ...inputStyle, color: "#1a1a1a", fontWeight: 600 }}>
            <option value="">コースを選択</option>
            {/* Show current value if not in list */}
            {form.serviceName && !SQUARE_SERVICES.find(s => s.name === form.serviceName) && (
              <option value={form.serviceName}>{form.serviceName}</option>
            )}
            {SQUARE_SERVICES.map(s => (
              <option key={s.name} value={s.name}>{s.name}</option>
            ))}
          </select>
          {form.serviceName && (
            <div style={{ fontSize: 11, color: "#0D4F4F", marginTop: 4, fontWeight: 600 }}>
              ✅ {form.serviceName} ({form.duration}分)
            </div>
          )}
        </Field>
        <Field label="セラピスト（ボディ）">
          <select value={form.therapist} onChange={e => set("therapist", e.target.value)} style={inputStyle}>
            <option value="">選択</option>
            {THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="開始時間">
          <input type="time" value={form.startTime} onChange={e => set("startTime", e.target.value)} style={inputStyle} />
        </Field>

        {/* Ticket toggle */}
        <div style={{ display: "flex", gap: 4 }}>
          <button onClick={() => setForm(f => ({...f, isTicket: false, isSameDayTicket: false, isGiftCard: false}))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${!form.isTicket && !form.isGiftCard ? "#C62828" : "#DDD"}`, background: !form.isTicket && !form.isGiftCard ? "#FFEBEE" : "#fff", cursor: "pointer", fontWeight: 700, color: !form.isTicket && !form.isGiftCard ? "#C62828" : "#888", fontSize: 11 }}>
            🔴 通常施術
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: false, isGiftCard: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#DDD"}`, background: form.isTicket && !form.isSameDayTicket ? "#E3F2FD" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isTicket && !form.isSameDayTicket ? "#1565C0" : "#888", fontSize: 11 }}>
            🔵 チケット消化
          </button>
          <button onClick={() => { setForm(f => ({...f, isTicket: true, isSameDayTicket: true, useToday: true, ticketCurrent: 1, isGiftCard: false})); autoDetectAndApplyTicket(); }} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isSameDayTicket ? "#2E7D32" : "#DDD"}`, background: form.isSameDayTicket ? "#E8F5E9" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isSameDayTicket ? "#2E7D32" : "#888", fontSize: 11 }}>
            🟢 当日購入
          </button>
          <button onClick={() => setForm(f => ({...f, isGiftCard: true, isTicket: false, isSameDayTicket: false}))} style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.isGiftCard ? "#F59E0B" : "#DDD"}`, background: form.isGiftCard ? "#FFFDE7" : "#fff", cursor: "pointer", fontWeight: 700, color: form.isGiftCard ? "#B45309" : "#888", fontSize: 11 }}>
            🎁 GC消化
          </button>
        </div>

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
                    else if (name.includes("facial") || name.includes("glow") || name.includes("sculpt")) prefix = "FA";
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

            {/* Same-day purchase payment — shown first so user enters package price before selecting course details */}
            {form.isSameDayTicket && (
              <div style={{ marginBottom: 10, background: "#E8F5E9", borderRadius: 8, padding: 10, border: "1px solid #A5D6A7" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#2E7D32", marginBottom: 8 }}>
                  💰 当日購入 — パッケージ代金（売上計上）
                  {form.useToday === false && <span style={{ color: "#888", fontWeight: 400 }}>（次回から使用）</span>}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="パッケージ代金 ($)">
                    <input type="number" value={form.packagePrice || ""} onChange={e => set("packagePrice", e.target.value)} style={inputStyle} placeholder="例: 432" />
                  </Field>
                  <Field label={form.useToday === false ? "チップ ($)" : "当日チップ ($)"}>
                    <input type="number" value={form.packageTip ?? ""} onChange={e => set("packageTip", Number(e.target.value))} style={inputStyle} placeholder="例: 91" />
                  </Field>
                </div>
                <div style={{ marginTop: 8 }}>
                  <Field label="支払方法（施術代金）"><PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} /></Field>
                </div>
                <div style={{ marginTop: 8 }}>
                  <Field label="チップ支払方法"><PaymentToggle value={form.tipPaymentType||"cash"} onChange={v => set("tipPaymentType", v)} /></Field>
                </div>
                {form.useToday !== false && (form.price > 0 || form.cavPrice > 0) && (
                  <div style={{ marginTop: 6, fontSize: 11, color: "#555", background: "#fff", borderRadius: 6, padding: 6 }}>
                    スタッフ振り分け（1回目）：{form.therapist} 施術${form.price||0} / チップ${form.tip||0}
                    {form.cavPrice > 0 ? ` ＋ ${form.cavTherapist} 施術$${form.cavPrice} / チップ$${form.cavTip||0}` : ""}
                  </div>
                )}
              </div>
            )}

            {/* Treatment menu selector */}
            <Field label="施術メニュー">
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
                      border: `2px solid ${selected ? "#1565C0" : "#DDD"}`,
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
            <Field label="料金バージョン">
              <div style={{ display: "flex", gap: 8 }}>
                {[["new","🆕 新料金（2月以降）"],["old","📋 旧料金（2月以前）"]].map(([v,label]) => (
                  <button key={v} onClick={() => { set("priceVersion", v); if(form.ticketMenu) applyTicketPrices(form.ticketMenu, form.ticketTotal, v, form.cavTherapist, form.therapist); }}
                    style={{ flex: 1, padding: "8px 4px", borderRadius: 8, border: `2px solid ${form.priceVersion===v?"#1565C0":"#DDD"}`, background: form.priceVersion===v?"#BBDEFB":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: form.priceVersion===v?"#1565C0":"#888" }}>
                    {label}
                  </button>
                ))}
              </div>
            </Field>

            {/* Total sessions */}
            <Field label="コース回数">
              <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                {[3,5].map(n => (
                  <button key={n} onClick={() => {
                    set("ticketTotal", n);
                    if(form.ticketMenu) {
                      const parts = form.ticketMenu.split("-");
                      applyTicketPrices(`${parts[0]}-${parts[1]}-${n}`, n, form.priceVersion||"new", form.cavTherapist, form.therapist);
                    }
                  }} style={{ flex: 1, padding: "8px", borderRadius: 8, border: `2px solid ${form.ticketTotal===n?"#1565C0":"#DDD"}`, background: form.ticketTotal===n?"#1565C0":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketTotal===n?"#fff":"#888" }}>
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
              <Field label="今日は何回目？">
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 }}>
                  {Array.from({ length: form.ticketTotal||3 }, (_,i) => i+1).map(n => (
                    <button key={n} onClick={() => set("ticketCurrent", n)}
                      style={{ padding: "6px 14px", borderRadius: 8, border: `2px solid ${form.ticketCurrent===n?"#0D4F4F":"#DDD"}`, background: form.ticketCurrent===n?"#0D4F4F":"#fff", cursor: "pointer", fontWeight: 700, color: form.ticketCurrent===n?"#fff":"#888" }}>
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
                    <Field label={`${form.therapist} 施術 ($)`}><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={inputStyle} /></Field>
                    <Field label={`${form.therapist} チップ ($)`}><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                    <Field label={`${form.cavTherapist} 施術 ($)`}><input type="number" value={form.cavPrice || ""} onChange={e => set("cavPrice", e.target.value)} style={inputStyle} /></Field>
                    <Field label={`${form.cavTherapist} チップ ($)`}><input type="number" value={form.cavTip || ""} onChange={e => set("cavTip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label="施術 合計 ($)"><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={inputStyle} /></Field>
                    <Field label="チップ 合計 ($)"><input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} /></Field>
                  </div>
                )}
              </div>
            )}

            {/* Extra tip for チケット消化 (not same-day purchase) */}
            {!form.isSameDayTicket && (
              <div style={{ marginTop: 10, background: "#FFF8E1", borderRadius: 8, padding: 10, border: "1px solid #FFE082" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#F57F17", marginBottom: 8 }}>💝 エクストラ（任意）</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                  <Field label="追加チップ ($)">
                    <input type="number" value={form.extraTip || ""} onChange={e => set("extraTip", Number(e.target.value) || 0)} style={inputStyle} placeholder="例: 20" />
                  </Field>
                  <Field label="物販（別途Retailに入力）">
                    <div style={{ fontSize: 11, color: "#888", padding: "10px 0" }}>下の物販セクションへ</div>
                  </Field>
                </div>
                {(form.extraTip > 0) && (
                  <div style={{ marginTop: 8 }}>
                    <Field label="チップ支払方法"><PaymentToggle value={form.extraTipPaymentType || "cash"} onChange={v => set("extraTipPaymentType", v)} /></Field>
                  </div>
                )}
              </div>
            )}

            {/* Ticket grand total display */}
            {form.ticketMenu && (() => {
              const svc = Number(form.price||0) + Number(form.cavPrice||0);
              const tip = Number(form.tip||0) + Number(form.cavTip||0);
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

          </div>
        )}

        {/* Regular service */}
        {!form.isTicket && (
          <>
            {/* Cav therapist picker for regular service */}
            {!isCavCapable(form.therapist) && (
              <Field label="機械担当セラピスト（任意）">
                <select value={form.cavTherapist} onChange={e => {
                  set("cavTherapist", e.target.value);
                  // recalc split when cav therapist changes
                  set("cavTherapist", e.target.value);
                }} style={inputStyle}>
                  <option value="">なし（機械なし）</option>
                  {CAV_CAPABLE.filter(t => t !== form.therapist).map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
            )}

            {/* Total service + tip input */}
            <div style={{ background: "#F9F9F9", borderRadius: 10, padding: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "#333", marginBottom: 10 }}>
                💆 施術料金・チップ（合計入力）
                {Number(form.depositApplied) > 0 && <span style={{ fontSize: 11, color: "#2E7D32", fontWeight: 600, marginLeft: 8 }}>← もらった金額を入力</span>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <Field label="施術合計 ($)">
                  <input type="number" value={form.totalServiceInput ?? (form.price || "")}
                    onChange={e => {
                      const total = Number(e.target.value);
                      const bodyMins = Number(form.bodyMins || form.duration);
                      const cavMins = Number(form.cavMins || 0);
                      const allMins = bodyMins + cavMins;
                      if (form.cavTherapist && cavMins > 0 && allMins > 0) {
                        set("price", Math.round(total * bodyMins / allMins * 10) / 10);
                        set("cavPrice", Math.round(total * cavMins / allMins * 10) / 10);
                      } else {
                        set("price", total);
                        set("cavPrice", 0);
                      }
                      set("totalServiceInput", total);
                    }}
                    style={inputStyle} placeholder="例: 158" />
                </Field>
                <Field label="チップ合計 ($)">
                  <input type="number" value={form.totalTipInput ?? (form.tip || "")}
                    onChange={e => {
                      const total = Number(e.target.value);
                      const bodyMins = Number(form.bodyMins || form.duration);
                      const cavMins = Number(form.cavMins || 0);
                      const allMins = bodyMins + cavMins;
                      if (form.cavTherapist && cavMins > 0 && allMins > 0) {
                        set("tip", Math.round(total * bodyMins / allMins * 10) / 10);
                        set("cavTip", Math.round(total * cavMins / allMins * 10) / 10);
                      } else {
                        set("tip", total);
                        set("cavTip", 0);
                      }
                      set("totalTipInput", total);
                    }}
                    style={inputStyle} placeholder="例: 30" />
                </Field>
              </div>

              {/* WL Tip Calculator — $116 fixed cav split */}
              {isWLService(form.serviceName) && (
                <div style={{ marginTop: 10, background: "#E8F0FE", borderRadius: 8, padding: 10, border: "2px dashed #3F51B5" }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#3F51B5", marginBottom: 8 }}>
                    ⚖️ ウェイトロス振り分け計算（キャビ${WL_CAV_FIXED}固定）
                  </div>
                  {(() => {
                    const dep = Number(form.depositApplied || 0);
                    const totalCash = Number(form.totalServiceInput ?? (Number(form.price || 0) + Number(form.cavPrice || 0)));
                    const totalSvc = totalCash + dep;
                    const totalTipAmt = Number(form.totalTipInput ?? (Number(form.tip || 0) + Number(form.cavTip || 0)));
                    const cavSvc = totalSvc > 0 ? Math.min(WL_CAV_FIXED, totalSvc) : 0;
                    const bodySvc = Math.max(0, totalSvc - cavSvc);
                    const tipRate = totalSvc > 0 ? totalTipAmt / totalSvc : 0;
                    const cavTipCalc = Math.round(cavSvc * tipRate * 10) / 10;
                    const bodyTipCalc = Math.round(bodySvc * tipRate * 10) / 10;
                    const tipPct = Math.round(tipRate * 1000) / 10;
                    const hasSplit = !!form.cavTherapist;
                    const cavCash = totalSvc > 0 ? Math.round(totalCash * cavSvc / totalSvc * 10) / 10 : 0;
                    const bodyCash = totalSvc > 0 ? Math.round(totalCash * bodySvc / totalSvc * 10) / 10 : 0;
                    return (
                      <>
                        {totalSvc > 0 && totalTipAmt > 0 && (
                          <div style={{ fontSize: 11, color: "#3F51B5", marginBottom: 8, fontWeight: 600 }}>
                            チップ率 <strong>{tipPct}%</strong>　（チップ${totalTipAmt} ÷ 施術${totalSvc}{dep > 0 ? `（デポ$${dep}含む）` : ''}）
                          </div>
                        )}
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                          <div style={{ background: "#F3E5F5", borderRadius: 8, padding: "8px 10px" }}>
                            <div style={{ fontWeight: 700, color: "#6A1B9A", fontSize: 11, marginBottom: 4 }}>
                              ⚡ キャビ（40min）{hasSplit && form.cavTherapist ? <span style={{ color: "#888", fontWeight: 400 }}> → {form.cavTherapist}</span> : form.therapist ? <span style={{ color: "#888", fontWeight: 400 }}> → {form.therapist}</span> : ""}
                            </div>
                            {dep > 0
                              ? <div><span style={{ background: "#6A1B9A", color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 800 }}>payroll ${cavSvc}</span></div>
                              : <div>施術 <strong style={{ color: "#6A1B9A" }}>${cavSvc}</strong></div>
                            }
                            {totalTipAmt > 0 && <div style={{ marginTop: 4 }}>チップ <strong style={{ color: "#6A1B9A" }}>${cavTipCalc}</strong></div>}
                          </div>
                          <div style={{ background: "#E8F5E9", borderRadius: 8, padding: "8px 10px" }}>
                            <div style={{ fontWeight: 700, color: "#2E7D32", fontSize: 11, marginBottom: 4 }}>
                              💚 ボディ（残り）{form.therapist ? <span style={{ color: "#888", fontWeight: 400 }}> → {form.therapist}</span> : ""}
                            </div>
                            {dep > 0
                              ? <div><span style={{ background: "#C62828", color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 800 }}>payroll ${bodySvc}</span></div>
                              : <div>施術 <strong style={{ color: "#2E7D32" }}>${bodySvc}</strong></div>
                            }
                            {totalTipAmt > 0 && <div style={{ marginTop: 4 }}>チップ <strong style={{ color: "#2E7D32" }}>${bodyTipCalc}</strong></div>}
                          </div>
                        </div>
                        {dep > 0 && totalSvc > 0 && (
                          <div style={{ marginTop: 6, fontSize: 11, color: "#888", borderTop: "1px solid #C5CAE9", paddingTop: 6 }}>
                            受取 ${totalCash} ＋ デポジット ${dep} ＝ 施術合計 <strong style={{ color: "#3F51B5" }}>${totalSvc}</strong>
                          </div>
                        )}
                        {totalSvc > 0 && hasSplit && (
                          <button
                            onClick={() => setForm(f => ({
                              ...f,
                              price: dep > 0 ? bodyCash : bodySvc,
                              cavPrice: dep > 0 ? cavCash : cavSvc,
                              tip: bodyTipCalc,
                              cavTip: cavTipCalc,
                              totalServiceInput: totalCash,
                              totalTipInput: totalTipAmt,
                            }))}
                            style={{ marginTop: 10, width: "100%", padding: "9px", borderRadius: 8, border: "none", background: "#3F51B5", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}
                          >
                            ✅ {form.cavTherapist}（キャビ）＋ {form.therapist}（ボディ）に設定する
                          </button>
                        )}
                        {totalSvc > 0 && !hasSplit && isDualLicense(form.therapist) && (
                          <div style={{ marginTop: 8, fontSize: 11, color: "#3F51B5" }}>
                            ※ {form.therapist}はデュアルライセンス → 全額{form.therapist}に計上（内訳参照用）
                          </div>
                        )}
                      </>
                    );
                  })()}
                </div>
              )}

              {/* Minutes split — only shown when cav therapist selected, not for WL (uses fixed $116 split instead) */}
              {form.cavTherapist && !isWLService(form.serviceName) && (
                <div style={{ marginTop: 10, background: "#EEF4FF", borderRadius: 8, padding: 10 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#1565C0", marginBottom: 8 }}>⏱️ 担当分数を入力 → 自動振り分け</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <Field label={`${form.therapist} 担当（分）`}>
                      <input type="number" value={form.bodyMins || ""}
                        onChange={e => {
                          const bodyMins = Number(e.target.value);
                          const cavMins = Number(form.duration) - bodyMins;
                          const allMins = Number(form.duration);
                          const svcTotal = Number(form.totalServiceInput ?? (Number(form.price) + Number(form.cavPrice || 0)));
                          const tipTotal = Number(form.totalTipInput ?? (Number(form.tip) + Number(form.cavTip || 0)));
                          setForm(f => ({
                            ...f,
                            bodyMins,
                            cavMins,
                            price: allMins > 0 ? Math.round(svcTotal * bodyMins / allMins * 10) / 10 : f.price,
                            cavPrice: allMins > 0 ? Math.round(svcTotal * cavMins / allMins * 10) / 10 : f.cavPrice,
                            tip: allMins > 0 ? Math.round(tipTotal * bodyMins / allMins * 10) / 10 : f.tip,
                            cavTip: allMins > 0 ? Math.round(tipTotal * cavMins / allMins * 10) / 10 : f.cavTip,
                          }));
                        }}
                        style={inputStyle} placeholder={`例: ${form.duration - 15}`} />
                    </Field>
                    <Field label={`${form.cavTherapist} 担当（分）`}>
                      <input type="number" value={form.cavMins || ""}
                        onChange={e => {
                          const cavMins = Number(e.target.value);
                          const bodyMins = Number(form.duration) - cavMins;
                          const allMins = Number(form.duration);
                          const svcTotal = Number(form.totalServiceInput ?? (Number(form.price) + Number(form.cavPrice || 0)));
                          const tipTotal = Number(form.totalTipInput ?? (Number(form.tip) + Number(form.cavTip || 0)));
                          setForm(f => ({
                            ...f,
                            bodyMins,
                            cavMins,
                            price: allMins > 0 ? Math.round(svcTotal * bodyMins / allMins * 10) / 10 : f.price,
                            cavPrice: allMins > 0 ? Math.round(svcTotal * cavMins / allMins * 10) / 10 : f.cavPrice,
                            tip: allMins > 0 ? Math.round(tipTotal * bodyMins / allMins * 10) / 10 : f.tip,
                            cavTip: allMins > 0 ? Math.round(tipTotal * cavMins / allMins * 10) / 10 : f.cavTip,
                          }));
                        }}
                        style={inputStyle} placeholder="例: 15" />
                    </Field>
                  </div>

                  {/* Show calculated split */}
                  {(form.bodyMins > 0 || form.cavMins > 0) && (() => {
                    const dep = Number(form.depositApplied || 0);
                    const bodyRcv = Number(form.price || 0);
                    const cavRcv = Number(form.cavPrice || 0);
                    const tipBody = Number(form.tip || 0);
                    const tipCav = Number(form.cavTip || 0);
                    const totalRcv = bodyRcv + cavRcv;
                    const fullPayroll = totalRcv + dep;
                    const bodyPayroll = dep > 0 && totalRcv > 0
                      ? Math.round(fullPayroll * bodyRcv / totalRcv * 10) / 10
                      : bodyRcv;
                    const cavPayroll = dep > 0 && totalRcv > 0
                      ? Math.round(fullPayroll * cavRcv / totalRcv * 10) / 10
                      : cavRcv;
                    return (
                      <div style={{ marginTop: 8, background: "#fff", borderRadius: 8, padding: 8, fontSize: 12 }}>
                        <div style={{ color: "#2E7D32", marginBottom: 6 }}>
                          💚 <strong>{form.therapist}</strong>（{form.bodyMins}分）：
                          {dep > 0
                            ? <><span style={{ background: "#C62828", color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 800 }}>payroll ${bodyPayroll}</span><span style={{ fontSize: 10, color: "#888", marginLeft: 6 }}>チップ ${tipBody}</span></>
                            : <> 施術 <strong>${bodyRcv}</strong> / チップ <strong>${tipBody}</strong></>
                          }
                        </div>
                        <div style={{ color: "#6A1B9A" }}>
                          ⚡ <strong>{form.cavTherapist}</strong>（{form.cavMins}分 機械）：
                          {dep > 0
                            ? <><span style={{ background: "#6A1B9A", color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 800 }}>payroll ${cavPayroll}</span><span style={{ fontSize: 10, color: "#888", marginLeft: 6 }}>チップ ${tipCav}</span></>
                            : <> 施術 <strong>${cavRcv}</strong> / チップ <strong>${tipCav}</strong></>
                          }
                        </div>
                        {dep > 0 && (
                          <div style={{ marginTop: 6, borderTop: "1px solid #EEE", paddingTop: 6, fontSize: 11, color: "#888" }}>
                            受取 ${totalRcv} ＋ デポジット ${dep} ＝ payroll合計 <strong style={{ color: "#C62828" }}>${fullPayroll}</strong>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              )}

              {!form.isGiftCard && (
              <div style={{ marginTop: 10, background: "#F0F4FF", borderRadius: 10, padding: "10px 12px" }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#1565C0", marginBottom: 8 }}>💳 支払い方法（施術・チップ別々に設定可）</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <Field label={Number(form.giftCardUsed) > 0 ? "施術（差額）支払方法" : "施術 支払方法"}>
                    <PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} />
                  </Field>
                  <Field label="チップ 支払方法">
                    <PaymentToggle value={form.tipPaymentType||"cash"} onChange={v => set("tipPaymentType", v)} />
                  </Field>
                </div>
              </div>
              )}

              {/* Gift Card Usage — hidden for GC消化 (entire appointment IS the GC) */}
              {!form.isGiftCard && <div style={{ background: "#E0F2F1", borderRadius: 10, padding: 12, marginTop: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 700, fontSize: 13, color: "#00796B" }}>🎁 ギフトカード使用</span>
                  <button onClick={() => set("giftCardUsed", Number(form.giftCardUsed) > 0 ? 0 : "")}
                    style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.giftCardUsed) > 0 ? "#00796B" : "#B2DFDB", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                    {Number(form.giftCardUsed) > 0 ? "ON ✓" : "OFF"}
                  </button>
                </div>
                {(Number(form.giftCardUsed) > 0 || form.giftCardUsed === "") && (
                  <div style={{ marginTop: 8 }}>
                    <div style={{ fontSize: 10, color: "#00796B", marginBottom: 4 }}>使用金額 ($)</div>
                    <input type="number" value={form.giftCardUsed === "" ? "" : form.giftCardUsed}
                      onChange={e => set("giftCardUsed", e.target.value)}
                      style={{ ...inputStyle, borderColor: "#80CBC4" }} placeholder="例：150" />
                    {(() => {
                      const gc = Number(form.giftCardUsed || 0);
                      const svc = Number(form.price || 0);
                      const tip = Number(form.tip || 0);
                      const retail = (form.purchaseTags||[]).includes("retail") ? Number(form.retailPurchaseAmount || 0) : 0;
                      const total = svc + tip + retail;
                      if (!gc || !total) return null;
                      const remainder = Math.max(0, total - gc);
                      return (
                        <div style={{ marginTop: 6, fontSize: 12, color: "#00796B" }}>
                          <div style={{ color: "#555", marginBottom: 2 }}>
                            施術${svc}{tip > 0 ? ` ＋ チップ$${tip}` : ""}{retail > 0 ? ` ＋ 物販$${retail}` : ""} ＝ <strong>${total}</strong>
                          </div>
                          🎁 GC <strong>${Math.min(gc, total)}</strong>
                          {remainder > 0
                            ? <> ＋ {form.paymentType === "card" ? "💳 カード" : "💵 現金"} <strong>${remainder}</strong></>
                            : <span style={{ color: "#2E7D32" }}> → 全額ギフトカード</span>}
                        </div>
                      );
                    })()}
                  </div>
                )}
              </div>}
            </div>

            {/* デポジット使用 */}
            {!form.isGiftCard && (
              <div style={{ background: "#E8F5E9", borderRadius: 10, padding: 12, border: "1.5px solid #A5D6A7" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontWeight: 700, fontSize: 13, color: "#2E7D32" }}>💰 デポジット使用</span>
                  <button onClick={() => set("depositApplied", Number(form.depositApplied) > 0 ? 0 : 20)}
                    style={{ padding: "4px 12px", borderRadius: 8, border: "none", background: Number(form.depositApplied) > 0 ? "#2E7D32" : "#C8E6C9", color: "#fff", fontWeight: 700, cursor: "pointer", fontSize: 12 }}>
                    {Number(form.depositApplied) > 0 ? "ON ✓" : "OFF"}
                  </button>
                </div>
                {Number(form.depositApplied) > 0 && (() => {
                  const receivedSvc = Number(form.price || 0) + Number(form.cavPrice || 0);
                  const dep = Number(form.depositApplied);
                  const payrollSvc = receivedSvc + dep;
                  const tip = Number(form.tip||0) + Number(form.cavTip||0);
                  return (
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: "#2E7D32", marginBottom: 4 }}>デポジット金額 ($)</div>
                      <input type="number" value={form.depositApplied}
                        onChange={e => set("depositApplied", Number(e.target.value) || 0)}
                        style={{ ...inputStyle, borderColor: "#81C784" }} placeholder="例：20" />
                    </div>
                  );
                })()}
              </div>
            )}

            {/* GC消化 banner — replaces payment type selectors */}
            {form.isGiftCard && (
              <div style={{ background: "#FFFDE7", borderRadius: 10, padding: 12, border: "2px solid #F59E0B", marginTop: 0 }}>
                <div style={{ fontWeight: 800, fontSize: 13, color: "#B45309" }}>🎁 GC消化 — 全額ギフトカードにて領収済</div>
                <div style={{ fontSize: 11, color: "#92400E", marginTop: 4 }}>事前購入のギフトカードを使用。本日の売上には計上されません。</div>
              </div>
            )}


            {/* Regular service grand total display */}
            {(() => {
              const svc = Number(form.price||0) + Number(form.cavPrice||0);
              const tip = Number(form.tip||0) + Number(form.cavTip||0);
              if (svc + tip === 0) return null;
              if (form.isGiftCard) {
                return (
                  <div style={{ background: "#78350F", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                    <div style={{ fontSize: 11, color: "#FDE68A" }}>
                      🎁 施術 <strong style={{ color: "#fff" }}>${svc}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}</strong>
                      {form.cavTherapist && !isDualLicense(form.therapist) && (
                        <span style={{ fontSize: 10, color: "#FDE68A", display: "block" }}>
                          {form.therapist} ${form.price||0} / {form.cavTherapist} ${form.cavPrice||0}
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
              return (
                <div style={{ background: "#0D4F4F", borderRadius: 8, padding: "10px 14px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                  <div style={{ fontSize: 11, color: "#B2EBF2" }}>
                    施術 <strong style={{ color: "#fff" }}>${svc}{form.paymentType==="card"?"💳":"💵"}</strong>　チップ <strong style={{ color: "#fff" }}>${tip}{form.tipPaymentType==="card"?"💳":"💵"}</strong>
                    {form.cavTherapist && !isDualLicense(form.therapist) && (
                      <span style={{ fontSize: 10, color: "#80CBC4", display: "block" }}>
                        {form.therapist} ${form.price||0} / {form.cavTherapist} ${form.cavPrice||0}
                      </span>
                    )}
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ color: "#B2EBF2", fontSize: 11 }}>合計　</span>
                    <span style={{ color: "#fff", fontSize: 22, fontWeight: 800 }}>${svc + tip}</span>
                  </div>
                </div>
              );
            })()}
          </>
        )}

        {/* ── 当日の追加購入 ── */}
        {(() => {
          const tags = form.purchaseTags || [];
          const hasTag = id => tags.includes(id);
          const toggleTag = id => set("purchaseTags", hasTag(id) ? tags.filter(t => t !== id) : [...tags, id]);
          const payBtns = (field, active, activeColor) => (
            <div style={{ display: "flex", gap: 5 }}>
              {[["cash","💵 Cash"],["card","💳 Card"],["check","🗒️ Check"]].map(([v,l]) => (
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

              {/* 🛍️ 物販購入 fields */}
              {hasTag("retail") && (
                <div style={{ background: "#F3E5F5", padding: "10px 12px", borderTop: "1px solid #E1BEE7", display: "flex", flexDirection: "column", gap: 8 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#6A1B9A" }}>🛍️ 物販購入 — 詳細</div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>商品名</div>
                    <select value={form.retailProductName || ""} onChange={e => {
                      const prod = RETAIL_PRODUCTS.find(p => p.name === e.target.value);
                      setForm(f => ({ ...f, retailProductName: e.target.value, retailPurchaseAmount: prod?.price > 0 ? prod.price : f.retailPurchaseAmount }));
                    }} style={{ ...inputStyle, borderColor: "#CE93D8" }}>
                      <option value="">— 商品を選択 —</option>
                      {RETAIL_PRODUCTS.map(p => (
                        <option key={p.name} value={p.name}>{p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
                      ))}
                      <option value="__other__">その他</option>
                    </select>
                    {form.retailProductName === "__other__" && (
                      <input type="text" placeholder="商品名を入力" style={{ ...inputStyle, borderColor: "#CE93D8", marginTop: 4 }}
                        onChange={e => set("retailProductName", e.target.value === "__other__" ? "" : e.target.value)} />
                    )}
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>金額 ($)</div>
                    <input type="number" value={form.retailPurchaseAmount || ""} onChange={e => set("retailPurchaseAmount", e.target.value)}
                      style={{ ...inputStyle, borderColor: "#CE93D8" }} placeholder="例：30" />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法</div>
                    {payBtns("retailPurchasePaymentType", true, "#6A1B9A")}
                  </div>
                  <div style={{ fontSize: 10, color: "#888" }}>担当：{form.therapist}{form.cavTherapist ? ` / ${form.cavTherapist}` : ""}</div>
                </div>
              )}

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

        <Field label="顧客タイプ">
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {CUSTOMER_TYPES.map(ct => {
              const colors = { RL:"#4CAF50", RT:"#2196F3", NL:"#FF9800", NT:"#9C27B0" };
              const labels = { RL:"リピ・ローカル", RT:"リピ・トラベラー", NL:"ニュー・ローカル", NT:"ニュー・トラベラー" };
              return (
                <button key={ct} onClick={() => set("customerType", ct)}
                  style={{ flex: 1, minWidth: 90, padding: "7px 4px", borderRadius: 8, border: `2px solid ${form.customerType===ct?colors[ct]:"#DDD"}`, background: form.customerType===ct?colors[ct]:"#fff", cursor: "pointer", fontWeight: 700, fontSize: 11, color: form.customerType===ct?"#fff":"#888" }}>
                  {ct}<br /><span style={{ fontWeight: 400, fontSize: 9 }}>{labels[ct]}</span>
                </button>
              );
            })}
          </div>
        </Field>

        {/* How they found us — shown for NT and NL */}
        {(form.customerType === "NT" || form.customerType === "NL") && (
          <Field label={form.customerType === "NT" ? "📍 どこで知りましたか？（NT）" : "📍 どこで知りましたか？（NL）"}>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {["Google / Website", "Google Map", "Instagram", "Yelp", "紹介"].map(src => (
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

        {/* ── オプション追加 ── */}
        <div style={{ background: "#F9F9F9", borderRadius: 10, padding: "10px 12px", border: "1.5px solid #E0E0E0" }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#555", marginBottom: 8 }}>➕ オプション追加（任意）</div>

          {/* Service picker */}
          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
            <select value={addonPick} onChange={e => setAddonPick(e.target.value)}
              style={{ ...inputStyle, flex: 1, fontSize: 12 }}>
              <option value="">— メニューを選択 —</option>
              {SQUARE_SERVICES.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
              <option value="__custom__">その他（カスタム入力）</option>
            </select>
            <button onClick={() => {
              if (!addonPick) return;
              set("addons", [...(form.addons||[]), {
                id: `${Date.now()}`,
                serviceName: addonPick === "__custom__" ? "" : addonPick,
                therapist: "",
                price: 0,
                tip: 0,
                paymentType: "cash",
                tipPaymentType: "cash",
              }]);
              setAddonPick("");
            }} style={{
              padding: "8px 14px", borderRadius: 8, border: "none",
              background: addonPick ? "#0D4F4F" : "#CCC",
              color: "#fff", cursor: addonPick ? "pointer" : "default", fontWeight: 700, fontSize: 13, whiteSpace: "nowrap"
            }}>＋ 追加</button>
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
                      {!addon.serviceName ? (
                        <input value={addon.serviceName||""} onChange={e => upd({ serviceName: e.target.value })}
                          style={{ ...inputStyle, flex: 1, fontSize: 12 }} placeholder="サービス名を入力" />
                      ) : (
                        <div style={{ fontWeight: 700, fontSize: 12, color: "#0D4F4F", flex: 1 }}>➕ {addon.serviceName}</div>
                      )}
                      <button onClick={() => set("addons", form.addons.filter((_,i) => i !== idx))}
                        style={{ border: "none", background: "none", cursor: "pointer", fontSize: 16, color: "#AAA", marginLeft: 6, flexShrink: 0 }}>✕</button>
                    </div>

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
                        <input type="number" value={addon.price||""} onChange={e => upd({ price: Number(e.target.value)||0 })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>チップ ($)</div>
                        <input type="number" value={addon.tip||""} onChange={e => upd({ tip: Number(e.target.value)||0 })}
                          style={{ ...inputStyle, fontSize: 12 }} placeholder="0" />
                      </div>
                    </div>

                    {/* Payment type */}
                    <div style={{ marginBottom: tipAmt > 0 ? 6 : 0 }}>
                      <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法（施術）</div>
                      <PaymentToggle value={addon.paymentType||"cash"} onChange={v => upd({ paymentType: v })} small />
                    </div>

                    {/* Tip payment type */}
                    {tipAmt > 0 && (
                      <div>
                        <div style={{ fontSize: 10, color: "#888", marginBottom: 3 }}>支払方法（チップ）</div>
                        <PaymentToggle value={addon.tipPaymentType||"cash"} onChange={v => upd({ tipPaymentType: v })} small />
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

        <Field label="メモ">
          <input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="備考など" />
        </Field>
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={() => onSave(form)} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#0D4F4F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        {appt.id && <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>}
      </div>
    </Modal>
  );
}

function RetailModal({ retail, onSave, onClose }) {
  const [form, setForm] = useState({ ...retail });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#6A1B9A" }}>🛍️ 物販</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="商品名"><input value={form.item} onChange={e => set("item", e.target.value)} style={inputStyle} placeholder="Muscle Rub など" /></Field>
        <Field label="金額 ($)"><input type="number" value={form.price || ""} onChange={e => set("price", e.target.value)} style={inputStyle} /></Field>
        <Field label="販売者"><select value={form.soldBy} onChange={e => set("soldBy", e.target.value)} style={inputStyle}><option value="">選択</option>{THERAPISTS.map(t => <option key={t} value={t}>{t}</option>)}</select></Field>
        <Field label="支払方法"><PaymentToggle value={form.paymentType} onChange={v => set("paymentType", v)} /></Field>
      </div>
      <button onClick={() => onSave(form)} style={{ width: "100%", marginTop: 16, padding: "12px", borderRadius: 10, border: "none", background: "#6A1B9A", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
    </Modal>
  );
}

function DepositModal({ deposit, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...deposit });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const isDeposit = form.type === "deposit";
  const isGiftCard = form.type === "giftcard";
  const payTypes = [["cash","💵 現金"],["card","💳 カード"],["check","🗒️ Check"]];
  return (
    <Modal onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h2 style={{ margin: 0, fontSize: 18, color: "#1565C0" }}>{isGiftCard ? "🎁 ギフトカード" : "💰 デポジット"}</h2>
        <button onClick={onClose} style={{ border: "none", background: "none", fontSize: 22, cursor: "pointer" }}>✕</button>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <Field label="種類">
          <div style={{ display: "flex", gap: 8 }}>
            {[["deposit","💰 デポジット"],["giftcard","🎁 ギフトカード"]].map(([val,label]) => (
              <button key={val} onClick={() => set("type", val)} style={{ flex: 1, padding: "10px", borderRadius: 8, border: `2px solid ${form.type===val?"#1565C0":"#DDD"}`, background: form.type===val?"#E3F2FD":"#fff", cursor: "pointer", fontWeight: 700, color: form.type===val?"#1565C0":"#888" }}>{label}</button>
            ))}
          </div>
        </Field>
        <Field label="お客様名"><input value={form.clientName} onChange={e => set("clientName", e.target.value)} style={inputStyle} placeholder="例：田中様" /></Field>
        {isDeposit && (
          <Field label="予約日時">
            <div style={{ display: "flex", gap: 8 }}>
              <input type="date" value={form.appointmentDate || ""} onChange={e => set("appointmentDate", e.target.value)} style={{ ...inputStyle, flex: 1 }} />
              <input type="time" value={form.appointmentTime || ""} onChange={e => set("appointmentTime", e.target.value)} style={{ ...inputStyle, width: 100 }} />
            </div>
            {form.appointmentDate && form.clientName && (
              <div style={{ fontSize: 11, color: "#1565C0", marginTop: 4 }}>
                {new Date(form.appointmentDate + "T00:00").toLocaleDateString("ja-JP", { month: "long", day: "numeric" })}
                {form.appointmentTime && ` ${form.appointmentTime}`}の{form.clientName}様のデポジット
              </div>
            )}
          </Field>
        )}
        <Field label="金額 ($)"><input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} /></Field>
        <div style={{ background: "#F0F4FF", borderRadius: 10, padding: "10px 12px" }}>
          <div style={{ fontWeight: 700, fontSize: 12, color: "#1565C0", marginBottom: 8 }}>💳 支払い方法（施術・チップ別々に設定可）</div>
          <Field label={isDeposit ? "デポジット 支払方法" : "金額 支払方法"}>
            <div style={{ display: "flex", gap: 6 }}>
              {payTypes.map(([val, label]) => (
                <button key={val} onClick={() => set("paymentType", val)}
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
              <Field label="チップ 支払方法">
                <div style={{ display: "flex", gap: 6 }}>
                  {payTypes.map(([val, label]) => (
                    <button key={val} onClick={() => set("tipPaymentType", val)}
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
        <button onClick={() => onSave(form)} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#1565C0", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function TicketPurchaseModal({ tp, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...tp });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const payTypes = [["cash","💵 Cash"],["card","💳 Card"],["check","🗒️ Check"]];
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
        <Field label="パッケージ名（任意）">
          <input value={form.packageName} onChange={e => set("packageName", e.target.value)} style={inputStyle} placeholder="例：IP-90 5回コース" />
        </Field>
        <Field label="施術料 ($)">
          <input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} placeholder="例：719" />
        </Field>
        <Field label="支払方法">
          <div style={{ display: "flex", gap: 6 }}>
            {payTypes.map(([val, label]) => (
              <button key={val} onClick={() => set("paymentType", val)}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#B71C1C":"#DDD"}`, background: form.paymentType===val?"#FFEBEE":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#B71C1C":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        <Field label="チップ ($)">
          <input type="number" value={form.tip || ""} onChange={e => set("tip", e.target.value)} style={inputStyle} placeholder="例：46" />
        </Field>
        {Number(form.tip) > 0 && (
          <Field label="チップ支払方法">
            <div style={{ display: "flex", gap: 6 }}>
              {payTypes.map(([val, label]) => (
                <button key={val} onClick={() => set("tipPaymentType", val)}
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
        <button onClick={() => onSave(form)} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#B71C1C", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function StaffPurchaseModal({ sp, onSave, onDelete, onClose }) {
  const [form, setForm] = useState({ ...sp });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
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
        <Field label="商品名">
          <select value={form.productName || ""} onChange={e => {
            const prod = RETAIL_PRODUCTS.find(p => p.name === e.target.value);
            setForm(f => ({ ...f, productName: e.target.value, amount: prod?.price > 0 ? prod.price : f.amount }));
          }} style={inputStyle}>
            <option value="">— 商品を選択 —</option>
            {RETAIL_PRODUCTS.map(p => (
              <option key={p.name} value={p.name}>{p.name}{p.price > 0 ? ` ($${p.price})` : ""}</option>
            ))}
            <option value="__other__">その他</option>
          </select>
          {form.productName === "__other__" && (
            <input type="text" placeholder="商品名を入力" style={{ ...inputStyle, marginTop: 4 }}
              onChange={e => set("productName", e.target.value)} />
          )}
        </Field>
        <Field label="金額 ($)"><input type="number" value={form.amount || ""} onChange={e => set("amount", e.target.value)} style={inputStyle} /></Field>
        <Field label="支払方法">
          <div style={{ display: "flex", gap: 6 }}>
            {[["cash","💵 現金"],["card","💳 カード"]].map(([val, label]) => (
              <button key={val} onClick={() => set("paymentType", val)}
                style={{ flex: 1, padding: "9px 4px", borderRadius: 8, border: `2px solid ${form.paymentType===val?"#37474F":"#DDD"}`, background: form.paymentType===val?"#ECEFF1":"#fff", cursor: "pointer", fontWeight: 700, fontSize: 12, color: form.paymentType===val?"#37474F":"#888" }}>
                {label}
              </button>
            ))}
          </div>
        </Field>
        <Field label="メモ"><input value={form.notes} onChange={e => set("notes", e.target.value)} style={inputStyle} placeholder="任意" /></Field>
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button onClick={() => onSave(form)} style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#37474F", color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer" }}>💾 保存</button>
        <button onClick={onDelete} style={{ padding: "12px 16px", borderRadius: 10, border: "none", background: "#FFEBEE", color: "#C62828", fontWeight: 700, cursor: "pointer" }}>🗑️</button>
      </div>
    </Modal>
  );
}

function SectionBox({ title, color, onAdd, children }) {
  return (
    <div style={{ marginTop: 20, background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
        <span style={{ fontWeight: 700, color, fontSize: 15 }}>{title}</span>
        <button onClick={onAdd} style={{ padding: "6px 14px", borderRadius: 8, background: color, color: "#fff", border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>+ 追加</button>
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

function Field({ label, children }) {
  return (
    <div>
      <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 5 }}>{label}</label>
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
  for (let d = 1; d <= lastDay; d++) {
    const dateStr = `${monthStr}-${String(d).padStart(2,"0")}`;
    const saved = localStorage.getItem(`spa-sheet-${dateStr}`);
    if (saved) {
      const data = JSON.parse(saved);
      const appts = (data.appointments || []).filter(a => !a.isCavSlot);
      const retails = data.retails || [];

      const cashTreatment = appts.filter(a => a.paymentType === "cash").reduce((s,a) => s + Number(a.isSameDayTicket ? (a.packagePrice||0) : (a.price||0)), 0);
      const cashProduct = retails.filter(r => r.paymentType === "cash").reduce((s,r) => s + Number(r.price||0), 0);
      const cashTip = appts.filter(a => a.tipPaymentType === "cash" || a.extraTipPaymentType === "cash").reduce((s,a) => {
        let t = 0;
        if (a.tipPaymentType === "cash") t += Number(a.isSameDayTicket ? (a.packageTip ?? a.tip ?? 0) : (a.tip||0));
        if (a.extraTipPaymentType === "cash") t += Number(a.extraTip||0);
        return s + t;
      }, 0);
      const cardTreatment = appts.filter(a => a.paymentType === "card").reduce((s,a) => s + Number(a.isSameDayTicket ? (a.packagePrice||0) : (a.price||0)), 0);
      const cardProduct = retails.filter(r => r.paymentType === "card").reduce((s,r) => s + Number(r.price||0), 0);
      const cardTip = appts.filter(a => a.tipPaymentType === "card" || a.extraTipPaymentType === "card").reduce((s,a) => {
        let t = 0;
        if (a.tipPaymentType === "card") t += Number(a.isSameDayTicket ? (a.packageTip ?? a.tip ?? 0) : (a.tip||0));
        if (a.extraTipPaymentType === "card") t += Number(a.extraTip||0);
        return s + t;
      }, 0);
      const totalTip = appts.reduce((s,a) => s + Number(a.isSameDayTicket ? (a.packageTip ?? a.tip ?? 0) : (a.tip||0)) + Number(a.extraTip||0), 0);
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
      });
    } else {
      dailyData.push({ date: d, totalSales: 0, clients: "", cashTreatment:"", cashProduct:"", totalCash:"", cashTip:"", cardTreatment:"", cardProduct:"", totalCard:"", cardTip:"", totalTip:"", grandTotal: 0 });
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

    for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
      const dateStr = d.toISOString().slice(0, 10);
      const saved = localStorage.getItem(`spa-sheet-${dateStr}`);
      if (saved) {
        const data = JSON.parse(saved);
        (data.appointments || []).forEach(a => {
          allAppts.push({ ...a, date: dateStr });
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

    allAppts.filter(a => !a.isCavSlot).forEach(a => {
      const t = a.therapist;
      if (!t || !byTherapist[t]) return;
      const depositAdd = Number(a.depositApplied || 0);
      const bodyReceived = Number(a.price || 0);
      let svc;
      if (depositAdd > 0) {
        // Find cav slot (if any) to get cav's received portion for proportional split
        const cavSlot = allAppts.find(s => s.isCavSlot && s.parentId === a.id);
        const cavReceived = cavSlot ? Number(cavSlot.price || 0) : 0;
        const totalReceived = bodyReceived + cavReceived;
        // Body therapist gets their proportion of full payroll (received + deposit)
        svc = totalReceived > 0
          ? Math.round((totalReceived + depositAdd) * bodyReceived / totalReceived * 10) / 10
          : bodyReceived + depositAdd;
      } else {
        svc = bodyReceived;
      }
      const tip = Number(a.tip || 0);
      const isCard = a.paymentType === "card";
      const isTipCard = a.tipPaymentType === "card";
      byTherapist[t].rows.push({
        date: a.date,
        client: a.clientName,
        isTicket: a.isTicket,
        isGiftCard: a.isGiftCard,
        ticketInfo: a.isTicket ? `${a.ticketMenu} ${a.ticketCurrent}/${a.ticketTotal}` : "",
        duration: a.duration,
        service: svc,
        tip,
        paymentType: a.isGiftCard ? "gc" : a.paymentType,
        tipPaymentType: a.isGiftCard ? "gc" : a.tipPaymentType,
        notes: depositAdd > 0 ? `${a.notes || ""}　💰deposit込`.trim() : (a.notes || ""),
      });
      byTherapist[t].totalService += svc;
      byTherapist[t].totalTip += tip;
      if (isCard) byTherapist[t].totalServiceCard += svc;
      if (isTipCard) byTherapist[t].totalTipCard += tip;
    });

    // Retail from appointments → body therapist
    allAppts.filter(a => !a.isCavSlot && a.purchaseTags?.includes("retail") && Number(a.retailPurchaseAmount || 0) > 0).forEach(a => {
      const t = a.therapist;
      if (!t || !byTherapist[t]) return;
      const retail = Number(a.retailPurchaseAmount);
      const isCard = a.retailPurchasePaymentType === "card";
      byTherapist[t].rows.push({
        date: a.date,
        client: a.clientName,
        isTicket: false,
        ticketInfo: "",
        duration: 0,
        service: 0,
        tip: 0,
        paymentType: a.retailPurchasePaymentType,
        tipPaymentType: "",
        retail,
        retailProduct: a.retailProductName || "",
        notes: `🛍️ 物販${a.retailProductName ? ` (${a.retailProductName})` : ""}`,
      });
      byTherapist[t].totalRetail += retail;
      if (isCard) byTherapist[t].totalRetailCard += retail;
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
          notes: `➕ ${addon.serviceName || addon.name || "オプション"}`,
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
      if (depositAdd > 0 && parent) {
        const bodyReceived = Number(parent.price || 0);
        const totalReceived = bodyReceived + cavReceived;
        // Cav therapist gets their proportion of full payroll (received + deposit)
        svc = totalReceived > 0
          ? Math.round((totalReceived + depositAdd) * cavReceived / totalReceived * 10) / 10
          : cavReceived;
      } else {
        svc = cavReceived;
      }
      const tip = Number(a.tip || 0);
      byTherapist[t].rows.push({
        date: a.date,
        client: a.clientName,
        isTicket: a.isTicket,
        ticketInfo: a.isTicket ? `${a.ticketMenu} ${a.ticketCurrent}/${a.ticketTotal}` : "",
        duration: 15,
        service: svc,
        tip,
        paymentType: a.paymentType,
        tipPaymentType: a.tipPaymentType,
        notes: depositAdd > 0 ? `⚡ 機械　💰deposit込` : `⚡ 機械`,
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
    const rows = [
      ["日付", "お客様", "分数", "施術料金", "支払方法", "チップ", "チップ支払", "物販", "チケット", "メモ"],
      ...data.rows.map(r => [
        r.date, r.client, r.duration || "", r.service || "",
        r.paymentType === "card" ? "カード" : (r.paymentType === "cash" ? "現金" : ""),
        r.tip || "",
        r.tipPaymentType === "card" ? "カード" : (r.tipPaymentType === "cash" ? "現金" : ""),
        r.retail || "",
        r.isTicket ? r.ticketInfo : (r.retail ? "物販" : "通常"),
        r.notes
      ]),
      [],
      ["", "", "合計", data.totalService, "", data.totalTip, "", data.totalRetail, "", ""],
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
    const allRows = [["担当者", "日付", "お客様", "分数", "施術料金", "支払方法", "チップ", "チップ支払", "物販", "チケット", "メモ"]];
    THERAPISTS.forEach(t => {
      const data = payrollData.byTherapist[t];
      if (!data || data.rows.length === 0) return;
      data.rows.forEach(r => {
        allRows.push([
          t, r.date, r.client, r.duration || "", r.service || "",
          r.paymentType === "card" ? "カード" : (r.paymentType === "cash" ? "現金" : ""),
          r.tip || "",
          r.tipPaymentType === "card" ? "カード" : (r.tipPaymentType === "cash" ? "現金" : ""),
          r.retail || "",
          r.isGiftCard ? "GC消化" : r.isTicket ? r.ticketInfo : (r.retail ? "物販" : "通常"),
          r.notes
        ]);
      });
      allRows.push([t, "", "小計", "", data.totalService, "", data.totalTip, "", data.totalRetail, "", ""]);
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
                    {["日付", "お客様", "分数", "施術", "支払", "チップ", "tip支払", "物販", "種別"].map(h => (
                      <th key={h} style={{ padding: "6px 8px", textAlign: "left", color: "#888", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.rows.sort((a, b) => a.date.localeCompare(b.date)).map((r, i) => (
                    <tr key={i} style={{ borderBottom: "1px solid #F5F5F5", background: r.isGiftCard ? "#FFFDE7" : r.isTicket ? "#F0F7FF" : "white" }}>
                      <td style={{ padding: "6px 8px", whiteSpace: "nowrap", color: "#888" }}>{r.date.slice(5)}</td>
                      <td style={{ padding: "6px 8px", fontWeight: 600 }}>{r.client}</td>
                      <td style={{ padding: "6px 8px", textAlign: "center" }}>{r.duration}分</td>
                      <td style={{ padding: "6px 8px", color: "#C62828", fontWeight: 700 }}>{formatCurrency(r.service)}</td>
                      <td style={{ padding: "6px 8px" }}>
                        {r.isGiftCard
                          ? <span style={{ fontSize: 10, background: "#FFFDE7", color: "#B45309", padding: "2px 6px", borderRadius: 8 }}>🎁GC</span>
                          : <span style={{ fontSize: 10, background: r.paymentType === "cash" ? "#E8F5E9" : "#E3F2FD", color: r.paymentType === "cash" ? "#2E7D32" : "#1565C0", padding: "2px 6px", borderRadius: 8 }}>
                              {r.paymentType === "cash" ? "現金" : "カード"}
                            </span>}
                      </td>
                      <td style={{ padding: "6px 8px", color: "#E65100", fontWeight: 700 }}>{r.tip > 0 ? formatCurrency(r.tip) : "—"}</td>
                      <td style={{ padding: "6px 8px" }}>
                        {r.tip > 0 && (r.isGiftCard
                          ? <span style={{ fontSize: 10, background: "#FFFDE7", color: "#B45309", padding: "2px 6px", borderRadius: 8 }}>🎁GC</span>
                          : <span style={{ fontSize: 10, background: r.tipPaymentType === "cash" ? "#E8F5E9" : "#E3F2FD", color: r.tipPaymentType === "cash" ? "#2E7D32" : "#1565C0", padding: "2px 6px", borderRadius: 8 }}>
                              {r.tipPaymentType === "cash" ? "現金" : "カード"}
                            </span>)}
                      </td>
                      <td style={{ padding: "6px 8px", color: "#6A1B9A", fontWeight: r.retail ? 700 : 400 }}>{r.retail ? formatCurrency(r.retail) : "—"}</td>
                      <td style={{ padding: "6px 8px", fontSize: 11, color: r.isGiftCard ? "#B45309" : r.isTicket ? "#1565C0" : (r.retail ? "#6A1B9A" : "#888") }}>
                        {r.isGiftCard ? "🎁 GC消化" : r.isTicket ? `🎟️ ${r.ticketInfo}` : r.notes || "通常"}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: "2px solid #EEE", background: "#FFF8E1" }}>
                    <td colSpan={3} style={{ padding: "8px", fontWeight: 700, color: "#555" }}>合計</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#C62828" }}>{formatCurrency(data.totalService)}</td>
                    <td />
                    <td style={{ padding: "8px", fontWeight: 800, color: "#E65100" }}>{formatCurrency(data.totalTip)}</td>
                    <td />
                    <td style={{ padding: "8px", fontWeight: 800, color: "#6A1B9A" }}>{data.totalRetail > 0 ? formatCurrency(data.totalRetail) : "—"}</td>
                    <td style={{ padding: "8px", fontWeight: 800, color: "#0D4F4F" }}>合計 {formatCurrency(data.totalService + data.totalTip + data.totalRetail)}</td>
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
