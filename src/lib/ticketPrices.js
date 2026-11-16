// Ticket package list prices (service and suggested tip), shared by the app (auto-fills a new
// ticket sale) and scripts/ticket-incentive.mjs (the incentive reference table and the
// "was the new-client $20 off applied?" check), so a price change only happens here.
export const TICKET_PACKAGE_PRICES = {
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

// Every new client (traveler or local) gets $20 off, including on a ticket package.
export const NEW_CLIENT_DISCOUNT = 20;
