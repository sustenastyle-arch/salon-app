// One day's gift card sales from Square, shared by the Vite dev middleware (vite.config.js) and
// the Vercel function (api/giftcard-activities.js) so the two can't drift apart.
//
// Each sale is a gift card activation (new card) or a load (value added to an existing card).
// `amount` is what the customer actually PAID for it, not the value put on the card: during a
// gift card discount (e.g. the 11/27–11/30/2026 10% off) a $200 card is sold for $180, and the
// sheet has to match Square's payments, so it records $180. `loadedAmount` keeps the card value.
// A sale whose order was refunded is left out (11/27/2026: a $90 eGift was refunded and its
// balance cleared a minute later).
const SQUARE = 'https://connect.squareup.com/v2';

export async function fetchGiftCardSales({ token, locationId, date }) {
  const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-10-17' };
  // Hawaii is UTC-10, no DST
  const beginTime = `${date}T00:00:00-10:00`;
  const next = new Date(`${date}T00:00:00Z`);
  next.setUTCDate(next.getUTCDate() + 1);
  const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`;

  const raw = [];
  for (const type of ['ACTIVATE', 'LOAD']) {
    let cursor = '';
    do {
      const params = new URLSearchParams({ location_id: locationId, type, begin_time: beginTime, end_time: endTime });
      if (cursor) params.set('cursor', cursor);
      const r = await fetch(`${SQUARE}/gift-cards/activities?${params}`, { headers });
      const data = await r.json();
      if (!r.ok) { const e = new Error('Square API error'); e.status = r.status; e.details = data; throw e; }
      raw.push(...(data.gift_card_activities || []));
      cursor = data.cursor || '';
    } while (cursor);
  }

  const orders = {};
  const getOrder = async (id) => {
    if (!(id in orders)) {
      try {
        const r = await fetch(`${SQUARE}/orders/${id}`, { headers });
        orders[id] = r.ok ? (await r.json()).order || null : null;
      } catch { orders[id] = null; }
    }
    return orders[id];
  };

  // A refund can show up only on the payment (the order itself still reads COMPLETED with no
  // refunds), so check each tender's payment.
  const fullyRefunded = async (order) => {
    if ((order.refunds || []).some(rf => rf.status !== 'REJECTED' && rf.status !== 'FAILED')) return true;
    const tenders = (order.tenders || []).filter(t => t.payment_id);
    if (!tenders.length) return false;
    for (const t of tenders) {
      try {
        const r = await fetch(`${SQUARE}/payments/${t.payment_id}`, { headers });
        const p = r.ok ? (await r.json()).payment : null;
        if (!p || (p.refunded_money?.amount || 0) < (p.total_money?.amount || 0)) return false;
      } catch { return false; }
    }
    return true;
  };

  const sales = [];
  for (const a of raw) {
    const details = a.activate_activity_details || a.load_activity_details || a.gift_card_activate_activity_details || {};
    const loaded = (details.amount_money?.amount || 0) / 100;
    const orderId = details.order_id || null;
    const order = orderId ? await getOrder(orderId) : null;
    if (order && await fullyRefunded(order)) continue;
    // Paid share of the gift card lines on the order (1.0 with no discount, 0.9 with 10% off).
    const gcLines = (order?.line_items || []).filter(li => li.item_type === 'GIFT_CARD');
    const gross = gcLines.reduce((s, li) => s + (li.gross_sales_money?.amount || 0), 0);
    const paid = gcLines.reduce((s, li) => s + (li.total_money?.amount || 0), 0);
    const ratio = gross > 0 ? paid / gross : 1;
    sales.push({
      id: a.id,
      type: a.type,
      giftCardId: a.gift_card_id,
      last4: (a.gift_card_gan || '').slice(-4),
      amount: Math.round(loaded * ratio * 100) / 100,
      loadedAmount: loaded,
      createdAt: a.created_at,
      orderId,
      paymentType: order?.tenders?.[0]?.type === 'CASH' ? 'cash' : 'card',
    });
  }
  return sales.sort((x, y) => x.createdAt.localeCompare(y.createdAt));
}
