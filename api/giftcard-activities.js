// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Keeps SQUARE_ACCESS_TOKEN server-side (never sent to the browser).
export default async function handler(req, res) {
  const token = process.env.SQUARE_ACCESS_TOKEN;
  const locationId = process.env.SQUARE_LOCATION_ID;
  if (!token || !locationId) {
    res.status(500).json({ error: 'SQUARE_ACCESS_TOKEN / SQUARE_LOCATION_ID not set' });
    return;
  }

  const date = req.query.date;
  if (!date) {
    res.status(400).json({ error: 'date query param required (YYYY-MM-DD)' });
    return;
  }

  try {
    // Hawaii is UTC-10, no DST
    const beginTime = `${date}T00:00:00-10:00`;
    const next = new Date(`${date}T00:00:00Z`);
    next.setUTCDate(next.getUTCDate() + 1);
    const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`;

    const params = new URLSearchParams({
      location_id: locationId,
      type: 'ACTIVATE',
      begin_time: beginTime,
      end_time: endTime,
    });
    const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' };

    const squareRes = await fetch(`https://connect.squareup.com/v2/gift-cards/activities?${params}`, { headers });
    const data = await squareRes.json();
    if (!squareRes.ok) {
      res.status(squareRes.status).json({ error: 'Square API error', details: data });
      return;
    }

    const activities = (data.gift_card_activities || []).map(a => {
      const details = a.activate_activity_details || a.gift_card_activate_activity_details || {};
      return {
        id: a.id,
        giftCardId: a.gift_card_id,
        amount: (details.amount_money?.amount || 0) / 100,
        createdAt: a.created_at,
        orderId: details.order_id || null,
      };
    });

    // Look up each order's actual tender (CASH vs CARD) — activations happen both in-person
    // (POS, often cash) and online (checkout, always card), so it can't be assumed.
    const tenderCache = {};
    for (const act of activities) {
      if (!act.orderId) { act.paymentType = 'card'; continue; }
      if (tenderCache[act.orderId] === undefined) {
        try {
          const orderRes = await fetch(`https://connect.squareup.com/v2/orders/${act.orderId}`, { headers });
          const orderData = await orderRes.json();
          const tenderType = orderRes.ok ? orderData.order?.tenders?.[0]?.type : null;
          tenderCache[act.orderId] = tenderType === 'CASH' ? 'cash' : 'card';
        } catch {
          tenderCache[act.orderId] = 'card';
        }
      }
      act.paymentType = tenderCache[act.orderId];
    }

    res.status(200).json({ activities });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
