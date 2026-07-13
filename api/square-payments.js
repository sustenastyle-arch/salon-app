// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Keeps SQUARE_ACCESS_TOKEN server-side (never sent to the browser).
//
// Pulls Square's actual completed payments for a day and totals them by tender
// (cash vs card), so the Daily Sheet's own manually-entered cash/card/tip totals
// can be checked against what Square really recorded — catching entry typos same-day
// instead of at month-end.
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
    const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' };

    // Package/ticket sales are sometimes rung up with the tip added as a manual line item
    // named "Tip" inside the order, instead of going through Square's tip-prompt flow — in
    // that case tip_money stays 0 even though real tip money is bundled into the total, so
    // it has to be recovered from the order's line items instead. Only checked for card
    // payments (the only tip breakdown this endpoint's caller actually compares).
    //
    // Use gross_sales_money (the line item's price as entered), not total_money — an
    // order-level discount gets auto-prorated across every line item including "Tip" by
    // Square, but a discount off the service price shouldn't silently shrink the tip figure
    // staff actually typed in.
    const orderTipCache = {};
    const getOrderTipLineItems = async (orderId) => {
      if (orderTipCache[orderId] !== undefined) return orderTipCache[orderId];
      try {
        const r = await fetch(`https://connect.squareup.com/v2/orders/${orderId}`, { headers });
        const d = await r.json();
        if (!r.ok) { orderTipCache[orderId] = 0; return 0; }
        const lineItems = d.order?.line_items || [];
        const tip = lineItems
          .filter(li => (li.name || '').trim().toLowerCase() === 'tip')
          .reduce((s, li) => s + (Number(li.gross_sales_money?.amount || 0)) / 100, 0);
        orderTipCache[orderId] = tip;
        return tip;
      } catch {
        orderTipCache[orderId] = 0;
        return 0;
      }
    };

    let cashTotal = 0, cardTotal = 0, cashTip = 0, cardTip = 0;
    let cursor;
    do {
      const params = new URLSearchParams({
        location_id: locationId,
        begin_time: beginTime,
        end_time: endTime,
        sort_order: 'ASC',
        limit: '100',
      });
      if (cursor) params.set('cursor', cursor);

      const squareRes = await fetch(`https://connect.squareup.com/v2/payments?${params}`, { headers });
      const data = await squareRes.json();
      if (!squareRes.ok) {
        res.status(squareRes.status).json({ error: 'Square API error', details: data });
        return;
      }

      for (const p of data.payments || []) {
        if (p.status !== 'COMPLETED') continue;
        const total = (p.total_money?.amount || 0) / 100;
        let tip = (p.tip_money?.amount || 0) / 100;
        if (p.source_type === 'CASH') {
          cashTotal += total;
          cashTip += tip;
        } else {
          if (tip === 0 && p.order_id) tip = await getOrderTipLineItems(p.order_id);
          cardTotal += total;
          cardTip += tip;
        }
      }
      cursor = data.cursor;
    } while (cursor);

    res.status(200).json({
      cashTotal: Math.round(cashTotal * 100) / 100,
      cardTotal: Math.round(cardTotal * 100) / 100,
      cashTip: Math.round(cashTip * 100) / 100,
      cardTip: Math.round(cardTip * 100) / 100,
    });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
