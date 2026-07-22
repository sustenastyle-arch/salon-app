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
    // it has to be recovered from the order's line items instead. Cash payments never go
    // through the tip-prompt flow at all (so tip_money is always 0 for them), but staff
    // sometimes still itemize a "Tip" line inside a cash order — checked for both tenders
    // for that reason. Always ADD the order's own "Tip" line item on top of tip_money rather
    // than only falling back to it when tip_money is 0 — a split-card sale can genuinely have
    // both (e.g. a same-day ticket purchase charged across two cards, with a separate top-up
    // tip line rung up on the second card alongside its own tip-prompt amount); treating them
    // as mutually exclusive silently dropped the line-item portion whenever tip_money was
    // already nonzero. The same order fetch also gives us the non-tip line item names, used
    // below to label refunded payments so staff can match a refund back to the
    // appointment/retail entry that needs correcting.
    //
    // Use gross_sales_money (the line item's price as entered), not total_money — an
    // order-level discount gets auto-prorated across every line item including "Tip" by
    // Square, but a discount off the service price shouldn't silently shrink the tip figure
    // staff actually typed in.
    const orderCache = {};
    const getOrder = async (orderId) => {
      if (orderCache[orderId] !== undefined) return orderCache[orderId];
      try {
        const r = await fetch(`https://connect.squareup.com/v2/orders/${orderId}`, { headers });
        const d = await r.json();
        orderCache[orderId] = r.ok ? (d.order || null) : null;
      } catch {
        orderCache[orderId] = null;
      }
      return orderCache[orderId];
    };
    const getOrderTipLineItems = async (orderId) => {
      const order = await getOrder(orderId);
      const lineItems = order?.line_items || [];
      return lineItems
        .filter(li => (li.name || '').trim().toLowerCase() === 'tip')
        .reduce((s, li) => s + (Number(li.gross_sales_money?.amount || 0)) / 100, 0);
    };
    const getOrderItemLabel = async (orderId) => {
      const order = await getOrder(orderId);
      const lineItems = order?.line_items || [];
      return lineItems
        .filter(li => (li.name || '').trim().toLowerCase() !== 'tip')
        .map(li => li.name)
        .filter(Boolean)
        .join(', ');
    };
    // Hawaii is UTC-10, no DST — same rule as the day-boundary math above.
    const formatHawaiiTime = (isoString) => {
      const d = new Date(isoString);
      const h = (d.getUTCHours() + 24 - 10) % 24;
      const m = d.getUTCMinutes();
      const ampm = h < 12 ? 'AM' : 'PM';
      return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${ampm}`;
    };

    let cashTotal = 0, cardTotal = 0, cashTip = 0, cardTip = 0;
    const refunds = [];
    // Every individual completed payment (net of any refund), so the client can try to match
    // each one against a sheet entry with the same tender + total — narrowing a reconciliation
    // mismatch down to specific transactions instead of just an aggregate total being off,
    // which otherwise has to be tracked down by hand.
    const payments = [];
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
        // Square keeps a refunded payment's status COMPLETED and total_money unchanged —
        // the refund only shows up in refunded_money — so without this a fully-refunded
        // sale still counted as real revenue here even though the customer got it all back.
        const refunded = (p.refunded_money?.amount || 0) / 100;
        const netTotal = Math.max(0, total - refunded);
        let tip = (p.tip_money?.amount || 0) / 100;
        if (p.order_id) tip += await getOrderTipLineItems(p.order_id);
        // A fully refunded payment kept no money at all, tip included.
        const netTip = netTotal > 0 ? tip : 0;
        const tender = p.source_type === 'CASH' ? 'cash' : 'card';
        if (tender === 'cash') {
          cashTotal += netTotal;
          cashTip += netTip;
        } else {
          cardTotal += netTotal;
          cardTip += netTip;
        }
        if (refunded > 0) {
          refunds.push({
            amount: Math.round(refunded * 100) / 100,
            tender,
            time: formatHawaiiTime(p.created_at),
            label: p.order_id ? await getOrderItemLabel(p.order_id) : '',
          });
        }
        if (netTotal > 0) {
          payments.push({
            amount: Math.round(netTotal * 100) / 100,
            tip: Math.round(netTip * 100) / 100,
            tender,
            time: formatHawaiiTime(p.created_at),
            label: p.order_id ? await getOrderItemLabel(p.order_id) : '',
          });
        }
      }
      cursor = data.cursor;
    } while (cursor);

    res.status(200).json({
      cashTotal: Math.round(cashTotal * 100) / 100,
      cardTotal: Math.round(cardTotal * 100) / 100,
      refunds,
      payments,
      cashTip: Math.round(cashTip * 100) / 100,
      cardTip: Math.round(cardTip * 100) / 100,
    });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
