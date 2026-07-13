// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Keeps SQUARE_ACCESS_TOKEN server-side (never sent to the browser).
//
// Deposits are collected via a Square Payment Link with a custom line item named "Deposit"
// (historically mistyped as "Desit") — there's no dedicated Square API for this the way gift
// cards have one, so it has to be found by scanning the day's completed payments and checking
// each one's order for a matching line item. Payment-Link orders stay in state "OPEN" even
// after being paid, so this can't filter by order state the way a normal search might.
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

  const DEPOSIT_NAMES = ['deposit', 'desit'];

  try {
    // Hawaii is UTC-10, no DST
    const beginTime = `${date}T00:00:00-10:00`;
    const next = new Date(`${date}T00:00:00Z`);
    next.setUTCDate(next.getUTCDate() + 1);
    const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`;
    const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' };

    const customerCache = {};
    const getCustomerName = async (id) => {
      if (!id) return '';
      if (customerCache[id] !== undefined) return customerCache[id];
      try {
        const r = await fetch(`https://connect.squareup.com/v2/customers/${id}`, { headers });
        const d = await r.json();
        const name = r.ok ? [d.customer?.given_name, d.customer?.family_name].filter(Boolean).join(' ') : '';
        customerCache[id] = name;
        return name;
      } catch {
        customerCache[id] = '';
        return '';
      }
    };

    const deposits = [];
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

      const payRes = await fetch(`https://connect.squareup.com/v2/payments?${params}`, { headers });
      const payData = await payRes.json();
      if (!payRes.ok) {
        res.status(payRes.status).json({ error: 'Square API error (payments)', details: payData });
        return;
      }

      for (const p of payData.payments || []) {
        if (p.status !== 'COMPLETED' || !p.order_id) continue;
        const orderRes = await fetch(`https://connect.squareup.com/v2/orders/${p.order_id}`, { headers });
        const orderData = await orderRes.json();
        if (!orderRes.ok) continue;
        const lineItems = orderData.order?.line_items || [];
        const depositItem = lineItems.find(li => DEPOSIT_NAMES.includes((li.name || '').trim().toLowerCase()));
        if (!depositItem) continue;

        const amount = (Number(depositItem.gross_sales_money?.amount || 0)) / 100;
        const name = p.customer_id
          ? await getCustomerName(p.customer_id)
          : [p.billing_address?.first_name, p.billing_address?.last_name].filter(Boolean).join(' ');

        deposits.push({
          id: p.id,
          amount,
          clientName: name || 'デポジット（お客様名不明）',
          paymentType: p.source_type === 'CASH' ? 'cash' : 'card',
          createdAt: p.created_at,
        });
      }
      cursor = payData.cursor;
    } while (cursor);

    res.status(200).json({ deposits });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
