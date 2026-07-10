// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Keeps SQUARE_ACCESS_TOKEN server-side (never sent to the browser).
const TEAM_MEMBER_MAP = {
  'TME2bzx3XoNFV-jw': 'Mami',
  'TMWQwOib7pJ7Sw22': 'Yuka',
  'TMez6BJMpuw8PpVZ': 'Megumi',
  'TMZUOeE7bTyj6JrW': 'Betsy',
  'TMhfddQBlrnZkR9n': 'Aya',
  'TMu1v7kb0cgbl5mn': 'Mai',
  'TMwBMik2I8Xpyog2': 'Maki',
  'TMW8eccMtM4pdMiM': 'Hitomi',
};

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
    const startAtMin = `${date}T00:00:00-10:00`;
    const next = new Date(`${date}T00:00:00Z`);
    next.setUTCDate(next.getUTCDate() + 1);
    const startAtMax = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`;
    const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' };

    const bookingsParams = new URLSearchParams({
      location_id: locationId,
      start_at_min: startAtMin,
      start_at_max: startAtMax,
      limit: '200',
    });
    const bookingsRes = await fetch(`https://connect.squareup.com/v2/bookings?${bookingsParams}`, { headers });
    const bookingsData = await bookingsRes.json();
    if (!bookingsRes.ok) {
      res.status(bookingsRes.status).json({ error: 'Square API error (bookings)', details: bookingsData });
      return;
    }
    const bookings = (bookingsData.bookings || []).filter(
      b => b.status !== 'CANCELLED_BY_SELLER' && b.status !== 'CANCELLED_BY_CUSTOMER'
    );

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

    const serviceCache = {};
    const getServiceName = async (variationId) => {
      if (!variationId) return '';
      if (serviceCache[variationId] !== undefined) return serviceCache[variationId];
      try {
        const r = await fetch(
          `https://connect.squareup.com/v2/catalog/object/${variationId}?include_related_objects=true`,
          { headers }
        );
        const d = await r.json();
        if (!r.ok) { serviceCache[variationId] = ''; return ''; }
        const variationName = d.object?.item_variation_data?.name || '';
        const itemId = d.object?.item_variation_data?.item_id;
        const itemObj = (d.related_objects || []).find(o => o.id === itemId);
        const itemName = itemObj?.item_data?.name || '';
        const full = itemName && variationName ? `${itemName} ${variationName}` : (itemName || variationName);
        serviceCache[variationId] = full;
        return full;
      } catch {
        serviceCache[variationId] = '';
        return '';
      }
    };

    const results = [];
    for (const b of bookings) {
      const seg = (b.appointment_segments || [])[0] || {};
      const clientName = await getCustomerName(b.customer_id);
      const serviceName = await getServiceName(seg.service_variation_id);
      const therapist = TEAM_MEMBER_MAP[seg.team_member_id] || '';
      const startAt = new Date(b.start_at);
      const hi = new Date(startAt.getTime() - 10 * 3600 * 1000);
      const hh = String(hi.getUTCHours()).padStart(2, '0');
      const mm = String(hi.getUTCMinutes()).padStart(2, '0');
      results.push({
        squareId: b.id,
        clientName,
        therapist,
        startTime: `${hh}:${mm}`,
        duration: seg.duration_minutes || 60,
        serviceName,
        notes: '',
      });
    }

    res.status(200).json({ bookings: results });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
}
