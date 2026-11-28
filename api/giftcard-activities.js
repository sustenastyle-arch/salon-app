// Vercel serverless function — production equivalent of the Vite dev-only middleware
// in vite.config.js. Keeps SQUARE_ACCESS_TOKEN server-side (never sent to the browser).
// The Square logic lives in api/_lib/giftCardSales.js, shared with the dev middleware.
import { fetchGiftCardSales } from './_lib/giftCardSales.js';

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
    res.status(200).json({ activities: await fetchGiftCardSales({ token, locationId, date }) });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.status ? 'Square API error' : String(e), details: e.details });
  }
}
