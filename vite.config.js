import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

// Dev-only middleware that proxies Square's Gift Card Activities API.
// Keeps SQUARE_ACCESS_TOKEN on the server side (never sent to the browser).
function squareGiftCardApi(env) {
  return {
    name: 'square-gift-card-api',
    configureServer(server) {
      server.middlewares.use('/api/giftcard-activities', async (req, res) => {
        const token = env.SQUARE_ACCESS_TOKEN
        const locationId = env.SQUARE_LOCATION_ID
        if (!token || !locationId) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'SQUARE_ACCESS_TOKEN / SQUARE_LOCATION_ID not set in .env' }))
          return
        }
        try {
          const url = new URL(req.url, 'http://localhost')
          const date = url.searchParams.get('date')
          if (!date) {
            res.statusCode = 400
            res.end(JSON.stringify({ error: 'date query param required (YYYY-MM-DD)' }))
            return
          }
          // Hawaii is UTC-10, no DST
          const beginTime = `${date}T00:00:00-10:00`
          const next = new Date(`${date}T00:00:00Z`)
          next.setUTCDate(next.getUTCDate() + 1)
          const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`

          const params = new URLSearchParams({
            location_id: locationId,
            type: 'ACTIVATE',
            begin_time: beginTime,
            end_time: endTime,
          })
          const squareRes = await fetch(`https://connect.squareup.com/v2/gift-cards/activities?${params}`, {
            headers: {
              Authorization: `Bearer ${token}`,
              'Square-Version': '2024-01-18',
            },
          })
          const data = await squareRes.json()
          if (!squareRes.ok) {
            res.statusCode = squareRes.status
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'Square API error', details: data }))
            return
          }

          const activities = (data.gift_card_activities || []).map(a => {
            const details = a.activate_activity_details || a.gift_card_activate_activity_details || {}
            return {
              id: a.id,
              giftCardId: a.gift_card_id,
              amount: (details.amount_money?.amount || 0) / 100,
              createdAt: a.created_at,
              orderId: details.order_id || null,
            }
          })

          // Look up each order's actual tender (CASH vs CARD) — activations happen both
          // in-person (POS, often cash) and online (checkout, always card), so the payment
          // method can't be assumed. Cache by order_id since a few activations can share one.
          const orderHeaders = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' }
          const tenderCache = {}
          for (const act of activities) {
            if (!act.orderId) { act.paymentType = 'card'; continue }
            if (tenderCache[act.orderId] === undefined) {
              try {
                const orderRes = await fetch(`https://connect.squareup.com/v2/orders/${act.orderId}`, { headers: orderHeaders })
                const orderData = await orderRes.json()
                const tenderType = orderRes.ok ? orderData.order?.tenders?.[0]?.type : null
                tenderCache[act.orderId] = tenderType === 'CASH' ? 'cash' : 'card'
              } catch {
                tenderCache[act.orderId] = 'card'
              }
            }
            act.paymentType = tenderCache[act.orderId]
          }

          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ activities }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })
    },
  }
}

// Dev-only middleware that proxies Square's Bookings/Customers/Catalog APIs so the Daily
// Sheet can sync a day's appointments (name, therapist, time, course) without exposing
// SQUARE_ACCESS_TOKEN to the browser.
function squareBookingsApi(env) {
  const TEAM_MEMBER_MAP = {
    'TME2bzx3XoNFV-jw': 'Mami',
    'TMWQwOib7pJ7Sw22': 'Yuka',
    'TMez6BJMpuw8PpVZ': 'Megumi',
    'TMZUOeE7bTyj6JrW': 'Betsy',
    'TMhfddQBlrnZkR9n': 'Aya',
    'TMu1v7kb0cgbl5mn': 'Mai',
    'TMwBMik2I8Xpyog2': 'Maki',
    'TMW8eccMtM4pdMiM': 'Hitomi',
  }
  return {
    name: 'square-bookings-api',
    configureServer(server) {
      server.middlewares.use('/api/square-bookings', async (req, res) => {
        const token = env.SQUARE_ACCESS_TOKEN
        const locationId = env.SQUARE_LOCATION_ID
        if (!token || !locationId) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'SQUARE_ACCESS_TOKEN / SQUARE_LOCATION_ID not set in .env' }))
          return
        }
        try {
          const url = new URL(req.url, 'http://localhost')
          const date = url.searchParams.get('date')
          if (!date) {
            res.statusCode = 400
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'date query param required (YYYY-MM-DD)' }))
            return
          }
          // Hawaii is UTC-10, no DST
          const startAtMin = `${date}T00:00:00-10:00`
          const next = new Date(`${date}T00:00:00Z`)
          next.setUTCDate(next.getUTCDate() + 1)
          const startAtMax = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`
          const headers = {
            Authorization: `Bearer ${token}`,
            'Square-Version': '2024-01-18',
          }

          const bookingsParams = new URLSearchParams({
            location_id: locationId,
            start_at_min: startAtMin,
            start_at_max: startAtMax,
            limit: '200',
          })
          const bookingsRes = await fetch(`https://connect.squareup.com/v2/bookings?${bookingsParams}`, { headers })
          const bookingsData = await bookingsRes.json()
          if (!bookingsRes.ok) {
            res.statusCode = bookingsRes.status
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'Square API error (bookings)', details: bookingsData }))
            return
          }
          const bookings = (bookingsData.bookings || []).filter(
            b => b.status !== 'CANCELLED_BY_SELLER' && b.status !== 'CANCELLED_BY_CUSTOMER'
          )

          const customerCache = {}
          const getCustomerName = async (id) => {
            if (!id) return ''
            if (customerCache[id] !== undefined) return customerCache[id]
            try {
              const r = await fetch(`https://connect.squareup.com/v2/customers/${id}`, { headers })
              const d = await r.json()
              const name = r.ok ? [d.customer?.given_name, d.customer?.family_name].filter(Boolean).join(' ') : ''
              customerCache[id] = name
              return name
            } catch {
              customerCache[id] = ''
              return ''
            }
          }

          const serviceCache = {}
          const getServiceName = async (variationId) => {
            if (!variationId) return ''
            if (serviceCache[variationId] !== undefined) return serviceCache[variationId]
            try {
              const r = await fetch(
                `https://connect.squareup.com/v2/catalog/object/${variationId}?include_related_objects=true`,
                { headers }
              )
              const d = await r.json()
              if (!r.ok) { serviceCache[variationId] = ''; return '' }
              const variationName = d.object?.item_variation_data?.name || ''
              const itemId = d.object?.item_variation_data?.item_id
              const itemObj = (d.related_objects || []).find(o => o.id === itemId)
              const itemName = itemObj?.item_data?.name || ''
              const full = itemName && variationName ? `${itemName} ${variationName}` : (itemName || variationName)
              serviceCache[variationId] = full
              return full
            } catch {
              serviceCache[variationId] = ''
              return ''
            }
          }

          const results = []
          for (const b of bookings) {
            const seg = (b.appointment_segments || [])[0] || {}
            const clientName = await getCustomerName(b.customer_id)
            const serviceName = await getServiceName(seg.service_variation_id)
            const therapist = TEAM_MEMBER_MAP[seg.team_member_id] || ''
            const startAt = new Date(b.start_at)
            const hi = new Date(startAt.getTime() - 10 * 3600 * 1000)
            const hh = String(hi.getUTCHours()).padStart(2, '0')
            const mm = String(hi.getUTCMinutes()).padStart(2, '0')
            results.push({
              squareId: b.id,
              clientName,
              therapist,
              startTime: `${hh}:${mm}`,
              duration: seg.duration_minutes || 60,
              serviceName,
              notes: '',
            })
          }

          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ bookings: results }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })
    },
  }
}

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return {
    plugins: [react(), squareGiftCardApi(env), squareBookingsApi(env)],
  }
})
