import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import { getDay, setDay, getDaysInRange, getAllDayBlobs, setDays, checkAuth, saveAutoBackup, listAutoBackups, restoreAutoBackup } from './api/_lib/dayDataStore.js'

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = ''
    req.on('data', (chunk) => { body += chunk })
    req.on('end', () => {
      try { resolve(body ? JSON.parse(body) : {}) } catch (e) { reject(e) }
    })
    req.on('error', reject)
  })
}

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

// Dev-only middleware that finds same-day deposit payments so the Daily Sheet can auto-fill
// the deposit section. Deposits are collected via a Square Payment Link with a custom line
// item named "Deposit" (historically mistyped as "Desit") — there's no dedicated Square API
// for this the way gift cards have one, so it has to be found by scanning the day's completed
// payments and checking each one's order for a matching line item. Payment-Link orders stay in
// state "OPEN" even after being paid, so this can't filter by order state the way a normal
// search might.
function squareDepositsApi(env) {
  const DEPOSIT_NAMES = ['deposit', 'desit']
  return {
    name: 'square-deposits-api',
    configureServer(server) {
      server.middlewares.use('/api/square-deposits', async (req, res) => {
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
          const beginTime = `${date}T00:00:00-10:00`
          const next = new Date(`${date}T00:00:00Z`)
          next.setUTCDate(next.getUTCDate() + 1)
          const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`
          const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' }

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

          const deposits = []
          let cursor
          do {
            const params = new URLSearchParams({
              location_id: locationId,
              begin_time: beginTime,
              end_time: endTime,
              sort_order: 'ASC',
              limit: '100',
            })
            if (cursor) params.set('cursor', cursor)

            const payRes = await fetch(`https://connect.squareup.com/v2/payments?${params}`, { headers })
            const payData = await payRes.json()
            if (!payRes.ok) {
              res.statusCode = payRes.status
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'Square API error (payments)', details: payData }))
              return
            }

            for (const p of payData.payments || []) {
              if (p.status !== 'COMPLETED' || !p.order_id) continue
              const orderRes = await fetch(`https://connect.squareup.com/v2/orders/${p.order_id}`, { headers })
              const orderData = await orderRes.json()
              if (!orderRes.ok) continue
              const lineItems = orderData.order?.line_items || []
              const depositItem = lineItems.find(li => DEPOSIT_NAMES.includes((li.name || '').trim().toLowerCase()))
              if (!depositItem) continue

              const amount = (Number(depositItem.gross_sales_money?.amount || 0)) / 100
              const name = p.customer_id
                ? await getCustomerName(p.customer_id)
                : [p.billing_address?.first_name, p.billing_address?.last_name].filter(Boolean).join(' ')

              deposits.push({
                id: p.id,
                amount,
                clientName: name || 'デポジット（お客様名不明）',
                paymentType: p.source_type === 'CASH' ? 'cash' : 'card',
                createdAt: p.created_at,
              })
            }
            cursor = payData.cursor
          } while (cursor)

          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ deposits }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })
    },
  }
}

// Dev-only middleware that proxies Square's Payments API so the Daily Sheet can compare
// its own manually-entered cash/card/tip totals against what Square actually recorded for
// the day — catching entry typos same-day instead of at month-end.
function squarePaymentsApi(env) {
  return {
    name: 'square-payments-api',
    configureServer(server) {
      server.middlewares.use('/api/square-payments', async (req, res) => {
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
          const beginTime = `${date}T00:00:00-10:00`
          const next = new Date(`${date}T00:00:00Z`)
          next.setUTCDate(next.getUTCDate() + 1)
          const endTime = `${next.toISOString().slice(0, 10)}T00:00:00-10:00`
          const headers = { Authorization: `Bearer ${token}`, 'Square-Version': '2024-01-18' }

          // Package/ticket sales are sometimes rung up with the tip added as a manual line
          // item named "Tip" inside the order, instead of going through Square's tip-prompt
          // flow — in that case tip_money stays 0 even though real tip money is bundled into
          // the total, so it has to be recovered from the order's line items instead. Cash
          // payments never go through the tip-prompt flow at all (so tip_money is always 0
          // for them), but staff sometimes still itemize a "Tip" line inside a cash order —
          // checked for both tenders for that reason. The same order fetch also gives us the
          // non-tip line item names, used below to label refunded payments so staff can match
          // a refund back to the appointment/retail entry that needs correcting.
          //
          // Use gross_sales_money (the line item's price as entered), not total_money — an
          // order-level discount gets auto-prorated across every line item including "Tip" by
          // Square, but a discount off the service price shouldn't silently shrink the tip
          // figure staff actually typed in.
          const orderCache = {}
          const getOrder = async (orderId) => {
            if (orderCache[orderId] !== undefined) return orderCache[orderId]
            try {
              const r = await fetch(`https://connect.squareup.com/v2/orders/${orderId}`, { headers })
              const d = await r.json()
              orderCache[orderId] = r.ok ? (d.order || null) : null
            } catch {
              orderCache[orderId] = null
            }
            return orderCache[orderId]
          }
          const getOrderTipLineItems = async (orderId) => {
            const order = await getOrder(orderId)
            const lineItems = order?.line_items || []
            return lineItems
              .filter(li => (li.name || '').trim().toLowerCase() === 'tip')
              .reduce((s, li) => s + (Number(li.gross_sales_money?.amount || 0)) / 100, 0)
          }
          const getOrderItemLabel = async (orderId) => {
            const order = await getOrder(orderId)
            const lineItems = order?.line_items || []
            return lineItems
              .filter(li => (li.name || '').trim().toLowerCase() !== 'tip')
              .map(li => li.name)
              .filter(Boolean)
              .join(', ')
          }
          // Hawaii is UTC-10, no DST — same rule as the day-boundary math above.
          const formatHawaiiTime = (isoString) => {
            const d = new Date(isoString)
            const h = (d.getUTCHours() + 24 - 10) % 24
            const m = d.getUTCMinutes()
            const ampm = h < 12 ? 'AM' : 'PM'
            return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${ampm}`
          }

          let cashTotal = 0, cardTotal = 0, cashTip = 0, cardTip = 0
          const refunds = []
          // Every individual completed payment (net of any refund), so the client can try to
          // match each one against a sheet entry with the same tender + total — narrowing a
          // reconciliation mismatch down to specific transactions instead of just an aggregate
          // total being off, which otherwise has to be tracked down by hand.
          const payments = []
          let cursor
          do {
            const params = new URLSearchParams({
              location_id: locationId,
              begin_time: beginTime,
              end_time: endTime,
              sort_order: 'ASC',
              limit: '100',
            })
            if (cursor) params.set('cursor', cursor)

            const squareRes = await fetch(`https://connect.squareup.com/v2/payments?${params}`, { headers })
            const data = await squareRes.json()
            if (!squareRes.ok) {
              res.statusCode = squareRes.status
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'Square API error', details: data }))
              return
            }

            for (const p of data.payments || []) {
              if (p.status !== 'COMPLETED') continue
              const total = (p.total_money?.amount || 0) / 100
              // Square keeps a refunded payment's status COMPLETED and total_money unchanged —
              // the refund only shows up in refunded_money — so without this a fully-refunded
              // sale still counted as real revenue here even though the customer got it all back.
              const refunded = (p.refunded_money?.amount || 0) / 100
              const netTotal = Math.max(0, total - refunded)
              let tip = (p.tip_money?.amount || 0) / 100
              if (tip === 0 && p.order_id) tip = await getOrderTipLineItems(p.order_id)
              // A fully refunded payment kept no money at all, tip included.
              const netTip = netTotal > 0 ? tip : 0
              const tender = p.source_type === 'CASH' ? 'cash' : 'card'
              if (tender === 'cash') {
                cashTotal += netTotal
                cashTip += netTip
              } else {
                cardTotal += netTotal
                cardTip += netTip
              }
              if (refunded > 0) {
                refunds.push({
                  amount: Math.round(refunded * 100) / 100,
                  tender,
                  time: formatHawaiiTime(p.created_at),
                  label: p.order_id ? await getOrderItemLabel(p.order_id) : '',
                })
              }
              if (netTotal > 0) {
                payments.push({
                  amount: Math.round(netTotal * 100) / 100,
                  tip: Math.round(netTip * 100) / 100,
                  tender,
                  time: formatHawaiiTime(p.created_at),
                  label: p.order_id ? await getOrderItemLabel(p.order_id) : '',
                })
              }
            }
            cursor = data.cursor
          } while (cursor)

          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({
            cashTotal: Math.round(cashTotal * 100) / 100,
            cardTotal: Math.round(cardTotal * 100) / 100,
            refunds,
            payments,
            cashTip: Math.round(cashTip * 100) / 100,
            cardTip: Math.round(cardTip * 100) / 100,
          }))
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
    'TMW8eccMtM4pdMiM': 'Hiromi',
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
          // DECLINED = a booking request the seller rejected — never actually happened, so it
          // must be excluded the same as a cancellation, not synced in as a real appointment.
          const bookings = (bookingsData.bookings || []).filter(
            b => b.status !== 'CANCELLED_BY_SELLER' && b.status !== 'CANCELLED_BY_CUSTOMER' && b.status !== 'DECLINED'
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
              const rawVariationName = d.object?.item_variation_data?.name || ''
              // "Regular" is Square's default variation name for single-variation items — it
              // carries no real info. Some of this business's variations are also just the
              // course name written in Japanese instead (e.g. "ロミロミ矯正90分") rather than
              // "通常" — staff want the schedule cards English-only, so any Japanese variation
              // name is dropped the same way, not just the literal default.
              const isDroppableVariation = rawVariationName.trim().toLowerCase() === 'regular'
                || /[぀-ヿ一-鿿]/.test(rawVariationName)
              const variationName = isDroppableVariation ? '' : rawVariationName
              const itemId = d.object?.item_variation_data?.item_id
              const itemObj = (d.related_objects || []).find(o => o.id === itemId)
              // Item names in this business's Square catalog carry extra junk that was never
              // meant to be in the Name field: a Japanese translation in parentheses, or — for
              // at least one item — a call-to-action with phone numbers typed straight into the
              // name (e.g. "HIFU(Full Face) -ハイフ全顔 Please Call us 808-922-5115 or Text us
              // 808-971-1267"). Staff want the schedule cards English-only, so all of it is
              // stripped at the source here rather than in every place serviceName gets displayed.
              const rawItemName = itemObj?.item_data?.name || ''
              let itemName = rawItemName.replace(/\s*\([^)]*[぀-ヿ一-鿿][^)]*\)/g, '')
              const ctaIdx = itemName.search(/(please\s+)?(call|text)\s+us|\d{3}[-.\s]\d{3}[-.\s]\d{4}/i)
              if (ctaIdx !== -1) itemName = itemName.slice(0, ctaIdx)
              itemName = itemName.split(/\s+/).filter(t => t && !/[぀-ヿ一-鿿]/.test(t)).join(' ').trim()
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

// Dev-only middleware — production equivalent is api/day-data.js, api/day-data-range.js,
// api/deposits.js, api/export-all.js, api/import-all.js. All of them (dev and prod alike)
// share the same Redis logic via api/_lib/dayDataStore.js, so only the request/response
// plumbing below is duplicated, not the actual store logic.
function dayDataApi(env) {
  return {
    name: 'day-data-api',
    configureServer(server) {
      // dayDataStore.js reads these from process.env (matching how Vercel injects them in
      // production) — under `vite dev` nothing copies .env into process.env automatically,
      // so do it here before any request can reach the (lazily constructed) Redis client.
      // Assigning `undefined` to a process.env property coerces it to the string "undefined"
      // (process.env values are always strings) — only assign when actually set, so an unset
      // VITE_APP_PASSWORD in local .env stays truly unset instead of becoming a truthy string.
      if (env.KV_REST_API_URL) process.env.KV_REST_API_URL = env.KV_REST_API_URL
      if (env.KV_REST_API_TOKEN) process.env.KV_REST_API_TOKEN = env.KV_REST_API_TOKEN
      if (env.VITE_APP_PASSWORD) process.env.VITE_APP_PASSWORD = env.VITE_APP_PASSWORD

      const authed = (req, res) => {
        if (!checkAuth(req)) {
          res.statusCode = 401
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'unauthorized' }))
          return false
        }
        return true
      }

      server.middlewares.use('/api/day-data', async (req, res) => {
        if (!authed(req, res)) return
        try {
          const url = new URL(req.url, 'http://localhost')
          if (req.method === 'GET') {
            const date = url.searchParams.get('date')
            if (!date) {
              res.statusCode = 400
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'date query param required (YYYY-MM-DD)' }))
              return
            }
            const data = await getDay(date)
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ data }))
            return
          }
          if (req.method === 'POST') {
            const body = await readJsonBody(req)
            if (!body.date || body.data === undefined) {
              res.statusCode = 400
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'date and data required' }))
              return
            }
            await setDay(body.date, body.data)
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ ok: true }))
            return
          }
          res.statusCode = 405
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'method not allowed' }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/day-data-range', async (req, res) => {
        if (!authed(req, res)) return
        try {
          const url = new URL(req.url, 'http://localhost')
          const start = url.searchParams.get('start')
          const end = url.searchParams.get('end')
          if (!start || !end) {
            res.statusCode = 400
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'start and end query params required (YYYY-MM-DD)' }))
            return
          }
          const days = await getDaysInRange(start, end)
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ days }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/deposits', async (req, res) => {
        if (!authed(req, res)) return
        try {
          const url = new URL(req.url, 'http://localhost')
          const mode = url.searchParams.get('mode')
          const allDays = await getAllDayBlobs()
          if (mode === 'date') {
            const date = url.searchParams.get('date')
            if (!date) {
              res.statusCode = 400
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'date query param required' }))
              return
            }
            const found = []
            for (const [recordedDate, d] of Object.entries(allDays)) {
              (d.deposits || []).forEach((dep) => {
                if (dep.appointmentDate === date && (dep.type === 'deposit' || dep.type === 'giftcard')) {
                  found.push({ ...dep, recordedDate })
                }
              })
            }
            found.sort((a, b) => (a.appointmentTime || '').localeCompare(b.appointmentTime || ''))
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ deposits: found }))
            return
          }
          if (mode === 'client') {
            const name = (url.searchParams.get('name') || '').toLowerCase().trim()
            if (!name) {
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ deposits: [] }))
              return
            }
            const found = []
            for (const [sheetDate, d] of Object.entries(allDays)) {
              (d.deposits || []).forEach((dep) => {
                if ((dep.clientName || '').toLowerCase().trim() === name && (dep.type === 'deposit' || dep.type === 'giftcard')) {
                  found.push({ ...dep, sheetDate })
                }
              })
            }
            found.sort((a, b) => a.sheetDate.localeCompare(b.sheetDate))
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ deposits: found }))
            return
          }
          res.statusCode = 400
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'mode query param required (date|client)' }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/export-all', async (req, res) => {
        if (!authed(req, res)) return
        try {
          const days = await getAllDayBlobs()
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ days }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/import-all', async (req, res) => {
        if (!authed(req, res)) return
        if (req.method !== 'POST') {
          res.statusCode = 405
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'method not allowed' }))
          return
        }
        try {
          const days = await readJsonBody(req)
          if (typeof days !== 'object' || Object.keys(days).length === 0) {
            res.statusCode = 400
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'body must be a non-empty {date: data} object' }))
            return
          }
          const count = await setDays(days)
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ count }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      // No Cron in local dev — hit this manually (e.g. curl) to test the snapshot logic.
      server.middlewares.use('/api/auto-backup', async (req, res) => {
        try {
          const result = await saveAutoBackup()
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify(result))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/auto-backup-list', async (req, res) => {
        if (!authed(req, res)) return
        try {
          const dates = await listAutoBackups()
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ dates }))
        } catch (e) {
          res.statusCode = 500
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: String(e) }))
        }
      })

      server.middlewares.use('/api/auto-backup-restore', async (req, res) => {
        if (!authed(req, res)) return
        if (req.method !== 'POST') {
          res.statusCode = 405
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ error: 'method not allowed' }))
          return
        }
        try {
          const body = await readJsonBody(req)
          if (!body.date) {
            res.statusCode = 400
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'date required' }))
            return
          }
          const count = await restoreAutoBackup(body.date)
          if (count == null) {
            res.statusCode = 404
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({ error: 'backup not found' }))
            return
          }
          res.setHeader('Content-Type', 'application/json')
          res.end(JSON.stringify({ count }))
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
    plugins: [react(), squareGiftCardApi(env), squareBookingsApi(env), squarePaymentsApi(env), squareDepositsApi(env), dayDataApi(env)],
  }
})
