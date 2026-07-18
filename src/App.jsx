import { useState } from 'react'
import SpaDailySheet from './spa-daily-sheet'

const APP_PASSWORD = import.meta.env.VITE_APP_PASSWORD
const AUTH_KEY = 'salon-app-authed'

// Simple shared-password gate for when the app is hosted somewhere reachable over the
// internet (not just localhost) — same "soft deterrent" model as the existing MANAGER_PIN:
// it stops a random person who finds the URL, not a determined attacker (the check runs
// client-side, so the password ships in the bundle). If VITE_APP_PASSWORD isn't set (e.g.
// local dev), the gate is skipped entirely.
function PasswordGate({ children }) {
  const [authed, setAuthed] = useState(() => localStorage.getItem(AUTH_KEY) === 'true')
  const [input, setInput] = useState('')
  const [error, setError] = useState(false)

  if (!APP_PASSWORD || authed) return children

  const handleSubmit = (e) => {
    e.preventDefault()
    if (input === APP_PASSWORD) {
      localStorage.setItem(AUTH_KEY, 'true')
      setAuthed(true)
    } else {
      setError(true)
    }
  }

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#F7F4EE', fontFamily: "'Inter','Helvetica Neue',sans-serif" }}>
      <form onSubmit={handleSubmit} style={{ background: '#fff', borderRadius: 16, padding: '32px 28px', width: 320, boxShadow: '0 4px 24px rgba(0,0,0,0.08)' }}>
        <div style={{ fontSize: 20, fontWeight: 800, color: '#0D4F4F', marginBottom: 4 }}>🌺 Spa Daily Sheet</div>
        <div style={{ fontSize: 13, color: '#888', marginBottom: 20 }}>Please enter the password</div>
        <input
          type="password"
          value={input}
          onChange={e => { setInput(e.target.value); setError(false) }}
          autoFocus
          style={{ width: '100%', padding: '10px 12px', borderRadius: 8, border: `2px solid ${error ? '#C62828' : '#DDD'}`, fontSize: 15, boxSizing: 'border-box' }}
          placeholder="Password"
        />
        {error && <div style={{ color: '#C62828', fontSize: 12, marginTop: 6 }}>Incorrect password</div>}
        <button type="submit" style={{ width: '100%', marginTop: 16, padding: '11px', borderRadius: 8, border: 'none', background: '#0D4F4F', color: '#fff', fontWeight: 700, fontSize: 14, cursor: 'pointer' }}>
          Open
        </button>
      </form>
    </div>
  )
}

function App() {
  return (
    <PasswordGate>
      <SpaDailySheet />
    </PasswordGate>
  )
}

export default App
