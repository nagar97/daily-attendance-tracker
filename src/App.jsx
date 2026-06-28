import React, { useState, useEffect, useCallback } from 'react'
import * as XLSX from 'xlsx'

const TEAM = [
  'Mustafa Kamel',
  'Mohamed Sobhy',
  'Nagy',
  'Mohamed Allam',
  'Ahmed Tantawy',
  'Ahmed Habib',
  'Mai Attia',
]

const STATUS_META = {
  Present: { bg: 'var(--green)',  color: 'var(--green-t)',  icon: '✓' },
  Absent:  { bg: 'var(--red)',    color: 'var(--red-t)',    icon: '✗' },
  Late:    { bg: 'var(--yellow)', color: 'var(--yellow-t)', icon: '⏰' },
  Excused: { bg: 'var(--purple)', color: 'var(--purple-t)', icon: '📋' },
}

const TODAY = () => new Date().toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' })

const STORAGE_LOG     = 'attendance_log'
const STORAGE_HISTORY = 'attendance_history'
const STORAGE_DATE    = 'attendance_date'

function initLog() {
  return TEAM.map(name => ({ name, status: 'Present', reason: '', notified: 'N/A', followUp: '', updatedAt: '' }))
}

export default function App() {
  const [log, setLog]         = useState(initLog)
  const [history, setHistory] = useState([])
  const [date, setDate]       = useState(TODAY())
  const [view, setView]       = useState('log')   // 'log' | 'history' | 'stats'
  const [toast, setToast]     = useState(null)
  const [confirmReset, setConfirmReset] = useState(false)

  // Load from localStorage
  useEffect(() => {
    try {
      const savedDate    = localStorage.getItem(STORAGE_DATE)
      const savedLog     = localStorage.getItem(STORAGE_LOG)
      const savedHistory = localStorage.getItem(STORAGE_HISTORY)

      if (savedHistory) setHistory(JSON.parse(savedHistory))

      const todayStr = TODAY()
      if (savedDate === todayStr && savedLog) {
        setLog(JSON.parse(savedLog))
        setDate(todayStr)
      } else if (savedDate && savedDate !== todayStr) {
        // New day detected
        setDate(todayStr)
        setLog(initLog())
      } else if (savedDate === todayStr) {
        setDate(todayStr)
      }
    } catch {}
  }, [])

  // Auto-save log on change
  useEffect(() => {
    localStorage.setItem(STORAGE_LOG, JSON.stringify(log))
    localStorage.setItem(STORAGE_DATE, date)
  }, [log, date])

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 3000)
  }

  const updateMember = (idx, field, value) => {
    setLog(prev => prev.map((m, i) => i === idx
      ? { ...m, [field]: value, updatedAt: field === 'status' ? new Date().toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit' }) : m.updatedAt }
      : m
    ))
  }

  const saveToHistory = () => {
    const alreadySaved = history.some(r => r.date === date)
    if (alreadySaved) {
      const newHistory = history.filter(r => r.date !== date)
      const entries = log.map(m => ({ ...m, date, savedAt: new Date().toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit' }) }))
      const updated = [...newHistory, ...entries]
      setHistory(updated)
      localStorage.setItem(STORAGE_HISTORY, JSON.stringify(updated))
      showToast(`Overwritten — ${date} saved to history`)
    } else {
      const entries = log.map(m => ({ ...m, date, savedAt: new Date().toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit' }) }))
      const updated = [...history, ...entries]
      setHistory(updated)
      localStorage.setItem(STORAGE_HISTORY, JSON.stringify(updated))
      showToast(`Saved — ${date} added to history`)
    }
  }

  const resetLog = () => {
    const isSaved = history.some(r => r.date === date)
    if (!isSaved) {
      setConfirmReset(true)
      return
    }
    doReset()
  }

  const doReset = () => {
    const newDate = TODAY()
    setLog(initLog())
    setDate(newDate)
    setConfirmReset(false)
    showToast(`Reset — ready for ${newDate}`)
  }

  const exportExcel = () => {
    const wb = XLSX.utils.book_new()
    // History sheet
    const histRows = [['Date','Name','Status','Reason','Notified','Follow-up','Saved At']]
    const grouped = {}
    history.forEach(r => { if (!grouped[r.date]) grouped[r.date] = []; grouped[r.date].push(r) })
    Object.keys(grouped).sort().forEach(d => {
      grouped[d].forEach(r => histRows.push([r.date, r.name, r.status, r.reason, r.notified, r.followUp, r.savedAt]))
    })
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(histRows), 'History')
    // Today sheet
    const todayRows = [['Name','Status','Reason','Notified','Follow-up','Last Updated']]
    log.forEach(m => todayRows.push([m.name, m.status, m.reason, m.notified, m.followUp, m.updatedAt]))
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(todayRows), 'Today')
    XLSX.writeFile(wb, `Attendance_${date.replace(/ /g,'-')}.xlsx`)
    showToast('Excel file downloaded')
  }

  // Stats
  const stats = TEAM.map(name => {
    const records = history.filter(r => r.name === name)
    const total   = records.length
    const present = records.filter(r => r.status === 'Present').length
    const absent  = records.filter(r => r.status === 'Absent').length
    const excused = records.filter(r => r.status === 'Excused').length
    const counted = present + absent + excused
    const rate    = counted ? Math.round((present + excused) / counted * 100) : null
    return { name, total, present, absent, excused, rate }
  })

  const counts = {
    Present: log.filter(m => m.status === 'Present').length,
    Absent:  log.filter(m => m.status === 'Absent').length,
    Late:    log.filter(m => m.status === 'Late').length,
    Excused: log.filter(m => m.status === 'Excused').length,
  }

  const uniqueDays = [...new Set(history.map(r => r.date))].length

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <header style={{ background: 'var(--navy)', color: '#fff', padding: '0 24px' }}>
        <div style={{ maxWidth: 900, margin: '0 auto', padding: '16px 0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '-0.3px' }}>Daily Meeting Tracker</div>
            <div style={{ fontSize: 12, color: 'rgba(255,255,255,.6)', marginTop: 2 }}>
              {uniqueDays} day{uniqueDays !== 1 ? 's' : ''} recorded · {history.length} entries
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: 12, color: 'rgba(255,255,255,.6)' }}>Today:</span>
            <span style={{ fontSize: 13, fontWeight: 500, background: 'rgba(255,255,255,.12)', padding: '4px 10px', borderRadius: 20 }}>{date}</span>
          </div>
        </div>
        {/* Nav */}
        <div style={{ maxWidth: 900, margin: '0 auto', display: 'flex', gap: 2, paddingBottom: 0 }}>
          {[['log','📋 Daily Log'],['history','🗂 History'],['stats','📊 Stats']].map(([k,label]) => (
            <button key={k} onClick={() => setView(k)} style={{
              padding: '10px 18px', fontSize: 13, fontWeight: 500, border: 'none',
              background: view === k ? '#fff' : 'transparent',
              color: view === k ? 'var(--navy)' : 'rgba(255,255,255,.7)',
              borderRadius: '8px 8px 0 0', transition: 'all .15s'
            }}>{label}</button>
          ))}
        </div>
      </header>

      {/* Main */}
      <main style={{ flex: 1, maxWidth: 900, margin: '0 auto', width: '100%', padding: '24px 16px' }}>

        {/* ── DAILY LOG ── */}
        {view === 'log' && (
          <div>
            {/* Action bar */}
            <div style={{ display: 'flex', gap: 8, marginBottom: 16, flexWrap: 'wrap' }}>
              <button onClick={saveToHistory} style={btnStyle('var(--blue)')}>
                💾 Save to History
              </button>
              <button onClick={resetLog} style={btnStyle('#475569')}>
                🔄 Reset for New Day
              </button>
              <button onClick={exportExcel} style={btnStyle('#065f46')}>
                ⬇ Export Excel
              </button>
            </div>

            {/* Summary chips */}
            <div style={{ display: 'flex', gap: 8, marginBottom: 20, flexWrap: 'wrap' }}>
              {Object.entries(counts).map(([s,n]) => (
                <div key={s} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 12px', borderRadius: 20, background: STATUS_META[s].bg, color: STATUS_META[s].color, fontSize: 13, fontWeight: 500 }}>
                  <span style={{ fontSize: 15 }}>{STATUS_META[s].icon}</span>
                  <span>{n} {s}</span>
                </div>
              ))}
              <div style={{ marginLeft: 'auto', padding: '5px 14px', borderRadius: 20, background: 'var(--navy)', color: '#fff', fontSize: 13, fontWeight: 500 }}>
                {counts.Present}/{TEAM.length} · {Math.round(counts.Present/TEAM.length*100)}%
              </div>
            </div>

            {/* Table */}
            <div style={{ background: 'var(--white)', borderRadius: var_radius, boxShadow: 'var(--shadow-md)', overflow: 'hidden' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr style={{ background: 'var(--navy)', color: '#fff' }}>
                    {['#','Team Member','Status','Reason / Notes','Notified?','Last Updated'].map(h => (
                      <th key={h} style={{ padding: '11px 14px', textAlign: 'left', fontSize: 12, fontWeight: 600, letterSpacing: '.3px', whiteSpace: 'nowrap' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {log.map((m, i) => {
                    const meta = STATUS_META[m.status]
                    const rowBg = i % 2 === 0 ? 'var(--white)' : 'var(--surface)'
                    return (
                      <tr key={m.name} style={{ background: rowBg, transition: 'background .1s' }}
                          onMouseEnter={e => e.currentTarget.style.background = 'var(--sky)'}
                          onMouseLeave={e => e.currentTarget.style.background = rowBg}>
                        <td style={{ padding: '10px 14px', color: 'var(--muted)', fontSize: 12 }}>{i+1}</td>
                        <td style={{ padding: '10px 14px', fontWeight: 500, fontSize: 13 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <Avatar name={m.name} />
                            {m.name}
                          </div>
                        </td>
                        <td style={{ padding: '10px 14px' }}>
                          <select
                            value={m.status}
                            onChange={e => updateMember(i, 'status', e.target.value)}
                            style={{ padding: '4px 8px', borderRadius: 20, border: 'none', fontSize: 12, fontWeight: 600, background: meta.bg, color: meta.color, appearance: 'none', paddingRight: 24, cursor: 'pointer' }}>
                            {Object.keys(STATUS_META).map(s => <option key={s}>{s}</option>)}
                          </select>
                        </td>
                        <td style={{ padding: '10px 14px' }}>
                          {m.status !== 'Present' ? (
                            <input
                              value={m.reason}
                              onChange={e => updateMember(i, 'reason', e.target.value)}
                              placeholder="Add reason..."
                              style={{ width: '100%', padding: '5px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12, background: 'transparent', color: 'var(--ink)', outline: 'none' }}
                            />
                          ) : <span style={{ color: 'var(--muted)', fontSize: 12 }}>—</span>}
                        </td>
                        <td style={{ padding: '10px 14px' }}>
                          <select
                            value={m.notified}
                            onChange={e => updateMember(i, 'notified', e.target.value)}
                            style={{ padding: '4px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12, background: 'transparent', cursor: 'pointer' }}>
                            {['N/A','Yes','No'].map(v => <option key={v}>{v}</option>)}
                          </select>
                        </td>
                        <td style={{ padding: '10px 14px', color: 'var(--muted)', fontSize: 11, fontFamily: 'JetBrains Mono, monospace' }}>
                          {m.updatedAt || '—'}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── HISTORY ── */}
        {view === 'history' && (
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 8 }}>
              <div style={{ fontSize: 15, fontWeight: 600 }}>Attendance History</div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button onClick={exportExcel} style={btnStyle('#065f46')}>⬇ Export Excel</button>
                {history.length > 0 && (
                  <button onClick={() => { if (window.confirm('Clear all history?')) { setHistory([]); localStorage.removeItem(STORAGE_HISTORY); showToast('History cleared') } }}
                    style={btnStyle('#9f1239')}>🗑 Clear All</button>
                )}
              </div>
            </div>

            {history.length === 0 ? (
              <EmptyState icon="🗂" title="No history yet" sub="Save today's log to start building your record." />
            ) : (() => {
              const grouped = {}
              history.forEach(r => { if (!grouped[r.date]) grouped[r.date] = []; grouped[r.date].push(r) })
              const dates = Object.keys(grouped).sort((a,b) => new Date(b) - new Date(a))
              return dates.map(d => {
                const rows = grouped[d]
                const p = rows.filter(r => r.status === 'Present').length
                return (
                  <div key={d} style={{ marginBottom: 16, background: 'var(--white)', borderRadius: 'var(--radius)', boxShadow: 'var(--shadow)', overflow: 'hidden' }}>
                    <div style={{ background: 'var(--navy)', color: '#fff', padding: '10px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <span style={{ fontWeight: 600, fontSize: 13 }}>{d}</span>
                      <span style={{ fontSize: 12, background: 'rgba(255,255,255,.15)', padding: '2px 10px', borderRadius: 20 }}>
                        {p}/{rows.length} present · {Math.round(p/rows.length*100)}%
                      </span>
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                        <tbody>
                          {rows.map((r,i) => {
                            const meta = STATUS_META[r.status]
                            return (
                              <tr key={i} style={{ borderBottom: '1px solid var(--border)', background: i%2===0?'#fff':'var(--surface)' }}>
                                <td style={{ padding: '8px 16px', fontWeight: 500 }}>{r.name}</td>
                                <td style={{ padding: '8px 16px' }}>
                                  <span style={{ background: meta.bg, color: meta.color, padding: '2px 10px', borderRadius: 20, fontSize: 12, fontWeight: 600 }}>
                                    {meta.icon} {r.status}
                                  </span>
                                </td>
                                <td style={{ padding: '8px 16px', color: 'var(--muted)' }}>{r.reason || '—'}</td>
                                <td style={{ padding: '8px 16px', color: 'var(--muted)', fontSize: 11, fontFamily: 'JetBrains Mono, monospace' }}>{r.savedAt}</td>
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )
              })
            })()}
          </div>
        )}

        {/* ── STATS ── */}
        {view === 'stats' && (
          <div>
            <div style={{ marginBottom: 16, fontSize: 15, fontWeight: 600 }}>Team Statistics <span style={{ fontSize: 12, fontWeight: 400, color: 'var(--muted)' }}>({uniqueDays} days recorded)</span></div>

            {uniqueDays === 0 ? (
              <EmptyState icon="📊" title="No data yet" sub="Save at least one day to see statistics." />
            ) : (
              <div style={{ background: 'var(--white)', borderRadius: 'var(--radius)', boxShadow: 'var(--shadow-md)', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ background: 'var(--navy)', color: '#fff' }}>
                      {['Team Member','Present','Absent','Excused','Attendance %'].map(h => (
                        <th key={h} style={{ padding: '11px 16px', textAlign: h === 'Team Member' ? 'left' : 'center', fontSize: 12, fontWeight: 600, letterSpacing: '.3px' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {stats.sort((a,b) => (b.rate ?? -1) - (a.rate ?? -1)).map((s,i) => (
                      <tr key={s.name} style={{ background: i%2===0?'var(--white)':'var(--surface)', borderBottom: '1px solid var(--border)' }}>
                        <td style={{ padding: '11px 16px', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 8 }}>
                          <Avatar name={s.name} />
                          {s.name}
                        </td>
                        <Num v={s.present} bg="var(--green)" color="var(--green-t)" />
                        <Num v={s.absent}  bg="var(--red)"   color="var(--red-t)" />
                        <Num v={s.excused} bg="var(--purple)"color="var(--purple-t)" />
                        <td style={{ padding: '11px 16px', textAlign: 'center' }}>
                          {s.rate !== null ? (
                            <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
                              <span style={{ fontWeight: 700, fontSize: 15, color: s.rate >= 80 ? 'var(--green-t)' : s.rate >= 60 ? 'var(--yellow-t)' : 'var(--red-t)' }}>{s.rate}%</span>
                              <div style={{ width: 60, height: 4, background: 'var(--border)', borderRadius: 4, overflow: 'hidden' }}>
                                <div style={{ width: `${s.rate}%`, height: '100%', background: s.rate >= 80 ? '#22c55e' : s.rate >= 60 ? '#eab308' : '#ef4444', borderRadius: 4, transition: 'width .5s' }} />
                              </div>
                            </div>
                          ) : <span style={{ color: 'var(--muted)', fontSize: 12 }}>—</span>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Toast */}
      {toast && (
        <div style={{ position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)', background: toast.type === 'error' ? '#7f1d1d' : 'var(--navy)', color: '#fff', padding: '12px 20px', borderRadius: 10, fontSize: 13, fontWeight: 500, boxShadow: '0 8px 24px rgba(0,0,0,.2)', zIndex: 1000, whiteSpace: 'nowrap', animation: 'fadeIn .2s' }}>
          {toast.msg}
        </div>
      )}

      {/* Confirm Reset Modal */}
      {confirmReset && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ background: 'var(--white)', borderRadius: 14, padding: 28, maxWidth: 380, width: '90%', boxShadow: '0 20px 60px rgba(0,0,0,.2)' }}>
            <div style={{ fontSize: 32, marginBottom: 12 }}>⚠️</div>
            <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>Today not saved yet</div>
            <div style={{ fontSize: 13, color: 'var(--muted)', marginBottom: 20, lineHeight: 1.6 }}>
              {date}'s attendance hasn't been saved to history. Resetting will lose all current data.
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => setConfirmReset(false)} style={{ flex: 1, padding: '10px', borderRadius: 8, border: '1px solid var(--border)', background: 'var(--white)', fontSize: 13, fontWeight: 500 }}>Cancel</button>
              <button onClick={doReset} style={{ flex: 1, padding: '10px', borderRadius: 8, border: 'none', background: '#9f1239', color: '#fff', fontSize: 13, fontWeight: 600 }}>Reset Anyway</button>
            </div>
          </div>
        </div>
      )}

      <style>{`
        @keyframes fadeIn { from { opacity:0; transform:translateX(-50%) translateY(8px) } to { opacity:1; transform:translateX(-50%) translateY(0) } }
        select { -webkit-appearance:none; appearance:none; }
        input:focus { border-color: var(--mid) !important; box-shadow: 0 0 0 3px rgba(46,117,182,.15) !important; }
        tr:last-child td { border-bottom: none !important; }
      `}</style>
    </div>
  )
}

const var_radius = 'var(--radius)'

function Avatar({ name }) {
  const initials = name.split(' ').map(w => w[0]).slice(0,2).join('')
  const hue = name.split('').reduce((a,c) => a + c.charCodeAt(0), 0) % 360
  return (
    <div style={{ width: 28, height: 28, borderRadius: '50%', background: `hsl(${hue},45%,55%)`, color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700, flexShrink: 0 }}>
      {initials}
    </div>
  )
}

function Num({ v, bg, color }) {
  return (
    <td style={{ padding: '11px 16px', textAlign: 'center' }}>
      {v > 0
        ? <span style={{ background: bg, color, padding: '2px 10px', borderRadius: 20, fontSize: 12, fontWeight: 700 }}>{v}</span>
        : <span style={{ color: 'var(--muted)', fontSize: 12 }}>—</span>}
    </td>
  )
}

function EmptyState({ icon, title, sub }) {
  return (
    <div style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--muted)' }}>
      <div style={{ fontSize: 40, marginBottom: 12 }}>{icon}</div>
      <div style={{ fontSize: 15, fontWeight: 600, color: 'var(--ink)', marginBottom: 6 }}>{title}</div>
      <div style={{ fontSize: 13 }}>{sub}</div>
    </div>
  )
}

function btnStyle(bg) {
  return {
    padding: '9px 16px', borderRadius: 8, border: 'none',
    background: bg, color: '#fff', fontSize: 13, fontWeight: 500,
    display: 'flex', alignItems: 'center', gap: 6,
    transition: 'opacity .15s', cursor: 'pointer',
  }
}
