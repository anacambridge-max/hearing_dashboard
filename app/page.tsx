'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { Upload, RefreshCcw, Search, FileSpreadsheet, AlertTriangle } from 'lucide-react'

type Row = Record<string, any>
type Model = { eci: Row[]; ps: Row[]; schedule: Row[]; dates: string[]; source: string; uploadedAt: string }

const fmt = (v: any) => new Intl.NumberFormat('en-IN').format(Number(v || 0))
const safe = (v: any) => v == null ? '' : String(v)
const num = (v: any) => Number(v || 0)
const dateKey = (v: any) => {
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 10)
  const d = new Date(safe(v))
  return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10)
}
const dateLabel = (v: string) => v ? new Intl.DateTimeFormat('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(v + 'T00:00:00')) : '—'

function sheet(wb: XLSX.WorkBook, name: string): Row[] {
  const ws = wb.Sheets[name]
  return ws ? XLSX.utils.sheet_to_json<Row>(ws, { defval: null, raw: true }) : []
}

function buildModel(wb: XLSX.WorkBook, source: string): Model {
  const eci = sheet(wb, 'ECI_INPUT').filter(r => num(r['AC Number']) === 34 || safe(r['Asmbly Name']).toUpperCase() === 'MATIALA')
  const ps = sheet(wb, 'PS_MASTER')
  const schedule = sheet(wb, 'HEARING_DATA').filter(r => num(r['PS No.']) > 0).map(r => ({ ...r, dateKey: dateKey(r.Date) }))
  const dates = [...new Set(schedule.map(r => r.dateKey).filter(Boolean))].sort()
  return { eci, ps, schedule, dates, source, uploadedAt: new Date().toLocaleString('en-IN') }
}

export default function Home() {
  const [model, setModel] = useState<Model | null>(null)
  const [selectedDate, setSelectedDate] = useState('')
  const [query, setQuery] = useState('')
  const [tab, setTab] = useState<'overview' | 'officers' | 'ps' | 'eci'>('overview')
  const [error, setError] = useState('')

  const load = async (file: File) => {
    try {
      setError('')
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true })
      for (const required of ['ECI_INPUT', 'PS_MASTER', 'HEARING_DATA']) {
        if (!wb.Sheets[required]) throw new Error(`${required} sheet not found. Upload the same workbook structure used for the ECI paste/update workflow.`)
      }
      const m = buildModel(wb, file.name)
      if (!m.eci.length) throw new Error('No AC-34 / MATIALA rows were found in ECI_INPUT.')
      setModel(m)
      setSelectedDate(m.dates[m.dates.length - 1] || '')
      setTab('overview')
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Unable to read workbook.')
    }
  }

  const active = useMemo(() => model?.schedule.filter(r => r.dateKey === selectedDate) || [], [model, selectedDate])
  const psMap = useMemo(() => new Map((model?.ps || []).map(p => [num(p['PS No.']), p])), [model])
  const enriched = useMemo(() => active.map(r => ({ ...r, ...(psMap.get(num(r['PS No.'])) || {}) })), [active, psMap])
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return enriched.filter(r => !q || [r.Officer, r['Hearing Centre'], r.BLO, r.Supervisor, r['PS No.'], r['Old PS No.'], r.Locality, r['Polling Area']].some(v => safe(v).toLowerCase().includes(q)))
  }, [enriched, query])

  const total = useMemo(() => enriched.reduce((a, r) => ({
    ps: a.ps + 1,
    scheduled: a.scheduled + num(r['Scheduled Notices for Hearing']),
    generated: a.generated + num(r['Notice Generated']),
    delivered: a.delivered + num(r['Notice Delivered']),
    pending: a.pending + num(r['Notice Pending Delivery']),
    held: a.held + num(r['Hearings Held'])
  }), { ps: 0, scheduled: 0, generated: 0, delivered: 0, pending: 0, held: 0 }), [enriched])

  const global = useMemo(() => (model?.eci || []).reduce((a, r) => ({
    generated: a.generated + num(r['Notice Generated']),
    delivered: a.delivered + num(r['Notice Delivered']),
    pending: a.pending + num(r['Notice Pending Delivery']),
    held: a.held + num(r['Hearings Held']),
    lapsed: a.lapsed + num(r['Hearing Date Lapsed']) + num(r['Reschedule Date Lapsed']),
    verified: a.verified + num(r['DEO-Status Verified']),
    notVerified: a.notVerified + num(r['DEO-Status Not Verified'])
  }), { generated: 0, delivered: 0, pending: 0, held: 0, lapsed: 0, verified: 0, notVerified: 0 }), [model])

  const officers = useMemo(() => {
    const map = new Map<string, Row>()
    for (const r of enriched) {
      const key = `${safe(r.Officer)}|||${safe(r['Hearing Centre'])}`
      const x = map.get(key) || { Officer: r.Officer, Mobile: r['Officer Mobile'], Centre: r['Hearing Centre'], PS: 0, Scheduled: 0, Generated: 0, Delivered: 0, Pending: 0, Held: 0 }
      x.PS++
      x.Scheduled += num(r['Scheduled Notices for Hearing'])
      x.Generated += num(r['Notice Generated'])
      x.Delivered += num(r['Notice Delivered'])
      x.Pending += num(r['Notice Pending Delivery'])
      x.Held += num(r['Hearings Held'])
      map.set(key, x)
    }
    return [...map.values()]
  }, [enriched])

  const pct = total.scheduled ? Math.round(total.delivered / total.scheduled * 100) : 0

  return <div className="app-shell">
    <header className="topbar"><div className="brand"><div><div className="title">AC-34 MATIALA — SIR 2026 HEARING & NOTICE DASHBOARD</div><div className="subtitle">Upload the latest ECI workbook and refresh the complete hearing position.</div></div>{model && <div className="updated">Loaded: {model.uploadedAt}<br />File: {model.source}</div>}</div></header>
    <main className="page">
      <div className="toolbar">
        <div className="field"><label>ECI Input Workbook</label><label className="upload"><Upload size={16} /> Upload latest Excel<input type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && load(e.target.files[0])} /></label></div>
        <div className="field"><label>Hearing Date</label><select disabled={!model} value={selectedDate} onChange={e => setSelectedDate(e.target.value)}><option value="">Select hearing date</option>{model?.dates.map(d => <option key={d} value={d}>{dateLabel(d)}</option>)}</select></div>
        <div className="field"><label>Search</label><div style={{ position: 'relative' }}><Search size={16} style={{ position: 'absolute', left: 11, top: 12, color: '#829ab1' }} /><input className="upload" style={{ paddingLeft: 34, width: '100%' }} value={query} onChange={e => setQuery(e.target.value)} placeholder="PS / officer / centre / BLO..." /></div></div>
        <button className="tab" onClick={() => model && setSelectedDate(model.dates[model.dates.length - 1] || '')}><RefreshCcw size={14} /> Latest date</button>
        <div className="status">{model ? `${fmt(model.eci.length)} ECI rows • ${fmt(model.ps.length)} PS master rows • ${model.dates.length} hearing dates` : 'Upload the ECI workbook to begin'}</div>
      </div>
      {error && <div className="error"><AlertTriangle size={15} style={{ verticalAlign: '-3px', marginRight: 6 }} />{error}</div>}

      {!model ? <div className="panel" style={{ marginTop: 16 }}><div className="empty"><FileSpreadsheet size={42} style={{ marginBottom: 10 }} /><h2 style={{ margin: '4px 0 8px' }}>Upload your latest ECI report</h2><p style={{ margin: 0 }}>Upload the complete workbook used in your current hearing workflow. The dashboard will recalculate the selected-date, officer-wise, PS-wise and ECI status views in the browser.</p></div></div> : <>
        <div className="section-title">Selected Hearing Date — {dateLabel(selectedDate)}</div>
        <div className="cards">
          {[["Polling Stations", total.ps, 'scheduled for this date'], ["Scheduled Notices", total.scheduled, 'hearing workload'], ["Notice Generated", total.generated, 'selected schedule'], ["Delivered", total.delivered, `${pct}% of scheduled`], ["Pending Delivery", total.pending, 'remaining notices'], ["Hearings Held", total.held, 'recorded by ECI']].map(([k, v, d]) => <div className="card" key={String(k)}><div className="kicker">{k}</div><div className="metric">{fmt(v)}</div><div className="delta">{d}</div></div>)}
        </div>
        {total.ps === 0 && <div className="error">No hearing is scheduled for the selected date.</div>}
        <div className="tabs">{(['overview', 'officers', 'ps', 'eci'] as const).map(t => <button key={t} className={`tab ${tab === t ? 'active' : ''}`} onClick={() => setTab(t)}>{t === 'ps' ? 'PS-wise Detail' : t === 'eci' ? 'ECI Status' : t === 'officers' ? 'Officer-wise' : 'Overview'}</button>)}</div>

        {tab === 'overview' && <>
          <div className="grid2"><div className="panel"><h3>Selected-date workload</h3><div className="details">{officers.map(o => <div className="detail" key={`${o.Officer}-${o.Centre}`}><b>{safe(o.Officer) || 'Officer not mapped'}</b>{safe(o.Centre)}<br />PS: {fmt(o.PS)} • Scheduled: {fmt(o.Scheduled)}<br />Delivered: {fmt(o.Delivered)} • Pending: {fmt(o.Pending)}</div>)}</div>{!officers.length && <div className="empty">No officer records for this date.</div>}</div>
          <div className="panel"><h3>Operational snapshot from latest ECI data</h3><div className="details"><div className="detail"><b>Total generated</b>{fmt(global?.generated)}</div><div className="detail"><b>Total delivered</b>{fmt(global?.delivered)}</div><div className="detail"><b>Total pending</b>{fmt(global?.pending)}</div><div className="detail"><b>Hearing + reschedule lapses</b>{fmt(global?.lapsed)}</div><div className="detail"><b>DEO verified</b>{fmt(global?.verified)}</div><div className="detail"><b>DEO not verified</b>{fmt(global?.notVerified)}</div></div></div></div>
        </>}

        {tab === 'officers' && <div className="panel"><h3>Officer-wise Hearing Summary</h3><div className="table-wrap"><table className="table"><thead><tr><th>Officer</th><th>Mobile</th><th>Hearing Centre</th><th>PS</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Delivery %</th></tr></thead><tbody>{officers.map((o, i) => <tr key={i}><td><b>{safe(o.Officer)}</b></td><td>{safe(o.Mobile)}</td><td>{safe(o.Centre)}</td><td>{fmt(o.PS)}</td><td className="right">{fmt(o.Scheduled)}</td><td className="right">{fmt(o.Generated)}</td><td className="right">{fmt(o.Delivered)}</td><td className="right">{fmt(o.Pending)}</td><td className="right">{fmt(o.Held)}</td><td>{o.Scheduled ? `${Math.round(o.Delivered / o.Scheduled * 100)}%` : '—'}</td></tr>)}<tr><td><b>TOTAL</b></td><td colSpan={2}></td><td>{fmt(total.ps)}</td><td className="right"><b>{fmt(total.scheduled)}</b></td><td className="right"><b>{fmt(total.generated)}</b></td><td className="right"><b>{fmt(total.delivered)}</b></td><td className="right"><b>{fmt(total.pending)}</b></td><td className="right"><b>{fmt(total.held)}</b></td><td><b>{pct}%</b></td></tr></tbody></table></div></div>}

        {tab === 'ps' && <div className="panel"><h3>PS-wise Detail — {dateLabel(selectedDate)}</h3><div className="table-wrap"><table className="table"><thead><tr><th>PS</th><th>Old PS</th><th>Officer</th><th>Hearing Centre</th><th>PS Address</th><th>BLO</th><th>Supervisor</th><th>Locality</th><th>Polling Area</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th></tr></thead><tbody>{filtered.map((r, i) => <tr key={i}><td><b>{safe(r['PS No.'])}</b></td><td>{safe(r['Old PS No.'])}</td><td>{safe(r.Officer)}</td><td>{safe(r['Hearing Centre'])}</td><td>{safe(r['PS Address'])}</td><td>{safe(r.BLO)}</td><td>{safe(r.Supervisor)}</td><td>{safe(r.Locality)}</td><td>{safe(r['Polling Area'])}</td><td className="right">{fmt(r['Scheduled Notices for Hearing'])}</td><td className="right">{fmt(r['Notice Generated'])}</td><td className="right">{fmt(r['Notice Delivered'])}</td><td className="right">{fmt(r['Notice Pending Delivery'])}</td><td className="right">{fmt(r['Hearings Held'])}</td></tr>)}</tbody></table></div><div className="footer-note">Showing {fmt(filtered.length)} of {fmt(enriched.length)} PS records.</div></div>}

        {tab === 'eci' && <div className="panel"><h3>Complete ECI Status — AC-34 Matiala</h3><div className="table-wrap"><table className="table"><thead><tr>{['POLLING STATION','Notice Generated','Pending for Notice Generation','Notice Delivered','Notice Pending Delivery','Hearings Held','Hearing Date Lapsed','Reschedule Date Lapsed','DEO-Status Total Pending','DEO-Status Pending GT 5 Days','DEO-Status Verified','DEO-Status Not Verified','ERO/AERO Status Found Ineligible For Final w.r.t. Notice Generated','ERO/AERO Status Parked For Final Publication','ERO/AERO Parked For Final Publication w.r.t. Others'].map(h => <th key={h}>{h}</th>)}</tr></thead><tbody>{model.eci.map((r, i) => <tr key={i}>{['POLLING STATION','Notice Generated','Pending for Notice Generation','Notice Delivered','Notice Pending Delivery','Hearings Held','Hearing Date Lapsed','Reschedule Date Lapsed','DEO-Status Total Pending','DEO-Status Pending GT 5 Days','DEO-Status Verified','DEO-Status Not Verified','ERO/AERO Status Found Ineligible For Final w.r.t. Notice Generated','ERO/AERO Status Parked For Final Publication','ERO/AERO Parked For Final Publication w.r.t. Others'].map(h => <td key={h}>{safe(r[h])}</td>)}</tr>)}</tbody></table></div></div>}
      </>}
    </main>
  </div>
}
