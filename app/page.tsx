'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { AlertTriangle, FileSpreadsheet, RefreshCcw, Search, Settings2, Upload } from 'lucide-react'
import { Bar, BarChart, CartesianGrid, Cell, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts'

type Row = Record<string, any>
type ScheduleRow = { dateKey: string; ps: number; scheduled: number }
type Master = { ps: Row[]; schedule: ScheduleRow[] }

const fmt = (v: unknown) => new Intl.NumberFormat('en-IN').format(Number(v || 0))
const safe = (v: unknown) => v == null ? '' : String(v)
const num = (v: unknown) => Number(v || 0)
const dateKey = (v: unknown) => {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0, 10)
  const s = safe(v).trim()
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) { const [dd, mm, yyyy] = s.split('.'); return `${yyyy}-${mm}-${dd}` }
  const d = new Date(s)
  return Number.isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10)
}
const dateLabel = (d: string) => d ? new Intl.DateTimeFormat('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${d}T00:00:00`)) : '—'
const readSheet = (wb: XLSX.WorkBook, name: string): Row[] => wb.Sheets[name] ? XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { defval: null, raw: true }) : []

function extractEci(wb: XLSX.WorkBook): Row[] {
  let rows = readSheet(wb, 'sirNoticeGenerate')
  if (!rows.length) rows = readSheet(wb, 'ECI_INPUT')
  if (!rows.length) {
    for (const name of wb.SheetNames) {
      const candidate = readSheet(wb, name)
      if (candidate[0] && ('Part No' in candidate[0] || 'POLLING STATION' in candidate[0])) { rows = candidate; break }
    }
  }
  return rows.filter(r => num(r['AC Number']) === 34 || safe(r['Asmbly Name']).toUpperCase() === 'MATIALA')
}

function extractMaster(wb: XLSX.WorkBook): Master {
  const ps = readSheet(wb, 'PS_MASTER')
  const hearing = readSheet(wb, 'HEARING_DATA')
  const schedule: ScheduleRow[] = hearing.filter(r => num(r['PS No.']) > 0 && dateKey(r.Date)).map(r => ({ dateKey: dateKey(r.Date), ps: num(r['PS No.']), scheduled: num(r['Scheduled Notices for Hearing']) }))
  return { ps, schedule }
}

function extractSchedule(wb: XLSX.WorkBook): ScheduleRow[] {
  let rows = readSheet(wb, 'Part Wise Hearing Summary')
  if (!rows.length) rows = readSheet(wb, 'HEARING_DATA')
  return rows.filter(r => num(r['Part No.'] ?? r['PS No.']) > 0 && dateKey(r['Hearing Date'] ?? r.Date)).map(r => ({ dateKey: dateKey(r['Hearing Date'] ?? r.Date), ps: num(r['Part No.'] ?? r['PS No.']), scheduled: num(r['Total Hearings'] ?? r['Scheduled Notices for Hearing']) }))
}

export default function Home() {
  const [eci, setEci] = useState<Row[]>([])
  const [master, setMaster] = useState<Master | null>(null)
  const [date, setDate] = useState('')
  const [query, setQuery] = useState('')
  const [tab, setTab] = useState<'overview' | 'officers' | 'ps' | 'eci'>('overview')
  const [error, setError] = useState('')
  const [source, setSource] = useState('')
  const [ready, setReady] = useState(false)

  useEffect(() => {
    try { const raw = localStorage.getItem('ac34_master_v2'); if (raw) setMaster(JSON.parse(raw) as Master) } catch { /* ignore corrupt cache */ }
    setReady(true)
  }, [])

  const saveMaster = (m: Master, labelText: string) => {
    setMaster(m); localStorage.setItem('ac34_master_v2', JSON.stringify(m)); setDate(m.schedule.map(x => x.dateKey).filter(Boolean).sort().at(-1) || ''); setSource(labelText)
  }

  const uploadDashboard = async (file: File) => {
    try { setError(''); const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true }); const m = extractMaster(wb); if (!m.ps.length || !m.schedule.length) throw new Error('The dashboard workbook must contain PS_MASTER and HEARING_DATA.'); saveMaster(m, `${file.name} — master/schedule loaded`) }
    catch (e) { setError(e instanceof Error ? e.message : 'Unable to read dashboard workbook.') }
  }

  const uploadSchedule = async (file: File) => {
    try { setError(''); const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true }); const schedule = extractSchedule(wb); if (!schedule.length) throw new Error('No part-wise hearing rows found. Expected Part No., Hearing Date and Total Hearings.'); const current: Master = master ? { ...master, schedule } : { ps: [], schedule }; saveMaster(current, `${file.name} — hearing schedule loaded`) }
    catch (e) { setError(e instanceof Error ? e.message : 'Unable to read hearing schedule.') }
  }

  const uploadEci = async (file: File) => {
    try { setError(''); const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true }); const rows = extractEci(wb); if (!rows.length) throw new Error('No AC-34 / MATIALA rows found. Please upload the ECI NOTICE_REPORT_PART_WISE Excel file.'); setEci(rows); setSource(file.name); if (master?.schedule.length) setDate(master.schedule.map(x => x.dateKey).filter(Boolean).sort().at(-1) || '') }
    catch (e) { setError(e instanceof Error ? e.message : 'Unable to read ECI Excel.') }
  }

  const dates = useMemo(() => [...new Set((master?.schedule || []).map(x => x.dateKey).filter(Boolean))].sort(), [master])
  const eciMap = useMemo(() => new Map(eci.map(r => [num(r['Part No'] ?? r['POLLING STATION']), r])), [eci])
  const psMap = useMemo(() => new Map((master?.ps || []).map(r => [num(r['PS No.']), r])), [master])
  const active: Row[] = useMemo(() => (master?.schedule || []).filter(s => s.dateKey === date).map((s): Row => ({ ...(psMap.get(s.ps) || {}), ...(eciMap.get(s.ps) || {}), 'PS No.': s.ps, 'Scheduled Notices for Hearing': s.scheduled, 'Date': s.dateKey })), [master, date, eciMap, psMap])
  const filtered: Row[] = useMemo(() => { const q = query.trim().toLowerCase(); return active.filter(r => !q || [r['PS No.'], r['Old PS No.'], r.Officer, r['Hearing Centre'], r.BLO, r.Supervisor, r.Locality, r['Polling Area']].some(v => safe(v).toLowerCase().includes(q))) }, [active, query])
  const total = useMemo(() => active.reduce((a, r) => ({ ps: a.ps + 1, scheduled: a.scheduled + num(r['Scheduled Notices for Hearing']), generated: a.generated + num(r['Notice Generated']), delivered: a.delivered + num(r['Notice Delivered']), pending: a.pending + num(r['Notice Pending Delivery']), held: a.held + num(r['Hearings Held']) }), { ps: 0, scheduled: 0, generated: 0, delivered: 0, pending: 0, held: 0 }), [active])
  const officers = useMemo(() => { const m = new Map<string, Row>(); active.forEach(r => { const k = `${safe(r.Officer)}|${safe(r['Hearing Centre'])}`; const x: Row = m.get(k) || { Officer: r.Officer, Mobile: r['Officer Mobile'], Centre: r['Hearing Centre'], PS: 0, Scheduled: 0, Generated: 0, Delivered: 0, Pending: 0, Held: 0 }; x.PS += 1; x.Scheduled += num(r['Scheduled Notices for Hearing']); x.Generated += num(r['Notice Generated']); x.Delivered += num(r['Notice Delivered']); x.Pending += num(r['Notice Pending Delivery']); x.Held += num(r['Hearings Held']); m.set(k, x) }); return [...m.values()] }, [active])
  const global = useMemo(() => eci.reduce((a, r) => ({ generated: a.generated + num(r['Notice Generated']), delivered: a.delivered + num(r['Notice Delivered']), pending: a.pending + num(r['Notice Pending Delivery']), held: a.held + num(r['Hearings Held']), hearingLapsed: a.hearingLapsed + num(r['Hearing Date Lapsed']), rescheduleLapsed: a.rescheduleLapsed + num(r['Reschedule Date Lapsed']), verified: a.verified + num(r['DEO-Status Verified']), notVerified: a.notVerified + num(r['DEO-Status Not Verified']) }), { generated: 0, delivered: 0, pending: 0, held: 0, hearingLapsed: 0, rescheduleLapsed: 0, verified: 0, notVerified: 0 }), [eci])
  const chart = officers.map(o => ({ name: safe(o.Officer).replace(/^SH\\. |^SMT\\. /, ''), Scheduled: o.Scheduled, Delivered: o.Delivered, Pending: o.Pending }))
  if (!ready) return null

  return <div className="app-shell"><header className="topbar"><div className="brand"><div><div className="title">AC-34 MATIALA — SIR 2026 HEARING & NOTICE DASHBOARD</div><div className="subtitle">Upload the latest ECI NOTICE_REPORT_PART_WISE file — no manual Excel formula work required.</div></div><div className="updated">{source && <>Source: {source}<br /></>}{eci.length ? `${fmt(eci.length)} ECI rows loaded` : master ? `${fmt(master.ps.length)} PS master records` : ''}</div></div></header><main className="page">
    <div className="toolbar"><div className="field"><label>ECI report — normal daily upload</label><label className="upload"><Upload size={16} /> Upload ECI Excel<input type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && uploadEci(e.target.files[0])} /></label></div><div className="field"><label>Hearing Date</label><select disabled={!master?.schedule.length} value={date} onChange={e => setDate(e.target.value)}><option value="">Select hearing date</option>{dates.map(d => <option key={d} value={d}>{dateLabel(d)}</option>)}</select></div><div className="field"><label>Search</label><div style={{ position: 'relative' }}><Search size={16} style={{ position: 'absolute', left: 11, top: 12, color: '#829ab1' }} /><input className="upload" style={{ paddingLeft: 34, width: '100%' }} value={query} onChange={e => setQuery(e.target.value)} placeholder="PS / officer / centre / BLO..." /></div></div><button className="tab" onClick={() => setDate(dates.at(-1) || '')}><RefreshCcw size={14} /> Latest date</button></div>
    {!master?.ps.length && <div className="panel setup"><Settings2 size={28} /><h2>One-time master setup</h2><p>Upload your existing AC-34 dashboard workbook once. It provides the 430-PS officer/BLO/supervisor master. After that, your routine is only the ECI Excel upload.</p><label className="upload"><Upload size={16} /> Upload dashboard workbook<input type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && uploadDashboard(e.target.files[0])} /></label></div>}
    {master?.ps.length && <div className="panel setup compact"><b>Hearing schedule:</b> {fmt(master.schedule.length)} part/date records loaded. <label className="upload"><Upload size={14} /> Update from Part Wise Hearing Summary<input type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && uploadSchedule(e.target.files[0])} /></label></div>}
    {error && <div className="error"><AlertTriangle size={15} /> {error}</div>}
    {master?.ps.length && eci.length > 0 && <><div className="section-title">Selected Hearing Date — {dateLabel(date)}</div><div className="cards"><div className="card"><div className="kicker">Polling Stations</div><div className="metric">{fmt(total.ps)}</div><div className="delta">scheduled</div></div><div className="card"><div className="kicker">Scheduled Hearings</div><div className="metric">{fmt(total.scheduled)}</div><div className="delta">from part-wise schedule</div></div><div className="card"><div className="kicker">Notice Generated</div><div className="metric">{fmt(total.generated)}</div></div><div className="card"><div className="kicker">Delivered</div><div className="metric">{fmt(total.delivered)}</div><div className="delta">{total.scheduled ? Math.round(total.delivered / total.scheduled * 100) : 0}%</div></div><div className="card"><div className="kicker">Pending Delivery</div><div className="metric">{fmt(total.pending)}</div></div><div className="card"><div className="kicker">Hearings Held</div><div className="metric">{fmt(total.held)}</div></div></div><div className="tabs"><button className={`tab ${tab === 'overview' ? 'active' : ''}`} onClick={() => setTab('overview')}>Overview</button><button className={`tab ${tab === 'officers' ? 'active' : ''}`} onClick={() => setTab('officers')}>Officer-wise</button><button className={`tab ${tab === 'ps' ? 'active' : ''}`} onClick={() => setTab('ps')}>PS-wise Detail</button><button className={`tab ${tab === 'eci' ? 'active' : ''}`} onClick={() => setTab('eci')}>ECI Status</button></div>
    {tab === 'overview' && <><div className="grid2"><div className="panel"><h3>Officer workload</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={chart} margin={{ top: 8, right: 15, left: -20, bottom: 30 }}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" angle={-25} textAnchor="end" interval={0} height={60} fontSize={10} /><YAxis fontSize={10} /><Tooltip /><Bar dataKey="Scheduled" /><Bar dataKey="Delivered" /><Bar dataKey="Pending" /></BarChart></ResponsiveContainer></div></div><div className="panel"><h3>Notice delivery status</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={[{ name: 'Delivered', value: total.delivered }, { name: 'Pending', value: total.pending }]} innerRadius={62} outerRadius={92} dataKey="value" label>{[0, 1].map(i => <Cell key={i} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer></div></div></div><div className="panel" style={{ marginTop: 14 }}><h3>Latest ECI operational snapshot</h3><div className="details"><div className="detail"><b>Notice generated</b>{fmt(global.generated)}</div><div className="detail"><b>Notice delivered</b>{fmt(global.delivered)}</div><div className="detail"><b>Pending delivery</b>{fmt(global.pending)}</div><div className="detail"><b>Hearings held</b>{fmt(global.held)}</div><div className="detail"><b>Hearing-date lapsed</b>{fmt(global.hearingLapsed)}</div><div className="detail"><b>Reschedule lapsed</b>{fmt(global.rescheduleLapsed)}</div><div className="detail"><b>DEO verified</b>{fmt(global.verified)}</div><div className="detail"><b>DEO not verified</b>{fmt(global.notVerified)}</div></div></div></>}
    {tab === 'officers' && <div className="panel"><h3>Officer-wise Hearing Summary</h3><div className="table-wrap"><table className="table"><thead><tr><th>Officer</th><th>Mobile</th><th>Hearing Centre</th><th>PS</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Delivery %</th></tr></thead><tbody>{officers.map((o, i) => <tr key={i}><td><b>{safe(o.Officer)}</b></td><td>{safe(o.Mobile)}</td><td>{safe(o.Centre)}</td><td>{fmt(o.PS)}</td><td>{fmt(o.Scheduled)}</td><td>{fmt(o.Generated)}</td><td>{fmt(o.Delivered)}</td><td>{fmt(o.Pending)}</td><td>{fmt(o.Held)}</td><td>{o.Scheduled ? `${Math.round(o.Delivered / o.Scheduled * 100)}%` : '—'}</td></tr>)}</tbody></table></div></div>}
    {tab === 'ps' && <div className="panel"><h3>PS-wise Detail — {fmt(filtered.length)} records</h3><div className="table-wrap"><table className="table"><thead><tr><th>PS</th><th>Old PS</th><th>Officer</th><th>Hearing Centre</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Locality</th><th>Polling Area</th><th>Anomaly</th><th>Mapping</th><th>Grand Total</th><th>Voters</th></tr></thead><tbody>{filtered.map((r, i) => <tr key={i}><td><b>{safe(r['PS No.'])}</b></td><td>{safe(r['Old PS No.'])}</td><td>{safe(r.Officer)}</td><td>{safe(r['Hearing Centre'])}</td><td>{safe(r.BLO)}</td><td>{safe(r.Supervisor)}</td><td>{fmt(r['Scheduled Notices for Hearing'])}</td><td>{fmt(r['Notice Generated'])}</td><td>{fmt(r['Notice Delivered'])}</td><td>{fmt(r['Notice Pending Delivery'])}</td><td>{fmt(r['Hearings Held'])}</td><td>{safe(r.Locality)}</td><td>{safe(r['Polling Area'])}</td><td>{fmt(r['Total Anomaly/Discrepancy'])}</td><td>{fmt(r['Total No Mapping'])}</td><td>{fmt(r['Grand Total'])}</td><td>{fmt(r['Total Voters'])}</td></tr>)}</tbody></table></div></div>}
    {tab === 'eci' && <div className="panel"><h3>Raw ECI Status — {fmt(eci.length)} AC-34 rows</h3><div className="table-wrap"><table className="table"><thead><tr><th>Part No</th><th>Notice Generated</th><th>Pending Generation</th><th>Delivered</th><th>Pending Delivery</th><th>Hearings Held</th><th>Hearing Lapsed</th><th>Reschedule Lapsed</th><th>DEO Pending</th><th>DEO &gt;5 Days</th><th>Verified</th><th>Not Verified</th></tr></thead><tbody>{eci.map((r, i) => <tr key={i}><td><b>{safe(r['Part No'] ?? r['POLLING STATION'])}</b></td><td>{fmt(r['Notice Generated'])}</td><td>{fmt(r['Pending for Notice Generation'])}</td><td>{fmt(r['Notice Delivered'])}</td><td>{fmt(r['Notice Pending Delivery'])}</td><td>{fmt(r['Hearings Held'])}</td><td>{fmt(r['Hearing Date Lapsed'])}</td><td>{fmt(r['Reschedule Date Lapsed'])}</td><td>{fmt(r['DEO-Status Total Pending'])}</td><td>{fmt(r['DEO-Status Pending GT 5 Days'])}</td><td>{fmt(r['DEO-Status Verified'])}</td><td>{fmt(r['DEO-Status Not Verified'])}</td></tr>)}</tbody></table></div></div>}
    </>}
    {!eci.length && master?.ps.length && <div className="panel empty"><FileSpreadsheet size={42} /><h2>Upload the latest ECI report</h2><p>Your master and hearing schedule are ready. Upload the ECI NOTICE_REPORT_PART_WISE Excel to populate the live figures.</p></div>}
    </main></div>
}
