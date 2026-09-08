'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { Upload, RefreshCcw, Search, FileSpreadsheet, AlertTriangle, Settings2 } from 'lucide-react'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'

type Row = Record<string, any>
type ScheduleRow = { dateKey: string; ps: number; scheduled: number }
type Master = { ps: Row[]; schedule: ScheduleRow[] }

const fmt=(v:unknown)=>new Intl.NumberFormat('en-IN').format(Number(v||0))
const safe=(v:unknown)=>v==null?'':String(v)
const num=(v:unknown)=>Number(v||0)

// Keep Excel calendar dates as local calendar dates. Never convert them with toISOString().
const key=(v:unknown)=>{
  if(v instanceof Date&&!Number.isNaN(v.getTime())) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`
  if(typeof v==='number'&&Number.isFinite(v)){
    const epoch=new Date(Date.UTC(1899,11,30)+v*86400000)
    return `${epoch.getUTCFullYear()}-${String(epoch.getUTCMonth()+1).padStart(2,'0')}-${String(epoch.getUTCDate()).padStart(2,'0')}`
  }
  const s=safe(v).trim(); let m=s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`
  m=s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/)
  if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`
  const d=new Date(s); if(Number.isNaN(d.getTime()))return ''
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
}
const label=(d:string)=>d?new Intl.DateTimeFormat('en-IN',{day:'2-digit',month:'short',year:'numeric'}).format(new Date(`${d}T00:00:00`)):'—'
const read=(wb:XLSX.WorkBook,name:string):Row[]=>wb.Sheets[name]?XLSX.utils.sheet_to_json<Row>(wb.Sheets[name],{defval:null,raw:true}):[]

function extractEci(wb:XLSX.WorkBook){
  let rows=read(wb,'sirNoticeGenerate')
  if(!rows.length)rows=read(wb,'ECI_INPUT')
  if(!rows.length)for(const s of wb.SheetNames){const r=read(wb,s);if(r[0]&&('Part No' in r[0]||'POLLING STATION' in r[0])){rows=r;break}}
  return rows.filter(r=>num(r['AC Number'])===34||safe(r['Asmbly Name']).toUpperCase()==='MATIALA')
}

// IMPORTANT: the master workbook supplies only fixed PS/master information.
// Hearing dates must come from the authoritative Part Wise Hearing Summary upload.
function extractMaster(wb:XLSX.WorkBook):Master{
  return {ps:read(wb,'PS_MASTER'),schedule:[]}
}

function extractSchedule(wb:XLSX.WorkBook):ScheduleRow[]{
  let rows=read(wb,'Part Wise Hearing Summary')
  if(!rows.length)rows=read(wb,'Part Wise Hearing Dates')
  if(!rows.length)rows=read(wb,'HEARING_DATA')
  return rows
    .map(r=>({dateKey:key(r['Hearing Date']??r.Date),ps:num(r['Part No.']??r['Part No']??r['PS No.']),scheduled:num(r['Total Hearings']??r['Scheduled Notices for Hearing'])}))
    .filter(r=>r.ps>0&&!!r.dateKey)
}

export default function Home(){
  const[eci,setEci]=useState<Row[]>([]),[master,setMaster]=useState<Master|null>(null),[date,setDate]=useState(''),[q,setQ]=useState(''),[tab,setTab]=useState<'overview'|'officers'|'ps'|'eci'>('overview'),[error,setError]=useState(''),[source,setSource]=useState(''),[ready,setReady]=useState(false)

  // v4 intentionally invalidates the old cached schedule. This prevents an old
  // 380-row HEARING_DATA schedule from surviving after the source was corrected.
  useEffect(()=>{
    try{
      const raw=localStorage.getItem('ac34_master_v4')
      if(raw){
        const m=JSON.parse(raw) as Master
        setMaster(m)
        const ds=(m.schedule||[]).map(x=>x.dateKey).filter(Boolean).sort()
        if(ds.length)setDate(ds.at(-1)||'')
      }
    }catch{}
    setReady(true)
  },[])

  const saveMaster=(m:Master,labelText:string)=>{
    setMaster(m)
    localStorage.setItem('ac34_master_v4',JSON.stringify(m))
    const ds=m.schedule.map(x=>x.dateKey).filter(Boolean).sort()
    setDate(ds.at(-1)||'')
    setSource(labelText)
  }

  const uploadDashboard=async(f:File)=>{
    try{
      setError('')
      const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true})
      const m=extractMaster(wb)
      if(!m.ps.length)throw new Error('Dashboard workbook must contain the PS_MASTER sheet.')
      saveMaster(m,`${f.name} — PS master loaded; hearing schedule must be loaded separately`)
    }catch(e){setError(e instanceof Error?e.message:'Unable to read dashboard workbook.')}
  }

  const uploadSchedule=async(f:File)=>{
    try{
      setError('')
      const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true})
      const schedule=extractSchedule(wb)
      if(!schedule.length)throw new Error('No valid hearing schedule found. Expected Part Wise Hearing Summary with Part No., Hearing Date and Total Hearings.')
      const unique=[...new Map(schedule.map(r=>[`${r.ps}|${r.dateKey}`,r])).values()]
      if(unique.length<schedule.length)setError(`Loaded ${fmt(unique.length)} unique part/date records; duplicate rows were ignored.`)
      saveMaster(master?{...master,schedule:unique}:{ps:[],schedule:unique},`${f.name} — authoritative hearing schedule loaded`)
    }catch(e){setError(e instanceof Error?e.message:'Unable to read hearing schedule.')}
  }

  const uploadEci=async(f:File)=>{
    try{
      setError('')
      const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true})
      const rows=extractEci(wb)
      if(!rows.length)throw new Error('No AC-34 / MATIALA rows found. This does not look like the ECI NOTICE_REPORT_PART_WISE file.')
      setEci(rows);setSource(f.name)
      const ds=(master?.schedule||[]).map(x=>x.dateKey).filter(Boolean).sort()
      if(ds.length)setDate(ds.at(-1)||'')
    }catch(e){setError(e instanceof Error?e.message:'Unable to read ECI file.')}
  }

  const dates=useMemo(()=>[...new Set((master?.schedule||[]).map(x=>x.dateKey).filter(Boolean))].sort(),[master])
  const eciMap=useMemo(()=>new Map(eci.map(r=>[num(r['Part No']??r['POLLING STATION']),r])),[eci])
  const psMap=useMemo(()=>new Map((master?.ps||[]).map(r=>[num(r['PS No.']),r])),[master])
  const active=useMemo(()=> (master?.schedule||[]).filter(r=>r.dateKey===date).map(s=>({...psMap.get(s.ps),...eciMap.get(s.ps),'PS No.':s.ps,'Scheduled Notices for Hearing':s.scheduled,Date:s.dateKey})),[master,date,eciMap,psMap])
  const filtered=useMemo(()=>{const x=q.trim().toLowerCase();return active.filter(r=>!x||[r['PS No.'],r['Old PS No.'],r.Officer,r['Hearing Centre'],r.BLO,r.Supervisor,r.Locality,r['Polling Area']].some(v=>safe(v).toLowerCase().includes(x)))},[active,q])
  const total=useMemo(()=>active.reduce((a,r)=>({ps:a.ps+1,scheduled:a.scheduled+num(r['Scheduled Notices for Hearing']),generated:a.generated+num(r['Notice Generated']),delivered:a.delivered+num(r['Notice Delivered']),pending:a.pending+num(r['Notice Pending Delivery']),held:a.held+num(r['Hearings Held'])}),{ps:0,scheduled:0,generated:0,delivered:0,pending:0,held:0}),[active])
  const officers=useMemo(()=>{const m=new Map<string,Row>();active.forEach(r=>{const k=`${safe(r.Officer)}|${safe(r['Hearing Centre'])}`,x=m.get(k)||{Officer:r.Officer,Mobile:r['Officer Mobile'],Centre:r['Hearing Centre'],PS:0,Scheduled:0,Generated:0,Delivered:0,Pending:0,Held:0};x.PS++;x.Scheduled+=num(r['Scheduled Notices for Hearing']);x.Generated+=num(r['Notice Generated']);x.Delivered+=num(r['Notice Delivered']);x.Pending+=num(r['Notice Pending Delivery']);x.Held+=num(r['Hearings Held']);m.set(k,x)});return[...m.values()]},[active])
  const global=useMemo(()=>eci.reduce((a,r)=>({generated:a.generated+num(r['Notice Generated']),delivered:a.delivered+num(r['Notice Delivered']),pending:a.pending+num(r['Notice Pending Delivery']),held:a.held+num(r['Hearings Held']),hearingLapsed:a.hearingLapsed+num(r['Hearing Date Lapsed']),rescheduleLapsed:a.rescheduleLapsed+num(r['Reschedule Date Lapsed']),verified:a.verified+num(r['DEO-Status Verified']),notVerified:a.notVerified+num(r['DEO-Status Not Verified'])}),{generated:0,delivered:0,pending:0,held:0,hearingLapsed:0,rescheduleLapsed:0,verified:0,notVerified:0}),[eci])
  const chart=officers.map(o=>({name:safe(o.Officer).replace(/^SH\. |^SMT\. /,''),Scheduled:o.Scheduled,Delivered:o.Delivered,Pending:o.Pending}))

  if(!ready)return null
  return <div className="app-shell">
    <header className="topbar"><div className="brand"><div><div className="title">AC-34 MATIALA — SIR 2026 HEARING & NOTICE DASHBOARD</div><div className="subtitle">Upload the latest ECI NOTICE_REPORT_PART_WISE file — no manual Excel formula work required.</div></div><div className="updated">{source&&<>Source: {source}<br/></>}{eci.length?`${fmt(eci.length)} ECI rows loaded`:master?`${fmt(master.ps.length)} PS master records`:''}</div></div></header>
    <main className="page">
      <div className="toolbar">
        <div className="field"><label>Latest ECI Report — normal daily upload</label><label className="upload"><Upload size={16}/> Upload ECI Excel<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&uploadEci(e.target.files[0])}/></label></div>
        <div className="field"><label>Hearing Date</label><select disabled={!dates.length} value={date} onChange={e=>setDate(e.target.value)}><option value="">Select hearing date</option>{dates.map(d=><option key={d} value={d}>{label(d)}</option>)}</select></div>
        <div className="field"><label>Search</label><div style={{position:'relative'}}><Search size={16} style={{position:'absolute',left:11,top:12,color:'#829ab1'}}/><input className="upload" style={{paddingLeft:34,width:'100%'}} value={q} onChange={e=>setQ(e.target.value)} placeholder="PS / officer / centre / BLO..."/></div></div>
        <button className="tab" disabled={!dates.length} onClick={()=>setDate(dates.at(-1)||'')}><RefreshCcw size={14}/> Latest date</button>
      </div>

      {!master&&<div className="panel setup"><Settings2 size={28}/><h2>One-time setup</h2><p>Upload your existing AC34_Matiala_Hearing_Dashboard workbook once for the fixed 430-PS master. Then upload the authoritative <b>Part Wise Hearing Summary</b>. After that, your normal workflow is only: <b>upload the latest ECI sheet</b>.</p><label className="upload"><Upload size={16}/> Upload dashboard workbook<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&uploadDashboard(e.target.files[0])}/></label></div>}
      {master&&<div className="panel setup compact"><b>Hearing schedule:</b> {fmt(master.schedule.length)} authoritative part/date records loaded. <label className="upload"><Upload size={14}/> Update from Part Wise Hearing Summary<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&uploadSchedule(e.target.files[0])}/></label></div>}
      {error&&<div className="error"><AlertTriangle size={15}/> {error}</div>}

      {master&&master.schedule.length===0&&<div className="panel empty"><FileSpreadsheet size={40}/><h2>Load the hearing schedule</h2><p>The old HEARING_DATA sheet is deliberately <b>not</b> used for hearing dates. Upload the current Part Wise Hearing Summary so the dashboard uses the authoritative hearing allocation.</p><label className="upload"><Upload size={16}/> Upload Part Wise Hearing Summary<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&uploadSchedule(e.target.files[0])}/></label></div>}

      {master&&eci.length>0&&dates.length>0&&<><div className="section-title">Selected Hearing Date — {label(date)}</div><div className="cards"><div className="card"><div className="kicker">Polling Stations</div><div className="metric">{fmt(total.ps)}</div><div className="delta">scheduled</div></div><div className="card"><div className="kicker">Scheduled Hearings</div><div className="metric">{fmt(total.scheduled)}</div><div className="delta">from authoritative part-wise schedule</div></div><div className="card"><div className="kicker">Notice Generated</div><div className="metric">{fmt(total.generated)}</div></div><div className="card"><div className="kicker">Delivered</div><div className="metric">{fmt(total.delivered)}</div><div className="delta">{total.scheduled?Math.round(total.delivered/total.scheduled*100):0}% of scheduled</div></div><div className="card"><div className="kicker">Pending Delivery</div><div className="metric">{fmt(total.pending)}</div></div><div className="card"><div className="kicker">Hearings Held</div><div className="metric">{fmt(total.held)}</div></div></div>
      <div className="tabs"><button className={`tab ${tab==='overview'?'active':''}`} onClick={()=>setTab('overview')}>Overview</button><button className={`tab ${tab==='officers'?'active':''}`} onClick={()=>setTab('officers')}>Officer-wise</button><button className={`tab ${tab==='ps'?'active':''}`} onClick={()=>setTab('ps')}>PS-wise Detail</button><button className={`tab ${tab==='eci'?'active':''}`} onClick={()=>setTab('eci')}>ECI Status</button></div>
      {tab==='overview'&&<><div className="grid2"><div className="panel"><h3>Officer workload</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={chart} margin={{top:8,right:15,left:-20,bottom:30}}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="name" angle={-25} textAnchor="end" interval={0} height={60} fontSize={10}/><YAxis fontSize={10}/><Tooltip/><Bar dataKey="Scheduled"/><Bar dataKey="Delivered"/><Bar dataKey="Pending"/></BarChart></ResponsiveContainer></div></div><div className="panel"><h3>Notice delivery status</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={[{name:'Delivered',value:total.delivered},{name:'Pending',value:total.pending}]} innerRadius={62} outerRadius={92} dataKey="value" label>{[0,1].map(i=><Cell key={i}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div></div></div><div className="panel" style={{marginTop:14}}><h3>Latest ECI operational snapshot</h3><div className="details"><div className="detail"><b>Notice generated</b>{fmt(global.generated)}</div><div className="detail"><b>Notice delivered</b>{fmt(global.delivered)}</div><div className="detail"><b>Pending delivery</b>{fmt(global.pending)}</div><div className="detail"><b>Hearings held</b>{fmt(global.held)}</div><div className="detail"><b>Hearing-date lapsed</b>{fmt(global.hearingLapsed)}</div><div className="detail"><b>Reschedule lapsed</b>{fmt(global.rescheduleLapsed)}</div><div className="detail"><b>DEO verified</b>{fmt(global.verified)}</div><div className="detail"><b>DEO not verified</b>{fmt(global.notVerified)}</div></div></div></>}
      {tab==='officers'&&<div className="panel"><h3>Officer-wise Hearing Summary</h3><div className="table-wrap"><table className="table"><thead><tr><th>Officer</th><th>Mobile</th><th>Hearing Centre</th><th>PS</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Delivery %</th></tr></thead><tbody>{officers.map((o,i)=><tr key={i}><td><b>{safe(o.Officer)}</b></td><td>{safe(o.Mobile)}</td><td>{safe(o.Centre)}</td><td>{o.PS}</td><td>{fmt(o.Scheduled)}</td><td>{fmt(o.Generated)}</td><td>{fmt(o.Delivered)}</td><td>{fmt(o.Pending)}</td><td>{fmt(o.Held)}</td><td>{o.Scheduled?Math.round(o.Delivered/o.Scheduled*100)+'%':'—'}</td></tr>)}</tbody></table></div></div>}
      {tab==='ps'&&<div className="panel"><h3>PS-wise Detail — {fmt(filtered.length)} records</h3><div className="table-wrap"><table className="table"><thead><tr><th>PS</th><th>Old PS</th><th>Officer</th><th>Hearing Centre</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Locality</th><th>Polling Area</th><th>Anomaly</th><th>Mapping</th><th>Grand Total</th><th>Voters</th></tr></thead><tbody>{filtered.map((r,i)=><tr key={i}><td><b>{safe(r['PS No.'])}</b></td><td>{safe(r['Old PS No.'])}</td><td>{safe(r.Officer)}</td><td>{safe(r['Hearing Centre'])}</td><td>{safe(r.BLO)}</td><td>{safe(r.Supervisor)}</td><td>{fmt(r['Scheduled Notices for Hearing'])}</td><td>{fmt(r['Notice Generated'])}</td><td>{fmt(r['Notice Delivered'])}</td><td>{fmt(r['Notice Pending Delivery'])}</td><td>{fmt(r['Hearings Held'])}</td><td>{safe(r.Locality)}</td><td>{safe(r['Polling Area'])}</td><td>{fmt(r['Total Anomaly/Discrepancy'])}</td><td>{fmt(r['Total No Mapping'])}</td><td>{fmt(r['Grand Total'])}</td><td>{fmt(r['Total Voters'])}</td></tr>)}</tbody></table></div></div>}
      {tab==='eci'&&<div className="panel"><h3>Raw ECI Status — {fmt(eci.length)} AC-34 rows</h3><div className="table-wrap"><table className="table"><thead><tr><th>Part No</th><th>Notice Generated</th><th>Pending Generation</th><th>Delivered</th><th>Pending Delivery</th><th>Hearings Held</th><th>Hearing Lapsed</th><th>Reschedule Lapsed</th><th>DEO Pending</th><th>DEO &gt;5 Days</th><th>Verified</th><th>Not Verified</th></tr></thead><tbody>{eci.map((r,i)=><tr key={i}><td><b>{safe(r['Part No']??r['POLLING STATION'])}</b></td><td>{fmt(r['Notice Generated'])}</td><td>{fmt(r['Pending for Notice Generation'])}</td><td>{fmt(r['Notice Delivered'])}</td><td>{fmt(r['Notice Pending Delivery'])}</td><td>{fmt(r['Hearings Held'])}</td><td>{fmt(r['Hearing Date Lapsed'])}</td><td>{fmt(r['Reschedule Date Lapsed'])}</td><td>{fmt(r['DEO-Status Total Pending'])}</td><td>{fmt(r['DEO-Status Pending GT 5 Days'])}</td><td>{fmt(r['DEO-Status Verified'])}</td><td>{fmt(r['DEO-Status Not Verified'])}</td></tr>)}</tbody></table></div></div>}
      </>}
      {master&&eci.length===0&&master.schedule.length>0&&<div className="panel empty"><FileSpreadsheet size={40}/><h2>Upload the latest ECI report</h2><p>Your authoritative hearing schedule is loaded. Upload the current ECI NOTICE_REPORT_PART_WISE Excel above to refresh the control position.</p></div>}
    </main>
  </div>
}
