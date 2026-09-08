'use client'

import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts'
import { Upload, RefreshCcw, Search, FileSpreadsheet, AlertTriangle, CheckCircle2 } from 'lucide-react'

const ECI_HEADERS = ['State Id','State Name','District Number','District Name','AC Number','Asmbly Name','POLLING STATION','Notice Generated','Pending for Notice Generation','Notice Delivered','Notice Pending Delivery','Hearings Held','Hearing Date Lapsed','Reschedule Date Lapsed','DEO-Status Total Pending','DEO-Status Pending GT 5 Days','DEO-Status Verified','DEO-Status Not Verified','ERO/AERO Status Found Ineligible For Final w.r.t. Notice Generated','ERO/AERO Status Parked For Final Publication','ERO/AERO Parked For Final Publication w.r.t. Others']
const PS_HEADERS = ['PS No.','Old PS No.','Officer','Officer Mobile','Hearing Centre','PS Address','BLO','BLO Mobile','Supervisor','Supervisor Mobile','Locality','Polling Area','Total Anomaly/Discrepancy','Total No Mapping','Grand Total','Total Voters']
const fmt = (n:number) => new Intl.NumberFormat('en-IN').format(Number(n||0))
const safe = (v:unknown) => v == null ? '' : String(v)
const dateKey = (v:unknown) => {
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0,10)
  const s=safe(v); const d=new Date(s); return isNaN(d.getTime()) ? '' : d.toISOString().slice(0,10)
}
const dateLabel = (key:string) => key ? new Intl.DateTimeFormat('en-IN',{day:'2-digit',month:'short',year:'numeric'}).format(new Date(key+'T00:00:00')) : '—'

type Row = Record<string, any>
type Model = { eci:Row[]; ps:Row[]; dates:string[]; schedule:Row[]; source:string; uploadedAt:string }

function readSheet(wb:XLSX.WorkBook, name:string){
  const ws=wb.Sheets[name]; if(!ws) return []
  return XLSX.utils.sheet_to_json<Row>(ws,{defval:null,raw:true})
}
function normaliseEci(rows:Row[]){
  return rows.filter(r=>Number(r['AC Number'])===34 || safe(r['Asmbly Name']).toUpperCase()==='MATIALA').map(r=>{
    const o={...r}; for(const h of ECI_HEADERS) if(!(h in o)) o[h]=null; return o
  })
}
function buildModel(wb:XLSX.WorkBook, source:string):Model{
  const eci=normaliseEci(readSheet(wb,'ECI_INPUT'))
  const psRaw=readSheet(wb,'PS_MASTER')
  const ps=psRaw.length?psRaw:eci.map(r=>({
    'PS No.':r['POLLING STATION'], 'Old PS No.':null, Officer:'', 'Officer Mobile':'', 'Hearing Centre':'', 'PS Address':'', BLO:'','BLO Mobile':'',Supervisor:'','Supervisor Mobile':'',Locality:'','Polling Area':'', 'Total Anomaly/Discrepancy':Number(r['Pending for Notice Generation']||0), 'Total No Mapping':Number(r['Notice Generated']||0), 'Grand Total':Number(r['Notice Generated']||0)+Number(r['Pending for Notice Generation']||0), 'Total Voters':0
  }))
  const byPs=new Map<number,Row>(); ps.forEach(p=>byPs.set(Number(p['PS No.']),p))
  const schedule=readSheet(wb,'HEARING_DATA').filter(r=>Number((r as any)['PS No.'])>0).map(r=>({...r, dateKey:dateKey(r.Date)}))
  const dates=[...new Set(schedule.map(r=>r.dateKey).filter(Boolean))].sort()
  return {eci,ps,dates,schedule,source,uploadedAt:new Date().toLocaleString('en-IN')}
}

export default function Home(){
  const [model,setModel]=useState<Model|null>(null)
  const [selectedDate,setSelectedDate]=useState('')
  const [query,setQuery]=useState('')
  const [tab,setTab]=useState<'overview'|'ps'|'eci'|'officers'>('overview')
  const [error,setError]=useState('')

  const loadFile=async(file:File)=>{
    try{
      setError(''); const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:'array',cellDates:true})
      if(!wb.Sheets['ECI_INPUT']) throw new Error('ECI_INPUT sheet not found. Upload the ECI workbook exported for AC-34 Matiala.')
      const m=buildModel(wb,file.name); setModel(m); setSelectedDate(m.dates[m.dates.length-1]||'')
    }catch(e){setError(e instanceof Error?e.message:'Unable to read workbook.')}
  }
  const activeRows=useMemo(()=>model?.schedule.filter(r=>r.dateKey===selectedDate)||[],[model,selectedDate])
  const psMap=useMemo(()=>new Map((model?.ps||[]).map(p=>[Number(p['PS No.']),p])),[model])
  const enriched=useMemo(()=>activeRows.map(r=>({...r, ...(psMap.get(Number(r['PS No.']))||{})})),[activeRows,psMap])
  const filtered=useMemo(()=>enriched.filter(r=>{const q=query.trim().toLowerCase(); if(!q)return true; return [r.Officer,r['Hearing Centre'],r['BLO'],r.Supervisor,r['PS No.'],r['Old PS No.'],r.Locality,r['Polling Area']].some(v=>safe(v).toLowerCase().includes(q))}),[enriched,query])
  const officerSummary=useMemo(()=>{
    const map=new Map<string,Row>()
    for(const r of enriched){ const k=`${safe(r.Officer)}|||${safe(r['Hearing Centre'])}`; const x=map.get(k)||{Officer:r.Officer,'Officer Mobile':r['Officer Mobile'],'Hearing Centre':r['Hearing Centre'],'No. of PS':0,Scheduled:0,Generated:0,Delivered:0,Pending:0,Held:0}; x['No. of PS']++; x.Scheduled+=Number(r['Scheduled Notices for Hearing']||0); x.Generated+=Number(r['Notice Generated']||0); x.Delivered+=Number(r['Notice Delivered']||0); x.Pending+=Number(r['Notice Pending Delivery']||0); x.Held+=Number(r['Hearings Held']||0); map.set(k,x) }
    return [...map.values()]
  },[enriched])
  const total=useMemo(()=>enriched.reduce((a,r)=>({ps:a.ps+1,scheduled:a.scheduled+Number(r['Scheduled Notices for Hearing']||0),generated:a.generated+Number(r['Notice Generated']||0),delivered:a.delivered+Number(r['Notice Delivered']||0),pending:a.pending+Number(r['Notice Pending Delivery']||0),held:a.held+Number(r['Hearings Held']||0)}),{ps:0,scheduled:0,generated:0,delivered:0,pending:0,held:0}),[enriched])
  const global=useMemo(()=>model?.eci.reduce((a,r)=>({generated:a.generated+Number(r['Notice Generated']||0),delivered:a.delivered+Number(r['Notice Delivered']||0),pending:a.pending+Number(r['Notice Pending Delivery']||0),held:a.held+Number(r['Hearings Held']||0),lapsed:a.lapsed+Number(r['Hearing Date Lapsed']||0)+Number(r['Reschedule Date Lapsed']||0),verified:a.verified+Number(r['DEO-Status Verified']||0),notVerified:a.notVerified+Number(r['DEO-Status Not Verified']||0)}),{generated:0,delivered:0,pending:0,held:0,lapsed:0,verified:0,notVerified:0}),null)
  ,[model])
  const chart=officerSummary.map(o=>({name:safe(o.Officer).replace(/^SH\\. |^SMT\\. /,''),Scheduled:o.Scheduled,Delivered:o.Delivered,Pending:o.Pending}))
  const pct=total.scheduled?Math.round(total.delivered/total.scheduled*100):0

  return <div className="app-shell">
    <header className="topbar"><div className="brand"><div><div className="title">AC-34 MATIALA — SIR 2026 HEARING & NOTICE DASHBOARD</div><div className="subtitle">Live analytical dashboard powered by the latest ECI NOTICE_REPORT_PART_WISE Excel export</div></div>{model&&<div className="updated">Loaded: {model.uploadedAt}<br/>File: {model.source}</div>}</div></header>
    <main className="page">
      <div className="toolbar">
        <div className="field"><label>ECI Input Workbook</label><label className="upload"><Upload size={16}/> Upload latest Excel<input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&loadFile(e.target.files[0])}/></label></div>
        <div className="field"><label>Hearing Date</label><select disabled={!model} value={selectedDate} onChange={e=>setSelectedDate(e.target.value)}><option value="">Select hearing date</option>{model?.dates.map(d=><option key={d} value={d}>{dateLabel(d)}</option>)}</select></div>
        <div className="field"><label>Search</label><div style={{position:'relative'}}><Search size={16} style={{position:'absolute',left:11,top:12,color:'#829ab1'}}/><input className="upload" style={{paddingLeft:34,width:'100%'}} value={query} onChange={e=>setQuery(e.target.value)} placeholder="PS / officer / centre / BLO..."/></div></div>
        <button className="tab" onClick={()=>{if(model)setSelectedDate(model.dates[model.dates.length-1]||'')}}><RefreshCcw size={14}/> Latest date</button>
        <div className="status">{model?`${fmt(model.eci.length)} ECI rows loaded • ${fmt(model.ps.length)} PS master rows • ${model.dates.length} hearing dates`: 'Upload the ECI Excel to begin'}</div>
      </div>
      {error&&<div className="error"><AlertTriangle size={15} style={{verticalAlign:'-3px',marginRight:6}}/>{error}</div>}
      {!model?<div className="panel" style={{marginTop:16}}><div className="empty"><FileSpreadsheet size={42} style={{marginBottom:10}}/><h2 style={{margin:'4px 0 8px'}}>Upload your latest ECI report</h2><p style={{margin:0}}>Upload the complete workbook containing the <b>ECI_INPUT</b> sheet. The dashboard recalculates the hearing schedule, notice delivery, officer summaries and PS-wise details in the browser.</p></div></div>:<>
      <div className="section-title">Selected Hearing Date — {dateLabel(selectedDate)}</div>
      <div className="cards">
        <div className="card"><div className="kicker">Polling Stations</div><div className="metric">{fmt(total.ps)}</div><div className="delta">scheduled for this date</div></div>
        <div className="card"><div className="kicker">Scheduled Notices</div><div className="metric">{fmt(total.scheduled)}</div><div className="delta">hearing workload</div></div>
        <div className="card"><div className="kicker">Notice Generated</div><div className="metric">{fmt(total.generated)}</div><div className="delta">against selected schedule</div></div>
        <div className="card"><div className="kicker">Delivered</div><div className="metric">{fmt(total.delivered)}</div><div className="delta">{pct}% of scheduled</div></div>
        <div className="card"><div className="kicker">Pending Delivery</div><div className="metric">{fmt(total.pending)}</div><div className="delta">remaining notices</div></div>
        <div className="card"><div className="kicker">Hearings Held</div><div className="metric">{fmt(total.held)}</div><div className="delta">recorded by ECI</div></div>
      </div>
      {total.ps===0&&<div className="error" style={{marginTop:12}}>No hearing is scheduled for the selected date.</div>}
      <div className="tabs"><button className={`tab ${tab==='overview'?'active':''}`} onClick={()=>setTab('overview')}>Overview</button><button className={`tab ${tab==='officers'?'active':''}`} onClick={()=>setTab('officers')}>Officer-wise</button><button className={`tab ${tab==='ps'?'active':''}`} onClick={()=>setTab('ps')}>PS-wise Detail</button><button className={`tab ${tab==='eci'?'active':''}`} onClick={()=>setTab('eci')}>ECI Status</button></div>
      {tab==='overview'&&<>
        <div className="grid2"><div className="panel"><h3>Officer workload for selected date</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={chart} margin={{top:8,right:15,left:-20,bottom:30}}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="name" angle={-25} textAnchor="end" interval={0} height={60} fontSize={10}/><YAxis fontSize={10}/><Tooltip/><Bar dataKey="Scheduled" name="Scheduled" fill="#315a7d"/><Bar dataKey="Delivered" name="Delivered" fill="#2f855a"/><Bar dataKey="Pending" name="Pending" fill="#c0841a"/></BarChart></ResponsiveContainer></div></div>
        <div className="panel"><h3>Notice delivery status</h3><div className="chart"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={[{name:'Delivered',value:total.delivered},{name:'Pending',value:total.pending}]} cx="50%" cy="50%" innerRadius={62} outerRadius={92} dataKey="value" label>{[{fill:'#2f855a'},{fill:'#c0841a'}].map((x,i)=><Cell key={i} fill={x.fill}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></div><div className="small" style={{textAlign:'center'}}>{fmt(total.delivered)} delivered • {fmt(total.pending)} pending</div></div></div>
        <div className="panel" style={{marginTop:14}}><h3>Operational snapshot from latest ECI data</h3><div className="details"><div className="detail"><b>Total notice generated</b>{fmt(global?.generated||0)}</div><div className="detail"><b>Total notice delivered</b>{fmt(global?.delivered||0)}</div><div className="detail"><b>Total pending delivery</b>{fmt(global?.pending||0)}</div><div className="detail"><b>Hearing + reschedule lapses</b>{fmt(global?.lapsed||0)}</div><div className="detail"><b>DEO verified</b>{fmt(global?.verified||0)}</div><div className="detail"><b>DEO not verified</b>{fmt(global?.notVerified||0)}</div><div className="detail"><b>PS master records</b>{fmt(model.ps.length)}</div><div className="detail"><b>Hearing dates</b>{fmt(model.dates.length)}</div></div></div>
      </>}
      {tab==='officers'&&<div className="panel"><h3>Officer-wise Hearing Summary</h3><div className="table-wrap"><table className="table"><thead><tr><th>Officer</th><th>Mobile</th><th>Hearing Centre</th><th>PS</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Delivery %</th></tr></thead><tbody>{officerSummary.map((o,i)=><tr key={i}><td><b>{safe(o.Officer)}</b></td><td>{safe(o['Officer Mobile'])}</td><td>{safe(o['Hearing Centre'])}</td><td>{fmt(o['No. of PS'])}</td><td className="right">{fmt(o.Scheduled)}</td><td className="right">{fmt(o.Generated)}</td><td className="right">{fmt(o.Delivered)}</td><td className="right">{fmt(o.Pending)}</td><td className="right">{fmt(o.Held)}</td><td>{o.Scheduled?<span className={`badge ${o.Delivered/o.Scheduled>.8?'good':o.Delivered/o.Scheduled>.4?'warn':'danger'}`}>{Math.round(o.Delivered/o.Scheduled*100)}%</span>:<span className="badge">—</span>}</td></tr>)}<tr><td><b>TOTAL</b></td><td colSpan={2}></td><td>{fmt(total.ps)}</td><td className="right"><b>{fmt(total.scheduled)}</b></td><td className="right"><b>{fmt(total.generated)}</b></td><td className="right"><b>{fmt(total.delivered)}</b></td><td className="right"><b>{fmt(total.pending)}</b></td><td className="right"><b>{fmt(total.held)}</b></td><td><b>{pct}%</b></td></tr></tbody></table></div></div>}
      {tab==='ps'&&<div className="panel"><h3>PS-wise Details — {dateLabel(selectedDate)}</h3><div className="small" style={{marginBottom:10}}>{filtered.length} records after search filter</div><div className="table-wrap"><table className="table"><thead><tr><th>Officer</th><th>Centre</th><th>PS No.</th><th>Old PS</th><th>BLO</th><th>Supervisor</th><th>Scheduled</th><th>Generated</th><th>Delivered</th><th>Pending</th><th>Held</th><th>Locality / Polling Area</th></tr></thead><tbody>{filtered.map((r,i)=><tr key={i}><td>{safe(r.Officer)}</td><td>{safe(r['Hearing Centre'])}</td><td><b>{safe(r['PS No.'])}</b></td><td>{safe(r['Old PS No.'])}</td><td>{safe(r.BLO)}<br/><span className="small">{safe(r['BLO Mobile'])}</span></td><td>{safe(r.Supervisor)}<br/><span className="small">{safe(r['Supervisor Mobile'])}</span></td><td className="right">{fmt(r['Scheduled Notices for Hearing'])}</td><td className="right">{fmt(r['Notice Generated'])}</td><td className="right">{fmt(r['Notice Delivered'])}</td><td className="right">{fmt(r['Notice Pending Delivery'])}</td><td className="right">{fmt(r['Hearings Held'])}</td><td>{safe(r.Locality)}<br/><span className="small">{safe(r['Polling Area'])}</span></td></tr>)}{!filtered.length&&<tr><td colSpan={12}><div className="empty">No matching PS records.</div></td></tr>}</tbody></table></div></div>}
      {tab==='eci'&&<div className="panel"><h3>ECI Status — Complete AC-34 input</h3><div className="small" style={{marginBottom:10}}>The cards above are based on the selected hearing schedule; this table gives the raw ECI rows and all status fields for auditing.</div><div className="table-wrap"><table className="table"><thead><tr>{ECI_HEADERS.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{model.eci.map((r,i)=><tr key={i}>{ECI_HEADERS.map(h=><td key={h}>{safe(r[h])}</td>)}</tr>)}</tbody></table></div></div>}
      <div className="footer-note"><CheckCircle2 size={13} style={{verticalAlign:'-2px',marginRight:5}}/>Data is calculated in the browser from the uploaded workbook. No ECI source data is required to be committed to GitHub.</div>
      </>}
    </main>
  </div>
}
