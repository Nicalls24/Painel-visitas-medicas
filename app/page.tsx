'use client'

import React, { useMemo, useState, useCallback, useEffect, useRef } from 'react'
import * as XLSX from 'xlsx'

// ─── Supabase ─────────────────────────────────────────────
const SB_URL = 'https://tdoubfwlpfcyxqfqyppb.supabase.co'
const SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRkb3ViZndscGZjeXhxZnF5cHBiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg1MjI1MDEsImV4cCI6MjA5NDA5ODUwMX0.stw5t_YY_IqyIEQv7sabN8lfpo5PF-9sKgHvYkRJLTI'
const TABLE  = 'visitas_medicas'

const sbFetch = (path: string, opts: RequestInit = {}) => {
  const { headers: extra, ...rest } = opts as any
  return fetch(`${SB_URL}/rest/v1/${path}`, {
    ...rest,
    headers: {
      'apikey':        SB_KEY,
      'Authorization': `Bearer ${SB_KEY}`,
      'Content-Type':  'application/json',
      ...extra,
    },
  })
}

// ─── Parser CSV ───────────────────────────────────────────
const parseCSV = (text: string, delimiter = ';'): string[][] =>
  text.split('\n')
    .map(line => line.split(delimiter).map(cell => cell.trim().replace(/^"|"$/g, '')))
    .filter(row => row.some(cell => cell !== ''))

// ─── Palette ──────────────────────────────────────────────
const C = {
  bg:      '#06080F',
  surface: '#0B1120',
  card:    '#0F1828',
  card2:   '#0D1525',
  border:  '#172438',
  accent:  '#00C6FF',
  accentB: '#0072FF',
  green:   '#10B981',
  yellow:  '#F59E0B',
  red:     '#EF4444',
  orange:  '#F97316',
  text:    '#EDF2FF',
  muted:   '#4A6A88',
  sub:     '#7FA8C4',
}
const META = 90

// ─── Helpers ──────────────────────────────────────────────
const extractDateFromName = (name: string): string | null => {
  // Formato: _DD_MM. ou _DD.MM. ou DD/MM ou qualquer variação
  const m1 = name.match(/_(\d{2})[_.](\d{2})[_.]/)
  if (m1) return `${new Date().getFullYear()}-${m1[2]}-${m1[1]}`
  const m2 = name.match(/\s(\d{2})\.(\d{2})\./)
  if (m2) return `${new Date().getFullYear()}-${m2[2]}-${m2[1]}`
  const m3 = name.match(/(\d{2})_(\d{2})/)
  if (m3) return `${new Date().getFullYear()}-${m3[2]}-${m3[1]}`
  // Tenta pegar data no formato YYYY-MM-DD ou DD-MM-YYYY
  const m4 = name.match(/(\d{4})-(\d{2})-(\d{2})/)
  if (m4) return `${m4[1]}-${m4[2]}-${m4[3]}`
  return null
}

const extractHourFromName = (name: string): number | null => {
  const m = name.match(/(\d{1,2})[Hh][\s_.\-]/i) || name.match(/[_\s](\d{1,2})[Hh]/i)
  return m ? parseInt(m[1]) : null
}

const isPrevistas = (name: string): boolean => {
  const h = extractHourFromName(name)
  // Arquivo de manhã (< 12h) = previstas; tarde = realizadas/pendentes
  if (h !== null) return h < 12
  // Fallback: nome contém "prev" ou "07"
  const low = name.toLowerCase()
  return low.includes('prev') || low.includes('07') || low.includes('agend')
}

const parseFile = async (file: File): Promise<Record<string, string>[]> => {
  const name = file.name
  const isXlsx = /\.(xlsx|xls)$/i.test(name)

  let raw: string[][] = []
  if (isXlsx) {
    const buf = await file.arrayBuffer()
    const wb  = XLSX.read(buf, { type: 'buffer' })
    const ws  = wb.Sheets[wb.SheetNames[0]]
    raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as string[][]
  } else {
    const text = await file.text()
    raw = parseCSV(text, ';')
  }

  // Encontra linha do cabeçalho
  const headerIdx = raw.findIndex(row =>
    row.some(cell => {
      const u = String(cell).trim().toUpperCase()
      return u === 'UF' || u === 'UNIDADE' || u === 'HOSPITAL' || u === 'NM_LOCAL'
    })
  )
  if (headerIdx < 0) return []

  const header   = raw[headerIdx].map(h => String(h).trim().toUpperCase())
  const dataRows = raw.slice(headerIdx + 1)

  return dataRows
    .map(r => {
      const obj: Record<string, string> = {}
      header.forEach((h, i) => { obj[h] = String(r[i] ?? '').trim() })
      return obj
    })
    .filter(r => {
      const unid = r['UNIDADE'] || r['NM_LOCAL'] || r['HOSPITAL'] || ''
      return unid && unid.toUpperCase() !== 'UNIDADE' && unid.toUpperCase() !== 'NM_LOCAL'
    })
    .map(r => {
      // Normaliza campos
      if (!r['UNIDADE'] && r['NM_LOCAL']) r['UNIDADE'] = r['NM_LOCAL']
      if (!r['UNIDADE'] && r['HOSPITAL']) r['UNIDADE'] = r['HOSPITAL']
      return r
    })
}

const localDateStr = (offset = 0): string => {
  const d = new Date(); d.setDate(d.getDate() + offset)
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
}
const fmtDate = (s: string | null): string => s ? s.split('-').reverse().join('/') : ''

// ─── Período ──────────────────────────────────────────────
const PERIODOS = [
  { key: 'TODOS',  label: 'Todos'  },
  { key: 'HOJE',   label: 'Hoje'   },
  { key: 'ONTEM',  label: 'Ontem'  },
  { key: 'SEMANA', label: 'Semana' },
  { key: 'MÊS',    label: 'Mês'    },
]

const filterByPeriodo = (dateStr: string, período: string): boolean => {
  if (período === 'TODOS' || !dateStr) return true
  const today = localDateStr(0), ontem = localDateStr(-1)
  if (período === 'HOJE')  return dateStr === today
  if (período === 'ONTEM') return dateStr === ontem
  if (período === 'SEMANA') {
    const d = new Date(dateStr + 'T00:00:00')
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 6); cutoff.setHours(0,0,0,0)
    return d >= cutoff
  }
  if (período === 'MÊS') return dateStr.slice(0,7) === today.slice(0,7)
  return true
}

// ─── SLA helpers ──────────────────────────────────────────
const slaColor = (pct: number) => pct >= META ? C.green : pct >= 70 ? C.yellow : C.red
const slaLabel = (pct: number) => pct >= META ? 'OK' : pct >= 70 ? 'Em Risco' : 'Crítico'

// ─── Types ────────────────────────────────────────────────
interface DbRow {
  id?: number
  data: string          // YYYY-MM-DD (date)
  data_ref?: string     // mesmo valor de data
  unidade: string
  uf: string
  previstas: number
  realizadas: number
  pendentes: number
  status?: string
}

interface UnidadeRow {
  uf: string; unidade: string
  previstas: number; pendentes: number; realizadas: number; sla: number
}

// ─── Sub-components ───────────────────────────────────────
function Card({ children, style = {} }: { children: React.ReactNode; style?: React.CSSProperties }) {
  return <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 18, ...style }}>{children}</div>
}

function KpiCard({ icon, label, value, sub, accent, note }: {
  icon: string; label: string; value: string | number; sub?: string; accent: string; note?: string
}) {
  return (
    <div style={{ background: C.card2, border: `1px solid ${C.border}`, borderRadius: 14,
      padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 6, position: 'relative', overflow: 'hidden' }}>
      <div style={{ position: 'absolute', top: -16, right: -16, width: 70, height: 70,
        borderRadius: '50%', background: accent, opacity: .08 }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <span style={{ fontSize: 16 }}>{icon}</span>
        <span style={{ fontSize: 9.5, color: C.muted, textTransform: 'uppercase', letterSpacing: '.1em', fontWeight: 700 }}>{label}</span>
      </div>
      <div style={{ fontSize: 32, fontWeight: 900, color: accent, lineHeight: 1 }}>{value}</div>
      {sub  && <div style={{ fontSize: 11, color: C.sub }}>{sub}</div>}
      {note && <div style={{ fontSize: 10, color: C.muted, fontStyle: 'italic', borderTop: `1px solid ${C.border}`, paddingTop: 5, marginTop: 2 }}>{note}</div>}
    </div>
  )
}

function Donut({ pct, size = 130 }: { pct: number; size?: number }) {
  const r = 46, cx = 60, cy = 60, circ = 2 * Math.PI * r
  const dash = (pct / 100) * circ, color = slaColor(pct)
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={C.border} strokeWidth="14" />
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={color} strokeWidth="14"
        strokeDasharray={`${dash} ${circ - dash}`} strokeDashoffset={circ / 4} strokeLinecap="round" />
      <text x="60" y="54" textAnchor="middle" fontSize="20" fontWeight="900" fill={color}>{pct.toFixed(1)}%</text>
      <text x="60" y="70" textAnchor="middle" fontSize="9" fill={C.muted}>SLA</text>
    </svg>
  )
}

function ProgressBar({ pct }: { pct: number }) {
  const color = slaColor(pct)
  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, fontSize: 11, color: C.muted }}>
        <span>0%</span><span style={{ color: C.accent }}>← Meta {META}%</span><span>100%</span>
      </div>
      <div style={{ position: 'relative', height: 14, background: C.border, borderRadius: 99, overflow: 'hidden' }}>
        <div style={{ height: '100%', borderRadius: 99, background: color, width: `${Math.min(pct,100)}%`, transition: 'width .8s ease' }} />
      </div>
      <div style={{ position: 'relative', height: 0 }}>
        <div style={{ position: 'absolute', left: `${META}%`, top: -14, width: 2, height: 14, background: C.accent, opacity: .8 }} />
      </div>
    </div>
  )
}

function UnidadeTable({ rows, showUF }: { rows: UnidadeRow[]; showUF: boolean }) {
  return (
    <div style={{ overflowY: 'auto', maxHeight: 460 }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead style={{ position: 'sticky', top: 0, background: C.card }}>
          <tr>
            {['#', showUF && 'UF', 'Unidade', 'Previstas', 'Realizadas', 'Pendentes', 'SLA', 'Status'].filter(Boolean).map(h => (
              <th key={String(h)} style={{ padding: '8px 10px', textAlign: 'left', borderBottom: `1px solid ${C.border}`,
                color: C.muted, fontWeight: 700, fontSize: 10, textTransform: 'uppercase', letterSpacing: '.07em', whiteSpace: 'nowrap' }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const color = slaColor(r.sla)
            return (
              <tr key={`${r.unidade}${i}`} style={{ borderBottom: `1px solid ${C.border}`, transition: 'background .1s' }}
                onMouseEnter={e => (e.currentTarget.style.background = '#0e1b2c')}
                onMouseLeave={e => (e.currentTarget.style.background = 'transparent')}>
                <td style={{ padding: '7px 10px', color: C.muted, fontSize: 10 }}>{i+1}</td>
                {showUF && <td style={{ padding: '7px 10px', color: C.sub, fontWeight: 700, fontSize: 11 }}>{r.uf}</td>}
                <td style={{ padding: '7px 10px', fontWeight: 500, color: C.text, maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.unidade}</td>
                <td style={{ padding: '7px 10px', color: C.sub, textAlign: 'center' }}>{r.previstas}</td>
                <td style={{ padding: '7px 10px', color: C.green, fontWeight: 700, textAlign: 'center' }}>{r.realizadas}</td>
                <td style={{ padding: '7px 10px', color: C.red, textAlign: 'center' }}>{r.pendentes}</td>
                <td style={{ padding: '7px 10px', fontWeight: 900, color, textAlign: 'center' }}>{r.sla.toFixed(1)}%</td>
                <td style={{ padding: '7px 10px' }}>
                  <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 99, background: color+'22', color, whiteSpace: 'nowrap' }}>
                    {slaLabel(r.sla)}
                  </span>
                </td>
              </tr>
            )
          })}
          {!rows.length && (
            <tr><td colSpan={8} style={{ padding: 16, color: C.muted, fontSize: 13 }}>Nenhuma unidade com dados para este filtro.</td></tr>
          )}
        </tbody>
      </table>
    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
// ─── Tendência Chart ──────────────────────────────────────
function TendenciaChart({ dbRows, filtDates, ufFiltro }: {
  dbRows: DbRow[]; filtDates: string[]; ufFiltro: string
}) {
  const chartRef  = useRef<HTMLCanvasElement>(null)
  const chartInst = useRef<any>(null)
  const [modoVis, setModoVis] = useState<'real'|'proj'>('real')

  const dayData = useMemo(() => {
    return filtDates.map(date => {
      const rows = dbRows.filter(r => r.data === date && (ufFiltro === 'TODOS' || r.uf === ufFiltro))
      const prev = rows.reduce((a, r) => a + r.previstas,  0)
      const real = rows.reduce((a, r) => a + r.realizadas, 0)
      const pend = rows.reduce((a, r) => a + r.pendentes,  0)
      const sla  = prev > 0 ? parseFloat(((real / prev) * 100).toFixed(1)) : 0
      return { date, sla, prev, real, pend }
    })
  }, [dbRows, filtDates, ufFiltro])

  const labels   = dayData.map(d => d.date.slice(5).split('-').reverse().join('/'))
  const slaVals  = dayData.map(d => d.sla)
  const pendVals = dayData.map(d => d.pend)
  const lastSla  = slaVals[slaVals.length - 1] ?? 0
  const firstSla = slaVals[0] ?? 0
  const trend    = parseFloat((lastSla - firstSla).toFixed(1))
  const trendColor = trend >= 0 ? '#10B981' : '#EF4444'

  // Calcula projeção: média móvel 3 dias + inclinação suavizada
  const projData = useMemo(() => {
    if (slaVals.length < 2) return []
    const slope = slaVals.length >= 3
      ? (slaVals[slaVals.length-1] - slaVals[slaVals.length-3]) / 2
      : (slaVals[slaVals.length-1] - slaVals[0])
    const result: (number|null)[] = [...Array(slaVals.length - 1).fill(null), lastSla]
    for (let i = 1; i <= 5; i++) {
      const last3 = [...slaVals, ...result.filter(v=>v!==null&&result.indexOf(v)>=slaVals.length-1) as number[]].slice(-3)
      const avg = last3.reduce((a,b)=>a+b,0)/last3.length
      result.push(parseFloat(Math.min(99, avg + slope * 0.5).toFixed(1)))
    }
    return result
  }, [slaVals, lastSla])

  const pendProjData = useMemo(() => {
    if (pendVals.length < 2) return []
    const slope = pendVals.length >= 3
      ? (pendVals[pendVals.length-1] - pendVals[pendVals.length-3]) / 2
      : (pendVals[pendVals.length-1] - pendVals[0])
    const result: (number|null)[] = [...Array(pendVals.length - 1).fill(null), pendVals[pendVals.length-1]]
    for (let i = 1; i <= 5; i++) {
      const prev = result.filter(v=>v!==null).slice(-1)[0] as number
      result.push(parseFloat(Math.max(0, prev + slope * 0.5).toFixed(0)))
    }
    return result
  }, [pendVals])

  // Gera labels para projeção (+5 dias)
  const projLabels = useMemo(() => {
    const last = filtDates[filtDates.length - 1]
    if (!last) return labels
    const extra = []
    for (let i = 1; i <= 5; i++) {
      const d = new Date(last + 'T00:00:00')
      d.setDate(d.getDate() + i)
      extra.push(`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`)
    }
    return [...labels, ...extra]
  }, [labels, filtDates])

  const projSlaAmanha  = projData[slaVals.length]  as number ?? lastSla
  const projSla5dias   = projData[projData.length-1] as number ?? lastSla
  const diasParaMeta   = lastSla < META && trend > 0
    ? Math.ceil((META - lastSla) / (trend / Math.max(filtDates.length-1,1)))
    : null

  useEffect(() => {
    if (!chartRef.current) return
    if (chartInst.current) { chartInst.current.destroy(); chartInst.current = null }
    const ctx = chartRef.current.getContext('2d')
    if (!ctx) return

    if (modoVis === 'real') {
      const g1 = ctx.createLinearGradient(0,0,0,280)
      g1.addColorStop(0,'#10B98155'); g1.addColorStop(1,'#10B98100')
      const g2 = ctx.createLinearGradient(0,0,0,280)
      g2.addColorStop(0,'#EF444433'); g2.addColorStop(1,'#EF444400')

      chartInst.current = new (window as any).Chart(ctx, {
        type: 'line',
        data: {
          labels,
          datasets: [
            { label:'SLA %',    data:slaVals,  borderColor:'#10B981', borderWidth:2.5, backgroundColor:g1, fill:true, tension:0.4, pointRadius:4, pointBackgroundColor:'#10B981', pointBorderColor:'#0F1828', pointBorderWidth:2, yAxisID:'y' },
            { label:'Meta 90%', data:Array(labels.length).fill(90), borderColor:'#00C6FF55', borderWidth:1.5, borderDash:[5,5], pointRadius:0, fill:false, yAxisID:'y' },
            { label:'Pendentes',data:pendVals, borderColor:'#EF4444', borderWidth:2, backgroundColor:g2, fill:true, tension:0.4, pointRadius:3, pointBackgroundColor:'#EF4444', pointBorderColor:'#0F1828', pointBorderWidth:2, yAxisID:'y2' },
          ]
        },
        options: {
          responsive:true, maintainAspectRatio:false,
          interaction:{mode:'index',intersect:false},
          plugins:{legend:{display:false},tooltip:{backgroundColor:'#0F1828',borderColor:'#172438',borderWidth:1,titleColor:'#7FA8C4',bodyColor:'#EDF2FF',callbacks:{label:(c:any)=>{
            if(c.dataset.label==='Meta 90%') return null
            if(c.dataset.label==='SLA %') return ` SLA: ${c.parsed.y.toFixed(1)}%`
            return ` Pendentes: ${c.parsed.y}`
          }}}},
          scales:{
            x:{ticks:{color:'#4A6A88',font:{size:11},maxRotation:45,autoSkip:false},grid:{color:'#172438'}},
            y:{position:'left',min:50,max:100,ticks:{color:'#4A6A88',font:{size:11},callback:(v:any)=>v+'%'},grid:{color:'#172438'}},
            y2:{position:'right',ticks:{color:'#EF444488',font:{size:11}},grid:{drawOnChartArea:false}}
          }
        }
      })
    } else {
      const realNull  = [...slaVals,  ...Array(5).fill(null)]
      const pendNull  = [...pendVals, ...Array(5).fill(null)]
      const g1 = ctx.createLinearGradient(0,0,0,280)
      g1.addColorStop(0,'#10B98133'); g1.addColorStop(1,'#10B98100')
      const g2 = ctx.createLinearGradient(0,0,0,280)
      g2.addColorStop(0,'#00C6FF44'); g2.addColorStop(1,'#00C6FF00')
      const g3 = ctx.createLinearGradient(0,0,0,280)
      g3.addColorStop(0,'#EF444422'); g3.addColorStop(1,'#EF444400')

      chartInst.current = new (window as any).Chart(ctx, {
        type: 'line',
        data: {
          labels: projLabels,
          datasets: [
            { label:'SLA Real',     data:realNull,    borderColor:'#10B98177', borderWidth:2, backgroundColor:g1, fill:true, tension:0.4, pointRadius:3, pointBackgroundColor:'#10B98177', pointBorderColor:'#0F1828', pointBorderWidth:1, yAxisID:'y' },
            { label:'Projeção SLA', data:projData,    borderColor:'#00C6FF',   borderWidth:2.5, borderDash:[7,4], backgroundColor:g2, fill:true, tension:0.4, pointRadius:5, pointStyle:'triangle', pointBackgroundColor:'#00C6FF', pointBorderColor:'#0F1828', pointBorderWidth:2, yAxisID:'y' },
            { label:'Meta 90%',     data:Array(projLabels.length).fill(90), borderColor:'#4A6A8866', borderWidth:1.5, borderDash:[4,4], pointRadius:0, fill:false, yAxisID:'y' },
            { label:'Pend. Real',   data:pendNull,    borderColor:'#EF444466', borderWidth:1.5, backgroundColor:g3, fill:true, tension:0.4, pointRadius:2, yAxisID:'y2' },
            { label:'Pend. Proj',   data:pendProjData,borderColor:'#EF4444',   borderWidth:2, borderDash:[6,4], tension:0.4, pointRadius:4, pointStyle:'triangle', pointBackgroundColor:'#EF4444', fill:false, yAxisID:'y2' },
          ]
        },
        options: {
          responsive:true, maintainAspectRatio:false,
          interaction:{mode:'index',intersect:false},
          plugins:{legend:{display:false},tooltip:{backgroundColor:'#0F1828',borderColor:'#172438',borderWidth:1,titleColor:'#7FA8C4',bodyColor:'#EDF2FF',callbacks:{label:(c:any)=>{
            if(c.dataset.label==='Meta 90%'||c.parsed.y==null) return null
            if(c.dataset.label.includes('SLA')) return ` ${c.dataset.label}: ${c.parsed.y.toFixed(1)}%`
            return ` ${c.dataset.label}: ${c.parsed.y}`
          }}}},
          scales:{
            x:{ticks:{color:'#4A6A88',font:{size:10},maxRotation:45,autoSkip:false},grid:{color:'#172438'}},
            y:{position:'left',min:50,max:100,ticks:{color:'#4A6A88',font:{size:11},callback:(v:any)=>v+'%'},grid:{color:'#172438'}},
            y2:{position:'right',ticks:{color:'#EF444488',font:{size:11}},grid:{drawOnChartArea:false}}
          }
        }
      })
    }
  }, [dayData, modoVis])

  return (
    <div>
      {/* Header com toggle */}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
        <div>
          <div style={{ fontSize:10, fontWeight:700, textTransform:'uppercase', letterSpacing:'.12em', color:'#4A6A88' }}>
            SLA Velocity · {filtDates.length} dia(s) · {ufFiltro === 'TODOS' ? 'Todos os estados' : ufFiltro}
          </div>
        </div>
        <div style={{ display:'flex', gap:0, background:'#0D1525', border:'1px solid #172438', borderRadius:99, padding:4 }}>
          <button onClick={()=>setModoVis('real')} style={{
            fontSize:11, fontWeight:700, padding:'5px 16px', borderRadius:99, border:'none', cursor:'pointer', transition:'all .2s',
            background: modoVis==='real' ? '#10B981' : 'transparent',
            color: modoVis==='real' ? '#000' : '#4A6A88'
          }}>REAL</button>
          <button onClick={()=>setModoVis('proj')} style={{
            fontSize:11, fontWeight:700, padding:'5px 16px', borderRadius:99, border:'none', cursor:'pointer', transition:'all .2s',
            background: modoVis==='proj' ? '#00C6FF' : 'transparent',
            color: modoVis==='proj' ? '#000' : '#4A6A88'
          }}>PROJEÇÃO</button>
        </div>
      </div>

      {/* KPIs */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:10, marginBottom:16 }}>
        {modoVis === 'real' ? [
          { label:'SLA Atual',  value:`${lastSla.toFixed(1)}%`, color: lastSla>=META ? '#10B981':'#F59E0B' },
          { label:'Variação',   value:`${trend>=0?'+':''}${trend.toFixed(1)}pp`, color:trendColor },
          { label:'Melhor Dia', value:`${slaVals.length?Math.max(...slaVals).toFixed(1):0}%`, color:'#10B981' },
          { label:'Pior Dia',   value:`${slaVals.length?Math.min(...slaVals).toFixed(1):0}%`, color:'#EF4444' },
        ] : [
          { label:'Projeção amanhã',   value:`${projSlaAmanha.toFixed(1)}%`,  color:'#00C6FF' },
          { label:`Projeção +5 dias`,  value:`${projSla5dias.toFixed(1)}%`,   color:'#00C6FF' },
          { label:'Tendência diária',  value:`${trend>=0?'+':''}${filtDates.length>1?(trend/(filtDates.length-1)).toFixed(1):0}pp`, color:trendColor },
          { label:'Atinge meta em',    value: diasParaMeta ? `~${diasParaMeta}d` : lastSla>=META ? 'Atingida!' : '—', color: diasParaMeta ? '#F59E0B' : lastSla>=META ? '#10B981':'#EF4444' },
        ]}.map(k => (
          <div key={k.label} style={{ background:'#0D1525', border:`1px solid ${modoVis==='proj'?'#00C6FF22':'#172438'}`, borderRadius:10, padding:'12px 14px' }}>
            <div style={{ fontSize:9, color:'#4A6A88', textTransform:'uppercase', letterSpacing:'.1em', marginBottom:6 }}>{k.label}</div>
            <div style={{ fontSize:22, fontWeight:900, color:k.color }}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Legenda */}
      <div style={{ display:'flex', gap:16, marginBottom:10, flexWrap:'wrap' }}>
        {modoVis === 'real' ? [
          { color:'#10B981', label:'SLA %', dash:false },
          { color:'#00C6FF', label:'Meta 90%', dash:true },
          { color:'#EF4444', label:'Pendentes (eixo direito)', dash:false },
        ] : [
          { color:'#10B98177', label:'SLA real', dash:false },
          { color:'#00C6FF',   label:'Projeção SLA', dash:true },
          { color:'#4A6A88',   label:'Meta 90%', dash:true },
          { color:'#EF4444',   label:'Projeção pendentes (eixo dir.)', dash:true },
        ]}.map(l => (
          <div key={l.label} style={{ display:'flex', alignItems:'center', gap:6 }}>
            <div style={{ width:20, height: l.dash?0:3, borderRadius:99, background:l.dash?'transparent':l.color, border:l.dash?`1.5px dashed ${l.color}`:'none' }} />
            <span style={{ fontSize:11, color:'#7FA8C4' }}>{l.label}</span>
          </div>
        ))}
      </div>

      {/* Gráfico principal */}
      <div style={{ position:'relative', height:280, marginBottom:16 }}>
        <canvas ref={chartRef} role="img" aria-label={`Tendência SLA ${modoVis} — ${filtDates.length} dias`} />
      </div>

      {/* Card de explicação projeção (só no modo PROJEÇÃO) */}
      {modoVis === 'proj' && (
        <div style={{ background:'#0D1525', border:'1px solid #00C6FF33', borderRadius:10, padding:'14px 16px', marginBottom:14 }}>
          <div style={{ fontSize:9, color:'#00C6FF', textTransform:'uppercase', letterSpacing:'.1em', fontWeight:700, marginBottom:10 }}>
            Como a projeção é calculada
          </div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:12 }}>
            {[
              { step:'Passo 1 — Coleta', desc:'Pega os SLAs dos últimos 3 dias reais disponíveis na base.' },
              { step:'Passo 2 — Média móvel', desc:'Calcula a média simples dos 3 valores. Ex: (73%+75%+74%) ÷ 3 = 74%' },
              { step:'Passo 3 — Tendência', desc:'Adiciona a inclinação média diária suavizada (×0.5) ao resultado.' },
            ].map(s => (
              <div key={s.step} style={{ borderLeft:'2px solid #00C6FF33', paddingLeft:10 }}>
                <div style={{ fontSize:11, color:'#7FA8C4', fontWeight:700, marginBottom:4 }}>{s.step}</div>
                <div style={{ fontSize:10, color:'#4A6A88', lineHeight:1.5 }}>{s.desc}</div>
              </div>
            ))}
          </div>
          <div style={{ marginTop:10, paddingTop:10, borderTop:'1px solid #172438', fontSize:10, color:'#4A6A88' }}>
            Fórmula:{' '}
            <span style={{ color:'#00C6FF', fontFamily:'monospace' }}>
              Projeção = média(últ. 3 dias) + inclinação_média × 0.5
            </span>
            {' '}· Quanto mais dias na base, mais precisa fica a projeção.
          </div>
        </div>
      )}

      {/* Rodapé tendência */}
      <div style={{ padding:'10px 14px', background:'#0D1525', borderRadius:10, border:`1px solid ${trendColor}33`, display:'flex', alignItems:'center', gap:8 }}>
        <span style={{ fontSize:16, color:trendColor }}>{trend>=0?'▲':'▼'}</span>
        <span style={{ fontSize:12, color:'#7FA8C4' }}>
          SLA {trend>=0?'subiu':'caiu'}{' '}
          <strong style={{ color:trendColor }}>{Math.abs(trend).toFixed(1)}pp</strong>
          {' '}no período selecionado.
          {lastSla < META && (
            <span style={{ color:'#F59E0B' }}>
              {' '}Faltam <strong>{(META-lastSla).toFixed(1)}pp</strong> para atingir {META}%.
              {modoVis==='proj' && diasParaMeta && ` Projeção: ~${diasParaMeta} dias.`}
            </span>
          )}
          {lastSla >= META && <span style={{ color:'#10B981' }}> Meta atingida!</span>}
        </span>
      </div>
    </div>
  )
}

export default function VisitasPage() {
  const [dbRows,     setDbRows]     = useState<DbRow[]>([])
  const [loading,    setLoading]    = useState(false)
  const [saving,     setSaving]     = useState(false)
  const [status,     setStatus]     = useState('')
  const [período,    setPeriodo]    = useState('MÊS')
  const [ufFiltro,   setUfFiltro]   = useState('TODOS')
  const [abaSel,     setAbaSel]     = useState('unidades')
  const [uploadInfo, setUploadInfo] = useState('')

  // ── Carrega do Supabase ao iniciar ─────────────────────
  useEffect(() => {
    const load = async () => {
      setStatus('Carregando dados…')
      try {
        const res = await sbFetch(`${TABLE}?select=*&order=id.asc`, {
          headers: { 'Range': '0-9999', 'Range-Unit': 'items' }
        })
        if (!res.ok) {
          const txt = await res.text()
          setStatus(`Erro ${res.status}: ${txt.slice(0,100)}`)
          return
        }
        const rows = await res.json()
        if (Array.isArray(rows) && rows.length > 0) {
          setDbRows(rows)
          const dates = [...new Set(rows.map((r: DbRow) => r.data).filter(Boolean))]
          setStatus(`☁ ${dates.length} dia(s) · ${rows.length} registros`)
          setTimeout(() => setStatus(''), 4000)
        } else {
          setStatus('')
        }
      } catch(e: any) {
        setStatus(`Erro: ${e.message}`)
      }
    }
    load()
  }, [])

  // ── Upload de arquivos ─────────────────────────────────
  const handleFiles = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[]
    if (!files.length) return
    setLoading(true)

    try {
      // Agrupa por data
      const byDate: Record<string, { previstas: Record<string,string>[]; pendentes: Record<string,string>[] }> = {}

      for (const file of files) {
        const fname = file.name.split('/').pop() || file.name
        const date  = extractDateFromName(fname)
        if (!date) { console.warn('Data não detectada:', fname); continue }

        const isPrev = isPrevistas(fname)
        const rows   = await parseFile(file)
        console.log(`[Visitas] ${fname} → date=${date} isPrev=${isPrev} rows=${rows.length}`)

        if (!byDate[date]) byDate[date] = { previstas: [], pendentes: [] }
        if (isPrev) byDate[date].previstas = rows
        else        byDate[date].pendentes = rows
      }

      if (!Object.keys(byDate).length) {
        setStatus('⚠ Nenhuma data detectada nos nomes dos arquivos')
        setLoading(false)
        return
      }

      // Converte para DbRows (1 linha por unidade por data por tipo)
      const newRows: DbRow[] = []
      for (const [date, snap] of Object.entries(byDate)) {
        // Conta previstas por unidade
        const prevCount: Record<string, { uf: string; count: number }> = {}
        for (const r of snap.previstas) {
          const unid = r['UNIDADE'] || r['NM_LOCAL'] || ''; if (!unid) continue
          const uf   = r['UF'] || r['UF_UNIDADE'] || ''
          if (!prevCount[unid]) prevCount[unid] = { uf, count: 0 }
          prevCount[unid].count++
        }
        // Conta pendentes por unidade
        const pendCount: Record<string, number> = {}
        for (const r of snap.pendentes) {
          const unid = r['UNIDADE'] || r['NM_LOCAL'] || ''; if (!unid) continue
          pendCount[unid] = (pendCount[unid] || 0) + 1
        }
        // Cria DbRow por unidade
        const allUnidades = new Set([...Object.keys(prevCount), ...Object.keys(pendCount)])
        for (const unid of allUnidades) {
          const prev = prevCount[unid]?.count || 0
          const pend = pendCount[unid] || 0
          const real = Math.max(0, prev - pend)
          newRows.push({
            data:       date,
            data_ref:   date,
            unidade:    unid,
            uf:         prevCount[unid]?.uf || '',
            previstas:  prev,
            realizadas: real,
            pendentes:  pend,
            status:     real >= prev * META / 100 ? 'OK' : real >= prev * 0.7 ? 'Em Risco' : 'Crítico',
          })
        }
      }

      // Salva no Supabase: apaga datas antigas e insere novas
      setSaving(true)
      const newDates = [...new Set(newRows.map(r => r.data))]
      setStatus(`Removendo ${newDates.length} dia(s) antigo(s)…`)

      for (const date of newDates) {
        await sbFetch(`${TABLE}?data=eq.${date}`, { method: 'DELETE' })
      }

      // Insere em lotes de 500
      const BATCH = 500
      const total = Math.ceil(newRows.length / BATCH)
      for (let i = 0; i < total; i++) {
        const slice = newRows.slice(i * BATCH, (i+1) * BATCH)
        setStatus(`Salvando lote ${i+1}/${total}…`)
        const res = await sbFetch(TABLE, {
          method: 'POST',
          headers: { 'Prefer': 'return=minimal' },
          body: JSON.stringify(slice),
        })
        if (!res.ok) {
          const txt = await res.text()
          setStatus(`⚠ Erro lote ${i+1}: ${txt.slice(0,120)}`)
          setSaving(false); setLoading(false)
          return
        }
      }

      // Recarrega tudo do Supabase
      setStatus('Recarregando…')
      const res2 = await sbFetch(`${TABLE}?select=*&order=id.asc`, {
        headers: { 'Range': '0-9999', 'Range-Unit': 'items' }
      })
      if (res2.ok) {
        const all = await res2.json()
        if (Array.isArray(all)) setDbRows(all)
      }

      const datesLoaded = newDates.map(fmtDate).join(', ')
      setUploadInfo(`${files.length} arquivo(s) · datas: ${datesLoaded}`)
      setStatus('✓ Salvo no Supabase')
      setTimeout(() => setStatus(''), 3000)
    } catch(e: any) {
      setStatus(`⚠ Erro: ${e.message}`)
    }
    setSaving(false); setLoading(false)
  }, [])

  // ── Limpar tudo ────────────────────────────────────────
  const clearAll = async () => {
    if (!confirm('Apagar todos os dados de visitas?')) return
    setSaving(true); setStatus('Apagando…')
    await sbFetch(`${TABLE}?id=gt.0`, { method: 'DELETE' })
    setDbRows([]); setUploadInfo(''); setStatus(''); setSaving(false)
  }

  // ── Filtra por período e UF ────────────────────────────
  const filtered = useMemo(() =>
    dbRows.filter(r => filterByPeriodo(r.data, período) && (ufFiltro === 'TODOS' || r.uf === ufFiltro)),
    [dbRows, período, ufFiltro])

  // ── Agrega por unidade ─────────────────────────────────
  const unidadeRows = useMemo((): UnidadeRow[] => {
    const m: Record<string, UnidadeRow> = {}
    for (const r of filtered) {
      if (!m[r.unidade]) m[r.unidade] = { uf: r.uf, unidade: r.unidade, previstas: 0, pendentes: 0, realizadas: 0, sla: 0 }
      m[r.unidade].previstas  += r.previstas
      m[r.unidade].pendentes  += r.pendentes
      m[r.unidade].realizadas += r.realizadas
    }
    return Object.values(m).map(r => ({
      ...r, sla: r.previstas > 0 ? (r.realizadas / r.previstas) * 100 : 0
    })).sort((a, b) => a.sla - b.sla)
  }, [filtered])

  const totalPrevistas  = unidadeRows.reduce((a, r) => a + r.previstas,  0)
  const totalPendentes  = unidadeRows.reduce((a, r) => a + r.pendentes,  0)
  const totalRealizadas = unidadeRows.reduce((a, r) => a + r.realizadas, 0)
  const slaGeral        = totalPrevistas > 0 ? (totalRealizadas / totalPrevistas) * 100 : 0
  const faltamMeta      = Math.max(0, Math.ceil(totalPrevistas * META / 100) - totalRealizadas)

  const criticas = unidadeRows.filter(r => r.sla < 70)
  const emRisco  = unidadeRows.filter(r => r.sla >= 70 && r.sla < META)
  const ok       = unidadeRows.filter(r => r.sla >= META)

  const ufRows = useMemo(() => {
    const m: Record<string, UnidadeRow> = {}
    for (const r of unidadeRows) {
      if (!m[r.uf]) m[r.uf] = { uf: r.uf, unidade: r.uf, previstas: 0, pendentes: 0, realizadas: 0, sla: 0 }
      m[r.uf].previstas  += r.previstas
      m[r.uf].pendentes  += r.pendentes
      m[r.uf].realizadas += r.realizadas
    }
    return Object.values(m).map(r => ({ ...r, sla: r.previstas > 0 ? (r.realizadas/r.previstas)*100 : 0 }))
      .sort((a, b) => a.sla - b.sla)
  }, [unidadeRows])

  const ufsDisp = useMemo(() => [...new Set(dbRows.map(r => r.uf).filter(u => u && u !== 'EMPTY'))].sort(), [dbRows])

  const allDates = useMemo(() => [...new Set(dbRows.map(r => r.data).filter(Boolean))].sort(), [dbRows])
  const filtDates = useMemo(() => [...new Set(filtered.map(r => r.data).filter(Boolean))].sort(), [filtered])

  const períodoLabel = filtDates.length === 0 ? 'sem dados'
    : filtDates.length === 1 ? fmtDate(filtDates[0])
    : `${fmtDate(filtDates[0])} → ${fmtDate(filtDates[filtDates.length-1])}`

  const hasData     = dbRows.length > 0
  const hasFiltered = filtered.length > 0

  return (
    <div style={{ background: C.bg, minHeight: '100vh', fontFamily: "'DM Sans','Segoe UI',sans-serif", color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;background:${C.bg}}
        ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
        select option{background:${C.surface};color:${C.text}}
        .ubtn:hover{opacity:.9;transform:translateY(-1px)}
        .tab:hover{background:${C.border}!important}
      `}</style>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js" />

      {/* ── Topbar ── */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`,
        padding: '0 36px', height: 60, display: 'flex', alignItems: 'center',
        justifyContent: 'space-between', position: 'sticky', top: 0, zIndex: 100 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, fontSize: 16,
            background: `linear-gradient(135deg,${C.accent},${C.accentB})`,
            display: 'flex', alignItems: 'center', justifyContent: 'center' }}>🏥</div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700 }}>Visitas Médicas · SLA Monitor</div>
            <div style={{ fontSize: 11, color: C.muted }}>
              Meta {META}% · {totalPrevistas.toLocaleString('pt-BR')} visitas previstas
              {uploadInfo && <span style={{ marginLeft: 8, color: C.accent }}>· {uploadInfo}</span>}
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          {/* Status */}
          {status && (
            <span style={{ fontSize: 11, color: status.startsWith('✓') || status.startsWith('☁') ? C.green : C.yellow, fontWeight: 600 }}>
              {status}
            </span>
          )}

          {/* Período */}
          <div style={{ display: 'flex', gap: 3, background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 4 }}>
            {PERIODOS.map(p => (
              <button key={p.key} onClick={() => setPeriodo(p.key)} style={{
                padding: '5px 12px', borderRadius: 7, border: 'none', fontSize: 12, fontWeight: 700, cursor: 'pointer', transition: 'all .15s',
                background: período === p.key ? `linear-gradient(135deg,${C.accent},${C.accentB})` : 'transparent',
                color: período === p.key ? '#000' : C.muted,
              }}>{p.label}</button>
            ))}
          </div>

          {/* UF */}
          <select value={ufFiltro} onChange={e => setUfFiltro(e.target.value)} style={{
            background: C.card, border: `1px solid ${C.border}`, borderRadius: 9,
            color: C.text, fontSize: 13, padding: '7px 12px', outline: 'none', cursor: 'pointer' }}>
            <option value="TODOS">Todos os Estados</option>
            {ufsDisp.map(u => <option key={u} value={u}>{u}</option>)}
          </select>

          {/* Limpar */}
          {hasData && !saving && (
            <button onClick={clearAll} style={{ background: 'transparent', border: `1px solid ${C.red}44`,
              borderRadius: 9, color: C.red, fontSize: 11, padding: '7px 11px', cursor: 'pointer', whiteSpace: 'nowrap' }}>
              🗑 Limpar
            </button>
          )}

          {/* Upload */}
          <label className="ubtn" style={{
            background: saving ? C.border : `linear-gradient(135deg,${C.accent},${C.accentB})`,
            color: '#000', fontWeight: 700, fontSize: 13, padding: '8px 18px', borderRadius: 9,
            cursor: saving ? 'default' : 'pointer', transition: 'all .2s', whiteSpace: 'nowrap' }}>
            {loading ? 'Lendo…' : saving ? 'Salvando…' : '+ Carregar Planilhas'}
            <input type="file" accept=".xlsx,.xls,.csv" multiple style={{ display: 'none' }}
              onChange={handleFiles} disabled={loading || saving} />
          </label>
        </div>
      </div>

      <div style={{ padding: '22px 36px' }}>
        {/* ── Sem dados ── */}
        {!hasData && (
          <div style={{ minHeight: 'calc(100vh - 110px)', display: 'flex', flexDirection: 'column',
            alignItems: 'center', justifyContent: 'center', gap: 16 }}>
            <div style={{ fontSize: 52 }}>📋</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>Nenhuma planilha carregada</div>
            <div style={{ color: C.muted, fontSize: 13, textAlign: 'center', maxWidth: 480, lineHeight: 1.8 }}>
              Carregue os arquivos <strong style={{ color: C.sub }}>07H</strong> (previstas) e{' '}
              <strong style={{ color: C.sub }}>12H</strong> (pendentes/realizadas) de um ou mais dias.<br/>
              O sistema detecta a data automaticamente pelo nome do arquivo.<br/>
              <span style={{ color: C.accent }}>Exemplo: VisitasMedicas_07H_12_05.xlsx</span>
            </div>
            {status && <div style={{ fontSize: 13, color: C.yellow, background: C.card,
              padding: '10px 20px', borderRadius: 10, border: `1px solid ${C.yellow}44` }}>{status}</div>}
          </div>
        )}

        {/* ── Dados mas sem filtro ── */}
        {hasData && !hasFiltered && (
          <div style={{ minHeight: 'calc(100vh - 110px)', display: 'flex', flexDirection: 'column',
            alignItems: 'center', justifyContent: 'center', gap: 14 }}>
            <div style={{ fontSize: 48 }}>🔍</div>
            <div style={{ fontSize: 18, fontWeight: 700 }}>Sem dados para este período</div>
            <div style={{ color: C.muted, fontSize: 13 }}>
              Os arquivos carregados cobrem:{' '}
              <strong style={{ color: C.sub }}>{allDates.map(fmtDate).join(', ')}</strong>
            </div>
            <button onClick={() => setPeriodo('TODOS')} style={{
              background: `linear-gradient(135deg,${C.accent},${C.accentB})`,
              color: '#000', fontWeight: 700, fontSize: 13, padding: '9px 20px', borderRadius: 9, border: 'none', cursor: 'pointer' }}>
              Ver Todos os Dados
            </button>
          </div>
        )}

        {/* ── Dashboard ── */}
        {hasData && hasFiltered && (<>
          {/* Hero SLA */}
          <div style={{ background: `linear-gradient(135deg,#0a1628,#0d1e35)`,
            border: `1px solid ${C.border}`, borderRadius: 16, padding: '28px 32px',
            marginBottom: 18, display: 'flex', gap: 32, alignItems: 'center' }}>
            <Donut pct={slaGeral} size={140} />
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                letterSpacing: '.12em', color: C.muted, marginBottom: 8 }}>
                SLA GERAL DE VISITAS MÉDICAS · {períodoLabel.toUpperCase()}
              </div>
              <div style={{ fontSize: 52, fontWeight: 900, color: slaColor(slaGeral), lineHeight: 1 }}>
                {slaGeral.toFixed(1)}%
              </div>
              <div style={{ fontSize: 14, color: C.sub, marginTop: 8 }}>
                <strong style={{ color: C.green }}>{totalRealizadas.toLocaleString('pt-BR')}</strong> realizadas de{' '}
                <strong style={{ color: C.text  }}>{totalPrevistas.toLocaleString('pt-BR')}</strong> previstas ·{' '}
                <strong style={{ color: C.red   }}>{totalPendentes.toLocaleString('pt-BR')}</strong> pendentes
              </div>
              <div style={{ marginTop: 12, display: 'flex', gap: 10, alignItems: 'center' }}>
                {faltamMeta > 0 ? (
                  <>
                    <span style={{ fontSize: 12, background: C.red+'22', color: C.red,
                      padding: '4px 12px', borderRadius: 99, fontWeight: 700 }}>
                      ⚠ Faltam {faltamMeta.toLocaleString('pt-BR')} visitas para atingir {META}%
                    </span>
                    <span style={{ fontSize: 11, fontWeight: 700, padding: '4px 10px', borderRadius: 99,
                      background: slaColor(slaGeral)+'22', color: slaColor(slaGeral) }}>{slaLabel(slaGeral)}</span>
                  </>
                ) : (
                  <span style={{ fontSize: 12, background: C.green+'22', color: C.green,
                    padding: '4px 12px', borderRadius: 99, fontWeight: 700 }}>✓ Meta {META}% atingida!</span>
                )}
              </div>
            </div>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                letterSpacing: '.12em', color: C.muted, marginBottom: 12 }}>Progresso vs Meta {META}%</div>
              <ProgressBar pct={slaGeral} />
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 10, marginTop: 18 }}>
                {[
                  { label: `OK ≥${META}%`, value: ok.length,      color: C.green  },
                  { label: 'Em Risco',      value: emRisco.length, color: C.yellow },
                  { label: 'Críticas',      value: criticas.length,color: C.red    },
                ].map(k => (
                  <div key={k.label} style={{ textAlign: 'center', background: k.color+'11',
                    border: `1px solid ${k.color}33`, borderRadius: 10, padding: '10px 8px' }}>
                    <div style={{ fontSize: 24, fontWeight: 900, color: k.color }}>{k.value}</div>
                    <div style={{ fontSize: 9.5, color: C.muted, textTransform: 'uppercase', letterSpacing: '.08em', marginTop: 4 }}>{k.label}</div>
                    <div style={{ fontSize: 9, color: C.muted }}>unidades</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* KPIs */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5,1fr)', gap: 12, marginBottom: 18 }}>
            <KpiCard icon="📋" label="Visitas Previstas (07H)" value={totalPrevistas.toLocaleString('pt-BR')}
              sub={`${unidadeRows.length} unidades · ${ufFiltro === 'TODOS' ? 'todas UFs' : ufFiltro}`} accent={C.accent} />
            <KpiCard icon="✅" label="Realizadas" value={totalRealizadas.toLocaleString('pt-BR')}
              sub={`${slaGeral.toFixed(1)}% do total`} accent={C.green} />
            <KpiCard icon="⏳" label="Pendentes (12H)" value={totalPendentes.toLocaleString('pt-BR')}
              sub={totalPrevistas > 0 ? `${(100-slaGeral).toFixed(1)}% não visitados` : '—'} accent={C.red} />
            <KpiCard icon="🏆" label="Unidades Atingiram Meta" value={`${ok.length}/${unidadeRows.length}`}
              sub={unidadeRows.length > 0 ? `${((ok.length/unidadeRows.length)*100).toFixed(0)}% das unidades` : '—'} accent={C.green} />
            <KpiCard icon="🎯" label="Faltam para 90%" value={faltamMeta > 0 ? faltamMeta.toLocaleString('pt-BR') : '—'}
              sub={faltamMeta > 0 ? 'visitas ainda necessárias' : 'Meta já atingida!'} accent={faltamMeta > 0 ? C.orange : C.green} />
          </div>

          {/* Abas */}
          <div style={{ display: 'flex', gap: 8, marginBottom: 14 }}>
            {[
              { key:'unidades', label:'📋 Por Unidade'     },
              { key:'criticas', label:'🚨 Postos Críticos' },
              { key:'ufs',      label:'📍 Por UF'          },
              { key:'tendencia',label:'📈 Tendência'       },
            ].map(t => (
              <button key={t.key} onClick={() => setAbaSel(t.key)} className="tab" style={{
                padding: '9px 18px', borderRadius: 10,
                border: `1px solid ${abaSel === t.key ? C.accent : C.border}`,
                background: abaSel === t.key ? C.accent+'22' : 'transparent',
                color: abaSel === t.key ? C.accent : C.sub,
                fontSize: 13, fontWeight: 700, cursor: 'pointer', transition: 'all .15s' }}>{t.label}</button>
            ))}
          </div>

          {/* Conteúdo das abas */}
          <Card style={{ marginBottom: 14 }}>
            {abaSel === 'unidades' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  Todas as Unidades — ordenadas por SLA (pior → melhor)
                </div>
                <UnidadeTable rows={unidadeRows} showUF={ufFiltro === 'TODOS'} />
              </>
            )}
            {abaSel === 'criticas' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  🚨 Postos Críticos (SLA &lt; {META}%) — {criticas.length + emRisco.length} unidades
                </div>
                <UnidadeTable rows={[...criticas, ...emRisco]} showUF={ufFiltro === 'TODOS'} />
              </>
            )}
            {abaSel === 'tendencia' && (
              <TendenciaChart dbRows={dbRows} filtDates={filtDates} ufFiltro={ufFiltro} />
            )}
            {abaSel === 'ufs' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  SLA por Estado (UF)
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 10 }}>
                  {ufRows.map(r => {
                    const color = slaColor(r.sla)
                    return (
                      <div key={r.uf} style={{ background: C.card2, border: `1px solid ${color}33`,
                        borderLeft: `4px solid ${color}`, borderRadius: 10, padding: '14px 16px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                          <span style={{ fontSize: 16, fontWeight: 900, color: C.text }}>{r.uf}</span>
                          <span style={{ fontSize: 18, fontWeight: 900, color }}>{r.sla.toFixed(1)}%</span>
                        </div>
                        <div style={{ background: C.border, borderRadius: 99, height: 6, overflow: 'hidden', marginBottom: 8 }}>
                          <div style={{ height: '100%', borderRadius: 99, background: color, width: `${Math.min(r.sla,100)}%` }} />
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: C.muted }}>
                          <span>✅ {r.realizadas}</span><span>📋 {r.previstas}</span><span style={{ color: C.red }}>⏳ {r.pendentes}</span>
                        </div>
                      </div>
                    )
                  })}
                </div>
              </>
            )}
          </Card>

          {/* Grid bolinhas */}
          <Card style={{ marginBottom: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '.12em', color: C.muted }}>
                SLA por Unidade — {unidadeRows.length} unidades
              </div>
              <div style={{ display: 'flex', gap: 14, alignItems: 'center' }}>
                {[{color:C.green,label:`≥${META}% OK`},{color:C.yellow,label:'70–89% Risco'},{color:C.red,label:'<70% Crítico'}].map(l => (
                  <div key={l.label} style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                    <div style={{ width: 12, height: 12, borderRadius: '50%', background: l.color }} />
                    <span style={{ fontSize: 10, color: C.muted }}>{l.label}</span>
                  </div>
                ))}
              </div>
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
              {[...unidadeRows].sort((a,b) => b.sla-a.sla).map((r, i) => {
                const color = slaColor(r.sla)
                return (
                  <div key={`${r.unidade}${i}`}
                    title={`${r.unidade}\nSLA: ${r.sla.toFixed(1)}%\n${r.realizadas}/${r.previstas} realizadas`}
                    style={{ width: 64, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 5, cursor: 'default' }}>
                    <div style={{ width: 52, height: 52, borderRadius: '50%', background: color+'22',
                      border: `3px solid ${color}`, display: 'flex', alignItems: 'center', justifyContent: 'center',
                      flexDirection: 'column', boxShadow: `0 0 12px ${color}44`, transition: 'transform .15s' }}
                      onMouseEnter={e => (e.currentTarget.style.transform='scale(1.15)')}
                      onMouseLeave={e => (e.currentTarget.style.transform='scale(1)')}>
                      <span style={{ fontSize: 9, fontWeight: 900, color, lineHeight: 1 }}>{r.sla.toFixed(1)}%</span>
                    </div>
                    <div style={{ fontSize: 8.5, color: C.muted, textAlign: 'center', lineHeight: 1.3, maxWidth: 62,
                      overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical' as const }}>
                      {r.unidade}
                    </div>
                  </div>
                )
              })}
            </div>
          </Card>

          <div style={{ textAlign: 'center', color: C.muted, fontSize: 11, paddingTop: 16, paddingBottom: 20 }}>
            Visitas Médicas SLA Monitor · Meta {META}% · {períodoLabel}
            {ufFiltro !== 'TODOS' && ` · UF: ${ufFiltro}`}
          </div>
        </>)}
      </div>
    </div>
  )
}
