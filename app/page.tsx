'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
// ─── Design tokens ────────────────────────────────────────
const C = {
  bg:       '#06090F',
  surface:  '#0A1221',
  card:     '#0D1829',
  card2:    '#0B1524',
  border:   '#162236',
  accent:   '#06B6D4',   // cyan-500
  green:    '#22C55E',
  yellow:   '#EAB308',
  red:      '#EF4444',
  orange:   '#F97316',
  indigo:   '#6366F1',
  text:     '#EFF6FF',
  sub:      '#7BA7C8',
  muted:    '#334E68',
}

const SLA_META = 90

// ─── Types ────────────────────────────────────────────────
interface Row {
  UF: string
  UNIDADE: string
  POSTO: string
  ATENDIMENTO: string | number
  PACIENTE: string
  'VISITA REALIZADA': string
}

interface UnitStat {
  unidade: string
  uf: string
  total: number
  pendentes: number
  realizadas: number
  sla: number
}

interface PostoStat {
  posto: string
  unidade: string
  total: number
  pendentes: number
  realizadas: number
  sla: number
}

// ─── Helpers ──────────────────────────────────────────────
const slaColor = (p: number) =>
  p >= SLA_META ? C.green : p >= 70 ? C.yellow : p >= 40 ? C.orange : C.red

const slaLabel = (p: number) =>
  p >= SLA_META ? 'Meta atingida' : p >= 70 ? 'Em risco' : p >= 40 ? 'Crítico' : 'Grave'

const readXlsx = async (file: File): Promise<Row[]> => {
  const buf = await file.arrayBuffer()
  const wb  = XLSX.read(buf, { type: 'buffer' })
  const ws  = wb.Sheets[wb.SheetNames[0]]
  return XLSX.utils.sheet_to_json<Row>(ws, { range: 1, defval: '' })
}

// ─── Sub-components ───────────────────────────────────────

// Donut ring com SLA %
function SlaRing({
  pct, size = 120, stroke = 10,
}: { pct: number; size?: number; stroke?: number }) {
  const r    = (size - stroke) / 2
  const cx   = size / 2
  const cy   = size / 2
  const circ = 2 * Math.PI * r
  const fill = Math.min(pct, 100) / 100 * circ
  const color = slaColor(pct)
  // target arc at 90%
  const targetOffset = circ * (1 - SLA_META / 100)

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      {/* track */}
      <circle cx={cx} cy={cy} r={r} fill="none"
        stroke={C.border} strokeWidth={stroke} />
      {/* target marker at 90% */}
      <circle cx={cx} cy={cy} r={r} fill="none"
        stroke={C.accent} strokeWidth={stroke} opacity={0.2}
        strokeDasharray={`2 ${circ - 2}`}
        strokeDashoffset={-circ * SLA_META / 100 + circ / 4}
        strokeLinecap="round" />
      {/* fill */}
      <circle cx={cx} cy={cy} r={r} fill="none"
        stroke={color} strokeWidth={stroke}
        strokeDasharray={`${fill} ${circ - fill}`}
        strokeDashoffset={circ / 4}
        strokeLinecap="round"
        style={{ transition: 'stroke-dasharray .8s ease' }} />
      <text x={cx} y={cy - 6} textAnchor="middle"
        fontSize={size * 0.2} fontWeight="800" fill={color}>
        {pct.toFixed(0)}%
      </text>
      <text x={cx} y={cy + 12} textAnchor="middle"
        fontSize={size * 0.09} fill={C.muted}>
        SLA
      </text>
    </svg>
  )
}

// Barra horizontal de SLA com linha de meta
function SlaBar({
  label, realizadas, total, pct, rank, compact = false,
}: {
  label: string; realizadas: number; total: number
  pct: number; rank?: number; compact?: boolean
}) {
  const color = slaColor(pct)
  return (
    <div style={{ marginBottom: compact ? 10 : 14 }}>
      <div style={{
        display: 'flex', justifyContent: 'space-between',
        alignItems: 'center', marginBottom: 5, gap: 8,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
          {rank != null && (
            <span style={{
              fontSize: 10, fontWeight: 800, minWidth: 22,
              color: rank < 3 ? C.red : C.muted,
            }}>#{rank + 1}</span>
          )}
          <span style={{
            fontSize: compact ? 11.5 : 12.5, color: C.text, fontWeight: 500,
            overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
          }}>{label}</span>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexShrink: 0 }}>
          <span style={{ fontSize: 11, color: C.sub }}>{realizadas}/{total}</span>
          <span style={{
            fontSize: 13, fontWeight: 800, color, minWidth: 44, textAlign: 'right',
          }}>{pct.toFixed(0)}%</span>
        </div>
      </div>
      <div style={{
        background: C.border, borderRadius: 99, height: compact ? 5 : 7,
        position: 'relative', overflow: 'hidden',
      }}>
        {/* meta line */}
        <div style={{
          position: 'absolute', left: `${SLA_META}%`, top: 0, bottom: 0,
          width: 1.5, background: C.accent, opacity: .6, zIndex: 2,
        }} />
        <div style={{
          height: '100%', borderRadius: 99, background: color,
          width: `${Math.min(pct, 100)}%`, transition: 'width .6s ease',
        }} />
      </div>
    </div>
  )
}

// KPI card com rótulo, valor e sub
function KpiCard({
  icon, label, value, sub, color, highlight = false,
}: {
  icon: string; label: string; value: string | number
  sub?: string; color: string; highlight?: boolean
}) {
  return (
    <div style={{
      background: highlight ? color + '12' : C.card,
      border: `1px solid ${highlight ? color + '40' : C.border}`,
      borderRadius: 14, padding: '18px 20px',
      display: 'flex', flexDirection: 'column', gap: 8,
      position: 'relative', overflow: 'hidden',
    }}>
      <div style={{
        position: 'absolute', top: -20, right: -20,
        width: 80, height: 80, borderRadius: '50%',
        background: color, opacity: .06,
      }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <span style={{ fontSize: 18 }}>{icon}</span>
        <span style={{
          fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
          letterSpacing: '.1em', color: C.muted,
        }}>{label}</span>
      </div>
      <div style={{
        fontSize: 34, fontWeight: 900, color, lineHeight: 1, letterSpacing: '-1px',
      }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: C.sub }}>{sub}</div>}
    </div>
  )
}

// Badge pill
function Badge({ label, color }: { label: string; color: string }) {
  return (
    <span style={{
      fontSize: 10, fontWeight: 700, padding: '3px 9px',
      borderRadius: 99, background: color + '1A', color,
      letterSpacing: '.04em', whiteSpace: 'nowrap',
    }}>{label}</span>
  )
}

// Section header
function SH({ children }: { children: React.ReactNode }) {
  return (
    <div style={{
      fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
      letterSpacing: '.12em', color: C.muted, marginBottom: 16,
    }}>{children}</div>
  )
}

// Card wrapper
function Card({ children, style = {} }: {
  children: React.ReactNode; style?: React.CSSProperties
}) {
  return (
    <div style={{
      background: C.card, border: `1px solid ${C.border}`,
      borderRadius: 14, padding: 22, ...style,
    }}>{children}</div>
  )
}

// Mini gauge strip  
function GaugeStrip({ pct }: { pct: number }) {
  const color = slaColor(pct)
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <div style={{
        flex: 1, background: C.border, borderRadius: 99, height: 6,
        position: 'relative', overflow: 'hidden',
      }}>
        <div style={{
          position: 'absolute', left: `${SLA_META}%`, top: 0, bottom: 0,
          width: 1.5, background: C.accent, opacity: .5,
        }} />
        <div style={{
          height: '100%', borderRadius: 99, background: color,
          width: `${Math.min(pct, 100)}%`,
        }} />
      </div>
      <span style={{ fontSize: 12, fontWeight: 800, color, minWidth: 38, textAlign: 'right' }}>
        {pct.toFixed(0)}%
      </span>
    </div>
  )
}

// ─── Main Component ───────────────────────────────────────
export default function Page() {
const [dados, setDados] = useState<any[]>([])
const [loading, setLoading] = useState(false)
const [ufFilter, setUfFilter] = useState('TODOS')

const [tab, setTab] =
  useState<'unidades' | 'postos' | 'ufs'>('unidades')

const [dateFilter, setDateFilter] = useState('hoje')

const [startDate, setStartDate] = useState('')
const [endDate, setEndDate] = useState('')
  useEffect(() => {

  async function carregarDados() {

    const { data, error } = await supabase
      .from('visitas_medicas')
      .select('*')

    if (!error && data) {
      setDados(data)
    }
  }

  carregarDados()

  const channel = supabase
    .channel('realtime-visitas')

    .on(
      'postgres_changes',
      {
        event: '*',
        schema: 'public',
        table: 'visitas_medicas'
      },
      () => {
        carregarDados()
      }
    )

    .subscribe()

  return () => {
    supabase.removeChannel(channel)
  }

}, [])

const dadosFiltrados = useMemo(() => {

  if (!dados.length) return []

  const hoje = new Date()

  return dados.filter((r: any) => {

    if (!r.created_at) return true

    const data = new Date(r.created_at)

    if (isNaN(data.getTime())) return true

    if (dateFilter === 'hoje') {
      return data.toDateString() === hoje.toDateString()
    }

    if (dateFilter === 'semana') {

      const inicioSemana = new Date(hoje)

      inicioSemana.setDate(hoje.getDate() - hoje.getDay())

      inicioSemana.setHours(0, 0, 0, 0)

      return data >= inicioSemana
    }

    if (dateFilter === 'mes') {
      return (
        data.getMonth() === hoje.getMonth() &&
        data.getFullYear() === hoje.getFullYear()
      )
    }

    if (dateFilter === 'ano') {
      return data.getFullYear() === hoje.getFullYear()
    }

    return true

  })

}, [dados, dateFilter])

const loaded = dados.length > 0

  // ── Upload handler
 const handleFiles = async (
  e: React.ChangeEvent<HTMLInputElement>
) => {

  const files = Array.from(e.target.files || [])

  if (files.length < 2) return

  setLoading(true)

  try {

    const f07 =
      files.find(f => f.name.includes('07')) || files[0]

    const f12 =
      files.find(f => f.name.includes('12')) || files[1]

    const [r07, r12] = await Promise.all([
      readXlsx(f07),
      readXlsx(f12)
    ])
console.log('07H', r07[0])
console.log('12H', r12[0])
    const previstasPorUnidade = new Map()
    const pendentesPorUnidade = new Map()

    r07.forEach((r: any) => {

const unidade = String(r.UNIDADE || '').trim()
const uf = String(r.UF || '').trim()

      const key = `${uf}-${unidade}`

      if (!previstasPorUnidade.has(key)) {
        previstasPorUnidade.set(key, {
          unidade,
          uf,
          previstas: 0
        })
      }

      previstasPorUnidade.get(key).previstas++
    })

    r12.forEach((r: any) => {

const unidade = String(r.UNIDADE || '').trim()
const uf = String(r.UF || '').trim()

      const key = `${uf}-${unidade}`

      if (!pendentesPorUnidade.has(key)) {
        pendentesPorUnidade.set(key, 0)
      }

      pendentesPorUnidade.set(
        key,
        pendentesPorUnidade.get(key) + 1
      )
    })

    const finalData = Array.from(
      previstasPorUnidade.entries()
    ).map(([key, value]: any) => {

      const pendentes =
        pendentesPorUnidade.get(key) || 0

      return {
        unidade: value.unidade,
        uf: value.uf,
        previstas: value.previstas,
        pendentes,
        realizadas:
          value.previstas - pendentes
      }
    })

    await supabase
      .from('visitas_medicas')
      .delete()
      .neq('id', 0)

    await supabase
      .from('visitas_medicas')
      .insert(finalData)

  } finally {

    setLoading(false)
  }
}

  const reset = () => {
  setUfFilter('TODOS')
}

  // ── UF list
 const ufs = useMemo(() =>
  [...new Set(
    dados.map(r => String(r.uf || '').trim()).filter(Boolean)
  )].sort(),
  [dados]
)

  // ── Filtered rows
const f07 = useMemo(() =>
  ufFilter === 'TODOS'
    ? dados
    : dados.filter(
        r => String(r.uf).trim() === ufFilter
      ),
  [dados, ufFilter]
)

 const f12 = useMemo(() =>
  ufFilter === 'TODOS'
    ? dados
    : dados.filter(
        r => String(r.uf).trim() === ufFilter
      ),
  [dados, ufFilter]
)

// ── Global KPIs
const gtotal = dadosFiltrados.reduce(
  (acc, r) => acc + Number(r.previstas || 0),
  0
)

const gpend = dadosFiltrados.reduce(
  (acc, r) => acc + Number(r.pendentes || 0),
  0
)

const greal = dadosFiltrados.reduce(
  (acc, r) => acc + Number(r.realizadas || 0),
  0
)

const gslaPct =
  gtotal > 0 ? (greal / gtotal) * 100 : 0

const gfalta = Math.max(
  0,
  Math.ceil(gtotal * SLA_META / 100) - greal
)

const gAtingiu = gslaPct >= SLA_META

  // ── Unit stats
const unitStats = useMemo((): UnitStat[] => {

  return dadosFiltrados.map((r: any) => {

    const total = Number(r.previstas || 0)
    const pendentes = Number(r.pendentes || 0)
    const realizadas = Number(r.realizadas || 0)

    return {
      unidade: r.unidade,
      uf: r.uf,
      total,
      pendentes,
      realizadas,
      sla: total > 0
        ? (realizadas / total) * 100
        : 0
    }
  })

}, [dados])

  const byPendDesc   = [...unitStats].sort((a, b) => b.pendentes - a.pendentes)
  const bySlaAsc     = [...unitStats].filter(u => u.sla < SLA_META).sort((a, b) => a.sla - b.sla)
  const bySlaDesc    = [...unitStats].sort((a, b) => b.sla - a.sla)
  const slaOk        = unitStats.filter(u => u.sla >= SLA_META).length
  const slaCrit      = unitStats.filter(u => u.sla <  40).length
  const slaRisk      = unitStats.filter(u => u.sla >= 40 && u.sla < SLA_META).length

  // ── UF stats
const ufStats = useMemo(() => {

  const mapa: any = {}

  dadosFiltrados.forEach((r: any) => {

    const uf = String(r.uf || '').trim()

    if (!mapa[uf]) {
      mapa[uf] = {
        total: 0,
        pend: 0,
        real: 0
      }
    }

    mapa[uf].total += Number(r.previstas || 0)
    mapa[uf].pend += Number(r.pendentes || 0)
    mapa[uf].real += Number(r.realizadas || 0)
  })

  return Object.entries(mapa).map(([uf, v]: any) => ({
    uf,
    total: v.total,
    pend: v.pend,
    real: v.real,
    sla:
      v.total > 0
        ? (v.real / v.total) * 100
        : 0
  }))

}, [dados])

  // ── Posto stats
const postoStats = useMemo((): PostoStat[] => {

 return dadosFiltrados.map((r: any) => {

    const total = Number(r.previstas || 0)
    const pendentes = Number(r.pendentes || 0)
    const realizadas = Number(r.realizadas || 0)

    return {
      posto: r.unidade,
      unidade: r.unidade,
      total,
      pendentes,
      realizadas,
      sla:
        total > 0
          ? (realizadas / total) * 100
          : 0
    }
  })
}, [dados])

  return (
    <div style={{
      background: C.bg, minHeight: '100vh', color: C.text,
      fontFamily: "'DM Sans','Segoe UI',sans-serif",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;background:${C.bg}}
        ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
        select option{background:${C.surface}}
        .ubtn:hover{filter:brightness(1.1);transform:translateY(-1px)}
        .trow:hover td{background:#0d1e30!important}
        input::placeholder{color:${C.muted}}
        .tab-btn{transition:all .15s}
      `}</style>

      {/* ── Topbar ── */}
      <div style={{
        background: C.surface, borderBottom: `1px solid ${C.border}`,
        padding: '0 40px', height: 62,
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        position: 'sticky', top: 0, zIndex: 99,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 32, height: 32, borderRadius: 8, fontSize: 16,
            background: `linear-gradient(135deg,${C.accent},${C.indigo})`,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>🏥</div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700 }}>Visitas Médicas · SLA Monitor</div>
            <div style={{ fontSize: 11, color: C.muted }}>
              {loaded
                ? `Meta ${SLA_META}% · ${gtotal.toLocaleString('pt-BR')} visitas previstas · ${unitStats.length} unidades`
                : `Meta ${SLA_META}% de visitas realizadas · 07h vs 12h`}
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          {loaded && (
            <>
              <select value={ufFilter} onChange={e => setUfFilter(e.target.value)} style={{
                background: C.card, border: `1px solid ${C.border}`, borderRadius: 9,
                color: C.text, fontSize: 13, padding: '7px 12px', outline: 'none', cursor: 'pointer',
              }}>
                <option value="TODOS">Todos os Estados</option>
                {ufs.map(u => <option key={u}>{u}</option>)}
              </select>

<select
  value={dateFilter}
  onChange={e => setDateFilter(e.target.value)}
  style={{
    background: C.card,
    border: `1px solid ${C.border}`,
    borderRadius: 9,
    color: C.text,
    fontSize: 13,
    padding: '7px 12px',
    outline: 'none',
    cursor: 'pointer',
  }}
>
<option value="hoje">Hoje</option>
<option value="semana">Semana Atual</option>
<option value="mes">Mês Atual</option>
<option value="ano">Ano Atual</option>
</select>

              <button onClick={reset} style={{
                background: 'transparent', border: `1px solid ${C.border}`,
                borderRadius: 9, color: C.sub, fontSize: 13,
                padding: '7px 14px', cursor: 'pointer',
              }}>✕ Redefinir</button>
            </>
          )}
          <label className="ubtn" style={{
            background: `linear-gradient(135deg,${C.accent},${C.indigo})`,
            color: '#000', fontWeight: 700, fontSize: 13,
            padding: '8px 20px', borderRadius: 9, cursor: 'pointer', transition: 'all .2s',
          }}>
            {loading ? 'Lendo…' : '+ Carregar Planilhas'}
            <input
  type="file"
  accept=".xlsx,.xls"
  multiple
  style={{ display: 'none' }}
  onChange={handleFiles}
/>
          </label>
        </div>
      </div>

      <div style={{ padding: '28px 40px' }}>

        {/* ── Empty state ── */}
        {!loaded && (
          <div style={{
            minHeight: 'calc(100vh - 130px)', display: 'flex',
            flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 16,
          }}>
            <div style={{ fontSize: 56 }}>🏥</div>
            <div style={{ fontSize: 22, fontWeight: 800 }}>Carregue as duas planilhas</div>
            <div style={{ color: C.muted, fontSize: 13, textAlign: 'center', maxWidth: 420, lineHeight: 1.7 }}>
              Selecione os dois arquivos juntos:<br />
              <strong style={{ color: C.sub }}>07h</strong> — total de visitas previstas do dia<br />
              <strong style={{ color: C.sub }}>12h</strong> — visitas pendentes (não realizadas)
            </div>
          </div>
        )}

        {loaded && (<>

          {/* ── Hero: SLA Geral ── */}
          <div style={{
            background: C.card,
            border: `1px solid ${gAtingiu ? C.green + '40' : C.red + '30'}`,
            borderRadius: 16, padding: '28px 36px', marginBottom: 20,
            display: 'flex', alignItems: 'center', gap: 40, flexWrap: 'wrap',
          }}>
            <SlaRing pct={gslaPct} size={130} stroke={12} />

            <div style={{ flex: 1, minWidth: 260 }}>
             <div style={{
  fontSize: 10,
  fontWeight: 700,
  textTransform: 'uppercase',
  letterSpacing: '.12em',
  color: C.muted,
  marginBottom: 10,
}}>
  {`SLA GERAL DE VISITAS MÉDICAS · ${new Date().toLocaleDateString('pt-BR')}`}
</div>

              <div style={{
                fontSize: 44, fontWeight: 900, color: slaColor(gslaPct),
                lineHeight: 1, letterSpacing: '-2px', marginBottom: 10,
              }}>
                {gslaPct.toFixed(1)}%
              </div>

              <div style={{ fontSize: 14, color: C.sub, marginBottom: 14 }}>
                <strong style={{ color: C.text }}>{greal.toLocaleString('pt-BR')}</strong> realizadas de{' '}
                <strong style={{ color: C.text }}>{gtotal.toLocaleString('pt-BR')}</strong> previstas ·{' '}
                <strong style={{ color: C.red }}>{gpend.toLocaleString('pt-BR')}</strong> pendentes
              </div>

              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                {gAtingiu
                  ? <Badge label="✓ Meta de 90% atingida" color={C.green} />
                  : <Badge label={`⚠ Faltam ${gfalta.toLocaleString('pt-BR')} visitas para atingir 90%`} color={C.red} />
                }
                <Badge label={slaLabel(gslaPct)} color={slaColor(gslaPct)} />
              </div>
            </div>

            {/* Barra visual de progresso vs meta */}
            <div style={{ flex: 1, minWidth: 260 }}>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 10 }}>
                Progresso vs Meta {SLA_META}%
              </div>
              <div style={{
                background: C.border, borderRadius: 99, height: 14,
                position: 'relative', overflow: 'hidden',
              }}>
                <div style={{
                  position: 'absolute', left: `${SLA_META}%`, top: 0, bottom: 0,
                  width: 2, background: C.accent, zIndex: 2,
                }} />
                <div style={{
                  height: '100%', borderRadius: 99,
                  background: `linear-gradient(90deg, ${slaColor(gslaPct)}, ${slaColor(gslaPct)}cc)`,
                  width: `${Math.min(gslaPct, 100)}%`, transition: 'width .8s ease',
                }} />
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 8 }}>
                <span style={{ fontSize: 11, color: C.muted }}>0%</span>
                <span style={{ fontSize: 11, color: C.accent, fontWeight: 700 }}>
                  ← Meta {SLA_META}%
                </span>
                <span style={{ fontSize: 11, color: C.muted }}>100%</span>
              </div>

              {/* mini stats abaixo */}
              <div style={{
                display: 'grid', gridTemplateColumns: '1fr 1fr 1fr',
                gap: 10, marginTop: 18,
              }}>
                {[
                  { label: 'OK ≥ 90%',  value: slaOk,   color: C.green  },
                  { label: 'Em risco',  value: slaRisk,  color: C.yellow },
                  { label: 'Críticas',  value: slaCrit,  color: C.red    },
                ].map(s => (
                  <div key={s.label} style={{
                    background: C.card2, borderRadius: 10, padding: '10px 12px',
                    border: `1px solid ${s.color}22`,
                  }}>
                    <div style={{ fontSize: 9.5, color: C.muted, textTransform: 'uppercase',
                      letterSpacing: '.08em', marginBottom: 4 }}>{s.label}</div>
                    <div style={{ fontSize: 24, fontWeight: 900, color: s.color }}>
                      {s.value}
                    </div>
                    <div style={{ fontSize: 10, color: C.muted }}>unidades</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* ── 5 KPI cards ── */}
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(5,1fr)',
            gap: 14, marginBottom: 20,
          }}>
            <KpiCard icon="📋" label="Visitas Previstas (07h)"
              value={gtotal.toLocaleString('pt-BR')} color={C.accent}
              sub={`${unitStats.length} unidades · ${ufFilter !== 'TODOS' ? ufFilter : 'todas as UFs'}`} />
            <KpiCard icon="✅" label="Realizadas"
              value={greal.toLocaleString('pt-BR')} color={C.green}
              sub={`${gslaPct.toFixed(1)}% do total`} highlight />
            <KpiCard icon="⏳" label="Pendentes (12h)"
              value={gpend.toLocaleString('pt-BR')} color={C.red}
              sub={`${(100 - gslaPct).toFixed(1)}% não visitados`} highlight />
            <KpiCard icon="🏢" label="Unidades atingiram meta"
              value={`${slaOk}/${unitStats.length}`} color={C.green}
              sub={`${((slaOk / unitStats.length) * 100).toFixed(0)}% das unidades`} />
            <KpiCard icon="🚨" label="Faltam para 90%"
              value={gfalta > 0 ? gfalta.toLocaleString('pt-BR') : '—'}
              color={gfalta > 0 ? C.red : C.green}
              sub={gfalta > 0 ? 'visitas ainda necessárias' : 'Meta já atingida!'} />
          </div>

          {/* ── Tabs ── */}
          <div style={{ display: 'flex', gap: 4, marginBottom: 16 }}>
            {[
              { key: 'unidades', label: '🏥 Por Unidade' },
              { key: 'postos',   label: '🚨 Postos Críticos' },
              { key: 'ufs',      label: '📍 Por UF' },
            ].map(t => (
              <button key={t.key} className="tab-btn"
                onClick={() => setTab(t.key as typeof tab)}
                style={{
                  padding: '9px 18px', borderRadius: 10, fontSize: 13, fontWeight: 600,
                  cursor: 'pointer', border: `1px solid ${tab === t.key ? C.accent : C.border}`,
                  background: tab === t.key ? C.accent + '18' : 'transparent',
                  color: tab === t.key ? C.accent : C.sub,
                }}>{t.label}</button>
            ))}
          </div>

          {/* ── Tab: Unidades ── */}
          {tab === 'unidades' && (
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 1fr',
              gap: 16, marginBottom: 16,
            }}>
              {/* Top pendências */}
              <Card>
                <SH>🔴 Unidades com mais pendências</SH>
                {byPendDesc.slice(0, 10).map((u, i) => (
                  <SlaBar key={u.unidade} rank={i}
                    label={u.unidade} realizadas={u.realizadas}
                    total={u.total} pct={u.sla} />
                ))}
              </Card>

              {/* Abaixo da meta */}
              <Card>
                <SH>⚠️ Abaixo da meta — ordenado por criticidade</SH>
                {bySlaAsc.slice(0, 10).map((u, i) => (
                  <SlaBar key={u.unidade} label={u.unidade}
                    realizadas={u.realizadas} total={u.total} pct={u.sla} />
                ))}
              </Card>
            </div>
          )}

          {/* Unidades — cards grid */}
          {tab === 'unidades' && bySlaAsc.length > 0 && (
            <Card style={{ marginBottom: 16 }}>
              <SH>📊 Mapa de SLA por Unidade</SH>
              <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
                gap: 12,
              }}>
                {bySlaDesc.map(u => (
                  <div key={u.unidade} style={{
                    background: C.card2, border: `1px solid ${slaColor(u.sla)}22`,
                    borderRadius: 12, padding: '14px 16px',
                    display: 'flex', alignItems: 'center', gap: 14,
                  }}>
                    <SlaRing pct={u.sla} size={60} stroke={6} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{
                        fontSize: 12, fontWeight: 600, color: C.text, marginBottom: 4,
                        overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                      }}>{u.unidade}</div>
                      <div style={{ fontSize: 11, color: C.sub, marginBottom: 6 }}>
                        {u.realizadas}/{u.total} · {u.pendentes} pendentes
                      </div>
                      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                        <Badge label={u.uf} color={C.muted} />
                        <Badge label={slaLabel(u.sla)} color={slaColor(u.sla)} />
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </Card>
          )}

          {/* ── Tab: Postos Críticos ── */}
          {tab === 'postos' && (
            <Card style={{ marginBottom: 16 }}>
              <div style={{ marginBottom: 4 }}>
                <SH>🚨 Postos abaixo de {SLA_META}% — {postoStats.length} postos críticos</SH>
              </div>
              <div style={{ fontSize: 11, color: C.sub, marginBottom: 18 }}>
                Ordenado por quantidade de pendências. Linha azul = meta de {SLA_META}%.
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12.5 }}>
                  <thead>
                    <tr>
                      {['#', 'Posto', 'Unidade', 'Previstas', 'Pendentes', 'Realizadas', 'SLA', 'Status'].map(h => (
                        <th key={h} style={{
                          padding: '10px 14px', textAlign: 'left',
                          borderBottom: `1px solid ${C.border}`,
                          color: C.muted, fontWeight: 600, fontSize: 10.5,
                          textTransform: 'uppercase', letterSpacing: '.07em',
                          whiteSpace: 'nowrap',
                        }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {postoStats.slice(0, 50).map((p, i) => {
                      const clr = slaColor(p.sla)
                      return (
                        <tr key={`${p.posto}-${i}`} className="trow">
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            color: C.muted, fontSize: 11 }}>{i + 1}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            color: C.sub, fontSize: 11.5, maxWidth: 300, overflow: 'hidden',
                            textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.posto}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            fontWeight: 500, color: C.text, whiteSpace: 'nowrap' }}>{p.unidade}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            textAlign: 'center', color: C.sub }}>{p.total}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            textAlign: 'center', color: C.red, fontWeight: 700 }}>{p.pendentes}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            textAlign: 'center', color: C.green }}>{p.realizadas}</td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}`,
                            minWidth: 140 }}>
                            <GaugeStrip pct={p.sla} />
                          </td>
                          <td style={{ padding: '11px 14px', borderBottom: `1px solid ${C.border}` }}>
                            <Badge label={slaLabel(p.sla)} color={clr} />
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </Card>
          )}

          {/* ── Tab: UFs ── */}
          {tab === 'ufs' && (
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 1fr',
              gap: 16, marginBottom: 16,
            }}>
              <Card>
                <SH>📍 SLA por Estado — do pior ao melhor</SH>
                {ufStats.map((u, i) => (
                  <SlaBar key={u.uf} label={`${u.uf}`}
                    realizadas={u.real} total={u.total} pct={u.sla} />
                ))}
              </Card>

              {/* Grid de UFs */}
              <Card>
                <SH>🗺️ Resumo por UF</SH>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {ufStats.map(u => (
                    <div key={u.uf} style={{
                      background: C.card2, borderRadius: 10,
                      padding: '12px 14px',
                      border: `1px solid ${slaColor(u.sla)}20`,
                      display: 'flex', alignItems: 'center', gap: 14,
                    }}>
                      <div style={{
                        fontSize: 15, fontWeight: 900, color: slaColor(u.sla),
                        minWidth: 36, textAlign: 'center',
                      }}>{u.uf}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <GaugeStrip pct={u.sla} />
                        <div style={{ fontSize: 10.5, color: C.muted, marginTop: 4 }}>
                          {u.real}/{u.total} realizadas · {u.pend} pendentes
                        </div>
                      </div>
                      <Badge label={slaLabel(u.sla)} color={slaColor(u.sla)} />
                    </div>
                  ))}
                </div>
              </Card>
            </div>
          )}

          {/* ── Footer ── */}
          <div style={{
            textAlign: 'center', color: C.muted, fontSize: 11,
            paddingBottom: 24, paddingTop: 8,
          }}>
            Visitas Médicas · 08/05 · Meta SLA {SLA_META}%
            {ufFilter !== 'TODOS' ? ` · Estado: ${ufFilter}` : ' · Todas as UFs'}
            {' · '}{gtotal.toLocaleString('pt-BR')} visitas previstas
          </div>

        </>)}
      </div>
    </div>
  )
}
