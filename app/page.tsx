'use client'

import { useMemo, useState, useCallback } from 'react'
import * as XLSX from 'xlsx'

// Parser CSV nativo (sem dependência externa)
const parseCSV = (text: string, delimiter = ';'): string[][] => {
  return text
    .split('\n')
    .map(line => line.split(delimiter).map(cell => cell.trim().replace(/^"|"$/g, '')))
    .filter(row => row.some(cell => cell !== ''))
}

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

// ─── Helpers ─────────────────────────────────────────────
// Extrai data do nome do arquivo: "VisitasMedicas_07H_12_05.xlsx" → '2026-05-12'
const extractDateFromName = (name: string): string | null => {
  const m = name.match(/_(\d{2})_(\d{2})/); // DD_MM
  if (!m) return null
  const day = m[1], month = m[2]
  const year = new Date().getFullYear()
  return `${year}-${month}-${day}`
}

// Extrai hora do nome: "07H" → 7, "12H" → 12
const extractHourFromName = (name: string): number | null => {
  const m = name.match(/_(\d{2})H_/)
  return m ? parseInt(m[1]) : null
}

// É arquivo de 07h (previstas) ou 12h (pendentes)?
const isPrevistas = (name: string): boolean => {
  const h = extractHourFromName(name)
  return h !== null && h < 12  // 07H = previstas
}

// Lê arquivo xlsx ou csv e retorna rows
const parseFile = async (file: File): Promise<Record<string, string>[]> => {
  const name = file.name
  const isXlsx = name.toLowerCase().endsWith('.xlsx') || name.toLowerCase().endsWith('.xls')

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

  // Encontra a linha do header: primeira linha que contém 'UF' ou 'UNIDADE'
  const headerIdx = raw.findIndex(row =>
    row.some(cell => String(cell).trim().toUpperCase() === 'UF' ||
                     String(cell).trim().toUpperCase() === 'UNIDADE')
  )
  if (headerIdx < 0) return []

  const header = raw[headerIdx].map(h => String(h).trim())
  const dataRows = raw.slice(headerIdx + 1)

  const rows: Record<string, string>[] = dataRows.map(r => {
    const obj: Record<string, string> = {}
    header.forEach((h, i) => { obj[h] = String(r[i] ?? '').trim() })
    return obj
  })

  // Filtra linhas válidas: UF preenchida e não é o próprio header repetido
  return rows.filter(r => r['UF'] && r['UF'] !== 'UF' && r['UF'].length <= 3)
}

// Data hoje e ontem no horário local
const localDateStr = (offset = 0): string => {
  const d = new Date()
  d.setDate(d.getDate() + offset)
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`
}
const fmtDate = (s: string | null): string => s ? s.split('-').reverse().join('/') : ''

// ─── Período filter ───────────────────────────────────────
const PERIODOS = [
  { key: 'TODOS',  label: 'Todos'  },
  { key: 'HOJE',   label: 'Hoje'   },
  { key: 'ONTEM',  label: 'Ontem'  },
  { key: 'SEMANA', label: 'Semana' },
  { key: 'MÊS',    label: 'Mês'    },
]

const filterByPeriodo = (dateStr: string, período: string): boolean => {
  if (período === 'TODOS' || !dateStr) return true
  const today    = localDateStr(0)
  const ontem    = localDateStr(-1)

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

// ─── SLA calculation helpers ─────────────────────────────
const slaColor = (pct: number): string =>
  pct >= META ? C.green : pct >= 70 ? C.yellow : C.red

const slaLabel = (pct: number): string =>
  pct >= META ? 'OK' : pct >= 70 ? 'Em Risco' : 'Crítico'

// ─── Types ────────────────────────────────────────────────
interface Snapshot {
  date: string
  previstas: Record<string, string>[]
  pendentes: Record<string, string>[]
}

interface UnidadeRow {
  uf: string
  unidade: string
  previstas: number
  pendentes: number
  realizadas: number
  sla: number
}

// ─── Sub-components ───────────────────────────────────────
function Card({ children, style = {} }: { children: React.ReactNode; style?: React.CSSProperties }) {
  return (
    <div style={{ background: C.card, border: `1px solid ${C.border}`,
      borderRadius: 14, padding: 18, ...style }}>
      {children}
    </div>
  )
}

function KpiCard({ icon, label, value, sub, accent, note }: {
  icon: string; label: string; value: string | number
  sub?: string; accent: string; note?: string
}) {
  return (
    <div style={{ background: C.card2, border: `1px solid ${C.border}`, borderRadius: 14,
      padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 6,
      position: 'relative', overflow: 'hidden' }}>
      <div style={{ position: 'absolute', top: -16, right: -16, width: 70, height: 70,
        borderRadius: '50%', background: accent, opacity: .08 }} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <span style={{ fontSize: 16 }}>{icon}</span>
        <span style={{ fontSize: 9.5, color: C.muted, textTransform: 'uppercase',
          letterSpacing: '.1em', fontWeight: 700 }}>{label}</span>
      </div>
      <div style={{ fontSize: 32, fontWeight: 900, color: accent, lineHeight: 1 }}>{value}</div>
      {sub  && <div style={{ fontSize: 11, color: C.sub }}>{sub}</div>}
      {note && <div style={{ fontSize: 10, color: C.muted, fontStyle: 'italic',
        borderTop: `1px solid ${C.border}`, paddingTop: 5, marginTop: 2 }}>{note}</div>}
    </div>
  )
}

// Donut SVG
function Donut({ pct, size = 130 }: { pct: number; size?: number }) {
  const r   = 46, cx = 60, cy = 60
  const circ = 2 * Math.PI * r
  const dash = (pct / 100) * circ
  const color = slaColor(pct)
  return (
    <svg width={size} height={size} viewBox="0 0 120 120">
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={C.border} strokeWidth="14" />
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={color} strokeWidth="14"
        strokeDasharray={`${dash} ${circ - dash}`}
        strokeDashoffset={circ / 4}
        strokeLinecap="round" />
      <text x="60" y="54" textAnchor="middle" fontSize="20" fontWeight="900"
        fill={color}>{pct.toFixed(1)}%</text>
      <text x="60" y="70" textAnchor="middle" fontSize="9" fill={C.muted}>SLA</text>
    </svg>
  )
}

// Barra de progresso vs meta
function ProgressBar({ pct }: { pct: number }) {
  const color = slaColor(pct)
  const metaPct = META
  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6, fontSize: 11, color: C.muted }}>
        <span>0%</span>
        <span style={{ color: C.accent }}>← Meta {META}%</span>
        <span>100%</span>
      </div>
      <div style={{ position: 'relative', height: 14, background: C.border, borderRadius: 99, overflow: 'hidden' }}>
        <div style={{ height: '100%', borderRadius: 99, background: color,
          width: `${Math.min(pct, 100)}%`, transition: 'width .8s ease' }} />
      </div>
      {/* Meta marker */}
      <div style={{ position: 'relative', height: 0 }}>
        <div style={{ position: 'absolute', left: `${metaPct}%`, top: -14,
          width: 2, height: 14, background: C.accent, opacity: .8 }} />
      </div>
    </div>
  )
}

// Tabela de unidades
function UnidadeTable({ rows, showUF }: { rows: UnidadeRow[]; showUF: boolean }) {
  return (
    <div style={{ overflowY: 'auto', maxHeight: 460 }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead style={{ position: 'sticky', top: 0, background: C.card }}>
          <tr>
            {['#', showUF && 'UF', 'Unidade', 'Previstas', 'Realizadas', 'Pendentes', 'SLA', 'Status'].filter(Boolean).map(h => (
              <th key={h} style={{ padding: '8px 10px', textAlign: 'left',
                borderBottom: `1px solid ${C.border}`, color: C.muted,
                fontWeight: 700, fontSize: 10, textTransform: 'uppercase',
                letterSpacing: '.07em', whiteSpace: 'nowrap' }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const color = slaColor(r.sla)
            return (
              <tr key={`${r.unidade}${i}`}
                style={{ borderBottom: `1px solid ${C.border}`, transition: 'background .1s' }}
                onMouseEnter={e => e.currentTarget.style.background = '#0e1b2c'}
                onMouseLeave={e => e.currentTarget.style.background = 'transparent'}>
                <td style={{ padding: '7px 10px', color: C.muted, fontSize: 10 }}>{i+1}</td>
                {showUF && <td style={{ padding: '7px 10px', color: C.sub, fontWeight: 700, fontSize: 11 }}>{r.uf}</td>}
                <td style={{ padding: '7px 10px', fontWeight: 500, color: C.text,
                  maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {r.unidade}
                </td>
                <td style={{ padding: '7px 10px', color: C.sub, textAlign: 'center' }}>{r.previstas}</td>
                <td style={{ padding: '7px 10px', color: C.green, fontWeight: 700, textAlign: 'center' }}>{r.realizadas}</td>
                <td style={{ padding: '7px 10px', color: C.red, textAlign: 'center' }}>{r.pendentes}</td>
                <td style={{ padding: '7px 10px', fontWeight: 900, color, textAlign: 'center' }}>{r.sla.toFixed(1)}%</td>
                <td style={{ padding: '7px 10px' }}>
                  <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 8px',
                    borderRadius: 99, background: color + '22', color, whiteSpace: 'nowrap' }}>
                    {slaLabel(r.sla)}
                  </span>
                </td>
              </tr>
            )
          })}
          {!rows.length && (
            <tr><td colSpan={8} style={{ padding: 16, color: C.muted, fontSize: 13 }}>
              Nenhuma unidade com dados para este filtro.
            </td></tr>
          )}
        </tbody>
      </table>
    </div>
  )
}

// ─── Main ─────────────────────────────────────────────────
export default function VisitasPage() {
  // snapshots: { date: 'YYYY-MM-DD', previstas: Row[], pendentes: Row[] }[]
  const [snapshots,  setSnapshots]  = useState<Snapshot[]>([])
  const [loading,    setLoading]    = useState(false)
  const [período,    setPeriodo]    = useState('HOJE')
  const [ufFiltro,   setUfFiltro]   = useState('TODOS')
  const [abaSel,     setAbaSel]     = useState('unidades') // unidades | criticas | ufs
  const [uploadInfo, setUploadInfo] = useState('')

  // ── Upload: aceita múltiplos arquivos
  const handleFiles = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[]
    if (!files.length) return
    setLoading(true)
    try {
      const newSnaps: Record<string, Snapshot> = {}

      for (const file of files) {
        // Usa só o nome base do arquivo (sem path)
        const fname = file.name.split('/').pop() || file.name
        const date  = extractDateFromName(fname)

        if (!date) {
          console.warn('Data não encontrada no nome:', fname)
          continue
        }

        const isPrev = isPrevistas(fname)
        const rows   = await parseFile(file)

        console.log(`[Visitas] ${fname} → date=${date} isPrev=${isPrev} rows=${rows.length}`)

        if (!newSnaps[date]) newSnaps[date] = { date, previstas: [], pendentes: [] }
        if (isPrev) newSnaps[date].previstas = rows
        else        newSnaps[date].pendentes = rows
      }

      // Mescla com snapshots anteriores
      setSnapshots(prev => {
        const merged: Record<string, Snapshot> = prev.reduce((a, s) => ({ ...a, [s.date]: s }), {} as Record<string, Snapshot>)
        Object.entries(newSnaps).forEach(([date, snap]) => {
          if (!merged[date]) merged[date] = snap
          else {
            if (snap.previstas.length) merged[date].previstas = snap.previstas
            if (snap.pendentes.length) merged[date].pendentes = snap.pendentes
          }
        })
        return Object.values(merged).sort((a, b) => a.date.localeCompare(b.date))
      })

      const datesLoaded = [...new Set(Object.keys(newSnaps).map(d => fmtDate(d)).filter(Boolean))].join(', ')
      setUploadInfo(`${files.length} arquivo(s) · datas: ${datesLoaded || '(nenhuma data detectada)'}`)
    } finally { setLoading(false) }
  }, [])

  // ── Snapshots filtrados pelo período selecionado
  const filteredSnaps = useMemo(() =>
    snapshots.filter(s => filterByPeriodo(s.date, período)),
    [snapshots, período])

  // ── Combina todos os dias filtrados em uma visão única
  // Se Hoje: só hoje. Se Semana: agrega todos os dias da semana.
  const agregado = useMemo(() => {
    if (!filteredSnaps.length) return { previstas: [], pendentes: [], dates: [] }

    // Para cada unidade, soma previstas e pendentes de todos os dias filtrados
    const prevMap: Record<string, { uf: string; unidade: string; count: number }> = {}
    const pendMap: Record<string, { uf: string; unidade: string; count: number }> = {}

    for (const snap of filteredSnaps) {
      for (const r of snap.previstas) {
        if (ufFiltro !== 'TODOS' && r.UF !== ufFiltro) continue
        const key = `${r.UF}||${r.UNIDADE}`
        if (!prevMap[key]) prevMap[key] = { uf: r.UF, unidade: r.UNIDADE, count: 0 }
        prevMap[key].count++
      }
      for (const r of snap.pendentes) {
        if (ufFiltro !== 'TODOS' && r.UF !== ufFiltro) continue
        const key = `${r.UF}||${r.UNIDADE}`
        if (!pendMap[key]) pendMap[key] = { uf: r.UF, unidade: r.UNIDADE, count: 0 }
        pendMap[key].count++
      }
    }

    return {
      prevMap,
      pendMap,
      dates: filteredSnaps.map(s => s.date),
      totalPrev: Object.values(prevMap).reduce((a, x) => a + x.count, 0),
      totalPend: Object.values(pendMap).reduce((a, x) => a + x.count, 0),
    }
  }, [filteredSnaps, ufFiltro])

  // ── Métricas gerais
  const totalPrevistas  = agregado.totalPrev  || 0
  const totalPendentes  = agregado.totalPend  || 0
  const totalRealizadas = Math.max(0, totalPrevistas - totalPendentes)
  const slaGeral        = totalPrevistas > 0 ? (totalRealizadas / totalPrevistas) * 100 : 0
  const faltamMeta      = Math.max(0, Math.ceil(totalPrevistas * META / 100) - totalRealizadas)

  // ── Por unidade
  const unidadeRows = useMemo(() => {
    const { prevMap = {}, pendMap = {} } = agregado
    const all = Object.entries(prevMap).map(([key, p]) => {
      const pend = pendMap[key]?.count || 0
      const real = Math.max(0, p.count - pend)
      const sla  = p.count > 0 ? (real / p.count) * 100 : 0
      return { uf: p.uf, unidade: p.unidade, previstas: p.count, pendentes: pend, realizadas: real, sla }
    })
    return all.sort((a, b) => a.sla - b.sla)
  }, [agregado])

  // Críticas (<META), em risco (70–META), ok (>=META)
  const criticas  = unidadeRows.filter(r => r.sla <  70)
  const emRisco   = unidadeRows.filter(r => r.sla >= 70 && r.sla < META)
  const ok        = unidadeRows.filter(r => r.sla >= META)

  // ── Por UF
  const ufRows = useMemo(() => {
    const m: Record<string, { uf: string; previstas: number; pendentes: number; realizadas: number }> = {}
    for (const r of unidadeRows) {
      if (!m[r.uf]) m[r.uf] = { uf: r.uf, previstas: 0, pendentes: 0, realizadas: 0 }
      m[r.uf].previstas  += r.previstas
      m[r.uf].pendentes  += r.pendentes
      m[r.uf].realizadas += r.realizadas
    }
    return Object.values(m).map(r => ({
      ...r, sla: r.previstas > 0 ? (r.realizadas / r.previstas) * 100 : 0
    })).sort((a, b) => a.sla - b.sla)
  }, [unidadeRows])

  // ── UFs disponíveis
  const ufsDisp = useMemo(() => {
    const s = new Set(snapshots.flatMap(s => [...s.previstas, ...s.pendentes].map(r => r.UF)).filter(u => u && u.trim()))
    return [...s].sort()
  }, [snapshots])

  // ── Label do período
  const períodoLabel = useMemo(() => {
    const dates = filteredSnaps.map(s => s.date).sort()
    if (!dates.length) return 'sem dados'
    if (dates.length === 1) return fmtDate(dates[0])
    return `${fmtDate(dates[0])} → ${fmtDate(dates[dates.length-1])}`
  }, [filteredSnaps])

  const hasData       = snapshots.length > 0
  const hasFiltered   = filteredSnaps.length > 0

  return (
    <div style={{ background: C.bg, minHeight: '100vh',
      fontFamily: "'DM Sans','Segoe UI',sans-serif", color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;background:${C.bg}}
        ::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
        select option{background:${C.surface};color:${C.text}}
        input::placeholder{color:${C.muted}}
        .ubtn:hover{opacity:.9;transform:translateY(-1px)}
        .tab:hover{background:${C.border}!important}
      `}</style>

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

        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          {/* Período */}
          <div style={{ display: 'flex', gap: 3, background: C.card,
            border: `1px solid ${C.border}`, borderRadius: 10, padding: 4 }}>
            {PERIODOS.map(p => (
              <button key={p.key} onClick={() => setPeriodo(p.key)} style={{
                padding: '5px 12px', borderRadius: 7, border: 'none',
                fontSize: 12, fontWeight: 700, cursor: 'pointer', transition: 'all .15s',
                background: período === p.key
                  ? `linear-gradient(135deg,${C.accent},${C.accentB})` : 'transparent',
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

          {/* Redefinir */}
          {(ufFiltro !== 'TODOS' || período !== 'HOJE') && (
            <button onClick={() => { setUfFiltro('TODOS'); setPeriodo('HOJE') }}
              style={{ background: 'transparent', border: `1px solid ${C.border}`, borderRadius: 9,
                color: C.muted, fontSize: 12, fontWeight: 700, padding: '7px 13px', cursor: 'pointer' }}>
              ✕ Redefinir
            </button>
          )}

          {/* Upload */}
          <label className="ubtn" style={{
            background: `linear-gradient(135deg,${C.accent},${C.accentB})`,
            color: '#000', fontWeight: 700, fontSize: 13,
            padding: '8px 18px', borderRadius: 9, cursor: 'pointer', transition: 'all .2s' }}>
            {loading ? 'Lendo…' : '+ Carregar Planilhas'}
            <input type="file" accept=".xlsx,.xls,.csv" multiple
              style={{ display: 'none' }} onChange={handleFiles} />
          </label>
        </div>
      </div>

      <div style={{ padding: '22px 36px' }}>
        {!hasData && (
          <div style={{ minHeight: 'calc(100vh - 110px)', display: 'flex',
            flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 16 }}>
            <div style={{ fontSize: 52 }}>📋</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>Nenhuma planilha carregada</div>
            <div style={{ color: C.muted, fontSize: 13, textAlign: 'center', maxWidth: 480, lineHeight: 1.8 }}>
              Carregue os arquivos <strong style={{ color: C.sub }}>07H</strong> (previstas) e{' '}
              <strong style={{ color: C.sub }}>12H</strong> (pendentes) de um ou mais dias.<br/>
              O sistema detecta a data automaticamente pelo nome do arquivo.<br/>
              <span style={{ color: C.accent }}>Exemplo: VisitasMedicas_07H_12_05.xlsx</span>
            </div>
          </div>
        )}

        {hasData && !hasFiltered && (
          <div style={{ minHeight: 'calc(100vh - 110px)', display: 'flex',
            flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 14 }}>
            <div style={{ fontSize: 48 }}>🔍</div>
            <div style={{ fontSize: 18, fontWeight: 700 }}>Sem dados para este período</div>
            <div style={{ color: C.muted, fontSize: 13 }}>
              Os arquivos carregados cobrem:{' '}
              <strong style={{ color: C.sub }}>
                {snapshots.map(s => fmtDate(s.date)).join(', ')}
              </strong>
            </div>
            <button onClick={() => setPeriodo('TODOS')} style={{
              background: `linear-gradient(135deg,${C.accent},${C.accentB})`,
              color: '#000', fontWeight: 700, fontSize: 13,
              padding: '9px 20px', borderRadius: 9, border: 'none', cursor: 'pointer',
            }}>Ver Todos os Dados</button>
          </div>
        )}

        {hasData && hasFiltered && (<>
          {/* ── Hero: SLA Geral ── */}
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
                    <span style={{ fontSize: 11, fontWeight: 700, padding: '4px 10px',
                      borderRadius: 99, background: slaColor(slaGeral)+'22',
                      color: slaColor(slaGeral) }}>{slaLabel(slaGeral)}</span>
                  </>
                ) : (
                  <span style={{ fontSize: 12, background: C.green+'22', color: C.green,
                    padding: '4px 12px', borderRadius: 99, fontWeight: 700 }}>
                    ✓ Meta {META}% atingida!
                  </span>
                )}
              </div>
            </div>

            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                letterSpacing: '.12em', color: C.muted, marginBottom: 12 }}>
                Progresso vs Meta {META}%
              </div>
              <ProgressBar pct={slaGeral} />
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 10, marginTop: 18 }}>
                {[
                  { label: `OK ≥ ${META}%`, value: ok.length,      color: C.green  },
                  { label: 'Em Risco',       value: emRisco.length, color: C.yellow },
                  { label: 'Críticas',       value: criticas.length,color: C.red    },
                ].map(k => (
                  <div key={k.label} style={{ textAlign: 'center', background: k.color+'11',
                    border: `1px solid ${k.color}33`, borderRadius: 10, padding: '10px 8px' }}>
                    <div style={{ fontSize: 24, fontWeight: 900, color: k.color }}>{k.value}</div>
                    <div style={{ fontSize: 9.5, color: C.muted, textTransform: 'uppercase',
                      letterSpacing: '.08em', marginTop: 4 }}>{k.label}</div>
                    <div style={{ fontSize: 9, color: C.muted }}>unidades</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* ── KPIs ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5,1fr)', gap: 12, marginBottom: 18 }}>
            <KpiCard icon="📋" label="Visitas Previstas (07H)"
              value={totalPrevistas.toLocaleString('pt-BR')}
              sub={`${unidadeRows.length} unidades · ${ufFiltro === 'TODOS' ? 'todas as UFs' : ufFiltro}`}
              accent={C.accent} />
            <KpiCard icon="✅" label="Realizadas"
              value={totalRealizadas.toLocaleString('pt-BR')}
              sub={`${slaGeral.toFixed(1)}% do total`} accent={C.green} />
            <KpiCard icon="⏳" label={`Pendentes (12H)`}
              value={totalPendentes.toLocaleString('pt-BR')}
              sub={totalPrevistas > 0 ? `${(100-slaGeral).toFixed(1)}% não visitados` : '—'}
              accent={C.red}
              note={totalPendentes === 0 && totalPrevistas > 0 ? 'Arquivo 12H ainda não carregado para este dia' : undefined} />
            <KpiCard icon="🏆" label="Unidades Atingiram Meta"
              value={`${ok.length}/${unidadeRows.length}`}
              sub={unidadeRows.length > 0 ? `${((ok.length/unidadeRows.length)*100).toFixed(0)}% das unidades` : 'NaN%'}
              accent={C.green} />
            <KpiCard icon="🎯" label="Faltam para 90%"
              value={faltamMeta > 0 ? faltamMeta.toLocaleString('pt-BR') : '—'}
              sub={faltamMeta > 0 ? 'visitas ainda necessárias' : 'Meta já atingida!'}
              accent={faltamMeta > 0 ? C.orange : C.green} />
          </div>

          {/* ── Abas ── */}
          <div style={{ display: 'flex', gap: 8, marginBottom: 14 }}>
            {[
              { key: 'unidades', label: '📋 Por Unidade' },
              { key: 'criticas', label: '🚨 Postos Críticos' },
              { key: 'ufs',      label: '📍 Por UF' },
            ].map(t => (
              <button key={t.key} onClick={() => setAbaSel(t.key)} className="tab" style={{
                padding: '9px 18px', borderRadius: 10,
                border: `1px solid ${abaSel === t.key ? C.accent : C.border}`,
                background: abaSel === t.key ? C.accent+'22' : 'transparent',
                color: abaSel === t.key ? C.accent : C.sub,
                fontSize: 13, fontWeight: 700, cursor: 'pointer', transition: 'all .15s',
              }}>{t.label}</button>
            ))}
          </div>

          {/* ── Conteúdo da aba ── */}
          <Card>
            {abaSel === 'unidades' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                  letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  Todas as Unidades — ordenadas por SLA (pior → melhor)
                </div>
                <UnidadeTable rows={unidadeRows} showUF={ufFiltro === 'TODOS'} />
              </>
            )}
            {abaSel === 'criticas' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                  letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  🚨 Postos Críticos (SLA &lt; {META}%) — {criticas.length + emRisco.length} unidades
                </div>
                <UnidadeTable rows={[...criticas, ...emRisco]} showUF={ufFiltro === 'TODOS'} />
              </>
            )}
            {abaSel === 'ufs' && (
              <>
                <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase',
                  letterSpacing: '.12em', color: C.muted, marginBottom: 14 }}>
                  SLA por Estado (UF)
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 10 }}>
                  {ufRows.map(r => {
                    const color = slaColor(r.sla)
                    const pct   = r.previstas > 0 ? (r.realizadas / r.previstas) * 100 : 0
                    return (
                      <div key={r.uf} style={{ background: C.card2,
                        border: `1px solid ${color}33`, borderLeft: `4px solid ${color}`,
                        borderRadius: 10, padding: '14px 16px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between',
                          alignItems: 'center', marginBottom: 10 }}>
                          <span style={{ fontSize: 16, fontWeight: 900, color: C.text }}>{r.uf}</span>
                          <span style={{ fontSize: 18, fontWeight: 900, color }}>{r.sla.toFixed(1)}%</span>
                        </div>
                        <div style={{ background: C.border, borderRadius: 99, height: 6, overflow: 'hidden', marginBottom: 8 }}>
                          <div style={{ height: '100%', borderRadius: 99, background: color,
                            width: `${Math.min(pct, 100)}%` }} />
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: C.muted }}>
                          <span>✅ {r.realizadas}</span>
                          <span>📋 {r.previstas}</span>
                          <span style={{ color: C.red }}>⏳ {r.pendentes}</span>
                        </div>
                      </div>
                    )
                  })}
                </div>
              </>
            )}
          </Card>

          {/* Footer */}
          <div style={{ textAlign: 'center', color: C.muted, fontSize: 11, paddingTop: 16, paddingBottom: 20 }}>
            Visitas Médicas SLA Monitor · Meta {META}% · {períodoLabel}
            {ufFiltro !== 'TODOS' && ` · UF: ${ufFiltro}`}
          </div>
        </>)}
      </div>
    </div>
  )
}
