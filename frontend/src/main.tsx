import { useEffect, useId, useRef, useState } from 'react'
import './style.css'

type IconName = 'file' | 'upload' | 'check' | 'arrow' | 'close' | 'clock'

function Icon({ name }: { name: IconName }) {
  const paths = {
    file: 'M13 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V10z M13 3v7h7 M8 15h8 M8 18h5',
    upload: 'M12 16V4 M7 9l5-5 5 5 M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3',
    check: 'M5 12l4 4L19 6',
    arrow: 'M5 12h14 M13 6l6 6-6 6',
    close: 'M6 6l12 12 M18 6L6 18',
    clock: 'M12 6v6l4 2 M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0z',
  }
  return <svg aria-hidden="true" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><path d={paths[name]} /></svg>
}

function formatSize(bytes: number) {
  return bytes >= 1024 * 1024 ? `${(bytes / (1024 * 1024)).toLocaleString('es', { maximumFractionDigits: 1 })} MB` : `${Math.ceil(bytes / 1024)} KB`
}

function formatDate(iso: string) {
  if (!iso) return ''
  try {
    return new Date(iso).toLocaleString('es', { dateStyle: 'short', timeStyle: 'short' })
  } catch {
    return iso
  }
}

function Picker({ label, hint, formats, files, onChange, multiple = false }: {
  label: string; hint: string; formats: string[]; files: File[]; onChange: (files: File[]) => void; multiple?: boolean
}) {
  const input = useRef<HTMLInputElement>(null)
  const id = useId()
  const [error, setError] = useState('')
  const [dragging, setDragging] = useState(false)

  function add(incoming: FileList | null) {
    if (!incoming?.length) return
    const valid = Array.from(incoming).filter(file => formats.some(ext => file.name.toLowerCase().endsWith(ext)))
    setError(!multiple && incoming.length > 1 ? 'Selecciona un solo archivo.' : valid.length < incoming.length ? `Solo se admiten archivos ${formats.join(', ')}.` : '')
    if (!valid.length || (!multiple && incoming.length > 1)) return
    onChange(multiple ? [...files, ...valid] : valid)
  }

  return <div className="picker">
    <input ref={input} type="file" hidden accept={formats.join(',')} multiple={multiple} onChange={e => { add(e.currentTarget.files); e.currentTarget.value = '' }} />
    <button className={`dropzone${dragging ? ' dragging' : ''}`} type="button" aria-label={label} aria-describedby={id}
      onClick={() => input.current?.click()} onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)} onDrop={e => { e.preventDefault(); setDragging(false); add(e.dataTransfer.files) }}>
      <span className="upload-icon"><Icon name="upload" /></span>
      <span><strong>{label}</strong><small id={id}>{hint}</small></span>
    </button>
    {error && <p className="error" role="alert">{error}</p>}
    {files.length > 0 && <ul className="files">{files.map((file, index) => <li key={`${file.name}-${index}`}>
      <span className="file-icon"><Icon name="file" /></span>
      <div><strong>{file.name}</strong><small>{formatSize(file.size)}</small></div>
      <button className="remove" type="button" aria-label={`Quitar ${file.name}`} onClick={() => onChange(files.filter((_, i) => i !== index))}><Icon name="close" /></button>
    </li>)}</ul>}
  </div>
}

type FindingItem = {
  id: string
  severity: 'critico' | 'importante' | 'menor'
  review_type: string
  location: string
  sheet?: string | null
  cell?: string | null
  page?: number | null
  field_reviewed?: string
  observation: string
  technical_explanation?: string
  current_value?: string | null
  expected_value?: string | null
  difference?: string | null
  evidence: string
  recommendation: string
  value_impact: string
  status: 'sustentado' | 'requiere_verificacion' | 'diferencia_explicada' | 'descartado'
  verification_note?: string | null
}

type ReviewResult = {
  case_id: string
  mode: 'excel' | 'pdf'
  created_at?: string
  duration_seconds?: number
  model?: string
  category_detected?: string
  summary?: string
  verified_independently?: boolean
  verification_summary?: string
  findings: number
  items?: FindingItem[]
  discarded?: FindingItem[]
  counts?: { critico?: number; importante?: number; menor?: number; sustentado?: number; requiere_verificacion?: number; descartado?: number }
  downloads: { word: string; excel?: string }
  coverage: string[]
  excel_warning?: string | null
  marked_cells?: number
}

type HistoryEntry = {
  case_id: string
  mode: 'excel' | 'pdf'
  created_at: string
  principal_label: string
  category_detected: string
  findings: number
  counts?: { critico?: number; importante?: number; menor?: number }
  verified_independently?: boolean
}

const STAGES = [
  { key: 'indexing', label: 'Indexación' },
  { key: 'analyzing', label: 'Análisis' },
  { key: 'verifying', label: 'Verificación' },
  { key: 'exporting', label: 'Entregables' },
] as const

const SEV_LABEL: Record<string, string> = { critico: 'Crítico', importante: 'Importante', menor: 'Menor' }
const STATUS_LABEL: Record<string, string> = {
  sustentado: 'Sustentado',
  requiere_verificacion: 'Por verificar',
  diferencia_explicada: 'Explicada',
  descartado: 'Descartado',
}
const IMPACT_LABEL: Record<string, string> = {
  demostrado: 'Impacto demostrado',
  posible: 'Impacto posible',
  ninguno_demostrado: 'Sin impacto demostrado',
  no_determinable: 'Impacto no determinable',
}

function stageIndex(stage: string) {
  const idx = STAGES.findIndex(s => s.key === stage)
  return idx === -1 ? 0 : idx
}

export default function App() {
  const [excel, setExcel] = useState<File[]>([])
  const [supports, setSupports] = useState<File[]>([])
  const [pdf, setPdf] = useState<File[]>([])
  const [mode, setMode] = useState<'excel' | 'pdf'>('excel')
  const [reviewing, setReviewing] = useState(false)
  const [jobId, setJobId] = useState<string | null>(null)
  const [stage, setStage] = useState('queued')
  const [statusMessage, setStatusMessage] = useState('')
  const [reviewError, setReviewError] = useState('')
  const [reviewResult, setReviewResult] = useState<ReviewResult | null>(null)
  const [filter, setFilter] = useState<string>('all')
  const [history, setHistory] = useState<HistoryEntry[]>([])
  const [showHistory, setShowHistory] = useState(false)
  const esRef = useRef<EventSource | null>(null)
  const principal = (mode === 'excel' ? excel : pdf)[0]

  async function refreshHistory() {
    try {
      const res = await fetch('/api/reviews/history')
      if (res.ok) {
        const data = await res.json()
        if (Array.isArray(data.history)) setHistory(data.history)
      }
    } catch { /* silent */ }
  }

  useEffect(() => {
    refreshHistory()
    return () => { esRef.current?.close() }
  }, [])

  function closeStream() {
    if (esRef.current) {
      esRef.current.close()
      esRef.current = null
    }
  }

  async function pollUntilDone(id: string) {
    while (true) {
      await new Promise(r => setTimeout(r, 1500))
      const res = await fetch(`/api/reviews/jobs/${id}`)
      if (!res.ok) throw new Error('No se pudo consultar el estado de la revisión.')
      const snap = await res.json()
      setStage(snap.stage || 'analyzing')
      if (snap.message) setStatusMessage(snap.message)
      if (snap.status === 'completed' && snap.result) {
        setReviewResult(snap.result)
        setFilter('all')
        setReviewing(false)
        setJobId(null)
        refreshHistory()
        return
      }
      if (snap.status === 'failed' || snap.status === 'cancelled') {
        setReviewError(snap.error || snap.message || 'La revisión no pudo completarse.')
        setReviewing(false)
        setJobId(null)
        return
      }
    }
  }

  async function reviewFiles() {
    if (!principal || reviewing) return
    closeStream()
    const form = new FormData()
    form.append('principal', principal)
    if (mode === 'excel') supports.forEach(file => form.append('respaldo', file))
    setReviewing(true)
    setReviewError('')
    setReviewResult(null)
    setStage('indexing')
    setStatusMessage('Cargando expediente y preparando análisis…')

    try {
      const response = await fetch(`/api/reviews/jobs/${mode}`, {
        method: 'POST', headers: { 'X-Avaluos-Review': 'local' }, body: form,
      })
      const started = await response.json()
      if (!response.ok || !started.job_id) {
        throw new Error(typeof started.detail === 'string' ? started.detail : 'No se pudo iniciar la revisión.')
      }
      const id = started.job_id as string
      setJobId(id)

      const es = new EventSource(`/api/reviews/jobs/${id}/events`)
      esRef.current = es
      es.addEventListener('progress', (event) => {
        const ev = JSON.parse((event as MessageEvent).data)
        if (ev.stage) setStage(ev.stage)
        if (ev.message) setStatusMessage(ev.message)
      })
      es.addEventListener('done', (event) => {
        const done = JSON.parse((event as MessageEvent).data)
        closeStream()
        setReviewing(false)
        setJobId(null)
        if (done.status === 'completed' && done.result) {
          setReviewResult(done.result)
          setFilter('all')
          refreshHistory()
        } else {
          setReviewError(done.error || done.message || 'La revisión no se completó.')
        }
      })
      es.onerror = () => {
        closeStream()
        pollUntilDone(id).catch((err) => {
          setReviewError((err as Error).message || 'Se perdió la conexión con el servidor local.')
          setReviewing(false)
          setJobId(null)
        })
      }
    } catch (error) {
      setReviewError(error instanceof SyntaxError || error instanceof TypeError
        ? 'No se pudo conectar con el servidor local.' : (error as Error).message)
      setReviewing(false)
      setJobId(null)
    }
  }

  async function cancelReview() {
    if (!jobId) return
    setStatusMessage('Cancelando revisión…')
    try {
      await fetch(`/api/reviews/jobs/${jobId}/cancel`, {
        method: 'POST', headers: { 'X-Avaluos-Review': 'local' },
      })
    } catch { /* handled by stream */ }
  }

  async function openHistoryCase(caseId: string) {
    setReviewError('')
    try {
      const res = await fetch(`/api/reviews/${caseId}/result`)
      if (!res.ok) throw new Error('No se encontró el resultado guardado.')
      const data = await res.json()
      setReviewResult(data)
      setFilter('all')
      setShowHistory(false)
    } catch (err) {
      setReviewError((err as Error).message)
    }
  }

  function switchMode(next: 'excel' | 'pdf') {
    if (reviewing) return
    setMode(next)
    setReviewError('')
  }

  const activeItems = reviewResult?.items ?? []
  const filteredItems = activeItems.filter(item => {
    if (filter === 'all') return true
    if (filter === 'sustentado' || filter === 'requiere_verificacion') return item.status === filter
    return item.severity === filter
  })
  const curStageIdx = stageIndex(stage)

  return <>
    <header className="topbar"><div className="topbar-inner">
      <div className="brand"><span className="brand-icon"><Icon name="file" /></span><span>Avalúos <strong>Revisión</strong></span></div>
      {history.length > 0 && <button type="button" className="ghost-btn" aria-expanded={showHistory} onClick={() => setShowHistory(v => !v)}>
        <Icon name="clock" /> Historial ({history.length})
      </button>}
    </div></header>

    <main>
      <div className="page-heading"><h1>Revisión de avalúos</h1></div>

      {showHistory && history.length > 0 && <section className="card history-panel" aria-label="Historial de revisiones">
        <div className="panel-head"><h2>Revisiones recientes</h2>
          <button type="button" className="remove" aria-label="Cerrar historial" onClick={() => setShowHistory(false)}><Icon name="close" /></button>
        </div>
        <ul className="history-list">{history.map(item => <li key={item.case_id}>
          <div>
            <strong>{item.principal_label}</strong>
            <small>{item.mode === 'pdf' ? 'PDF final' : 'Excel y respaldos'}{item.category_detected ? ` · ${item.category_detected}` : ''} · {formatDate(item.created_at)}</small>
          </div>
          <div className="history-meta">
            <span className="pill">{item.findings} obs.</span>
            <button type="button" className="ghost-btn" onClick={() => openHistoryCase(item.case_id)}>Abrir</button>
          </div>
        </li>)}</ul>
      </section>}

      <div className="switcher" role="group" aria-label="Tipo de revisión">
        <button type="button" disabled={reviewing} aria-pressed={mode === 'excel'} onClick={() => switchMode('excel')}>Avalúo y respaldos</button>
        <button type="button" disabled={reviewing} aria-pressed={mode === 'pdf'} onClick={() => switchMode('pdf')}>PDF final</button>
      </div>

      <div className="content">
        <section className="card" aria-label={mode === 'excel' ? 'Avalúo y respaldos' : 'PDF final'}>
          {mode === 'excel' ? <>
            <div className="step"><span className="step-number" aria-hidden="true">01</span><div><h3>Excel del avalúo <span className="tag required">Obligatorio</span></h3>
              <Picker key="excel" label={excel.length ? 'Reemplazar Excel' : 'Seleccionar Excel'} hint="Archivo .xlsx · cualquier tamaño o categoría" formats={['.xlsx']} files={excel} onChange={setExcel} />
            </div></div>
            <div className="step"><span className="step-number" aria-hidden="true">02</span><div><h3>Respaldos <span className="tag">Opcionales</span></h3>
              <Picker key="supports" label="Añadir respaldos" hint="PDF, DOCX, XLSX, PNG, JPG o WEBP · sin límite práctico de archivos" formats={['.pdf', '.docx', '.xlsx', '.png', '.jpg', '.jpeg', '.webp']} files={supports} onChange={setSupports} multiple />
            </div></div>
          </> : <>
            <div className="step"><span className="step-number" aria-hidden="true">01</span><div><h3>PDF final <span className="tag required">Obligatorio</span></h3>
              <Picker key="pdf" label={pdf.length ? 'Reemplazar PDF' : 'Seleccionar PDF'} hint="Archivo .pdf · texto e inspección visual de páginas" formats={['.pdf']} files={pdf} onChange={setPdf} />
            </div></div>
          </>}
          <footer className="card-footer">
            <p id="local-status" role="status">{reviewing ? statusMessage : 'Apoyo técnico profesional. Requiere validación del perito.'}</p>
            <div className="footer-actions">
              {reviewing && jobId && <button className="ghost-btn danger" type="button" onClick={cancelReview}>Cancelar</button>}
              <button className="action" type="button" disabled={!principal || reviewing} aria-describedby="local-status" onClick={reviewFiles}>
                {reviewing ? 'Revisando…' : mode === 'excel' ? 'Iniciar revisión' : 'Verificar PDF'} <Icon name="arrow" />
              </button>
            </div>
          </footer>
          {reviewing && <div className="progress-box" aria-live="polite">
            <ol className="stepper">{STAGES.map((s, idx) => <li key={s.key} className={idx < curStageIdx ? 'done' : idx === curStageIdx ? 'active' : ''}>
              <span>{idx + 1}</span>{s.label}
            </li>)}</ol>
          </div>}
          {reviewError && <p className="error" role="alert">{reviewError}</p>}
        </section>

        <aside className="guide" aria-labelledby="guide-heading">
          <h2 id="guide-heading">Entregables</h2>
          <ul>{(mode === 'excel'
            ? [
                'Excel marcado con notas por celda y hoja REVISIÓN TÉCNICA',
                'Informe Word técnico con evidencias y verificación independiente',
                'Contraste completo de libro, fórmulas y respaldos',
              ]
            : [
                'Informe Word por página con severidad y recomendaciones',
                'Revisión de texto, cálculos e inspección visual',
                'Control independiente del Excel',
              ]).map(text => <li key={text}><Icon name="check" />{text}</li>)}</ul>
        </aside>
      </div>

      {reviewResult && <section className="card result-card" aria-labelledby="result-heading">
        <div className="result-header">
          <div>
            <h2 id="result-heading">Resultados de la revisión</h2>
            <div className="result-badges">
              {reviewResult.category_detected && <span className="pill">{reviewResult.category_detected}</span>}
              <span className={`pill ${reviewResult.verified_independently ? 'ok' : ''}`}>
                {reviewResult.verified_independently ? 'Verificación independiente completada' : 'Sin segunda verificación'}
              </span>
              {reviewResult.duration_seconds !== undefined && <span className="pill">{reviewResult.duration_seconds}s</span>}
            </div>
          </div>
          <div className="downloads">
            <a href={reviewResult.downloads.word} className="action">Descargar informe Word</a>
            {reviewResult.downloads.excel && <a href={reviewResult.downloads.excel} className="action">Descargar Excel marcado</a>}
          </div>
        </div>

        {reviewResult.excel_warning && <p className="error" role="alert">{reviewResult.excel_warning}</p>}

        <div className="kpi-grid">
          <div className="kpi critico"><strong>{reviewResult.counts?.critico ?? 0}</strong><span>Críticos</span></div>
          <div className="kpi importante"><strong>{reviewResult.counts?.importante ?? 0}</strong><span>Importantes</span></div>
          <div className="kpi menor"><strong>{reviewResult.counts?.menor ?? 0}</strong><span>Menores</span></div>
          <div className="kpi ok"><strong>{reviewResult.counts?.sustentado ?? 0}</strong><span>Sustentados</span></div>
          <div className="kpi"><strong>{reviewResult.counts?.requiere_verificacion ?? 0}</strong><span>Por verificar</span></div>
        </div>

        {reviewResult.summary && <p className="summary-text">{reviewResult.summary}</p>}
        {reviewResult.verification_summary && <p className="verif-text"><strong>Verificación independiente:</strong> {reviewResult.verification_summary}</p>}

        {activeItems.length > 0 && <>
          <div className="filters" role="group" aria-label="Filtrar observaciones">
            {[
              ['all', `Todas (${activeItems.length})`],
              ['critico', `Críticos (${reviewResult.counts?.critico ?? 0})`],
              ['importante', `Importantes (${reviewResult.counts?.importante ?? 0})`],
              ['menor', `Menores (${reviewResult.counts?.menor ?? 0})`],
              ['sustentado', `Sustentados (${reviewResult.counts?.sustentado ?? 0})`],
              ['requiere_verificacion', `Por verificar (${reviewResult.counts?.requiere_verificacion ?? 0})`],
            ].map(([key, lbl]) => <button key={key} type="button" className={`filter-btn${filter === key ? ' active' : ''}`} onClick={() => setFilter(key)}>{lbl}</button>)}
          </div>

          <ul className="findings-list">{filteredItems.map((item, idx) => <li key={item.id || `${item.location}-${idx}`} className="finding-item">
            <div className="finding-top">
              <div className="finding-tags">
                <span className={`badge sev-${item.severity}`}>{SEV_LABEL[item.severity] || item.severity}</span>
                <span className={`badge st-${item.status}`}>{STATUS_LABEL[item.status] || item.status}</span>
                <span className="loc-tag">{item.location}</span>
                <span className="type-tag">{item.review_type}</span>
              </div>
              <span className="impact-tag">{IMPACT_LABEL[item.value_impact] || item.value_impact}</span>
            </div>
            <p className="finding-obs">{item.observation}</p>
            {(item.current_value || item.expected_value || item.difference) && <div className="value-compare">
              {item.current_value && <span><strong>Actual:</strong> {item.current_value}</span>}
              {item.expected_value && <span><strong>Referencia:</strong> {item.expected_value}</span>}
              {item.difference && <span><strong>Diferencia:</strong> {item.difference}</span>}
            </div>}
            <dl className="finding-meta">
              {item.technical_explanation && <><dt>Análisis</dt><dd>{item.technical_explanation}</dd></>}
              <dt>Evidencia</dt><dd>{item.evidence}</dd>
              <dt>Recomendación</dt><dd>{item.recommendation}</dd>
              {item.verification_note && <><dt>Verificación</dt><dd>{item.verification_note}</dd></>}
            </dl>
          </li>)}</ul>
        </>}

        {reviewResult.discarded && reviewResult.discarded.length > 0 && <details className="coverage">
          <summary>Descartadas o explicadas en verificación ({reviewResult.discarded.length})</summary>
          <ul>{reviewResult.discarded.map((d, idx) => <li key={`${d.location}-${idx}`}>
            <strong>[{STATUS_LABEL[d.status] || d.status}] {d.location}:</strong> {d.observation}{d.verification_note ? ` — ${d.verification_note}` : ''}
          </li>)}</ul>
        </details>}

        <details className="coverage"><summary>Cobertura y límites</summary>
          <ul>{reviewResult.coverage.map((item, idx) => <li key={`${idx}-${item}`}>{item}</li>)}</ul>
        </details>
      </section>}
    </main>
  </>
}
