import { useId, useRef, useState } from 'react'
import './style.css'

type IconName = 'file' | 'upload' | 'check' | 'arrow' | 'close'

function Icon({ name }: { name: IconName }) {
  const paths = {
    file: 'M13 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V10z M13 3v7h7 M8 15h8 M8 18h5',
    upload: 'M12 16V4 M7 9l5-5 5 5 M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3',
    check: 'M5 12l4 4L19 6',
    arrow: 'M5 12h14 M13 6l6 6-6 6',
    close: 'M6 6l12 12 M18 6L6 18',
  }
  return <svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.65" strokeLinecap="round" strokeLinejoin="round"><path d={paths[name]} /></svg>
}

function formatSize(bytes: number) {
  return bytes >= 1024 * 1024 ? `${(bytes / (1024 * 1024)).toLocaleString('es', { maximumFractionDigits: 1 })} MB` : `${Math.ceil(bytes / 1024)} KB`
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

type ReviewResult = {
  case_id: string
  findings: number
  downloads: { word: string; excel?: string }
  coverage: string[]
  excel_warning?: string
}

export default function App() {
  const [excel, setExcel] = useState<File[]>([])
  const [supports, setSupports] = useState<File[]>([])
  const [pdf, setPdf] = useState<File[]>([])
  const [mode, setMode] = useState<'excel' | 'pdf'>('excel')
  const [reviewing, setReviewing] = useState(false)
  const [reviewError, setReviewError] = useState('')
  const [reviewResult, setReviewResult] = useState<ReviewResult | null>(null)
  const principal = (mode === 'excel' ? excel : pdf)[0]

  async function reviewFiles() {
    if (!principal || reviewing) return
    const form = new FormData()
    form.append('principal', principal)
    if (mode === 'excel') supports.forEach(file => form.append('respaldo', file))
    setReviewing(true)
    setReviewError('')
    setReviewResult(null)
    try {
      const response = await fetch(`/api/reviews/${mode}`, {
        method: 'POST', headers: { 'X-Avaluos-Review': 'local' }, body: form,
      })
      const result = await response.json()
      if (!response.ok || !result.downloads?.word) {
        throw new Error(typeof result.detail === 'string' ? result.detail : 'No se pudo completar la revisión.')
      }
      setReviewResult(result)
    } catch (error) {
      setReviewError(error instanceof SyntaxError || error instanceof TypeError
        ? 'No se pudo conectar con el servidor local.' : (error as Error).message)
    } finally {
      setReviewing(false)
    }
  }

  function switchMode(next: 'excel' | 'pdf') {
    if (reviewing) return
    setMode(next)
    setReviewError('')
  }

  return <>
    <header className="topbar"><div className="topbar-inner">
      <div className="brand"><span className="brand-icon"><Icon name="file" /></span><span>Avalúos <strong>Revisión</strong></span></div>
    </div></header>
    <main>
      <div className="page-heading"><h1>Revisión de avalúos</h1></div>
      <div className="switcher" role="group" aria-label="Tipo de revisión">
        <button type="button" disabled={reviewing} aria-pressed={mode === 'excel'} onClick={() => switchMode('excel')}>Avalúo y respaldos</button>
        <button type="button" disabled={reviewing} aria-pressed={mode === 'pdf'} onClick={() => switchMode('pdf')}>PDF final</button>
      </div>
      <div className="content">
        <section className="card" aria-label={mode === 'excel' ? 'Avalúo y respaldos' : 'PDF final'}>
          {mode === 'excel' ? <>
            <div className="step"><span className="step-number" aria-hidden="true">01</span><div><h3>Excel del avalúo <span className="tag required">Obligatorio</span></h3>
              <Picker key="excel" label={excel.length ? 'Reemplazar Excel' : 'Seleccionar Excel'} hint="Archivo .xlsx" formats={['.xlsx']} files={excel} onChange={setExcel} />
            </div></div>
            <div className="step"><span className="step-number" aria-hidden="true">02</span><div><h3>Respaldos <span className="tag">Opcionales</span></h3>
              <Picker key="supports" label="Añadir respaldos" hint="PDF, DOCX o XLSX" formats={['.pdf', '.docx', '.xlsx']} files={supports} onChange={setSupports} multiple />
            </div></div>
          </> : <>
            <div className="step"><span className="step-number" aria-hidden="true">01</span><div><h3>PDF final <span className="tag required">Obligatorio</span></h3>
              <Picker key="pdf" label={pdf.length ? 'Reemplazar PDF' : 'Seleccionar PDF'} hint="Archivo .pdf" formats={['.pdf']} files={pdf} onChange={setPdf} />
            </div></div>
          </>}
          <footer className="card-footer">
            <p id="local-status" role="status">{reviewing ? 'Revisando documentos…' : 'Apoyo técnico. Requiere validación del perito.'}</p>
            <button className="action" type="button" disabled={!principal || reviewing} aria-describedby="local-status" onClick={reviewFiles}>
              {mode === 'excel' ? 'Iniciar revisión' : 'Verificar PDF'} <Icon name="arrow" />
            </button>
          </footer>
          {reviewError && <p className="error" role="alert">{reviewError}</p>}
        </section>
        <aside className="guide" aria-labelledby="guide-heading">
          <h2 id="guide-heading">Entregables</h2>
          <ul>{(mode === 'excel'
            ? ['Excel marcado con notas por celda', 'Informe Word con observaciones']
            : ['Informe Word por página', 'Control independiente del Excel']).map(text => <li key={text}><Icon name="check" />{text}</li>)}</ul>
        </aside>
      </div>
      {reviewResult && <section className="card result-card" aria-labelledby="result-heading">
        <div className="card-heading"><h2 id="result-heading">Resultados</h2>
          <p role="status">{reviewResult.findings} observaciones. Requieren validación del perito.</p></div>
        <div className="downloads">
          <a href={reviewResult.downloads.word} className="action">Descargar informe Word</a>
          {reviewResult.downloads.excel && <a href={reviewResult.downloads.excel} className="action">Descargar Excel marcado</a>}
        </div>
        {reviewResult.excel_warning && <p className="error" role="alert">{reviewResult.excel_warning}</p>}
        <details className="coverage"><summary>Cobertura</summary>
          <ul>{reviewResult.coverage.map(item => <li key={item}>{item}</li>)}</ul>
        </details>
      </section>}
    </main>
  </>
}
