import React from 'react'
import { createRoot } from 'react-dom/client'
import './style.css'

function App() {
  return <div className="layout"><aside><div className="brand"><span className="mark">A</span><span>Auditor?a<br/><small>DOCUMENTAL</small></span></div><p className="nav-label">ESPACIO DE TRABAJO</p><nav aria-label="Navegaci?n principal"><a href="#expedientes" aria-current="page">? <span>Expedientes</span></a></nav><div className="aside-note">Uso interno<br/><strong>Revisi?n asistida</strong></div></aside><main id="expedientes"><header><div><p className="eyebrow">?REA DE TRABAJO / EXPEDIENTES</p><h1>Expedientes</h1><p className="intro">Organice y revise las auditor?as documentales en un solo lugar.</p></div><span className="badge">Base inicial</span></header><section className="panel"><div className="panel-head"><div><h2>Auditor?as</h2><p>Expedientes y revisiones</p></div></div><div className="empty"><div className="empty-icon" aria-hidden="true">?</div><h3>Su espacio de auditor?a</h3><p>La gesti?n de expedientes estar? disponible en una pr?xima etapa. Aqu? podr? consultar documentos, hallazgos y revisiones.</p></div></section><p className="footnote">La revisi?n profesional es necesaria para validar los resultados.</p></main></div>
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>)
