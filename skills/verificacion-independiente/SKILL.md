---
name: verificacion-independiente
description: "Trigger: segunda comprobación independiente de hallazgos antes de exportar. Reabre las fuentes originales, descarta falsos positivos y detecta omisiones críticas."
license: Apache-2.0
metadata:
  author: "avaluosrev"
  version: "2.0"
---

## Activation Contract

Carga esta skill en una **segunda ejecución completamente separada** antes de emitir entregables del Excel o del PDF. Recibe los candidatos del primer agente y acceso MCP de solo lectura al expediente original completo.

## Hard Rules

- No aceptes ningún hallazgo del primer agente por confianza ni por buena redacción.
- Tampoco descartes un hallazgo real sin verificarlo primero directamente en las fuentes originales.
- No abras enlaces externos, no edites archivos y no inventes reglas bancarias no suministradas.
- La decisión final de aprobación o entrega corresponde siempre al perito revisor humano.

## Decision Gates

- **Fuente localizable y evidencia comprobada que sustenta la discrepancia, error de cálculo, omisión legal o error concreto de redacción:** confirma como `sustentado`.
- **Ubicación válida y señal real, pero falta documentación o criterio técnico del perito para confirmar error definitivo:** clasifica como `requiere_verificacion` con motivo preciso.
- **Otra hoja o fuente legítima explica la diferencia** (unidad, concepto distinto, fecha o redondeo documentado): clasifica como `diferencia_explicada` (no se marca en rojo en Excel, pero queda en trazabilidad).
- **Falso positivo, celda inexistente/oculta/bloque opcional vacío, o evidencia que desmiente la afirmación:** clasifica como `descartado` explicando el motivo.

## Execution Steps

1. Para **cada candidato** recibido, reabre la fuente original vía MCP (`leer_rango_excel`/`leer_hoja_excel`, `rastrear_celda_excel`, `leer_paginas_pdf`/`ver_pagina_pdf`, `leer_documento_docx`, `ver_imagen`). Verifica literalidad, unidades, fechas y fórmulas.
2. Repite los cálculos aritméticos sobre los operandos originales y comprueba con `rastrear_celda_excel` si la incidencia alegada sobre el valor comercial o de realización es realmente `demostrado`, `posible`, `ninguno_demostrado` o `no_determinable`.
3. En Excel, comprueba que cada celda a marcar sea visible, no vacía y aplicable (nunca oculta ni bloque opcional sin uso); si el origen está oculto, redirige al resultado visible afectado. Si una omisión legal no tiene celda destino clara, deja `sheet` y `cell` nulos para que conste solo en el informe Word.
4. **Control de cobertura contra revisiones superficiales:** revisa brevemente el atlas y las hojas/respaldos clave (identificación, áreas, gravámenes, estudio de mercado, valoración final y conclusiones). Si detectas un error crítico, discrepancia documental material u omisión legal importante que el primer agente pasó por alto, incorpóralo en `additional_findings` con evidencia directamente comprobada.
5. Consolida duplicados sin perder evidencias y verifica que la conclusión recomendada sea coherente con los hallazgos finales sustentados.

## Output Contract

Devuelve una decisión motivada por cada `finding_id`, `additional_findings` comprobados cuando el primer agente omitió hallazgos relevantes, `additional_limitations` y `verification_summary` según el [contrato común](../README.md).
