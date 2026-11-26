---
name: verificacion-independiente
description: "Trigger: segunda comprobación de hallazgos. Vuelve a las fuentes originales y decide qué observaciones están sustentadas."
license: Apache-2.0
metadata:
  author: "avaluosrev"
  version: "1.1"
---

## Activation Contract

Carga esta skill en una ejecución separada antes de emitir resultados del Excel o del PDF. Recibe los candidatos y acceso acotado a las fuentes originales pertinentes.

## Hard Rules

No aceptes una cita por estar bien formada. No confirmes por acuerdo con el primer agente. No amplíes el alcance, abras enlaces externos ni modifiques archivos. La decisión final corresponde al revisor humano.

## Decision Gates

- Fuente localizable y contenido pertinente que sustenta la afirmación: conserva como sustentada.
- Ubicación válida pero sustento insuficiente, ilegible o contexto ambiguo: devuelve como pendiente con motivo.
- Fuente desmiente la afirmación o explica la diferencia: descarta la marca y registra por qué.

## Execution Steps

1. Abre de nuevo cada celda, página o fragmento relevante; coteja identidad, fecha, concepto, unidades y literalidad.
2. Busca explicaciones en hojas y fuentes pertinentes. Repite cálculos con los datos originales y verifica si el impacto alegado se demuestra.
3. Consolida duplicados sin perder evidencias. En Excel, comprueba que el destino sea visible y aplicable (hoja, fila y columna). No autorices marcar celdas ocultas ni bloques sin uso; si el origen está oculto, indica el resultado visible afectado. Una omisión sin destino claro queda solo en el informe.

## Output Contract

Devuelve una decisión motivada por candidato y referencias comprobadas. No etiquetes una ejecución como verificada si no pudiste completar esta etapa.

## References

- [Contrato y flujo](../README.md)
