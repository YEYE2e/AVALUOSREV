---
name: revision-tecnica-excel
description: "Trigger: avalúo Excel, cálculos, mercado u homologación. Revisa la coherencia técnica del libro y rastrea su impacto."
license: Apache-2.0
metadata:
  author: "avaluosrev"
  version: "1.1"
---

## Activation Contract

Carga esta skill en la auditoría inicial del Excel, con o sin respaldos. Usa los hallazgos documentales como candidatos, no como hechos confirmados.

## Hard Rules

No edites el libro ni ejecutes macros o vínculos. No inventes factores, tolerancias, reglas bancarias ni metodologías. Distingue fórmula, valor guardado y cálculo independiente.
El objetivo es comprobar que los datos usados en el avalúo sean correctos, no solo que las fórmulas calculen. Usa el contraste documental para distinguir datos cotejados de datos que solo son coherentes dentro del libro.

## Decision Gates

- Diferencia explicada por unidad, fecha, redondeo documentado o característica: no marques por la sola diferencia.
- Factor o referencia atípica sin evidencia suficiente: solicita verificación; no elimines la referencia ni adoptes otro factor.
- Error de redacción concreto: propone el texto corregido; excluye preferencias estéticas.
- Hoja, fila o columna oculta: no la revises como objetivo ni la marques. Consulta una dependencia oculta solo para explicar un efecto en una celda visible y aplicable; cita el origen, pero señala el destino visible.
- Celda vacía o bloque opcional sin usar: no marques por estar vacío ni por un error guardado aislado. Si un dato obligatorio y aplicable falta, informa la omisión con evidencia de su obligatoriedad. Si un cálculo aplicable o resultado visible falla por esa ausencia, señala ese resultado y explica la causa.

## Execution Steps

1. Revisa coherencia entre hojas, identificación, áreas, características, ajustes, depreciación y conclusiones. Busca errores de copia y pega.
   Comprueba que cada dato material corresponda al inmueble, concepto, unidad y período correctos. Contrasta los valores de origen, no solo el resultado; si falta respaldo, limita la conclusión a la coherencia interna.
2. Recalcula precio de oferta / área si área y precio corresponden y el divisor es válido. Comprueba homologación según las fórmulas y factores disponibles; si no puedes reproducirla, explica por qué.
3. Sigue desde las referencias y el valor unitario adoptado hasta terreno, construcciones, complementos, ajustes y valor comercial o de realización, según aplique. Localiza la celda de origen; separa impacto demostrado, posible y no determinable.

## Output Contract

Devuelve candidatos y limitaciones. Incluye operandos, unidades, fórmula, resultado y diferencia cuando el cálculo sea reproducible. No presentes un total como recalculado si solo leíste su valor guardado.

## References

- [Contrato y flujo](../README.md)
