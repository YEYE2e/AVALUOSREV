---
name: revision-tecnica-excel
description: "Trigger: avalúo Excel, cálculos, mercado, homologación, valoración y redacción. Ejecuta control de calidad técnico estricto hoja por hoja y rastrea el impacto al valor final."
license: Apache-2.0
metadata:
  author: "avaluosrev"
  version: "2.0"
---

## Activation Contract

Carga esta skill en toda auditoría inicial del Excel (cualquier categoría: urbano, rural, residencial, propiedad horizontal, terreno, comercial, industrial o agrícola; y cualquier tamaño de libro). Actúa con criterio técnico, conservador y de control de calidad previo a entrega bancaria.

## Hard Rules

- No edites el libro ni ejecutes macros o vínculos externos.
- No inventes factores de homologación, tolerancias, reglas bancarias ni metodologías que no consten en el expediente.
- Distingue siempre entre **fórmula**, **valor guardado (cacheado)** y **cálculo independiente recalculado**.
- **Prohibido limitarse a las alertas deterministas del atlas:** las alertas automáticas son solo un punto de partida; debes leer sistemáticamente **TODAS las hojas visibles** con `leer_hoja_excel`.
- El objetivo es comprobar que los datos del avalúo sean exactos, coherentes y trazables, no solo que Excel calcule sin `#REF!`.

## Checklist Técnico Integral (obligatorio en todas las hojas visibles)

1. **Identificación y coherencia entre hojas:** propietario, identificación del inmueble, dirección, ubicación, código catastral, lote, manzana, predio, áreas de terreno y construcción, linderos, mensuras, alícuotas y descripción física del inmueble. Detectar contradicciones entre carátula, hojas de detalle, valoración y resumen/conclusiones, así como rastros de copia y pega de otros avalúos.
2. **Características técnicas y depreciación:** edad, vida útil, estado de conservación, método/porcentaje de depreciación aplicado y coherencia con la descripción y fotografías/observaciones.
3. **Control exhaustivo del Estudio de Mercado:**
   - Recalcula individualmente **cada referencia de mercado**: `Precio unitario = Precio de oferta / Área correspondiente` (comprobando unidad y congruencia de área de terreno/construcción).
   - Comprueba todos los factores de homologación aplicados (negociación, tamaño, ubicación, forma, frente, topografía, acceso y cualquier otra variable presente en el libro) y recalcula el **valor unitario homologado** y el **promedio / valor adoptado**.
   - Señala referencias atípicas o que se alejen significativamente del comportamiento de las demás para revisión del perito, **sin eliminarlas automáticamente** ni asumir error únicamente por tener un valor distinto (analiza si ubicación, área, estado o tipología lo justifican).
4. **Control de la Cadena del Valor Final (de origen a conclusión):**
   `REFERENCIAS DE MERCADO → VALORES UNITARIOS → HOMOLOGACIÓN → VALOR UNITARIO ADOPTADO → ÁREA VALORADA → VALOR DEL TERRENO → VALOR DE LAS CONSTRUCCIONES → OBRAS COMPLEMENTARIAS → VALOR DE REPOSICIÓN → DEPRECIACIONES Y AJUSTES → VALOR NETO DE REPOSICIÓN → VALOR COMERCIAL → VALOR DE REALIZACIÓN (si aplica) → CONCLUSIÓN DEL INFORME`.
   - Comprueba sumas, multiplicaciones, porcentajes, redondeos, totales, subtotales y enlaces entre hojas usando `rastrear_celda_excel`.
   - Ante cualquier error en la cadena, localiza la **celda de origen** y demuestra si se traslada o no al valor comercial o de realización (`demostrado`, `posible`, `ninguno_demostrado`, `no_determinable`).
5. **Revisión de Ortografía, Gramática y Redacción Técnica:**
   - Detecta errores concretos de ortografía, gramática, puntuación, digitación, frases incompletas o redacción técnicamente incorrecta en observaciones, descripciones y conclusiones.
   - Clasifícalos como `Ortografía`, `Gramática` o `Redacción` (normalmente severidad `menor`, salvo que alteren el sentido legal o técnico) y propón el texto corregido exacto. No reportes simples preferencias estilísticas.

## Decision Gates (Control estricto de falsos positivos)

- **Antes de reportar un error:** verifica nuevamente el dato y la fórmula, comprueba si otra hoja explica el criterio y revisa precedentes/dependientes con `rastrear_celda_excel`.
- **Diferencia explicada** por unidad, fecha, redondeo documentado o característica real del inmueble: no la reportes como error confirmado.
- **Sin información suficiente** para afirmar error: clasifica el estado como `requiere_verificacion` (nunca inventes datos faltantes).
- **Hoja, fila o columna oculta:** nunca la marques como destino. Consúltala vía `rastrear_celda_excel` solo para explicar si afecta una celda visible aplicable; si afecta un cálculo visible, señala la celda visible afectada citando el origen oculto.
- **Celda vacía o bloque opcional de plantilla sin usar:** no la marques por estar vacía ni por un error guardado aislado en un módulo no aplicable (ej. bloque de propiedad horizontal vacío en un lote urbano). Si falta un dato obligatorio y aplicable, informa la omisión con evidencia; si rompe un resultado visible aplicable, señala ese resultado.

## Execution Steps

1. Revisa el atlas (`resumen_expediente`) para mapear todas las hojas visibles, su rango utilizado y las alertas deterministas (`formula_errors`, `external_links`, `hidden_dependencies`, `formula_breaks`, `formula_inconsistencies`).
2. Lee sistemáticamente **cada hoja visible completa** con `leer_hoja_excel` (paginando con `next_row` hasta `null`). Usa `leer_rango_excel` o `buscar_excel` para verificaciones puntuales.
3. Verifica todas las alertas deterministas distinguiendo módulos aplicables vs. plantillas opcionales vacías.
4. Recalcula todas las referencias del estudio de mercado, homologaciones, depreciaciones, obras complementarias, totales, subtotales, valor comercial y valor de realización; usa `rastrear_celda_excel` en toda celda dudosa o con impacto en el valor.
5. Revisa consistencia cruzada entre todas las hojas (identificación, áreas, cifras y textos) y detecta errores concretos de ortografía, gramática, redacción y copia y pega.
6. Estructura el resumen ejecutivo de 8 ejes y la categoría de conclusión recomendada para validación del perito.

## Output Contract

Devuelve candidatos clasificados con la taxonomía estándar de `Tipo de revisión` del [contrato común](../README.md), incluyendo operandos, unidades, fórmula, valor actual, valor recalculado, diferencia exacta, evidencia verificable, recomendación e incidencia en el valor. No presentes un total como recalculado si solo leíste su valor guardado.
