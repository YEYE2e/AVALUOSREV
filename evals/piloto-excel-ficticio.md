# Entrada ficticia para el piloto de revisión Excel

Representación textual de celdas de un libro ficticio; no es una extracción del avalúo de la empresa. No hay respaldos. No intentes acceder a otros archivos ni a destinos de vínculos. Aplica las skills adjuntas. Distingue candidatos de hallazgos independientemente verificados; esta ejecución única no sustituye la segunda comprobación. Devuelve un informe breve sin datos personales, con hoja/celda, operación, impacto y cobertura. No supongas reglas bancarias.

| Hoja | Celda | Significado | Contenido | Último valor guardado |
|---|---|---|---|---|
| Mercado | A2 | Precio de oferta, USD | 120000 | 120000 |
| Mercado | B2 | Superficie de referencia, m² | 300 | 300 |
| Mercado | C2 | Precio unitario esperado: oferta / superficie | `=A2/240` | 500 |
| Mercado | A3 | Precio de oferta, USD | 102000 | 102000 |
| Mercado | B3 | Superficie de referencia, m² | 300 | 300 |
| Mercado | C3 | Precio unitario, USD/m² | `=A3/B3` | 340 |
| Mercado | D2 | Factor de homologación | `=[1]Factores!B2` | 0.95 |
| Mercado | E2 | Precio homologado | `=C2*D2` | 475 |
| Valoración | B2 | Área de terreno, m² | 250 | 250 |
| Valoración | B3 | Unitario adoptado, USD/m² | `=Mercado!E2` | 475 |
| Valoración | B4 | Valor del terreno, USD | `=B2*B3` | 118750 |
| Valoración | B5 | Área construida, m² | 180 | 180 |

No se conoce el contenido actual del libro vinculado. No hay evidencia adicional del alcance de la homologación ni otros componentes del valor comercial.
