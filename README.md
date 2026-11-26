# Auditor?a documental de aval?os

Cimientos de una aplicaci?n interna de apoyo a la revisi?n profesional. Esta versi?n no permite cargar expedientes ni ejecutar auditor?as: solo contiene contratos, una primitiva de almacenamiento, una interfaz de proveedor y una pantalla inicial. No utilizar con documentaci?n real ni exponer en red.

## Preparaci?n local

Requisitos: Python 3.11 o posterior y Node.js compatible con Vite 6. En PowerShell:

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
.\.venv\Scripts\python -m pytest
.\.venv\Scripts\python -m uvicorn app.main:app --reload
```

En otra terminal:

```powershell
cd frontend
npm.cmd install
npm.cmd run dev
```

Para validar frontend: `npm.cmd run build`. Para generar el contrato JSON desde el backend: `python -m app.schema > audit-result.schema.json` (la salida est?ndar de PowerShell puede depender de su codificaci?n). No hay datos de ejemplo ni conexi?n a una base de datos implementada.

## Piloto de skills sin acceso a expedientes

Desde la raíz, con `omp` instalado y autenticado para `openai-codex`, ejecutar en PowerShell:

```powershell
omp -p --model openai-codex/gpt-6.1-terra --thinking medium --max-time 90 --offline --no-session --no-tools --no-context-files --no-skills --no-rules '@skills/README.md' '@skills/contraste-documental/SKILL.md' '@skills/revision-tecnica-excel/SKILL.md' '@skills/verificacion-independiente/SKILL.md' '@evals/piloto-excel-ficticio.md'
```

Envía únicamente las instrucciones seleccionadas y una tabla ficticia al proveedor. El modelo no dispone de herramientas para abrir archivos o consultar otros destinos; `--offline` no es un cortafuegos, limita los servicios auxiliares de omp. Esta prueba mide razonamiento sobre celdas, no lectura nativa de `.xlsx`, verificación en una segunda ejecución ni exportación. No sustituir la entrada ficticia por el libro original.

## Informe local de errores guardados en Excel

Con Excel de escritorio en Windows, ejecutar `powershell -NoProfile -File scripts/revisar-errores-excel.ps1 -Libro 'ruta-al-libro.xlsx'`. Genera una copia marcada y un CSV bajo `data/` sin modificar el original. Marca errores que ya aparecen guardados como tales en el archivo; no envía datos a un modelo, no recalcula, no actualiza vínculos y no verifica causas, impacto ni respaldos. No usar el resultado como aprobación de entrega.

## MCP local de revisión documental

Instalar desde `backend`: `.\.venv\Scripts\python.exe -m pip install -e ".[dev,documents]"`.
Preparar un expediente ficticio con `Workspace.save_original`: cada documento se guarda como `RAIZ/UUID_EXPEDIENTE/originals/UUID_DOCUMENTO`, sin extensión. La raíz no debe ser escribible por el agente ni por otros usuarios. No usar con datos reales ni exponer en red.

Para configurarlo en omp, añadir esta entrada a `.omp/mcp.json` usando rutas absolutas y el UUID del expediente de prueba (conservar las demás entradas si existen):

```json
{
  "mcpServers": {
    "avaluos-lectura": {
      "type": "stdio",
      "command": "C:/proyecto/backend/.venv/Scripts/python.exe",
      "cwd": "C:/proyecto/backend",
      "args": ["-m", "app.document_mcp", "--root", "C:/proyecto/data", "--case", "00000000-0000-0000-0000-000000000001"]
    }
  }
}
```

Usar el inventario y luego pedir rangos Excel, texto de páginas PDF o bloques DOCX. Se omiten celdas ocultas y no se actualizan vínculos ni se recalcula. PDF sin capa de texto requiere OCR separado. DOCX se lee por orden de bloques del cuerpo, no por páginas.

La publicación se configura en **otro proceso**: agregar `--mode report` a una segunda entrada MCP en la sesión de publicación. `crear_informe` crea un borrador DOCX; `marcar_excel` recibe `document_id` y hallazgos revisados (`sheet`, `cell`, `location`, `evidence`, `observation`, `recommendation`, `status`: `sustentado` o `requiere_verificacion`). Genera una copia XLSX con relleno rojo, notas y hoja `REVISIÓN TÉCNICA`. Rechaza destinos ocultos, vacíos o repetidos y guarda bajo `reports/` sin editar el original. La escritura XLSX requiere Excel de escritorio en Windows. Mantener este modo fuera de la sesión de revisión automática: el exportador no decide por sí solo si un bloque aplica ni verifica el sustento semántico. MCP limita sus propias operaciones, pero no el tráfico del cliente o de su proveedor.

Prueba local del protocolo: desde `backend`, `.\.venv\Scripts\python.exe -m pytest tests/test_document_mcp.py`.

## Revisión local desde la web

Copia `.env.example` a `.env` en la raíz. Configura `AVALUOS_STORAGE_ROOT` como ruta absoluta a una carpeta privada. `AVALUOS_AGENT_EXECUTABLE`, `AVALUOS_AGENT_MODEL` y `AVALUOS_AGENT_EFFORT` seleccionan el CLI compatible con omp, el modelo y el esfuerzo. Desde `backend`, arranca `.\.venv\Scripts\python.exe -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000`. Instala `.[dev,documents]` y autentica omp primero. Desde `frontend`, ejecuta `npm.cmd run dev` y abre `http://127.0.0.1:5174/`.

Cada carga registra originales sin modificar en `UUID_EXPEDIENTE/originals` y resultados en `UUID_EXPEDIENTE/reports`. El navegador permite descargar Word y, para un avalúo, la copia XLSX anotada cuando Excel de escritorio está disponible en Windows. Los identificadores se generan en el servidor. Los archivos permanecen en disco hasta que el operador los retire.
El XLSX abre en `REVISIÓN TÉCNICA`. **Ir a la celda** lleva a la marca roja. Excel muestra la explicación, evidencia y recomendación en **Revisar → Comentarios** cuando la instalación admite comentarios encadenados. Si no los admite, se guarda una **Nota** visible en esa misma celda. El exportador vuelve a abrir la copia y comprueba que el detalle quedó guardado.

Usar solo material ficticio o anonimizado. No desplegar en red. El servidor consulta las herramientas MCP localmente y solo envía los fragmentos del muestreo al agente, sin darle herramientas ni rutas a los originales. El informe declara cobertura: Excel visible en rangos A1:L32, PDF hasta 20 páginas de texto sin OCR ni inspección visual y DOCX por bloques, sujetos al presupuesto de contexto. No se presenta como revisión integral ni como comprobación independiente. Se admite una ejecución a la vez, hasta 64 MiB por archivo, 96 MiB y 64 archivos por carga. La prueba sin archivos puede activarse por separado con `AVALUOS_LOCAL_PILOT=1`.

Consulte [arquitectura](docs/architecture.md) y [roadmap](docs/roadmap.md).
