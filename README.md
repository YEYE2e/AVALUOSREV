# Revisión de avalúos

Aplicación local de apoyo a la revisión técnica y documental de avalúos inmobiliarios (cualquier categoría y tamaño), compatible con **macOS, Linux y Windows**. El sistema asiste al perito revisor con agentes conectados por MCP de solo lectura y una segunda etapa de verificación independiente; no sustituye el criterio profesional ni autoriza la entrega bancaria.

## Arquitectura del flujo

1. **Ingesta e indexación determinista (100 % del expediente):** guarda originales intactos bajo UUID y escanea todas las hojas visibles, filas y columnas del Excel (errores de fórmula, vínculos externos, dependencias ocultas y quiebres de fórmula), todas las páginas de cada PDF y todos los bloques/tablas DOCX e imágenes de respaldo.
2. **Auditoría técnica y contraste documental (Agente 1 + MCP solo lectura):** ejecuta `omp` en un directorio temporal aislado con `--tools=` y acceso exclusivo al servidor MCP `avaluos-lectura` (búsqueda completa en Excel/PDF/DOCX, lectura de rangos, rastreo de precedentes y dependientes, e inspección visual de páginas PDF e imágenes).
3. **Verificación independiente (Agente 2 en sesión separada):** reabre las fuentes originales mediante MCP para confirmar, matizar o descartar cada candidato antes de generar entregables.
4. **Entregables profesionales:**
   - **Excel marcado (`avaluo-marcado.xlsx`):** copia sin alterar fórmulas ni valores originales, con celdas señaladas en rojo, notas consolidadas (conservando notas previas) y hoja `REVISIÓN TÉCNICA` de 16 columnas con enlaces `Ir a la celda`. Usa Excel de escritorio (COM) en Windows y el motor multiplataforma `openpyxl` en macOS/Linux (configurable con `AVALUOS_EXCEL_ENGINE`).
   - **Informe Word (`informe-revision.docx`):** resumen ejecutivo, KPIs por severidad y estado, tabla consolidada, detalle técnico con evidencias de ambos extremos, trazabilidad de descartados/explicados y límites de cobertura.

## Comprobación rápida de dependencias (macOS / Linux / Windows)

Un solo comando verifica Python, paquetes declarados en `backend/pyproject.toml`, Node/npm, `frontend/node_modules`, CLI del agente (`omp`), `pdftoppm`, motor Excel y `.env`:

```bash
# macOS / Linux
python3 scripts/check_env.py

# Windows
python scripts/check_env.py
```

## Instalación y ejecución local

### macOS / Linux

```bash
cp .env.example .env
# Edita .env y define una ruta absoluta privada en AVALUOS_STORAGE_ROOT (ej. /Users/tuusuario/avaluos-data)
brew install poppler  # En macOS (en Ubuntu/Debian: sudo apt install poppler-utils)

python3 -m venv backend/.venv
backend/.venv/bin/python -m pip install -r requirements.txt
(cd frontend && npm install)
python3 scripts/check_env.py
```

Arranque (en dos terminales):

```bash
# Terminal 1 (backend)
cd backend && ./.venv/bin/python -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000

# Terminal 2 (frontend)
cd frontend && npm run dev
```

### Windows (PowerShell)

```powershell
Copy-Item .env.example .env
# Edita .env y define una ruta absoluta privada en AVALUOS_STORAGE_ROOT

python -m venv backend\.venv
.\backend\.venv\Scripts\python.exe -m pip install -r requirements.txt
npm.cmd --prefix frontend install
python scripts\check_env.py
```

Arranque (en dos terminales):

```powershell
# Terminal 1 (backend)
cd backend
.\.venv\Scripts\python.exe -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000

# Terminal 2 (frontend)
cd frontend
npm.cmd run dev
```

Abre `http://127.0.0.1:5174/`.

## Validación y contratos

- Diagnóstico multiplataforma: `python3 scripts/check_env.py` (o `--json`).
- Pruebas backend: `pytest -v` (dentro de `backend/` usando `.venv`).
- Compilación frontend: `npm run build` (dentro de `frontend/`).
- Esquemas JSON sincronizados: `python -m app.schema` (dentro de `backend/`).
- Skills versionadas: [skills/README.md](skills/README.md) · Evaluaciones: [evals/README.md](evals/README.md).
