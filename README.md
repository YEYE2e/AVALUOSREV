# Revisión de avalúos

Aplicación local de apoyo a la revisión técnica y documental de avalúos inmobiliarios (cualquier categoría y tamaño), compatible con **macOS, Linux y Windows**. El sistema asiste al perito revisor con agentes conectados por MCP de solo lectura y una segunda etapa de verificación independiente; no sustituye el criterio profesional ni autoriza la entrega bancaria.

## Arquitectura del flujo

1. **Ingesta e indexación determinista (100 % del expediente):** guarda originales intactos bajo UUID y escanea todas las hojas visibles, filas y columnas del Excel (errores de fórmula, vínculos externos, dependencias ocultas, quiebres de fórmula e inconsistencias de patrón de fórmula), todas las páginas de cada PDF y todos los bloques/tablas DOCX e imágenes de respaldo.
2. **Auditoría técnica y contraste documental (Agente 1 + MCP solo lectura):** ejecuta `omp` en un directorio temporal aislado con `--tools=` y acceso exclusivo al servidor MCP `avaluos-lectura` (lectura completa hoja por hoja `leer_hoja_excel`, lectura por bloques `leer_paginas_pdf` y `leer_documento_docx`, búsquedas globales, rastreo de precedentes/dependientes e inspección visual de páginas PDF e imágenes).
3. **Verificación independiente (Agente 2 en sesión separada):** reabre las fuentes originales mediante MCP para confirmar, matizar o descartar cada candidato y auditar cobertura crítica antes de generar entregables.
4. **Entregables profesionales:**
   - **Excel marcado (`avaluo-marcado.xlsx`):** copia sin alterar fórmulas ni valores originales, con celdas señaladas en rojo, notas con prefijo técnico (`REVISIÓN:`, `ERROR DE CÁLCULO:`, `INCONSISTENCIA:`, `REDACCIÓN:`, `REQUIERE VERIFICACIÓN:`) conservando notas previas, y hoja `REVISIÓN TÉCNICA` de 16 columnas con enlaces `Ir a la celda`. Usa Excel de escritorio (COM) en Windows y el motor multiplataforma `openpyxl` en macOS/Linux (configurable con `AVALUOS_EXCEL_ENGINE`).
   - **Informe Word (`informe-revision.docx`):** dictamen documental, resumen ejecutivo en 8 ejes, KPIs por severidad y estado, tabla consolidada, detalle técnico con evidencias de ambos extremos, `CONCLUSIÓN DE LA REVISIÓN`, trazabilidad de descartados/explicados y límites de cobertura.

---

## Requisitos y dependencias al mover el proyecto de equipo

Al clonar o copiar este repositorio en otra máquina (**macOS, Linux o Windows**), el entorno virtual (`.venv`), `node_modules` y el archivo `.env` no se versionan en Git. Instala las siguientes dependencias en el equipo destino:

### 1. Dependencias del sistema (previas)

| Componente | Obligatorio | Para qué sirve | Instalación rápida |
| --- | --- | --- | --- |
| **Git** | Sí | Clonar y actualizar el repositorio | macOS: `xcode-select --install` · Windows: `winget install Git.Git` |
| **Python 3.11 o superior** | Sí | Backend FastAPI, analizador determinista y servidor MCP | macOS: `brew install python@3.12` · Windows: `winget install Python.Python.3.12` · Linux: `sudo apt install python3 python3-venv` |
| **Node.js 20+ (LTS) y npm** | Sí | Interfaz web Astro + React | macOS: `brew install node` · Windows: `winget install OpenJS.NodeJS.LTS` |
| **CLI del agente (`omp`)** | Sí | Ejecuta las etapas de auditoría y verificación independiente con credenciales autenticadas | Instalar `omp` en el `PATH` del sistema (o indicar su ruta absoluta en `AVALUOS_AGENT_EXECUTABLE`) y verificar sesión iniciada antes de levantar el backend |
| **Poppler (`pdftoppm`)** | Recomendado | Renderiza páginas PDF a PNG para inspección visual (PDF escaneados, planos, cuadros, fotos y control final PDF) | macOS: `brew install poppler` · Ubuntu/Debian: `sudo apt install poppler-utils` · Windows: `winget install oschwartz10612.Poppler` (o paquete Poppler en `PATH`) |
| **Microsoft Excel** | Opcional (solo Windows) | Motor COM nativo para conservar estilos complejos en Windows. En macOS/Linux (o Windows sin Excel) se usa automáticamente `openpyxl` | Solo aplica en Windows con `AVALUOS_EXCEL_ENGINE=auto` o `com` |

### 2. Dependencias del proyecto (Python y Node)

- **Backend Python:** declaradas en `backend/pyproject.toml` y delegadas desde `requirements.txt` (`fastapi`, `uvicorn[standard]`, `python-multipart`, `pydantic`, `mcp`, `openpyxl`, `pypdf`, `python-docx`). Se instalan con `pip install -r requirements.txt` dentro de `backend/.venv`.
- **Frontend:** declaradas en `frontend/package.json` (`astro`, `@astrojs/react`, `react`, `react-dom`, `typescript`). Se instalan con `npm install` dentro de `frontend/`.

---

## Configuración del archivo `.env`

Crea `.env` en la raíz del proyecto a partir de `.env.example` y ajusta las rutas para tu equipo:

| Variable | Obligatoria | Valor por defecto | Significado |
| --- | --- | --- | --- |
| `AVALUOS_STORAGE_ROOT` | **Sí** | *(sin defecto seguro)* | Carpeta **absoluta y privada** donde se guardan expedientes (originales intactos, `result.json` e informes). Debe existir o ser creable y **no** estar dentro del repositorio. Ej. macOS: `/Users/tuusuario/avaluos-data` · Windows: `C:/Users/tuusuario/avaluos-data`. |
| `AVALUOS_AGENT_EXECUTABLE` | Sí | `omp` | Comando o ruta absoluta al binario CLI del agente compatible con `omp`. |
| `AVALUOS_AGENT_MODEL` | Sí | `openai-codex/gpt-6.1-terra` | Modelo usado por el Agente 1 (auditoría principal). |
| `AVALUOS_AGENT_EFFORT` | Sí | `medium` | Nivel de razonamiento del Agente 1 (`low`, `medium`, `high`, `xhigh`, `max`). Para avalúos extensos se recomienda `medium` o `high`. |
| `AVALUOS_ENABLE_VERIFIER` | No | `1` | `1` activa la segunda verificación independiente en sesión separada (recomendado); `0` la desactiva. |
| `AVALUOS_VERIFIER_MODEL` | No | *(hereda `AVALUOS_AGENT_MODEL`)* | Modelo específico para el Agente 2 (verificador independiente). Déjalo vacío para usar el mismo que el Agente 1. |
| `AVALUOS_VERIFIER_EFFORT` | No | *(hereda `AVALUOS_AGENT_EFFORT`)* | Esfuerzo específico para el Agente 2. Déjalo vacío para heredar el del Agente 1. |
| `AVALUOS_EXCEL_ENGINE` | No | `auto` | Motor de marcado Excel: `auto` (COM en Windows con respaldo automático a `openpyxl`; `openpyxl` en macOS/Linux), `com` o `openpyxl`. |
| `AVALUOS_AGENT_MAX_TIME` | No | `420` | Tiempo máximo en segundos por etapa de agente (rango permitido: `60`–`1200`). Súbelo si revisas libros muy grandes con múltiples respaldos. |
| `AVALUOS_AGENT_MAX_COST` | No | `2.00` | Presupuesto máximo en USD por etapa de agente (rango permitido: `0.05`–`15.00`). |
| `AVALUOS_LOCAL_PILOT` | No | `0` | Déjalo en `0` en uso normal; solo `1` habilita pruebas piloto internas locales. |

---

## Instalación y puesta en marcha paso a paso

### macOS y Linux

```bash
# 1. Clonar el proyecto
git clone https://github.com/YEYE2e/AVALUOSREV.git
cd AVALUOSREV

# 2. Instalar Poppler (inspección visual de PDFs)
brew install poppler          # macOS
# sudo apt install poppler-utils   # Ubuntu / Debian

# 3. Crear carpeta privada de expedientes y configurar .env
mkdir -p "$HOME/avaluos-data"
cp .env.example .env
# Edita .env y coloca: AVALUOS_STORAGE_ROOT=/Users/TU_USUARIO/avaluos-data (o /home/TU_USUARIO/avaluos-data)

# 4. Instalar dependencias Python y Node
python3 -m venv backend/.venv
backend/.venv/bin/python -m pip install -r requirements.txt
(cd frontend && npm install)

# 5. Verificar que todo el entorno está listo (Python, paquetes, Node, omp, pdftoppm, .env)
python3 scripts/check_env.py
```

Levantar el servicio (en **dos terminales** desde la raíz del proyecto):

```bash
# Terminal 1 — Backend API (puerto 8000)
cd backend && ./.venv/bin/python -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000

# Terminal 2 — Frontend web (puerto 5174)
cd frontend && npm run dev
```

### Windows (PowerShell)

```powershell
# 1. Clonar el proyecto
git clone https://github.com/YEYE2e/AVALUOSREV.git
cd AVALUOSREV

# 2. Crear carpeta privada de expedientes y configurar .env
New-Item -ItemType Directory -Force "$HOME\avaluos-data" | Out-Null
Copy-Item .env.example .env
# Edita .env y coloca: AVALUOS_STORAGE_ROOT=C:/Users/TU_USUARIO/avaluos-data

# 3. Instalar dependencias Python y Node
python -m venv backend\.venv
.\backend\.venv\Scripts\python.exe -m pip install -r requirements.txt
npm.cmd --prefix frontend install

# 4. Verificar todo el entorno antes de levantar
python scripts\check_env.py
```

Levantar el servicio (en **dos terminales PowerShell**):

```powershell
# Terminal 1 — Backend API (puerto 8000)
cd backend
.\.venv\Scripts\python.exe -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000

# Terminal 2 — Frontend web (puerto 5174)
cd frontend
npm.cmd run dev
```

### Comprobación de que el servicio quedó levantado

1. Abre la interfaz web en **`http://127.0.0.1:5174/`**.
2. El frontend reenvía automáticamente `/api/*` al backend en `http://127.0.0.1:8000`. Puedes comprobar el estado del backend en `http://127.0.0.1:8000/api/health`.
3. **Importante:** ambos procesos (backend en `8000` y frontend en `5174`) deben permanecer ejecutándose mientras revisas expedientes. Si cambias el `.env`, reinicia la terminal del backend.

---

## Validación, pruebas y contratos

- Diagnóstico multiplataforma completo: `python3 scripts/check_env.py` (o `python3 scripts/check_env.py --json`).
- Pruebas unitarias y de integración del backend:
  - macOS/Linux: `(cd backend && ./.venv/bin/python -m pytest -v)`
  - Windows: `cd backend; .\.venv\Scripts\python.exe -m pytest -v`
- Compilación de producción del frontend: `npm --prefix frontend run build` (`npm.cmd` en Windows).
- Sincronización de esquemas JSON: `python -m app.schema` (desde `backend/` usando `.venv`).
- Skills versionadas: [skills/README.md](skills/README.md) · Evaluaciones: [evals/README.md](evals/README.md).
