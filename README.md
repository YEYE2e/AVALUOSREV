# Revisión de avalúos

Aplicación local de apoyo a la revisión técnica y documental de avalúos inmobiliarios (cualquier categoría y tamaño). El sistema asiste al perito revisor con agentes conectados por MCP de solo lectura y una segunda etapa de verificación independiente; no sustituye el criterio profesional ni autoriza la entrega bancaria.

## Arquitectura del flujo

1. **Ingesta e indexación determinista (100 % del expediente):** guarda originales intactos bajo UUID y escanea todas las hojas visibles, filas y columnas del Excel (errores de fórmula, vínculos externos, dependencias ocultas y quiebres de fórmula), todas las páginas de cada PDF y todos los bloques/tablas DOCX e imágenes de respaldo.
2. **Auditoría técnica y contraste documental (Agente 1 + MCP solo lectura):** ejecuta `omp` en un directorio temporal aislado con `--tools=` y acceso exclusivo al servidor MCP `avaluos-lectura` (búsqueda completa en Excel/PDF/DOCX, lectura de rangos, rastreo de precedentes y dependientes, e inspección visual de páginas PDF e imágenes).
3. **Verificación independiente (Agente 2 en sesión separada):** reabre las fuentes originales mediante MCP para confirmar, matizar o descartar cada candidato antes de generar entregables.
4. **Entregables profesionales:**
   - **Excel marcado (`avaluo-marcado.xlsx`):** copia sin alterar fórmulas ni valores originales, con celdas señaladas en rojo, notas visibles consolidadas (conservando notas previas) y hoja `REVISIÓN TÉCNICA` de 16 columnas con enlaces `Ir a la celda`.
   - **Informe Word (`informe-revision.docx`):** resumen ejecutivo, KPIs por severidad y estado, tabla consolidada, detalle técnico con evidencias de ambos extremos, trazabilidad de descartados/explicados y límites de cobertura.

## Ejecución local

1. Copia `.env.example` a `.env` en la raíz y ajusta `AVALUOS_STORAGE_ROOT` (ruta absoluta), modelo y límites.
2. Instala y arranca el backend (PowerShell):

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev,documents]"
.\.venv\Scripts\python.exe -m pytest -v
.\.venv\Scripts\python.exe -m uvicorn app.main:app --env-file ../.env --host 127.0.0.1 --port 8000
```

3. En otra terminal, arranca el frontend Astro:

```powershell
cd frontend
npm.cmd install
npm.cmd run dev
```

Abre `http://127.0.0.1:5174/`.

## Requisitos opcionales recomendados

- **Excel de escritorio en Windows:** necesario para generar la copia `.xlsx` marcada con notas visibles y hoja `REVISIÓN TÉCNICA`.
- **`pdftoppm` en `PATH`:** habilita la inspección visual multimodal de páginas PDF escaneadas, planos y detección de cortes en el control final PDF.

## Validación y contratos

- Pruebas backend: `.\.venv\Scripts\python.exe -m pytest -v` (en `backend/`).
- Compilación frontend: `npm.cmd run build` (en `frontend/`).
- Esquemas JSON sincronizados: `.\.venv\Scripts\python.exe -m app.schema` (en `backend/`).
- Skills versionadas: [skills/README.md](skills/README.md) · Evaluaciones: [evals/README.md](evals/README.md).
