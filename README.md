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

Consulte [arquitectura](docs/architecture.md) y [roadmap](docs/roadmap.md).
