#!/usr/bin/env python3
"""Cross-platform environment and dependency checker for AVALUOSREV (macOS, Linux, Windows).

Uses Python standard library only so it can run even before creating a virtual environment:
    python3 scripts/check_env.py          # macOS / Linux
    python scripts/check_env.py           # Windows
    python3 scripts/check_env.py --json   # Machine-readable output
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

if sys.version_info >= (3, 11):
    import tomllib
else:
    tomllib = None

ROOT = Path(__file__).resolve().parents[1]
BACKEND_DIR = ROOT / "backend"
FRONTEND_DIR = ROOT / "frontend"
PYPROJECT_PATH = BACKEND_DIR / "pyproject.toml"
ENV_PATH = ROOT / ".env"

# Map distribution names from pyproject.toml to their importable Python module names.
DIST_TO_IMPORT = {
    "fastapi": "fastapi",
    "uvicorn": "uvicorn",
    "pydantic": "pydantic",
    "pytest": "pytest",
    "httpx": "httpx",
    "mcp": "mcp",
    "openpyxl": "openpyxl",
    "defusedxml": "defusedxml",
    "pypdf": "pypdf",
    "python-docx": "docx",
    "python-dotenv": "dotenv",
    "python-multipart": "multipart",
}

# Runtime modules required directly by FastAPI forms and uvicorn --env-file.
EXTRA_RUNTIME_DISTS = ("python-dotenv", "python-multipart")

_DEP_NAME_RE = re.compile(r"^\s*([A-Za-z0-9_.-]+)")
_WINDOWS_DRIVE_RE = re.compile(r"^[A-Za-z]:[\\/]")


def _os_family() -> str:
    system = platform.system().lower()
    if system == "darwin":
        return "macos"
    if system == "windows":
        return "windows"
    return "linux"


def _parse_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for raw_line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        values[key.strip()] = val.strip().strip('"').strip("'")
    return values


def _declare_pyproject_packages() -> list[str]:
    names: list[str] = []
    if tomllib is None or not PYPROJECT_PATH.is_file():
        return list(DIST_TO_IMPORT.keys())
    data = tomllib.loads(PYPROJECT_PATH.read_text(encoding="utf-8"))
    project = data.get("project", {})
    raw_specs: list[str] = list(project.get("dependencies", []))
    for group_specs in project.get("optional-dependencies", {}).values():
        raw_specs.extend(group_specs)
    for spec in raw_specs:
        match = _DEP_NAME_RE.match(spec)
        if match:
            dist = match.group(1).lower()
            if dist not in names:
                names.append(dist)
    for extra in EXTRA_RUNTIME_DISTS:
        if extra not in names:
            names.append(extra)
    return names


def _resolve_target_python() -> tuple[Path, str]:
    venv_python = (
        BACKEND_DIR / ".venv" / "Scripts" / "python.exe"
        if os.name == "nt"
        else BACKEND_DIR / ".venv" / "bin" / "python"
    )
    if venv_python.is_file():
        return venv_python, "backend/.venv"
    return Path(sys.executable), "actual (sin backend/.venv)"


def _probe_python_env(python_exe: Path, dist_names: list[str]) -> tuple[str | None, list[str], str | None]:
    modules = {dist: DIST_TO_IMPORT.get(dist, dist.replace("-", "_")) for dist in dist_names}
    probe = (
        "import importlib.util, json, sys;"
        f"mods = {json.dumps(modules)};"
        "missing = [d for d, m in mods.items() if importlib.util.find_spec(m) is None];"
        "print(json.dumps({'version': f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}', 'missing': missing}))"
    )
    try:
        proc = subprocess.run(
            [str(python_exe), "-c", probe],
            capture_output=True,
            text=True,
            timeout=15,
            check=True,
        )
        payload = json.loads(proc.stdout.strip())
        return payload.get("version"), payload.get("missing", []), None
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError) as exc:
        return None, dist_names, str(exc)


def _run_version(executable: str, args: list[str] | None = None) -> tuple[bool, str]:
    resolved = shutil.which(executable)
    if not resolved:
        return False, ""
    try:
        proc = subprocess.run(
            [resolved, *(args or ["--version"])],
            capture_output=True,
            text=True,
            timeout=10,
        )
        out = (proc.stdout or proc.stderr or "").strip().splitlines()
        return proc.returncode == 0, (out[0].strip() if out else resolved)
    except (OSError, subprocess.SubprocessError):
        return False, resolved


def _check_windows_excel_com() -> bool:
    if os.name != "nt":
        return False
    powershell = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    if not powershell.is_file():
        return False
    try:
        proc = subprocess.run(
            [
                str(powershell),
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "if ([type]::GetTypeFromProgID('Excel.Application')) { exit 0 } else { exit 1 }",
            ],
            capture_output=True,
            timeout=10,
        )
        return proc.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def run_checks() -> dict[str, Any]:
    os_name = _os_family()
    checks: list[dict[str, Any]] = []

    def add(
        key: str,
        label: str,
        status: str,
        detail: str,
        *,
        required: bool = True,
        hint: str = "",
    ) -> None:
        checks.append(
            {
                "key": key,
                "label": label,
                "status": status,  # ok | warn | fail
                "required": required,
                "detail": detail,
                "hint": hint,
            }
        )

    # 1. Python version + backend packages (source of truth: backend/pyproject.toml)
    py_exe, py_source = _resolve_target_python()
    dists = _declare_pyproject_packages()
    py_ver, missing_dists, probe_err = _probe_python_env(py_exe, dists)
    if probe_err or not py_ver:
        add(
            "python",
            "Python >= 3.11",
            "fail",
            f"No se pudo inspeccionar {py_exe}: {probe_err}",
            hint="Instala Python 3.11+ y crea el entorno virtual en backend/.venv.",
        )
    else:
        major, minor = (int(part) for part in py_ver.split(".")[:2])
        if (major, minor) >= (3, 11):
            add("python", "Python >= 3.11", "ok", f"Python {py_ver} ({py_source})")
        else:
            add(
                "python",
                "Python >= 3.11",
                "fail",
                f"Python {py_ver} detectado en {py_source} (se requiere >= 3.11)",
                hint="Instala Python 3.11 o superior (en macOS: brew install python@3.12).",
            )

        if not missing_dists:
            add(
                "python_deps",
                "Dependencias Python (pyproject.toml)",
                "ok",
                f"Todas instaladas ({len(dists)} paquetes verificados en {py_source})",
            )
        else:
            install_cmd = (
                ".\\.venv\\Scripts\\python.exe -m pip install -r ..\\requirements.txt"
                if os_name == "windows"
                else "python3 -m pip install -r requirements.txt"
            )
            add(
                "python_deps",
                "Dependencias Python (pyproject.toml)",
                "fail",
                f"Faltan: {', '.join(missing_dists)}",
                hint=f"Crea/activa backend/.venv y ejecuta: {install_cmd}",
            )

    # 2. Node.js, npm y frontend/node_modules
    node_ok, node_ver = _run_version("node", ["-v"])
    npm_cmd = "npm.cmd" if os_name == "windows" else "npm"
    npm_ok, npm_ver = _run_version(npm_cmd, ["-v"])
    if node_ok and npm_ok:
        add("node", "Node.js y npm", "ok", f"Node {node_ver} · npm {npm_ver}")
    else:
        node_hint = "brew install node" if os_name == "macos" else "Instala Node.js LTS (https://nodejs.org)"
        add("node", "Node.js y npm", "fail", "No se encontró node o npm en PATH", hint=node_hint)

    if (FRONTEND_DIR / "node_modules" / "astro").is_dir():
        add("frontend_deps", "Dependencias frontend (Astro/React)", "ok", "frontend/node_modules instalado")
    else:
        add(
            "frontend_deps",
            "Dependencias frontend (Astro/React)",
            "fail",
            "Falta instalar dependencias en frontend/node_modules",
            hint="Ejecuta: cd frontend && npm install",
        )

    # 3. Configuración .env y AVALUOS_STORAGE_ROOT
    env_vals = _parse_env_file(ENV_PATH)
    agent_exec = os.getenv("AVALUOS_AGENT_EXECUTABLE") or env_vals.get("AVALUOS_AGENT_EXECUTABLE") or "omp"
    storage_raw = os.getenv("AVALUOS_STORAGE_ROOT") or env_vals.get("AVALUOS_STORAGE_ROOT") or ""

    if not ENV_PATH.is_file() and not storage_raw:
        add(
            "env",
            "Archivo .env y almacenamiento",
            "fail",
            "No existe .env en la raíz ni AVALUOS_STORAGE_ROOT en el entorno",
            hint=(
                "Copia .env.example a .env y define AVALUOS_STORAGE_ROOT "
                + ("(ej. /Users/tuusuario/avaluos-data)" if os_name == "macos" else "(ruta absoluta privada)")
            ),
        )
    elif not storage_raw:
        add(
            "env",
            "Archivo .env y almacenamiento",
            "fail",
            "AVALUOS_STORAGE_ROOT está vacío en .env",
            hint="Define una ruta absoluta privada para AVALUOS_STORAGE_ROOT en .env.",
        )
    elif os_name != "windows" and _WINDOWS_DRIVE_RE.match(storage_raw):
        add(
            "env",
            "Archivo .env y almacenamiento",
            "fail",
            f"AVALUOS_STORAGE_ROOT apunta a una ruta de Windows ({storage_raw}) en {os_name}",
            hint="Actualiza AVALUOS_STORAGE_ROOT en .env a una ruta POSIX absoluta (ej. /Users/tuusuario/avaluos-data).",
        )
    else:
        storage_path = Path(storage_raw).expanduser()
        if not storage_path.is_absolute():
            add(
                "env",
                "Archivo .env y almacenamiento",
                "fail",
                f"AVALUOS_STORAGE_ROOT no es una ruta absoluta: {storage_raw}",
                hint="Usa una ruta absoluta en AVALUOS_STORAGE_ROOT.",
            )
        else:
            add("env", "Archivo .env y almacenamiento", "ok", f"AVALUOS_STORAGE_ROOT={storage_path}")

    # 4. Agente CLI (omp)
    agent_ok, agent_ver = _run_version(agent_exec, ["--version"])
    if agent_ok:
        add("agent_cli", f"Agente CLI ({agent_exec})", "ok", agent_ver)
    else:
        add(
            "agent_cli",
            f"Agente CLI ({agent_exec})",
            "fail",
            f"No se encontró '{agent_exec}' en PATH",
            hint=f"Instala y autentica '{agent_exec}' o ajusta AVALUOS_AGENT_EXECUTABLE en .env.",
        )

    # 5. pdftoppm / Poppler (recomendado para inspección visual PDF)
    pdf_ok, pdf_ver = _run_version("pdftoppm", ["-v"])
    if pdf_ok:
        add("pdftoppm", "Inspección visual PDF (pdftoppm)", "ok", pdf_ver, required=False)
    else:
        poppler_hint = {
            "macos": "brew install poppler",
            "linux": "sudo apt install poppler-utils",
            "windows": "Instala Poppler o MiKTeX y agrega pdftoppm al PATH",
        }[os_name]
        add(
            "pdftoppm",
            "Inspección visual PDF (pdftoppm)",
            "warn",
            "No encontrado: los PDF escaneados y el control visual de cortes quedan deshabilitados",
            required=False,
            hint=poppler_hint,
        )

    # 6. Motor de exportación Excel marcado
    if os_name == "windows":
        if _check_windows_excel_com():
            add(
                "excel_export",
                "Exportación Excel marcado",
                "ok",
                "Excel de escritorio (COM) disponible + respaldo multiplataforma openpyxl",
                required=False,
            )
        else:
            add(
                "excel_export",
                "Exportación Excel marcado",
                "warn",
                "Excel de escritorio no detectado; usará automáticamente el motor multiplataforma openpyxl",
                required=False,
            )
    else:
        has_openpyxl = "openpyxl" not in missing_dists
        add(
            "excel_export",
            "Exportación Excel marcado",
            "ok" if has_openpyxl else "fail",
            (
                "Motor multiplataforma openpyxl activo (macOS/Linux: celdas rojas, notas y hoja REVISIÓN TÉCNICA)"
                if has_openpyxl
                else "Falta openpyxl para exportar copias marcadas XLSX en macOS/Linux"
            ),
            required=True,
            hint="" if has_openpyxl else "Instala las dependencias Python con: pip install -r requirements.txt",
        )

    failed = [c for c in checks if c["status"] == "fail"]
    warnings = [c for c in checks if c["status"] == "warn"]
    return {
        "platform": platform.platform(),
        "os": os_name,
        "ready": len(failed) == 0,
        "failed_count": len(failed),
        "warning_count": len(warnings),
        "checks": checks,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Verifica dependencias y configuración de AVALUOSREV en macOS, Linux y Windows.")
    parser.add_argument("--json", action="store_true", help="Imprime el diagnóstico en formato JSON.")
    args = parser.parse_args()

    report = run_checks()
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["ready"] else 1

    print(f"AVALUOSREV — Verificación de entorno ({report['os']} · {report['platform']})")
    print("-" * 72)
    tags = {"ok": "[ OK ]", "warn": "[WARN]", "fail": "[FAIL]"}
    for item in report["checks"]:
        print(f"{tags[item['status']]} {item['label']}: {item['detail']}")
        if item["hint"]:
            print(f"       -> {item['hint']}")
    print("-" * 72)
    if report["ready"]:
        suffix = f" ({report['warning_count']} aviso(s) opcional(es))" if report["warning_count"] else ""
        print(f"Entorno listo para ejecutar el sistema{suffix}.")
        return 0
    print(f"Faltan {report['failed_count']} requisito(s) obligatorio(s). Corrige los puntos [FAIL] y vuelve a ejecutar.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
