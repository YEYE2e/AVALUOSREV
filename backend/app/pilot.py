"""Opt-in local agent smoke test. It never accepts documents from HTTP."""

import os
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Lock

from fastapi import APIRouter, HTTPException, Request
from starlette.concurrency import run_in_threadpool

router = APIRouter()
MODEL = "openai-codex/gpt-6.1-terra"
EFFORT = "medium"
ROOT = Path(__file__).resolve().parents[2]
INPUTS = (
    "skills/README.md",
    "skills/contraste-documental/SKILL.md",
    "skills/revision-tecnica-excel/SKILL.md",
    "skills/verificacion-independiente/SKILL.md",
    "evals/piloto-excel-ficticio.md",
)
RUN_LOCK = Lock()


def enabled() -> bool:
    return os.environ.get("AVALUOS_LOCAL_PILOT") == "1"


@router.get("/api/pilot")
def pilot_status() -> dict:
    return {"enabled": enabled(), "model": MODEL, "effort": EFFORT}


@router.post("/api/pilot")
async def run_pilot(request: Request) -> dict:
    if not enabled():
        raise HTTPException(503, "Piloto local deshabilitado")
    if (request.headers.get("origin") != "http://127.0.0.1:5174"
            or request.headers.get("x-avaluos-pilot") != "synthetic"
            or request.headers.get("host") not in {"127.0.0.1:5174", "127.0.0.1:8000"}):
        raise HTTPException(403, "Solicitud de piloto no autorizada")
    async for chunk in request.stream():
        if chunk:
            raise HTTPException(403, "El piloto no acepta archivos ni otros datos")
    return await run_in_threadpool(run_synthetic_review)


def run_synthetic_review() -> dict:
    if not RUN_LOCK.acquire(blocking=False):
        raise HTTPException(409, "Ya hay una prueba en ejecución")
    try:
        with TemporaryDirectory() as cwd:
            result = subprocess.run(
                ["omp", "-p", "--model", MODEL, "--thinking", EFFORT, "--max-time", "90",
                 "--max-cost", "0.10", "--offline", "--no-session", "--no-tools", "--no-context-files",
                 "--no-skills", "--no-rules", "--no-extensions", *(f"@{ROOT / p}" for p in INPUTS)],
                cwd=cwd, stdin=subprocess.DEVNULL, timeout=100,
                capture_output=True, encoding="utf-8", errors="replace",
            )
    except subprocess.TimeoutExpired:
        raise HTTPException(504, "El agente superó el tiempo de ejecución") from None
    except OSError:
        raise HTTPException(503, "omp no se puede ejecutar desde el servidor") from None
    finally:
        RUN_LOCK.release()
    if result.returncode or not result.stdout.strip():
        raise HTTPException(502, "omp no devolvió un resultado. Revisa autenticación, cuota y configuración local.")
    return {"model": MODEL, "effort": EFFORT, "report": result.stdout.strip()[:24000]}
