"""Local appraisal review HTTP API: staged immutable uploads, async jobs, SSE progress, cancellation, and history."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal
from uuid import UUID, uuid4

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import UploadFile
from starlette.responses import FileResponse

from .analyzer import MAX_FILE_BYTES, sanitize_label
from .orchestrator import AuditCancelledError, AuditExecutionError, execute_audit_pipeline
from .pilot import RUN_LOCK

router = APIRouter()
MAX_REQUEST = 512 * 1024**2
MAX_FILE = MAX_FILE_BYTES
MAX_UPLOAD_FILES = 128

LOCAL_HOSTS = {"127.0.0.1:5174", "localhost:5174", "127.0.0.1:8000", "localhost:8000"}
LOCAL_ORIGINS = {"http://127.0.0.1:5174", "http://localhost:5174"}


@dataclass
class ReviewJob:
    job_id: str
    case_id: UUID
    mode: str
    status: str = "queued"
    stage: str = "queued"
    message: str = "Expediente recibido. Iniciando revisión…"
    stats: dict[str, Any] = field(default_factory=dict)
    events: list[dict[str, Any]] = field(default_factory=list)
    result: dict[str, Any] | None = None
    error: str | None = None
    created_at: float = field(default_factory=time.time)
    cancel_event: threading.Event = field(default_factory=threading.Event)

    def push_event(self, stage: str, message: str, stats: dict[str, Any] | None = None) -> None:
        self.stage = stage
        self.status = stage
        self.message = message
        if stats is not None:
            self.stats = dict(stats)
        self.events.append({
            "seq": len(self.events) + 1,
            "stage": stage,
            "message": message,
            "stats": self.stats,
            "timestamp": datetime.now(timezone.utc).isoformat(),
        })

    def snapshot(self) -> dict[str, Any]:
        return {
            "job_id": self.job_id,
            "case_id": str(self.case_id),
            "mode": self.mode,
            "status": self.status,
            "stage": self.stage,
            "message": self.message,
            "stats": self.stats,
            "events": self.events,
            "result": self.result,
            "error": self.error,
        }


_JOBS: dict[str, ReviewJob] = {}
_JOBS_LOCK = threading.Lock()


def storage(request: Request) -> Path:
    configured = os.getenv("AVALUOS_STORAGE_ROOT")
    if not configured:
        raise HTTPException(503, "Configura AVALUOS_STORAGE_ROOT para activar la revisión local.")
    if (
        request.headers.get("host") not in LOCAL_HOSTS
        or request.client is None
        or request.client.host not in {"127.0.0.1", "::1"}
    ):
        raise HTTPException(403, "La revisión solo está disponible en este equipo.")
    root = Path(configured).expanduser()
    if not root.is_absolute():
        raise HTTPException(503, "AVALUOS_STORAGE_ROOT debe ser una ruta absoluta.")
    root.mkdir(parents=True, exist_ok=True)
    return root.resolve(strict=True)


def _check_local_post(request: Request) -> None:
    if (
        request.headers.get("origin") not in LOCAL_ORIGINS
        or request.headers.get("x-avaluos-review") not in {"local", "anonymized"}
    ):
        raise HTTPException(403, "Usa la interfaz local para gestionar revisiones.")


def _copy_and_hash(source_file, dest_path: Path) -> tuple[int, str]:
    hasher = hashlib.sha256()
    total = 0
    with dest_path.open("xb") as destination:
        while True:
            chunk = source_file.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_FILE:
                raise HTTPException(413, "Cada archivo debe pesar como máximo 128 MiB.")
            hasher.update(chunk)
            destination.write(chunk)
    return total, hasher.hexdigest()


async def _stage_uploads(request: Request, root: Path, mode: Literal["excel", "pdf"]) -> tuple[UUID, str]:
    if not request.headers.get("content-type", "").startswith("multipart/form-data;"):
        raise HTTPException(403, "Usa la interfaz local para iniciar la revisión.")
    length = request.headers.get("content-length", "")
    if (
        not length.isascii()
        or not length.isdecimal()
        or len(length) > 12
        or not 0 < int(length) <= MAX_REQUEST
        or "transfer-encoding" in request.headers
    ):
        raise HTTPException(413, "La carga supera el límite permitido (512 MiB).")

    case_id = uuid4()
    originals = root / str(case_id) / "originals"
    metadata_docs: dict[str, dict[str, Any]] = {}
    principal_id = ""

    async with request.form(max_files=MAX_UPLOAD_FILES, max_fields=0) as form:
        principal = form.getlist("principal")
        supports = form.getlist("respaldo")
        if len(principal) != 1 or (mode == "pdf" and supports) or len(form.multi_items()) != 1 + len(supports):
            raise HTTPException(422, "Selecciona un archivo principal y solo los respaldos admitidos.")
        uploads = principal + supports
        if any(not isinstance(f, UploadFile) or not f.size or f.size > MAX_FILE for f in uploads):
            raise HTTPException(413, "Cada archivo debe tener contenido y pesar como máximo 128 MiB.")
        originals.mkdir(parents=True, exist_ok=False)
        for index, upload in enumerate(uploads):
            document_id = str(uuid4())
            if index == 0:
                principal_id = document_id
            role = "principal" if index == 0 else "respaldo"
            label = sanitize_label(upload.filename, f"{role}-{index + 1}")
            dest = originals / document_id
            size_bytes, sha256 = await run_in_threadpool(_copy_and_hash, upload.file, dest)
            metadata_docs[document_id] = {
                "role": role,
                "label": label,
                "size_bytes": size_bytes,
                "sha256": sha256,
            }

    meta_path = root / str(case_id) / "metadata.json"
    meta_path.write_text(
        json.dumps(
            {
                "case_id": str(case_id),
                "mode": mode,
                "created_at": datetime.now(timezone.utc).isoformat(),
                "documents": metadata_docs,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return case_id, principal_id


def _prune_old_jobs() -> None:
    cutoff = time.time() - 3600
    with _JOBS_LOCK:
        stale = [
            jid
            for jid, job in _JOBS.items()
            if job.status in {"completed", "failed", "cancelled"} and job.created_at < cutoff
        ]
        for jid in stale:
            _JOBS.pop(jid, None)


def _run_job_thread(job: ReviewJob, root: Path, principal_id: str) -> None:
    try:
        result = asyncio.run(
            execute_audit_pipeline(
                root=root,
                case_id=job.case_id,
                mode=job.mode,
                principal_id=principal_id,
                progress=job.push_event,
                cancel_event=job.cancel_event,
            )
        )
        job.result = result
        job.push_event("completed", f"Revisión completada: {result['findings']} observación(es) activa(s).", result.get("stats"))
    except AuditCancelledError as exc:
        job.error = str(exc)
        job.push_event("cancelled", str(exc))
    except AuditExecutionError as exc:
        job.error = str(exc)
        job.push_event("failed", str(exc))
    except ImportError:
        msg = "Instala el extra documents del backend."
        job.error = msg
        job.push_event("failed", msg)
    except Exception:
        msg = "No se pudo leer el expediente o completar la revisión. Comprueba que los archivos sean legibles y no estén cifrados."
        job.error = msg
        job.push_event("failed", msg)
    finally:
        RUN_LOCK.release()


@router.post("/api/reviews/jobs/{mode}")
async def start_review_job(request: Request, mode: Literal["excel", "pdf"]):
    root = storage(request)
    _check_local_post(request)
    _prune_old_jobs()
    if not RUN_LOCK.acquire(blocking=False):
        raise HTTPException(409, "Ya hay una revisión en ejecución.")
    try:
        case_id, principal_id = await _stage_uploads(request, root, mode)
    except Exception:
        RUN_LOCK.release()
        raise

    job = ReviewJob(job_id=str(uuid4()), case_id=case_id, mode=mode)
    job.push_event("queued", "Expediente guardado sin alterar originales. Preparando análisis…")
    with _JOBS_LOCK:
        _JOBS[job.job_id] = job

    worker = threading.Thread(target=_run_job_thread, args=(job, root, principal_id), daemon=True)
    worker.start()
    return JSONResponse(status_code=202, content={"job_id": job.job_id, "case_id": str(case_id), "status": job.status})


@router.get("/api/reviews/jobs/{job_id}")
def get_review_job(request: Request, job_id: UUID):
    storage(request)
    with _JOBS_LOCK:
        job = _JOBS.get(str(job_id))
    if not job:
        raise HTTPException(404, "Trabajo de revisión no encontrado.")
    return job.snapshot()


@router.get("/api/reviews/jobs/{job_id}/events")
async def stream_review_job_events(request: Request, job_id: UUID):
    storage(request)
    with _JOBS_LOCK:
        job = _JOBS.get(str(job_id))
    if not job:
        raise HTTPException(404, "Trabajo de revisión no encontrado.")

    async def event_generator():
        sent = 0
        heartbeat = 0
        while True:
            if await request.is_disconnected():
                break
            events = job.events
            while sent < len(events):
                ev = events[sent]
                sent += 1
                yield f"id: {ev['seq']}\nevent: progress\ndata: {json.dumps(ev, ensure_ascii=False)}\n\n"
            if job.status in {"completed", "failed", "cancelled"}:
                final_payload = {
                    "status": job.status,
                    "result": job.result,
                    "error": job.error,
                    "message": job.message,
                }
                yield f"event: done\ndata: {json.dumps(final_payload, ensure_ascii=False)}\n\n"
                break
            heartbeat += 1
            if heartbeat % 20 == 0:
                yield ": keepalive\n\n"
            await asyncio.sleep(0.35)

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@router.post("/api/reviews/jobs/{job_id}/cancel")
def cancel_review_job(request: Request, job_id: UUID):
    storage(request)
    _check_local_post(request)
    with _JOBS_LOCK:
        job = _JOBS.get(str(job_id))
    if not job:
        raise HTTPException(404, "Trabajo de revisión no encontrado.")
    if job.status not in {"completed", "failed", "cancelled"}:
        job.cancel_event.set()
    return {"job_id": job.job_id, "status": "cancelling" if job.status not in {"completed", "failed", "cancelled"} else job.status}


@router.get("/api/reviews/history")
def list_review_history(request: Request):
    root = storage(request)
    items: list[dict[str, Any]] = []
    for case_dir in root.iterdir():
        if not case_dir.is_dir():
            continue
        res_file = case_dir / "result.json"
        if not res_file.is_file():
            continue
        try:
            data = json.loads(res_file.read_text(encoding="utf-8"))
            principal = next((d["label"] for d in data.get("documents", []) if d.get("role") == "principal"), case_dir.name)
            items.append({
                "case_id": data.get("case_id", case_dir.name),
                "mode": data.get("mode", "excel"),
                "created_at": data.get("created_at", ""),
                "principal_label": principal,
                "category_detected": data.get("category_detected", ""),
                "findings": data.get("findings", 0),
                "counts": data.get("counts", {}),
                "verified_independently": data.get("verified_independently", False),
                "downloads": data.get("downloads", {}),
            })
        except (OSError, json.JSONDecodeError):
            continue
    items.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return {"history": items[:20]}


@router.get("/api/reviews/{case_id}/result")
def get_case_result(request: Request, case_id: UUID):
    root = storage(request)
    path = root / str(case_id) / "result.json"
    if not path.is_file() or path.resolve() != path:
        raise HTTPException(404, "Resultado de expediente no encontrado.")
    return json.loads(path.read_text(encoding="utf-8"))


@router.post("/api/reviews/{mode}")
async def review_upload(request: Request, mode: Literal["excel", "pdf"]):
    """Synchronous review endpoint retained for direct API calls and tests."""
    root = storage(request)
    _check_local_post(request)
    if not RUN_LOCK.acquire(blocking=False):
        raise HTTPException(409, "Ya hay una revisión en ejecución.")
    try:
        case_id, principal_id = await _stage_uploads(request, root, mode)
        return await execute_audit_pipeline(root=root, case_id=case_id, mode=mode, principal_id=principal_id)
    except HTTPException:
        raise
    except AuditExecutionError as exc:
        raise HTTPException(exc.status_code, str(exc)) from None
    except AuditCancelledError as exc:
        raise HTTPException(499, str(exc)) from None
    except ImportError:
        raise HTTPException(503, "Instala el extra documents del backend.") from None
    except Exception:
        raise HTTPException(
            422,
            "No se pudo leer el expediente o generar el informe. Comprueba que los archivos sean legibles y no estén cifrados.",
        ) from None
    finally:
        RUN_LOCK.release()


@router.get("/api/reviews/{case_id}/{artifact_id}/{extension}")
def download_review(request: Request, case_id: UUID, artifact_id: UUID, extension: Literal["docx", "xlsx"]):
    root = storage(request)
    path = root / str(case_id) / "reports" / f"{artifact_id}.{extension}"
    if not path.is_file() or path.resolve() != path:
        raise HTTPException(404, "Informe no encontrado.")
    name = "informe-revision.docx" if extension == "docx" else "avaluo-marcado.xlsx"
    return FileResponse(path, filename=name, headers={"Cache-Control": "no-store", "X-Content-Type-Options": "nosniff"})
