from pathlib import Path
from typing import Protocol
from uuid import UUID
from .models import AuditResult

class AgentProvider(Protocol):
    async def run(self, run_id: UUID, workspace: Path) -> AuditResult:
        """Return unverified results; caller is responsible for independent verification."""
        ...
    async def cancel(self, run_id: UUID) -> None: ...

class CodexProvider:
    async def run(self, run_id: UUID, workspace: Path) -> AuditResult:
        raise NotImplementedError("Codex CLI no est? habilitado en esta etapa")
    async def cancel(self, run_id: UUID) -> None:
        raise NotImplementedError("No hay ejecuciones habilitadas")
