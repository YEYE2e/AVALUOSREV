from datetime import datetime
from enum import StrEnum
from uuid import UUID
from pydantic import BaseModel, ConfigDict, Field, HttpUrl, model_validator

class Contract(BaseModel):
    model_config = ConfigDict(extra="forbid")

class RunStatus(StrEnum):
    queued = "queued"
    running = "running"
    verifying = "verifying"
    completed = "completed"
    failed = "failed"
    cancelled = "cancelled"

class Evidence(Contract):
    document_id: UUID | None = None
    url: HttpUrl | None = None
    page: int | None = Field(default=None, ge=1)
    sheet: str | None = None
    cell: str | None = None
    excerpt: str | None = None
    @model_validator(mode="after")
    def validate_source(self):
        if (self.document_id is None) == (self.url is None):
            raise ValueError("Indique exactamente un documento o una URL")
        if self.url and (self.page or self.sheet or self.cell):
            raise ValueError("La URL no admite localizadores de documento")
        if self.cell and not self.sheet:
            raise ValueError("La celda requiere una hoja")
        return self

class Finding(Contract):
    category: str = Field(pattern="^(confirmed|discrepancy|unverifiable|human_review)$")
    severity: str = Field(pattern="^(low|medium|high)$")
    description: str = Field(min_length=1)
    evidence: list[Evidence] = Field(default_factory=list)
    @model_validator(mode="after")
    def evidence_for_claim(self):
        if self.category in ("confirmed", "discrepancy") and not self.evidence:
            raise ValueError("La afirmaci?n requiere evidencia")
        return self

class AuditResult(Contract):
    findings: list[Finding] = Field(default_factory=list)
    limitations: list[str] = Field(default_factory=list)

class Case(Contract):
    id: UUID
    name: str
    created_at: datetime

class Document(Contract):
    id: UUID
    case_id: UUID
    original_name: str
    sha256: str
    size_bytes: int = Field(ge=0)

class AuditRun(Contract):
    id: UUID
    case_id: UUID
    status: RunStatus
    provider: str
    model: str | None = None
    created_at: datetime
    error: str | None = None
    usage: dict[str, int] | None = None

class RunEvent(Contract):
    run_id: UUID
    sequence: int = Field(ge=1)
    timestamp: datetime
    type: str
    message: str
