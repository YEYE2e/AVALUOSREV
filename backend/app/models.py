from datetime import datetime
from enum import StrEnum
from typing import Literal
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, HttpUrl, model_validator


class Contract(BaseModel):
    model_config = ConfigDict(extra="forbid")


class RunStatus(StrEnum):
    queued = "queued"
    indexing = "indexing"
    analyzing = "analyzing"
    verifying = "verifying"
    exporting = "exporting"
    running = "running"
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
            raise ValueError("La afirmación requiere evidencia")
        return self


class AuditResult(Contract):
    findings: list[Finding] = Field(default_factory=list)
    limitations: list[str] = Field(default_factory=list)


SeverityLevel = Literal["critico", "importante", "menor"]
FindingStatus = Literal[
    "sustentado",
    "requiere_verificacion",
    "diferencia_explicada",
    "descartado",
]
ValueImpact = Literal[
    "demostrado",
    "posible",
    "ninguno_demostrado",
    "no_determinable",
]
DocumentaryVerdict = Literal[
    "concuerda",
    "existen_diferencias",
    "parcialmente_verificable",
    "no_verificable_sin_respaldos",
    "no_aplica",
]
ConclusionCategory = Literal[
    "conforme_para_revision_del_perito",
    "conforme_con_correcciones_menores",
    "requiere_correcciones_tecnicas",
    "requiere_documentacion_o_verificacion",
    "no_recomendable_en_estado_actual",
]


class ExecutiveSummary(BaseModel):
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    errors_by_type: str = Field(default="", max_length=800)
    main_inconsistencies: str = Field(default="", max_length=1200)
    value_impact_risks: str = Field(default="", max_length=1000)
    market_references_to_review: str = Field(default="", max_length=1000)
    missing_documentation: str = Field(default="", max_length=1000)
    calculation_errors: str = Field(default="", max_length=1000)
    commercial_and_realization_differences: str = Field(default="", max_length=1000)
    banking_review_observations: str = Field(default="", max_length=1200)


class StructuredFinding(BaseModel):
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    id: str = Field(default="", max_length=32)
    severity: SeverityLevel = "importante"
    review_type: str = Field(default="Revisión técnica", min_length=1, max_length=120)
    location: str = Field(min_length=1, max_length=240)
    sheet: str | None = Field(default=None, max_length=31)
    cell: str | None = Field(default=None, pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
    page: int | None = Field(default=None, ge=1)
    field_reviewed: str = Field(default="", max_length=200)
    observation: str = Field(min_length=1, max_length=800)
    technical_explanation: str = Field(default="", max_length=1200)
    current_value: str | None = Field(default=None, max_length=300)
    expected_value: str | None = Field(default=None, max_length=300)
    difference: str | None = Field(default=None, max_length=300)
    evidence: str = Field(min_length=1, max_length=1200)
    recommendation: str = Field(min_length=1, max_length=1000)
    value_impact: ValueImpact = "no_determinable"
    status: FindingStatus = "requiere_verificacion"
    verification_note: str | None = Field(default=None, max_length=800)


class AgentReviewDraft(BaseModel):
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    category_detected: str = Field(default="No especificada", max_length=160)
    summary: str = Field(min_length=1, max_length=3000)
    documentary_verdict: DocumentaryVerdict = "no_aplica"
    conclusion_category: ConclusionCategory = "requiere_documentacion_o_verificacion"
    conclusion_detail: str = Field(default="", max_length=1500)
    executive_summary: ExecutiveSummary | None = None
    findings: list[StructuredFinding] = Field(default_factory=list, max_length=150)
    limitations: list[str] = Field(default_factory=list, max_length=60)


class VerifierDecision(BaseModel):
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    finding_id: str = Field(min_length=1, max_length=32)
    decision: FindingStatus
    severity: SeverityLevel | None = None
    value_impact: ValueImpact | None = None
    sheet: str | None = Field(default=None, max_length=31)
    cell: str | None = Field(default=None, pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
    page: int | None = Field(default=None, ge=1)
    evidence_verified: str | None = Field(default=None, max_length=1200)
    verification_note: str = Field(min_length=1, max_length=800)


class VerifierOutput(BaseModel):
    model_config = ConfigDict(extra="ignore", str_strip_whitespace=True)

    verification_summary: str = Field(min_length=1, max_length=2000)
    decisions: list[VerifierDecision] = Field(default_factory=list, max_length=160)
    additional_findings: list[StructuredFinding] = Field(default_factory=list, max_length=40)
    additional_limitations: list[str] = Field(default_factory=list, max_length=40)


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
