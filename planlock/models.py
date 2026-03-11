from __future__ import annotations

from enum import Enum
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field


class Stage(str, Enum):
    OCR = "OCR"
    DATA_ENTRY = "Data Entry"
    FINANCIAL_CALCULATIONS = "Financial Calculations"


class Severity(str, Enum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


class ValueKind(str, Enum):
    STRING = "string"
    NUMBER = "number"
    DATE = "date"
    BOOLEAN = "boolean"
    FORMULA = "formula"


class ExtractedFigure(BaseModel):
    label: str
    value: str
    units: str | None = None
    qualifier: str | None = None
    source_excerpt: str
    confidence: float = Field(ge=0.0, le=1.0)


class ExtractedTable(BaseModel):
    title: str
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class PageOcrResult(BaseModel):
    page_number: int
    summary: str
    raw_text: str
    source_snippets: list[str] = Field(default_factory=list)
    figures: list[ExtractedFigure] = Field(default_factory=list)
    tables: list[ExtractedTable] = Field(default_factory=list)
    recommendations: list[str] = Field(default_factory=list)
    confidence: float = Field(ge=0.0, le=1.0)


class FieldCandidate(BaseModel):
    target_key: str
    value: str | float | int | bool | None
    value_kind: ValueKind
    page_number: int | None = None
    source_excerpt: str
    confidence: float = Field(ge=0.0, le=1.0)
    comment: str | None = None


class ExpenseCandidate(BaseModel):
    category: str
    label: str = "PlanLock Import"
    monthly_amount: float | None = None
    yearly_amount: float | None = None
    discretionary: bool | None = None
    page_number: int | None = None
    source_excerpt: str
    confidence: float = Field(ge=0.0, le=1.0)
    comment: str | None = None


class AccountCandidate(BaseModel):
    account_type: str | None = None
    owner_name: str | None = None
    account_identifier: str | None = None
    apy: float | None = None
    institution: str | None = None
    balance: float | None = None
    monthly_contribution: float | None = None
    last_updated: str | None = None
    notes: str | None = None
    page_number: int | None = None
    source_excerpt: str
    confidence: float = Field(ge=0.0, le=1.0)


class HoldingCandidate(BaseModel):
    sheet_name: Literal["Retirement Accounts", "Taxable Accounts", "Education Accounts"]
    owner_section: Literal["client_1", "client_2", "education"]
    account_name: str
    holding_name: str
    symbol: str | None = None
    category: str | None = None
    expense_ratio: float | None = None
    yield_pct: float | None = None
    one_year_return_pct: float | None = None
    five_year_return_pct: float | None = None
    shares: float | None = None
    price: float | None = None
    purchase_price: float | None = None
    page_number: int | None = None
    source_excerpt: str
    confidence: float = Field(ge=0.0, le=1.0)
    comment: str | None = None


class PageMappingResult(BaseModel):
    page_number: int
    mapped_fields: list[FieldCandidate] = Field(default_factory=list)
    expenses: list[ExpenseCandidate] = Field(default_factory=list)
    accounts: list[AccountCandidate] = Field(default_factory=list)
    holdings: list[HoldingCandidate] = Field(default_factory=list)
    unmapped_items: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class QuestionOption(BaseModel):
    label: str
    value: str
    description: str | None = None


class AgentQuestion(BaseModel):
    id: str
    sheet_name: str
    prompt: str
    rationale: str
    affected_targets: list[str] = Field(default_factory=list)
    options: list[QuestionOption] = Field(default_factory=list)
    allow_free_text: bool = True
    pdf_rereviewed: bool = False


class AgentAnswer(BaseModel):
    question_id: str
    sheet_name: str
    answer: str
    source: Literal["option", "free_text", "agent", "raw_pdf_review"] = "option"
    affected_targets: list[str] = Field(default_factory=list)


class SheetEntryResult(BaseModel):
    sheet_name: str
    mapped_fields: list[FieldCandidate] = Field(default_factory=list)
    expenses: list[ExpenseCandidate] = Field(default_factory=list)
    accounts: list[AccountCandidate] = Field(default_factory=list)
    holdings: list[HoldingCandidate] = Field(default_factory=list)
    unresolved_supported_targets: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    question: AgentQuestion | None = None


class ImportWarning(BaseModel):
    code: str
    message: str
    severity: Severity
    stage: Stage
    page_numbers: list[int] = Field(default_factory=list)
    details: dict[str, Any] = Field(default_factory=dict)


class CanonicalPlanDocument(BaseModel):
    fields: dict[str, FieldCandidate] = Field(default_factory=dict)
    expenses: list[ExpenseCandidate] = Field(default_factory=list)
    accounts: list[AccountCandidate] = Field(default_factory=list)
    holdings: list[HoldingCandidate] = Field(default_factory=list)
    unmapped_items: list[str] = Field(default_factory=list)
    assumptions: list[str] = Field(default_factory=list)


class CellAssignment(BaseModel):
    sheet_name: str
    cell: str
    value: str | float | int | bool | None
    value_kind: ValueKind
    semantic_key: str
    source_pages: list[int] = Field(default_factory=list)
    comment: str | None = None


class TemplateDriftCheckResult(BaseModel):
    passed: bool
    violations: list[str] = Field(default_factory=list)


class CalculationValidationResult(BaseModel):
    passed: bool
    recalc_attempted: bool
    recalc_completed: bool
    formula_cells_template: int
    formula_cells_output: int
    warnings: list[str] = Field(default_factory=list)


class SheetEntrySummary(BaseModel):
    sheet_name: str
    status: Literal["completed", "needs_input", "skipped"] = "completed"
    mapped_count: int = 0
    unresolved_count: int = 0
    touched_cells: list[str] = Field(default_factory=list)
    message: str | None = None


class CoverageSummary(BaseModel):
    supported_target_count: int = 0
    unresolved_supported_target_count: int = 0
    coverage_ratio: float = Field(default=1.0, ge=0.0, le=1.0)
    critical_sheet_names: list[str] = Field(default_factory=list)
    unresolved_critical_sheet_names: list[str] = Field(default_factory=list)


class EntrySessionState(BaseModel):
    job_id: str
    template_sha256: str
    workbook_path: Path
    ocr_results_path: Path
    current_sheet_index: int = 0
    sheet_order: list[str] = Field(default_factory=list)
    pending_question: AgentQuestion | None = None
    questions_asked: list[AgentQuestion] = Field(default_factory=list)
    user_answers: list[AgentAnswer] = Field(default_factory=list)
    sheet_results: list[SheetEntryResult] = Field(default_factory=list)
    sheet_summaries: list[SheetEntrySummary] = Field(default_factory=list)
    mapped_assignments: list[CellAssignment] = Field(default_factory=list)
    unmapped_items: list[str] = Field(default_factory=list)
    assumptions: list[str] = Field(default_factory=list)
    coverage_summary: CoverageSummary = Field(default_factory=CoverageSummary)
    review_required_reasons: list[str] = Field(default_factory=list)
    completed: bool = False


class ReviewReport(BaseModel):
    job_id: str
    template_sha256: str
    success: bool
    warnings: list[ImportWarning] = Field(default_factory=list)
    mapped_assignments: list[CellAssignment] = Field(default_factory=list)
    unmapped_items: list[str] = Field(default_factory=list)
    assumptions: list[str] = Field(default_factory=list)
    sheet_summaries: list[SheetEntrySummary] = Field(default_factory=list)
    user_answers: list[AgentAnswer] = Field(default_factory=list)
    questions_asked: list[AgentQuestion] = Field(default_factory=list)
    coverage_summary: CoverageSummary | None = None
    review_required_reasons: list[str] = Field(default_factory=list)
    drift_check: TemplateDriftCheckResult | None = None
    calculation_validation: CalculationValidationResult | None = None


class ImportArtifacts(BaseModel):
    success: bool
    job_id: str
    job_dir: Path
    output_workbook_path: Path | None = None
    review_report_path: Path | None = None
    review_report: ReviewReport | None = None
    entry_state_path: Path | None = None
    ocr_results_path: Path | None = None
    pending_question: AgentQuestion | None = None


class PhaseOneCache(BaseModel):
    source_filename: str
    page_total: int = 0
    ocr_results: list[PageOcrResult] = Field(default_factory=list)


class RunEvent(BaseModel):
    stage: Stage
    message: str
    detail_message: str | None = None
    severity: Severity = Severity.INFO
    stage_completed: int = 0
    stage_total: int = 1
    page_number: int | None = None
    page_total: int | None = None
    pipe_number: int | None = None
    pipe_total: int | None = None
    attempt_number: int | None = None
    max_attempts: int | None = None
    retry_delay_seconds: float | None = None
    retry_reason: Literal["rate_limit", "transient"] | None = None
    phase: Literal["start", "retry", "complete", "failed", "heartbeat", "paused"] | None = None
    artifacts: ImportArtifacts | None = None
