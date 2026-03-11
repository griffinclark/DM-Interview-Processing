from __future__ import annotations

import json
from pathlib import Path
from typing import Protocol, TypedDict

from langchain_core.messages import HumanMessage, SystemMessage
from langgraph.graph import END, StateGraph
from openpyxl import load_workbook
from pydantic import BaseModel, Field

from planlock.config import Settings
from planlock.llm_pipeline import RetryNotifier, StructuredOutputClient, invoke_with_retries
from planlock.models import (
    AgentAnswer,
    AgentQuestion,
    CellAssignment,
    CoverageSummary,
    EntrySessionState,
    PageMappingResult,
    PageOcrResult,
    QuestionOption,
    SheetEntryResult,
    SheetEntrySummary,
)
from planlock.template_schema import (
    ALLOWED_WRITE_CELLS_BY_SHEET,
    TEMPLATE_SHEET_ORDER,
    sheet_reference_for_prompt,
)

SHEET_ENTRY_SYSTEM_PROMPT = """
You fill one locked workbook sheet at a time using OCR data, workbook state, and prior user answers.

Rules:
- Only emit candidates for the active workbook sheet.
- You may inspect the full OCR corpus for context, but only map values supported by the active sheet schema.
- Ask a user question only when a material ambiguity blocks a supported target on the active sheet.
- If the user explicitly delegates a decision back to you, choose the most defensible supported value from the available evidence and do not ask the same question again unless the sheet still cannot proceed safely.
- Emit literal values only for directly observed constants or user-provided answers.
- If a value depends on extracted constants, do not emit a computed literal. Emit the source constants and let the workbook formulas compute downstream values inside the sheet.
- Do not duplicate a derived value in both raw and computed form.
- Keep unsupported or still-ambiguous items in unresolved_supported_targets or warnings.
- Return structured output only.
""".strip()


RAW_PDF_REREVIEW_SYSTEM_PROMPT = """
You review raw OCR text before a planner-facing question is shown.

Rules:
- Use only the raw OCR text and source snippets provided in the prompt.
- Try to answer the pending question if the raw OCR contains a defensible answer.
- If the raw OCR still does not support a defensible answer, set answer_found to false.
- Do not invent facts, compute new values, or rely on information outside the prompt.
- Return structured output only.
""".strip()


class TemplateEntryAgent(Protocol):
    def advance(
        self,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        retry_notifier: RetryNotifier | None = None,
    ) -> EntrySessionState:
        ...


class _GraphState(TypedDict):
    sheet_name: str
    prompt: str
    retry_notifier: RetryNotifier | None
    result: SheetEntryResult | None


class RawPdfQuestionReview(BaseModel):
    answer_found: bool = False
    answer: str | None = None
    rationale: str = ""
    source_page_numbers: list[int] = Field(default_factory=list)


def persist_ocr_results(path: Path, ocr_results: list[PageOcrResult]) -> None:
    path.write_text(
        json.dumps([result.model_dump(mode="json") for result in ocr_results], indent=2),
        encoding="utf-8",
    )


def load_ocr_results(path: Path) -> list[PageOcrResult]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    return [PageOcrResult.model_validate(item) for item in payload]


def save_entry_state(path: Path, state: EntrySessionState) -> None:
    path.write_text(json.dumps(state.model_dump(mode="json"), indent=2), encoding="utf-8")


def load_entry_state(path: Path) -> EntrySessionState:
    return EntrySessionState.model_validate_json(path.read_text(encoding="utf-8"))


def read_sheet_context(
    workbook_path: Path,
    sheet_name: str,
    touched_cells: list[str] | None = None,
) -> str:
    workbook = load_workbook(workbook_path, data_only=False)
    sheet = workbook[sheet_name]
    allowed_cells = sorted(ALLOWED_WRITE_CELLS_BY_SHEET.get(sheet_name, set()))
    cells_to_read = list(dict.fromkeys([*(touched_cells or []), *allowed_cells]))
    lines = [f"Workbook context for {sheet_name}:"]
    populated_lines: list[str] = []
    for cell_ref in cells_to_read:
        cell = sheet[cell_ref]
        value = cell.value
        if value in (None, ""):
            continue
        populated_lines.append(f"- {cell_ref} = {value}")
    if populated_lines:
        lines.extend(populated_lines[:200])
    else:
        lines.append("- No writable cells are populated yet.")
    return "\n".join(lines)


def touched_cells_for_assignments(assignments: list[CellAssignment], sheet_name: str) -> list[str]:
    return [assignment.cell for assignment in assignments if assignment.sheet_name == sheet_name]


def coverage_summary_for_state(state: EntrySessionState) -> CoverageSummary:
    supported = 0
    unresolved = 0
    unresolved_critical: list[str] = []
    critical_sheets = ["Data Input", "Expenses"]
    for result in state.sheet_results:
        mapped_count = (
            len(result.mapped_fields)
            + len(result.expenses)
            + len(result.accounts)
            + len(result.holdings)
        )
        unresolved_count = len(result.unresolved_supported_targets)
        supported += mapped_count + unresolved_count
        unresolved += unresolved_count
        if result.sheet_name in critical_sheets and mapped_count == 0 and unresolved_count > 0:
            unresolved_critical.append(result.sheet_name)
    ratio = 1.0 if supported == 0 else (supported - unresolved) / supported
    return CoverageSummary(
        supported_target_count=supported,
        unresolved_supported_target_count=unresolved,
        coverage_ratio=ratio,
        critical_sheet_names=critical_sheets,
        unresolved_critical_sheet_names=sorted(set(unresolved_critical)),
    )


def _upsert_sheet_summary(state: EntrySessionState, summary: SheetEntrySummary) -> None:
    for index, existing in enumerate(state.sheet_summaries):
        if existing.sheet_name == summary.sheet_name:
            state.sheet_summaries[index] = summary
            return
    state.sheet_summaries.append(summary)


def sheet_result_to_page_mapping_result(result: SheetEntryResult) -> PageMappingResult:
    return PageMappingResult(
        page_number=0,
        mapped_fields=result.mapped_fields,
        expenses=result.expenses,
        accounts=result.accounts,
        holdings=result.holdings,
        unmapped_items=result.unresolved_supported_targets,
        warnings=result.warnings,
    )


class LangGraphTemplateEntryAgent:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._llm = StructuredOutputClient(settings, model=settings.model_mapping)
        graph = StateGraph(_GraphState)
        graph.add_node("draft_sheet", self._draft_sheet)
        graph.set_entry_point("draft_sheet")
        graph.add_edge("draft_sheet", END)
        self._graph = graph.compile()

    def _invoke_structured_with_retries(
        self,
        *,
        operation_name: str,
        schema,
        messages,
        retry_notifier: RetryNotifier | None = None,
    ):
        return invoke_with_retries(
            self._settings,
            operation_name,
            lambda: self._llm.invoke(
                schema=schema,
                messages=messages,
                operation_name=operation_name,
            ),
            retry_notifier=retry_notifier,
        )

    def _draft_sheet(self, state: _GraphState) -> _GraphState:
        result = self._invoke_structured_with_retries(
            operation_name=f"Workbook entry for {state['sheet_name']}",
            schema=SheetEntryResult,
            messages=[
                SystemMessage(content=SHEET_ENTRY_SYSTEM_PROMPT),
                HumanMessage(content=state["prompt"]),
            ],
            retry_notifier=state.get("retry_notifier"),
        )
        result.sheet_name = state["sheet_name"]
        return {
            "sheet_name": state["sheet_name"],
            "prompt": state["prompt"],
            "retry_notifier": state.get("retry_notifier"),
            "result": result,
        }

    def _build_prompt(
        self,
        *,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        sheet_name: str,
    ) -> str:
        explicit_answers_payload = [
            answer.model_dump(mode="json")
            for answer in session_state.user_answers
            if answer.sheet_name == sheet_name and answer.source in {"option", "free_text"}
        ]
        agent_answers_payload = [
            answer.model_dump(mode="json")
            for answer in session_state.user_answers
            if answer.sheet_name == sheet_name and answer.source in {"agent", "raw_pdf_review"}
        ]
        prior_assignments = [
            assignment.model_dump(mode="json")
            for assignment in session_state.mapped_assignments
            if assignment.sheet_name == sheet_name
        ]
        workbook_context = read_sheet_context(
            session_state.workbook_path,
            sheet_name,
            touched_cells=touched_cells_for_assignments(session_state.mapped_assignments, sheet_name),
        )
        return (
            f"Active workbook sheet: {sheet_name}\n\n"
            f"{sheet_reference_for_prompt(sheet_name)}\n\n"
            f"{workbook_context}\n\n"
            "Confirmed user answers for this sheet:\n"
            f"{json.dumps(explicit_answers_payload, indent=2)}\n\n"
            "Agent-sourced decisions for this sheet:\n"
            f"{json.dumps(agent_answers_payload, indent=2)}\n\n"
            "Agent-sourced decisions can come from a prior raw PDF re-review or an explicit user delegation. If they are present, resolve those targets yourself and avoid asking the same question again unless the sheet still cannot proceed safely.\n\n"
            "Prior mapped assignments already written for this sheet:\n"
            f"{json.dumps(prior_assignments, indent=2)}\n\n"
            "Full OCR output JSON for the document:\n"
            f"{json.dumps([result.model_dump(mode='json') for result in ocr_results], indent=2)}"
        )

    def _build_raw_pdf_rereview_prompt(
        self,
        *,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        sheet_name: str,
        question: AgentQuestion,
    ) -> str:
        explicit_answers_payload = [
            answer.model_dump(mode="json")
            for answer in session_state.user_answers
            if answer.sheet_name == sheet_name and answer.source in {"option", "free_text"}
        ]
        workbook_context = read_sheet_context(
            session_state.workbook_path,
            sheet_name,
            touched_cells=touched_cells_for_assignments(session_state.mapped_assignments, sheet_name),
        )
        raw_ocr_payload = [
            {
                "page_number": result.page_number,
                "raw_text": result.raw_text,
                "source_snippets": result.source_snippets,
            }
            for result in ocr_results
        ]
        return (
            f"Active workbook sheet: {sheet_name}\n\n"
            f"{sheet_reference_for_prompt(sheet_name)}\n\n"
            f"{workbook_context}\n\n"
            "Prior explicit user answers for this sheet:\n"
            f"{json.dumps(explicit_answers_payload, indent=2)}\n\n"
            "Pending question JSON:\n"
            f"{json.dumps(question.model_dump(mode='json'), indent=2)}\n\n"
            "Raw OCR text and source snippets for the full document:\n"
            f"{json.dumps(raw_ocr_payload, indent=2)}\n\n"
            "If the raw OCR gives a defensible answer to the pending question, return answer_found=true and provide the answer text exactly as it should be used. If not, return answer_found=false."
        )

    def _review_question_against_raw_pdf(
        self,
        *,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        sheet_name: str,
        question: AgentQuestion,
        retry_notifier: RetryNotifier | None = None,
    ) -> AgentAnswer | None:
        review = self._invoke_structured_with_retries(
            operation_name=f"Raw PDF re-review for {sheet_name}",
            schema=RawPdfQuestionReview,
            messages=[
                SystemMessage(content=RAW_PDF_REREVIEW_SYSTEM_PROMPT),
                HumanMessage(
                    content=self._build_raw_pdf_rereview_prompt(
                        session_state=session_state,
                        ocr_results=ocr_results,
                        sheet_name=sheet_name,
                        question=question,
                    )
                ),
            ],
            retry_notifier=retry_notifier,
        )
        if not review.answer_found or review.answer is None or not review.answer.strip():
            return None
        return answer_from_question(
            question,
            answer=review.answer.strip(),
            source="raw_pdf_review",
        )

    @staticmethod
    def _sheet_has_raw_pdf_rereview_answer(
        session_state: EntrySessionState,
        sheet_name: str,
    ) -> bool:
        return any(
            answer.sheet_name == sheet_name and answer.source == "raw_pdf_review"
            for answer in session_state.user_answers
        )

    def advance(
        self,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        retry_notifier: RetryNotifier | None = None,
    ) -> EntrySessionState:
        return self._advance(
            session_state,
            ocr_results,
            allow_raw_pdf_rereview=True,
            retry_notifier=retry_notifier,
        )

    def _advance(
        self,
        session_state: EntrySessionState,
        ocr_results: list[PageOcrResult],
        *,
        allow_raw_pdf_rereview: bool,
        retry_notifier: RetryNotifier | None,
    ) -> EntrySessionState:
        if session_state.current_sheet_index >= len(session_state.sheet_order):
            session_state.pending_question = None
            session_state.completed = True
            session_state.coverage_summary = coverage_summary_for_state(session_state)
            return session_state

        sheet_name = session_state.sheet_order[session_state.current_sheet_index]
        if not ALLOWED_WRITE_CELLS_BY_SHEET.get(sheet_name):
            _upsert_sheet_summary(
                session_state,
                SheetEntrySummary(
                    sheet_name=sheet_name,
                    status="skipped",
                    message="No writable targets configured for this sheet.",
                )
            )
            session_state.current_sheet_index += 1
            session_state.coverage_summary = coverage_summary_for_state(session_state)
            session_state.completed = session_state.current_sheet_index >= len(session_state.sheet_order)
            return session_state

        graph_state = self._graph.invoke(
            {
                "sheet_name": sheet_name,
                "prompt": self._build_prompt(
                    session_state=session_state,
                    ocr_results=ocr_results,
                    sheet_name=sheet_name,
                ),
                "retry_notifier": retry_notifier,
                "result": None,
            }
        )
        result = graph_state["result"]
        if result is None:
            raise RuntimeError(f"Template entry graph did not produce a result for {sheet_name}.")

        result.sheet_name = sheet_name
        if result.question is not None and allow_raw_pdf_rereview:
            reviewed_answer = self._review_question_against_raw_pdf(
                session_state=session_state,
                ocr_results=ocr_results,
                sheet_name=sheet_name,
                question=result.question,
                retry_notifier=retry_notifier,
            )
            if reviewed_answer is not None:
                session_state.user_answers.append(reviewed_answer)
                return self._advance(
                    session_state,
                    ocr_results,
                    allow_raw_pdf_rereview=False,
                    retry_notifier=retry_notifier,
                )
            result.question.pdf_rereviewed = True

        if (
            result.question is not None
            and self._sheet_has_raw_pdf_rereview_answer(session_state, sheet_name)
        ):
            result.question.pdf_rereviewed = True

        if result.question is not None:
            session_state.pending_question = result.question
            session_state.questions_asked.append(result.question)
            _upsert_sheet_summary(
                session_state,
                SheetEntrySummary(
                    sheet_name=sheet_name,
                    status="needs_input",
                    mapped_count=(
                        len(result.mapped_fields)
                        + len(result.expenses)
                        + len(result.accounts)
                        + len(result.holdings)
                    ),
                    unresolved_count=len(result.unresolved_supported_targets),
                    message=result.question.prompt,
                )
            )
            session_state.coverage_summary = coverage_summary_for_state(session_state)
            return session_state

        session_state.pending_question = None
        session_state.sheet_results.append(result)
        _upsert_sheet_summary(
            session_state,
            SheetEntrySummary(
                sheet_name=sheet_name,
                status="completed",
                mapped_count=(
                    len(result.mapped_fields)
                    + len(result.expenses)
                    + len(result.accounts)
                    + len(result.holdings)
                ),
                unresolved_count=len(result.unresolved_supported_targets),
                message="Sheet entry completed.",
            )
        )
        session_state.current_sheet_index += 1
        session_state.coverage_summary = coverage_summary_for_state(session_state)
        session_state.completed = session_state.current_sheet_index >= len(session_state.sheet_order)
        return session_state


def default_sheet_order() -> list[str]:
    return list(TEMPLATE_SHEET_ORDER)


def answer_from_question(
    question: AgentQuestion,
    *,
    answer: str,
    source: str,
) -> AgentAnswer:
    if source == "free_text":
        normalized_source = "free_text"
    elif source == "agent":
        normalized_source = "agent"
    elif source == "raw_pdf_review":
        normalized_source = "raw_pdf_review"
    else:
        normalized_source = "option"
    return AgentAnswer(
        question_id=question.id,
        sheet_name=question.sheet_name,
        answer=answer,
        source=normalized_source,
        affected_targets=question.affected_targets,
    )


def fallback_question(
    *,
    sheet_name: str,
    prompt: str,
    rationale: str,
    affected_targets: list[str],
    options: list[tuple[str, str, str]],
) -> AgentQuestion:
    return AgentQuestion(
        id=f"{sheet_name.lower().replace(' ', '_')}-question",
        sheet_name=sheet_name,
        prompt=prompt,
        rationale=rationale,
        affected_targets=affected_targets,
        options=[
            QuestionOption(label=label, value=value, description=description)
            for label, value, description in options
        ],
    )
