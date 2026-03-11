from __future__ import annotations

from pathlib import Path

import planlock.template_entry_agent as entry_agent_module
from planlock.models import (
    AgentQuestion,
    EntrySessionState,
    FieldCandidate,
    PageOcrResult,
    SheetEntryResult,
    ValueKind,
)
from planlock.template_entry_agent import LangGraphTemplateEntryAgent, answer_from_question


class FakeGraph:
    def __init__(self, results: list[SheetEntryResult]) -> None:
        self._results = list(results)
        self.prompts: list[str] = []

    def invoke(self, state):
        self.prompts.append(state["prompt"])
        return {
            "sheet_name": state["sheet_name"],
            "prompt": state["prompt"],
            "result": self._results.pop(0),
        }


def _session_state(tmp_path: Path) -> EntrySessionState:
    return EntrySessionState(
        job_id="job-123",
        template_sha256="template-sha",
        workbook_path=tmp_path / "output.xlsx",
        ocr_results_path=tmp_path / "ocr_results.json",
        sheet_order=["Data Input"],
    )


def _ocr_results() -> list[PageOcrResult]:
    return [
        PageOcrResult(
            page_number=1,
            summary="Client profile summary",
            raw_text="Client 1 first name: Taylor",
            source_snippets=["Client 1 first name: Taylor"],
            confidence=0.99,
        )
    ]


def _pending_question() -> AgentQuestion:
    return AgentQuestion(
        id="client-1-name",
        sheet_name="Data Input",
        prompt="Which first name should populate client 1?",
        rationale="The OCR appears to contain multiple candidate names.",
        affected_targets=["profile.client_1.first_name"],
    )


def test_advance_reuses_raw_pdf_rereview_answer_before_asking(tmp_path: Path) -> None:
    question = _pending_question()
    initial_result = SheetEntryResult(sheet_name="Data Input", question=question)
    resolved_result = SheetEntryResult(
        sheet_name="Data Input",
        mapped_fields=[
            FieldCandidate(
                target_key="profile.client_1.first_name",
                value="Taylor",
                value_kind=ValueKind.STRING,
                page_number=1,
                source_excerpt="Client 1 first name: Taylor",
                confidence=0.94,
            )
        ],
    )

    agent = object.__new__(LangGraphTemplateEntryAgent)
    agent._graph = FakeGraph([initial_result, resolved_result])  # type: ignore[attr-defined]
    agent._build_prompt = lambda **kwargs: "sheet prompt"  # type: ignore[attr-defined]
    agent._review_question_against_raw_pdf = lambda **kwargs: answer_from_question(  # type: ignore[attr-defined]
        question,
        answer="Taylor",
        source="raw_pdf_review",
    )

    session_state = _session_state(tmp_path)
    updated_state = agent.advance(session_state, _ocr_results())

    assert updated_state.completed is True
    assert updated_state.pending_question is None
    assert len(updated_state.sheet_results) == 1
    assert updated_state.sheet_results[0].mapped_fields[0].value == "Taylor"
    assert len(updated_state.user_answers) == 1
    assert updated_state.user_answers[0].source == "raw_pdf_review"
    assert updated_state.user_answers[0].answer == "Taylor"
    assert updated_state.questions_asked == []
    assert len(agent._graph.prompts) == 2  # type: ignore[attr-defined]


def test_advance_marks_question_after_raw_pdf_rereview_when_no_answer_found(tmp_path: Path) -> None:
    initial_result = SheetEntryResult(sheet_name="Data Input", question=_pending_question())

    agent = object.__new__(LangGraphTemplateEntryAgent)
    agent._graph = FakeGraph([initial_result])  # type: ignore[attr-defined]
    agent._build_prompt = lambda **kwargs: "sheet prompt"  # type: ignore[attr-defined]
    agent._review_question_against_raw_pdf = lambda **kwargs: None  # type: ignore[attr-defined]

    session_state = _session_state(tmp_path)
    updated_state = agent.advance(session_state, _ocr_results())

    assert updated_state.completed is False
    assert updated_state.pending_question is not None
    assert updated_state.pending_question.pdf_rereviewed is True
    assert len(updated_state.questions_asked) == 1
    assert updated_state.questions_asked[0].pdf_rereviewed is True
    assert len(agent._graph.prompts) == 1  # type: ignore[attr-defined]


def test_invoke_structured_with_retries_uses_shared_retry_helper(monkeypatch) -> None:
    observed: dict[str, object] = {}

    class FakeLLM:
        def invoke(self, *, schema, messages, operation_name=None):
            observed["schema"] = schema
            observed["messages"] = messages
            observed["operation_name"] = operation_name
            return SheetEntryResult(sheet_name="Data Input")

    def fake_invoke_with_retries(settings, operation_name, invoke_fn, retry_notifier=None):
        observed["settings"] = settings
        observed["operation_name"] = operation_name
        observed["retry_notifier"] = retry_notifier
        return invoke_fn()

    monkeypatch.setattr(entry_agent_module, "invoke_with_retries", fake_invoke_with_retries)

    agent = object.__new__(LangGraphTemplateEntryAgent)
    agent._settings = object()
    agent._llm = FakeLLM()

    notifier = object()
    result = agent._invoke_structured_with_retries(
        operation_name="Workbook entry for Data Input",
        schema=SheetEntryResult,
        messages=["message"],
        retry_notifier=notifier,
    )

    assert isinstance(result, SheetEntryResult)
    assert observed["schema"] is SheetEntryResult
    assert observed["settings"] is agent._settings
    assert observed["operation_name"] == "Workbook entry for Data Input"
    assert observed["retry_notifier"] is notifier
    assert observed["messages"] == ["message"]
