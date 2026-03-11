from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from queue import Queue
import threading
import time

import fitz
from openpyxl import load_workbook
import pytest

from planlock.config import Settings
from planlock.job_runner import JobRunner
from planlock.models import (
    AgentQuestion,
    ExpenseCandidate,
    FieldCandidate,
    PageOcrResult,
    QuestionOption,
    SheetEntryResult,
    Stage,
    ValueKind,
)


class FakeExtractionClient:
    def ocr_page(self, page, retry_notifier=None):
        return PageOcrResult(
            page_number=page.page_number,
            summary="Synthetic OCR result",
            raw_text="Target personal spending of $1,250 per month for travel.",
            source_snippets=["Travel target $1,250"],
            figures=[],
            tables=[],
            recommendations=["Use suggested travel target."],
            confidence=0.99,
        )


class FakeRateLimitError(Exception):
    status_code = 429


class SilentRateLimitError(Exception):
    status_code = 429


class TrackingTemplateEntryAgent:
    def __init__(self) -> None:
        self.sheet_calls: list[str] = []

    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        sheet_name = state.sheet_order[state.current_sheet_index]
        self.sheet_calls.append(sheet_name)
        if sheet_name == "Data Input":
            state.sheet_results.append(
                SheetEntryResult(
                    sheet_name=sheet_name,
                    mapped_fields=[
                        FieldCandidate(
                            target_key="profile.client_1.first_name",
                            value="Taylor",
                            value_kind=ValueKind.STRING,
                            page_number=1,
                            source_excerpt="Taylor",
                            confidence=0.95,
                        )
                    ],
                )
            )
        elif sheet_name == "Expenses":
            state.sheet_results.append(
                SheetEntryResult(
                    sheet_name=sheet_name,
                    expenses=[
                        ExpenseCandidate(
                            category="Travel",
                            monthly_amount=1250.0,
                            label="Imported Target",
                            page_number=1,
                            source_excerpt="Travel target $1,250",
                            confidence=0.96,
                            comment="Suggested target chosen over current value.",
                        )
                    ],
                )
            )
        state.current_sheet_index += 1
        state.completed = state.current_sheet_index >= len(state.sheet_order)
        return state


class QuestioningTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        sheet_name = state.sheet_order[state.current_sheet_index]
        self.sheet_calls.append(sheet_name)
        if sheet_name == "Data Input" and not any(answer.sheet_name == "Data Input" for answer in state.user_answers):
            state.pending_question = AgentQuestion(
                id="data-input-name",
                sheet_name=sheet_name,
                prompt="Which first name should populate client 1?",
                rationale="The OCR shows two possible names for client 1.",
                affected_targets=["profile.client_1.first_name"],
                options=[
                    QuestionOption(label="Taylor", value="Taylor"),
                    QuestionOption(label="Tyler", value="Tyler"),
                ],
            )
            return state

        if sheet_name == "Data Input":
            chosen_name = next(answer.answer for answer in state.user_answers if answer.sheet_name == "Data Input")
            state.sheet_results.append(
                SheetEntryResult(
                    sheet_name=sheet_name,
                    mapped_fields=[
                        FieldCandidate(
                            target_key="profile.client_1.first_name",
                            value=chosen_name,
                            value_kind=ValueKind.STRING,
                            page_number=1,
                            source_excerpt=chosen_name,
                            confidence=0.95,
                        )
                    ],
                )
            )
        elif sheet_name == "Expenses":
            state.sheet_results.append(
                SheetEntryResult(
                    sheet_name=sheet_name,
                    expenses=[
                        ExpenseCandidate(
                            category="Travel",
                            monthly_amount=1250.0,
                            label="Imported Target",
                            page_number=1,
                            source_excerpt="Travel target $1,250",
                            confidence=0.96,
                        )
                    ],
                )
            )
        state.current_sheet_index += 1
        state.completed = state.current_sheet_index >= len(state.sheet_order)
        return state


class LowCoverageTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        sheet_name = state.sheet_order[state.current_sheet_index]
        self.sheet_calls.append(sheet_name)
        if sheet_name == "Data Input":
            state.sheet_results.append(
                SheetEntryResult(
                    sheet_name=sheet_name,
                    unresolved_supported_targets=["profile.client_1.first_name"],
                    warnings=["Need user confirmation before writing profile values."],
                )
            )
        state.current_sheet_index += 1
        state.completed = state.current_sheet_index >= len(state.sheet_order)
        return state


class RetryingTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        if retry_notifier is not None:
            retry_notifier(
                "Workbook entry for Data Input",
                2,
                7,
                10.0,
                FakeRateLimitError(
                    "Error code: 429 - {'type': 'error', 'error': {'type': 'rate_limit_error'}}"
                ),
            )
        time.sleep(0.2)
        return super().advance(state, ocr_results, retry_notifier=retry_notifier)


class SilentRetryingTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        if retry_notifier is not None:
            retry_notifier(
                "Workbook entry for Data Input",
                2,
                7,
                10.0,
                SilentRateLimitError("HTTP 429"),
            )
        time.sleep(0.2)
        return super().advance(state, ocr_results, retry_notifier=retry_notifier)


class TokenReportingTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        if usage_notifier is not None:
            usage_notifier(
                "Workbook entry for Data Input",
                {
                    "input_tokens": 9,
                    "output_tokens": 6,
                    "total_tokens": 15,
                },
            )
        time.sleep(0.2)
        return super().advance(state, ocr_results, retry_notifier=retry_notifier)


class ProgressReportingTemplateEntryAgent(TrackingTemplateEntryAgent):
    def advance(self, state, ocr_results, retry_notifier=None, usage_notifier=None, progress_notifier=None):
        if progress_notifier is not None:
            progress_notifier("Workbook entry for Data Input", "Reviewing household profile evidence.")
            progress_notifier("Workbook entry for Data Input", "Reviewing household profile evidence.")
            progress_notifier(
                "Workbook entry for Data Input",
                "Matching workbook fields to the strongest OCR evidence.",
            )
        time.sleep(0.2)
        return super().advance(state, ocr_results, retry_notifier=retry_notifier)


class StaggeredConcurrencyExtractionClient(FakeExtractionClient):
    def __init__(self, page_delays: dict[int, float]) -> None:
        self._lock = threading.Lock()
        self.active_calls = 0
        self.peak_calls = 0
        self.page_delays = page_delays

    def ocr_page(self, page, retry_notifier=None):
        with self._lock:
            self.active_calls += 1
            self.peak_calls = max(self.peak_calls, self.active_calls)
        try:
            time.sleep(self.page_delays.get(page.page_number, 0.05))
            return FakeExtractionClient.ocr_page(self, page, retry_notifier=retry_notifier)
        finally:
            with self._lock:
                self.active_calls -= 1


class SlowExtractionClient(FakeExtractionClient):
    def ocr_page(self, page, retry_notifier=None):
        time.sleep(0.25)
        return super().ocr_page(page, retry_notifier=retry_notifier)


class FailingExtractionClient:
    def ocr_page(self, page, retry_notifier=None):
        raise RuntimeError("OCR page 1: 1 validation error for PageOcrResult")


def _make_pdf_bytes(page_count: int = 1) -> bytes:
    document = fitz.open()
    for page_number in range(1, page_count + 1):
        page = document.new_page()
        page.insert_text(
            (72, 72),
            f"Page {page_number}: Target personal spending of $1,250 per month for travel.",
        )
    return document.tobytes()


def _settings(tmp_path: Path, **kwargs) -> Settings:
    settings = replace(
        Settings.from_env(),
        jobs_dir=tmp_path / "jobs",
        **kwargs,
    )
    settings.ensure_runtime_dirs()
    return settings


def test_job_runner_emits_all_three_stages_and_keeps_derived_values_as_formulas(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    agent = TrackingTemplateEntryAgent()
    runner = JobRunner(settings, extraction_client=FakeExtractionClient(), entry_agent=agent)
    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))

    stages = [event.stage for event in events]
    assert Stage.OCR in stages
    assert Stage.DATA_ENTRY in stages
    assert Stage.FINANCIAL_CALCULATIONS in stages
    assert events[-1].artifacts is not None
    assert events[-1].artifacts.review_report is not None
    assert events[-1].artifacts.review_report.success
    workbook = load_workbook(events[-1].artifacts.output_workbook_path)
    assert workbook["Data Input"]["C6"].value == "Taylor"
    assert workbook["Expenses"]["D55"].value == 1250.0
    assert workbook["Expenses"]["C55"].value == '=IF(D55="","",D55*12)'


def test_job_runner_refills_finished_lanes_without_waiting_for_a_batch(tmp_path: Path) -> None:
    settings = _settings(tmp_path, ocr_parallel_workers=5)
    client = StaggeredConcurrencyExtractionClient(
        page_delays={
            1: 0.01,
            2: 0.2,
            3: 0.2,
            4: 0.2,
            5: 0.2,
            6: 0.01,
        }
    )
    runner = JobRunner(settings, extraction_client=client, entry_agent=TrackingTemplateEntryAgent())
    events = list(runner.run(_make_pdf_bytes(page_count=6), "test.pdf"))

    ocr_start_events = [event for event in events if event.stage == Stage.OCR and event.phase == "start"]

    assert len(ocr_start_events) == 6
    assert not any(event.phase in {"batch_start", "batch_complete"} for event in events)
    assert {event.pipe_number for event in ocr_start_events[:5]} == {1, 2, 3, 4, 5}
    page_six_start_index = next(
        index
        for index, event in enumerate(events)
        if event.stage == Stage.OCR and event.phase == "start" and event.page_number == 6
    )
    late_complete_index = min(
        index
        for index, event in enumerate(events)
        if event.stage == Stage.OCR and event.phase == "complete" and event.page_number in {2, 3, 4, 5}
    )
    assert next(event.pipe_number for event in ocr_start_events if event.page_number == 6) == 1
    assert page_six_start_index < late_complete_index
    assert client.peak_calls >= 2
    assert client.peak_calls <= settings.ocr_parallel_workers


def test_job_runner_emits_heartbeat_events_while_ocr_is_waiting(tmp_path: Path) -> None:
    settings = _settings(tmp_path, ocr_parallel_workers=1)
    runner = JobRunner(settings, extraction_client=SlowExtractionClient(), entry_agent=TrackingTemplateEntryAgent())

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))

    assert any(event.phase == "heartbeat" for event in events)


def test_job_runner_streams_agent_token_totals_on_workbook_heartbeats(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=TokenReportingTemplateEntryAgent(),
    )

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    heartbeat_events = [
        event
        for event in events
        if event.stage == Stage.DATA_ENTRY and event.phase == "heartbeat" and event.agent_total_tokens is not None
    ]

    assert heartbeat_events
    assert heartbeat_events[0].agent_total_tokens == 15


def test_job_runner_streams_reasoning_summaries_on_workbook_heartbeats(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=ProgressReportingTemplateEntryAgent(),
    )

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    heartbeat_events = [
        event
        for event in events
        if event.stage == Stage.DATA_ENTRY and event.phase == "heartbeat" and event.progress_message is not None
    ]

    assert heartbeat_events
    assert heartbeat_events[0].progress_message == "Matching workbook fields to the strongest OCR evidence."


def test_job_runner_persists_ocr_results_and_resumes_after_question(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=QuestioningTemplateEntryAgent(),
    )

    start_events = list(runner.start_job(_make_pdf_bytes(), "test.pdf"))
    paused_event = start_events[-1]

    assert paused_event.phase == "paused"
    assert paused_event.artifacts is not None
    assert paused_event.artifacts.pending_question is not None
    assert paused_event.artifacts.entry_state_path is not None
    assert paused_event.artifacts.entry_state_path.exists()
    assert paused_event.artifacts.ocr_results_path is not None
    assert paused_event.artifacts.ocr_results_path.exists()

    resume_events = list(runner.resume_job(paused_event.artifacts.job_id, "Taylor"))
    assert resume_events[-1].artifacts is not None
    assert resume_events[-1].artifacts.review_report is not None
    assert resume_events[-1].artifacts.review_report.user_answers[0].answer == "Taylor"
    assert resume_events[-1].artifacts.review_report.success


def test_job_runner_can_delegate_question_back_to_agent(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=QuestioningTemplateEntryAgent(),
    )

    start_events = list(runner.start_job(_make_pdf_bytes(), "test.pdf"))
    paused_event = start_events[-1]

    resume_events = list(runner.resume_job(paused_event.artifacts.job_id, "", source="agent"))

    assert resume_events[0].message == "Letting the agent resolve Data Input. Resuming workbook entry."
    assert resume_events[-1].artifacts is not None
    assert resume_events[-1].artifacts.review_report is not None
    assert resume_events[-1].artifacts.review_report.user_answers[0].source == "agent"
    assert resume_events[-1].artifacts.review_report.success
def test_job_runner_traverses_template_in_sheet_order(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    agent = TrackingTemplateEntryAgent()
    runner = JobRunner(settings, extraction_client=FakeExtractionClient(), entry_agent=agent)

    list(runner.run(_make_pdf_bytes(), "test.pdf"))

    assert agent.sheet_calls == [
        "Data Input",
        "Net Worth",
        "Transactions Raw",
        "Expenses",
        "Retirement Accounts",
        "Taxable Accounts",
        "Education Accounts",
    ]


def test_job_runner_emits_sheet_start_events_for_workbook_entry(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(settings, extraction_client=FakeExtractionClient(), entry_agent=TrackingTemplateEntryAgent())

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    start_events = [event for event in events if event.stage == Stage.DATA_ENTRY and event.phase == "start"]

    assert start_events
    assert start_events[0].sheet_name == "Data Input"
    assert start_events[0].message == "LangGraph is working on Data Input."
    assert "Reviewing OCR evidence" in str(start_events[0].detail_message)


def test_job_runner_marks_low_coverage_runs_for_review(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=LowCoverageTemplateEntryAgent(),
    )

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    report = events[-1].artifacts.review_report

    assert not report.success
    assert report.review_required_reasons
    assert report.coverage_summary is not None
    assert report.coverage_summary.unresolved_supported_target_count > 0


def test_job_runner_surfaces_workbook_rate_limit_retries(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=RetryingTemplateEntryAgent(),
    )

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    retry_events = [
        event
        for event in events
        if event.stage == Stage.DATA_ENTRY and event.phase == "retry"
    ]

    assert retry_events
    assert retry_events[0].retry_reason == "rate_limit"
    assert retry_events[0].retry_delay_seconds == 10.0
    assert retry_events[0].attempt_number == 2
    assert retry_events[0].max_attempts == 7
    assert "OpenAI rate limit hit while filling Data Input." in retry_events[0].message
    assert "rate_limit_error" in str(retry_events[0].detail_message)


def test_job_runner_classifies_status_code_429_as_rate_limit_without_text(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FakeExtractionClient(),
        entry_agent=SilentRetryingTemplateEntryAgent(),
    )

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))
    retry_events = [
        event
        for event in events
        if event.stage == Stage.DATA_ENTRY and event.phase == "retry"
    ]

    assert retry_events
    assert retry_events[0].retry_reason == "rate_limit"
    assert "OpenAI rate limit hit while filling Data Input." in retry_events[0].message
    assert retry_events[0].detail_message == "Workbook entry for Data Input: HTTP 429"


def test_retry_event_keeps_raw_error_in_detail_message() -> None:
    retry_queue: Queue[dict[str, object]] = Queue()
    retry_queue.put(
        {
            "page_number": 1,
            "pipe_number": 1,
            "attempt_number": 2,
            "max_attempts": 3,
            "retry_delay_seconds": 2.0,
            "detail_message": "OCR page 1: 1 validation error for PageOcrResult",
            "retry_reason": "transient",
        }
    )

    events = list(JobRunner._drain_retry_queue(retry_queue, completed_pages=0, total_pages=23, pipe_total=3))

    assert len(events) == 1
    assert events[0].message == "Restarting lane 1 for page 1/23 (pass 2/3) after 2.0s."
    assert events[0].detail_message == "OCR page 1: 1 validation error for PageOcrResult"


def test_timeout_retry_event_mentions_escalated_timeout() -> None:
    retry_queue: Queue[dict[str, object]] = Queue()
    retry_queue.put(
        {
            "page_number": 1,
            "pipe_number": 1,
            "attempt_number": 2,
            "max_attempts": 3,
            "retry_delay_seconds": 0.0,
            "detail_message": "OCR page 1: request timed out after 120.0s",
            "retry_reason": "timeout",
            "current_timeout_seconds": 120.0,
            "next_timeout_seconds": 180.0,
        }
    )

    events = list(JobRunner._drain_retry_queue(retry_queue, completed_pages=0, total_pages=23, pipe_total=3))

    assert len(events) == 1
    assert (
        events[0].message
        == "Page 1/23 in lane 1 exceeded the 120s processing limit. Retrying immediately with 180s timeout (pass 2/3)."
    )
    assert events[0].retry_reason == "timeout"


def test_job_runner_raises_sanitized_page_failure(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    runner = JobRunner(
        settings,
        extraction_client=FailingExtractionClient(),
        entry_agent=TrackingTemplateEntryAgent(),
    )
    events = runner.run(_make_pdf_bytes(), "test.pdf")
    emitted = []

    with pytest.raises(RuntimeError, match=r"Page review failed for page 1/1 in lane 1\.") as exc_info:
        for event in events:
            emitted.append(event)

    assert "validation error" not in str(exc_info.value)
    failed_event = emitted[-1]
    assert failed_event.phase == "failed"
    assert failed_event.message == "Page review failed for page 1/1 in lane 1."
    assert failed_event.detail_message == "OCR page 1: 1 validation error for PageOcrResult"
