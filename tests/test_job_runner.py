from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import threading
import time

import fitz

from planlock.config import Settings
from planlock.job_runner import JobRunner
from planlock.models import (
    ExpenseCandidate,
    FieldCandidate,
    PageMappingResult,
    PageOcrResult,
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

    def map_page(self, page, ocr_result, retry_notifier=None):
        return PageMappingResult(
            page_number=page.page_number,
            mapped_fields=[
                FieldCandidate(
                    target_key="profile.client_1.first_name",
                    value="Taylor",
                    value_kind=ValueKind.STRING,
                    page_number=page.page_number,
                    source_excerpt="Taylor",
                    confidence=0.95,
                )
            ],
            expenses=[
                ExpenseCandidate(
                    category="Travel",
                    monthly_amount=1250.0,
                    label="Imported Target",
                    page_number=page.page_number,
                    source_excerpt="Travel target $1,250",
                    confidence=0.96,
                    comment="Suggested target chosen over current value.",
                )
            ],
            accounts=[],
            holdings=[],
            unmapped_items=[],
            warnings=[],
        )


class ConcurrencyTrackingExtractionClient(FakeExtractionClient):
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.active_calls = 0
        self.peak_calls = 0

    def ocr_page(self, page, retry_notifier=None):
        with self._lock:
            self.active_calls += 1
            self.peak_calls = max(self.peak_calls, self.active_calls)
        try:
            time.sleep(0.05)
            return super().ocr_page(page, retry_notifier=retry_notifier)
        finally:
            with self._lock:
                self.active_calls -= 1


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


def _make_pdf_bytes(page_count: int = 1) -> bytes:
    document = fitz.open()
    for page_number in range(1, page_count + 1):
        page = document.new_page()
        page.insert_text(
            (72, 72),
            f"Page {page_number}: Target personal spending of $1,250 per month for travel.",
        )
    return document.tobytes()


def test_job_runner_emits_all_three_stages(tmp_path: Path) -> None:
    settings = Settings.from_env()
    settings.ensure_runtime_dirs()
    runner = JobRunner(settings, extraction_client=FakeExtractionClient())
    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))

    stages = [event.stage for event in events]
    assert Stage.OCR in stages
    assert Stage.DATA_ENTRY in stages
    assert Stage.FINANCIAL_CALCULATIONS in stages
    assert events[-1].artifacts is not None
    assert events[-1].artifacts.review_report is not None
    assert events[-1].artifacts.review_report.success


def test_job_runner_refills_finished_lanes_without_waiting_for_a_batch(tmp_path: Path) -> None:
    settings = replace(Settings.from_env(), ocr_parallel_workers=5)
    settings.ensure_runtime_dirs()
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
    runner = JobRunner(settings, extraction_client=client)
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
    settings = replace(Settings.from_env(), ocr_parallel_workers=1)
    settings.ensure_runtime_dirs()
    runner = JobRunner(settings, extraction_client=SlowExtractionClient())

    events = list(runner.run(_make_pdf_bytes(), "test.pdf"))

    assert any(event.phase == "heartbeat" for event in events)
