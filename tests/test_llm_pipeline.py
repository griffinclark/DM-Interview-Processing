from __future__ import annotations

from planlock.models import PageOcrResult
from planlock.pdf_renderer import RenderedPage
from planlock.config import Settings
from planlock.llm_pipeline import ClaudeExtractionClient


class FakeRateLimitError(Exception):
    status_code = 429


def _page_ocr_result(page_number: int, raw_text: str) -> PageOcrResult:
    return PageOcrResult(
        page_number=page_number,
        summary=f"Summary for page {page_number}",
        raw_text=raw_text,
        source_snippets=[raw_text],
        figures=[],
        tables=[],
        recommendations=[],
        confidence=0.99,
    )


def test_invoke_with_retries_succeeds_after_retry(monkeypatch) -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    attempts = {"count": 0}

    def flaky():
        attempts["count"] += 1
        if attempts["count"] < 2:
            raise TimeoutError("transient timeout")
        return "ok"

    result = ClaudeExtractionClient._invoke_with_retries(client, "test operation", flaky)

    assert result == "ok"
    assert attempts["count"] == 2


def test_invoke_with_retries_raises_after_exhaustion(monkeypatch) -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    attempts = {"count": 0}

    def always_fail():
        attempts["count"] += 1
        raise TimeoutError("still failing")

    try:
        ClaudeExtractionClient._invoke_with_retries(client, "test operation", always_fail)
    except RuntimeError as exc:
        assert "failed after" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == settings.llm_max_retries + 1


def test_invoke_with_retries_reports_retry_metadata(monkeypatch) -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    notifications: list[tuple[int, int, float]] = []

    def flaky():
        raise TimeoutError("transient timeout")

    def notifier(operation_name, attempt_number, max_attempts, retry_delay_seconds, exc) -> None:
        notifications.append((attempt_number, max_attempts, retry_delay_seconds))

    try:
        ClaudeExtractionClient._invoke_with_retries(
            client,
            "test operation",
            flaky,
            retry_notifier=notifier,
        )
    except RuntimeError:
        pass

    assert notifications
    assert notifications[0][0] == 2
    assert notifications[0][1] == settings.llm_max_retries + 1


def test_invoke_with_retries_waits_longer_for_rate_limits(monkeypatch) -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]

    sleep_calls: list[float] = []
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", sleep_calls.append)

    attempts = {"count": 0}
    notifications: list[tuple[int, int, float]] = []

    def rate_limited():
        attempts["count"] += 1
        raise FakeRateLimitError(
            "Error code: 429 - {'type': 'error', 'error': {'type': 'rate_limit_error'}}"
        )

    def notifier(operation_name, attempt_number, max_attempts, retry_delay_seconds, exc) -> None:
        notifications.append((attempt_number, max_attempts, retry_delay_seconds))

    try:
        ClaudeExtractionClient._invoke_with_retries(
            client,
            "test operation",
            rate_limited,
            retry_notifier=notifier,
        )
    except RuntimeError as exc:
        assert "failed after 7 attempt(s)" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == 7
    assert sleep_calls == [10.0, 20.0, 40.0, 80.0, 160.0, 300.0]
    assert notifications[0] == (2, 7, 10.0)


def test_invoke_with_retries_caps_exponential_backoff_at_five_minutes() -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]

    delays = [
        ClaudeExtractionClient._backoff_seconds(client, TimeoutError("transient timeout"), attempt)
        for attempt in range(1, 8)
    ]

    assert delays == [10.0, 20.0, 40.0, 80.0, 160.0, 300.0, 300.0]


def test_invoke_with_retries_caps_retry_after_at_five_minutes() -> None:
    settings = Settings.from_env()
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]

    class RetryAfterError(Exception):
        status_code = 429

        def __init__(self) -> None:
            super().__init__("rate limit")
            self.headers = {"Retry-After": "600"}

    delay = ClaudeExtractionClient._backoff_seconds(client, RetryAfterError(), 1)

    assert delay == 300.0


def test_build_mapping_prompt_includes_current_page_and_full_document_context() -> None:
    page = RenderedPage(page_number=2, image_bytes=b"", native_text="Native text for page 2")
    page_one_result = _page_ocr_result(1, "Extracted text for page 1")
    page_two_result = _page_ocr_result(2, "Extracted text for page 2")

    prompt = ClaudeExtractionClient._build_mapping_prompt(
        page,
        page_two_result,
        [page_one_result, page_two_result],
    )

    assert "Target page for this mapping pass: 2" in prompt
    assert "Current page OCR output JSON:" in prompt
    assert "Full document OCR output JSON (reference context across all pages):" in prompt
    assert '"page_number": 1' in prompt
    assert '"raw_text": "Extracted text for page 1"' in prompt
    assert '"page_number": 2' in prompt
    assert prompt.index("Current page OCR output JSON:") < prompt.index(
        "Full document OCR output JSON (reference context across all pages):"
    )
