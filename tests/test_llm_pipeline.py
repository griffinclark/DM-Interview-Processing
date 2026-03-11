from __future__ import annotations

from dataclasses import replace

from planlock.models import PageOcrResult
from planlock.pdf_renderer import RenderedPage
from planlock.config import LLM_PROVIDER_ANTHROPIC, Settings
from planlock.llm_pipeline import ClaudeExtractionClient, StructuredOutputClient


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


def _client(settings: Settings):
    client = object.__new__(ClaudeExtractionClient)
    client._settings = settings  # type: ignore[attr-defined]
    return client


def _disable_shared_throttle(monkeypatch, settings: Settings) -> None:
    monkeypatch.setattr(settings.request_throttle, "wait_for_availability", lambda: None)


def test_invoke_with_retries_succeeds_after_retry(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    attempts = {"count": 0}
    observed_timeouts: list[float] = []

    def flaky(timeout_seconds: float):
        observed_timeouts.append(timeout_seconds)
        attempts["count"] += 1
        if attempts["count"] < 2:
            raise TimeoutError("transient timeout")
        return "ok"

    result = ClaudeExtractionClient._invoke_with_retries(client, "test operation", flaky)

    assert result == "ok"
    assert attempts["count"] == 2
    assert observed_timeouts == [60.0, 90.0]


def test_invoke_with_retries_raises_after_exhaustion(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    attempts = {"count": 0}
    observed_timeouts: list[float] = []

    def always_fail(timeout_seconds: float):
        observed_timeouts.append(timeout_seconds)
        attempts["count"] += 1
        raise TimeoutError("still failing")

    try:
        ClaudeExtractionClient._invoke_with_retries(client, "test operation", always_fail)
    except RuntimeError as exc:
        assert "failed after 3 attempt(s)" in str(exc)
        assert "request timed out after 120.0s" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == 3
    assert observed_timeouts == [60.0, 90.0, 120.0]


def test_invoke_with_retries_reports_retry_metadata(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", lambda _: None)

    notifications: list[tuple[int, int, float]] = []

    def flaky(timeout_seconds: float):
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
    assert notifications[0][1] == 3
    assert notifications == [(2, 3, 0.0), (3, 3, 0.0)]


def test_invoke_with_retries_waits_longer_for_rate_limits(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)

    sleep_calls: list[float] = []
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", sleep_calls.append)

    attempts = {"count": 0}
    notifications: list[tuple[int, int, float]] = []
    observed_timeouts: list[float] = []

    def rate_limited(timeout_seconds: float):
        observed_timeouts.append(timeout_seconds)
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
    assert observed_timeouts == [60.0, 90.0, 120.0, 120.0, 120.0, 120.0, 120.0]
    assert sleep_calls == [10.0, 20.0, 40.0, 80.0, 160.0, 300.0]
    assert notifications[0] == (2, 7, 10.0)


def test_invoke_with_retries_caps_exponential_backoff_at_five_minutes() -> None:
    settings = Settings.from_env()
    client = _client(settings)

    delays = [
        ClaudeExtractionClient._backoff_seconds(client, TimeoutError("transient timeout"), attempt)
        for attempt in range(1, 8)
    ]

    assert delays == [10.0, 20.0, 40.0, 80.0, 160.0, 300.0, 300.0]


def test_invoke_with_retries_caps_retry_after_at_five_minutes() -> None:
    settings = Settings.from_env()
    client = _client(settings)

    class RetryAfterError(Exception):
        status_code = 429

        def __init__(self) -> None:
            super().__init__("rate limit")
            self.headers = {"Retry-After": "600"}

    delay = ClaudeExtractionClient._backoff_seconds(client, RetryAfterError(), 1)

    assert delay == 300.0


def test_invoke_with_retries_uses_openai_reset_headers(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)

    class ResetHeaderError(Exception):
        status_code = 429

        def __init__(self) -> None:
            super().__init__("HTTP 429")
            self.headers = {
                "x-ratelimit-reset-requests": "1m30s",
                "x-ratelimit-reset-tokens": "45s",
            }

    delay = ClaudeExtractionClient._backoff_seconds(client, ResetHeaderError(), 1)

    assert delay == 90.0


def test_invoke_with_retries_does_not_retry_quota_exhaustion(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    _disable_shared_throttle(monkeypatch, settings)

    sleep_calls: list[float] = []
    monkeypatch.setattr("planlock.llm_pipeline.time.sleep", sleep_calls.append)

    attempts = {"count": 0}

    class QuotaError(Exception):
        status_code = 429

        def __init__(self) -> None:
            super().__init__(
                "Error code: 429 - {'type': 'insufficient_quota', 'message': 'You exceeded your current quota, please check your plan and billing details.'}"
            )

    def out_of_quota(timeout_seconds: float):
        attempts["count"] += 1
        raise QuotaError()

    try:
        ClaudeExtractionClient._invoke_with_retries(client, "test operation", out_of_quota)
    except RuntimeError as exc:
        assert "failed after 1 attempt(s)" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == 1
    assert sleep_calls == []


def test_request_throttle_coordinator_waits_for_existing_cooldown(monkeypatch) -> None:
    settings = Settings.from_env()
    monotonic_values = iter([100.0, 100.0, 110.0])
    sleep_calls: list[float] = []

    monkeypatch.setattr("planlock.throttle.time.monotonic", lambda: next(monotonic_values))
    monkeypatch.setattr("planlock.throttle.time.sleep", sleep_calls.append)

    settings.request_throttle.impose_cooldown(10.0)
    settings.request_throttle.wait_for_availability()

    assert sleep_calls == [10.0]


def test_structured_output_client_does_not_set_openai_max_tokens(monkeypatch) -> None:
    settings = Settings.from_env()
    observed: dict[str, object] = {}

    class FakeChatOpenAI:
        def __init__(self, **kwargs) -> None:
            observed.update(kwargs)

    monkeypatch.setattr("planlock.llm_pipeline.ChatOpenAI", FakeChatOpenAI)

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    client._build_llm(timeout_seconds=45.0)

    assert observed["timeout"] == 45.0
    assert "max_tokens" not in observed


def test_structured_output_client_does_not_set_anthropic_max_tokens(monkeypatch) -> None:
    settings = replace(
        Settings.from_env(),
        llm_provider=LLM_PROVIDER_ANTHROPIC,
        anthropic_api_key="test-key",
    )
    observed: dict[str, object] = {}

    class FakeChatAnthropic:
        def __init__(self, **kwargs) -> None:
            observed.update(kwargs)

    monkeypatch.setattr("planlock.llm_pipeline.ChatAnthropic", FakeChatAnthropic)

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    client._build_llm(timeout_seconds=45.0)

    assert observed["timeout"] == 45.0
    assert "max_tokens" not in observed


def test_ocr_page_includes_full_native_pdf_text(monkeypatch) -> None:
    settings = Settings.from_env()
    client = _client(settings)
    observed: dict[str, object] = {}

    class FakeStructuredOutputClient:
        def invoke(self, *, schema, messages, operation_name=None, timeout_seconds=None):
            observed["schema"] = schema
            observed["messages"] = messages
            observed["operation_name"] = operation_name
            observed["timeout_seconds"] = timeout_seconds
            return _page_ocr_result(1, "Extracted text")

    client._ocr_llm = FakeStructuredOutputClient()  # type: ignore[attr-defined]
    client._invoke_with_retries = (  # type: ignore[attr-defined]
        lambda operation_name, invoke_fn, retry_notifier=None: invoke_fn(45.0)
    )

    native_text = ("native line\n" * 2000) + "tail-marker"
    page = RenderedPage(page_number=1, image_bytes=b"", native_text=native_text)

    result = ClaudeExtractionClient.ocr_page(client, page)

    assert result.page_number == 1
    assert observed["timeout_seconds"] == 45.0
    system_prompt = observed["messages"][0].content  # type: ignore[index]
    prompt = observed["messages"][1].content[0]["text"]  # type: ignore[index]
    assert "populate a locked Excel workbook" in system_prompt
    assert "Pull usable data out of charts and graphs" in system_prompt
    assert "populate a locked planner workbook" in prompt
    assert "chart or graph values" in prompt
    assert native_text in prompt
    assert prompt.endswith(native_text)


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
