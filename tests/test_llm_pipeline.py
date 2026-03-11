from __future__ import annotations

from dataclasses import replace
from types import SimpleNamespace

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_core.tools import StructuredTool

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
    assert observed_timeouts == [120.0, 180.0]


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
        assert "request timed out after 240.0s" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == 3
    assert observed_timeouts == [120.0, 180.0, 240.0]


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
    assert observed_timeouts == [120.0, 180.0, 240.0, 240.0, 240.0, 240.0, 240.0]
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


def test_structured_output_client_streams_openai_reasoning_summaries() -> None:
    settings = Settings.from_env()
    observed: dict[str, object] = {}
    call_order: list[str] = []

    class FakeStream:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb) -> None:
            return None

        def __iter__(self):
            call_order.append("stream")
            yield SimpleNamespace(
                type="response.reasoning_summary_text.delta",
                summary_index=0,
                delta="Reviewing the current workbook context.",
            )

        def get_final_response(self):
            call_order.append("final_response")
            return SimpleNamespace(
                usage=SimpleNamespace(
                    input_tokens=11,
                    output_tokens=7,
                    total_tokens=18,
                ),
                output_parsed=_page_ocr_result(1, "Extracted text"),
            )

    class FakeResponses:
        def stream(self, **kwargs):
            observed.update(kwargs)
            return FakeStream()

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    client._openai_client = SimpleNamespace(responses=FakeResponses())  # type: ignore[attr-defined]

    usage_calls: list[tuple[str, dict[str, int]]] = []
    progress_calls: list[tuple[str, str]] = []
    result = client.invoke(
        schema=PageOcrResult,
        messages=[
            SystemMessage(content="System guidance"),
            HumanMessage(content="User prompt"),
        ],
        operation_name="Workbook entry for Data Input",
        timeout_seconds=45.0,
        usage_notifier=lambda operation_name, usage: usage_calls.append((operation_name, usage)),
        progress_notifier=lambda operation_name, message: (
            call_order.append("progress"),
            progress_calls.append((operation_name, message)),
        ),
    )

    assert isinstance(result, PageOcrResult)
    assert observed["model"] == settings.model_ocr
    assert observed["input"] == [
        {"role": "system", "content": "System guidance"},
        {"role": "user", "content": "User prompt"},
    ]
    assert observed["text_format"] is PageOcrResult
    assert observed["reasoning"] == {"summary": "concise"}
    assert observed["temperature"] == 0
    assert observed["timeout"] == 45.0
    assert progress_calls == [
        (
            "Workbook entry for Data Input",
            "Reviewing the current workbook context.",
        )
    ]
    assert call_order.index("progress") < call_order.index("final_response")
    assert usage_calls == [
        (
            "Workbook entry for Data Input",
            {
                "input_tokens": 11,
                "output_tokens": 7,
                "total_tokens": 18,
            },
        )
    ]


def test_structured_output_client_handles_openai_streams_without_reasoning_summaries() -> None:
    settings = Settings.from_env()

    class FakeStream:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb) -> None:
            return None

        def __iter__(self):
            yield SimpleNamespace(type="response.output_text.delta", delta="{")

        def get_final_response(self):
            return SimpleNamespace(
                usage=None,
                output_parsed=_page_ocr_result(1, "Extracted text"),
            )

    class FakeResponses:
        def stream(self, **kwargs):
            return FakeStream()

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    client._openai_client = SimpleNamespace(responses=FakeResponses())  # type: ignore[attr-defined]

    progress_calls: list[tuple[str, str]] = []
    result = client.invoke(
        schema=PageOcrResult,
        messages=[HumanMessage(content="User prompt")],
        operation_name="Workbook entry for Data Input",
        timeout_seconds=45.0,
        progress_notifier=lambda operation_name, message: progress_calls.append((operation_name, message)),
    )

    assert isinstance(result, PageOcrResult)
    assert progress_calls == []


def test_structured_output_client_reports_token_usage_from_raw_message_for_anthropic(monkeypatch) -> None:
    settings = replace(
        Settings.from_env(),
        llm_provider=LLM_PROVIDER_ANTHROPIC,
        anthropic_api_key="test-key",
    )
    observed: dict[str, object] = {}

    class FakeRunnable:
        def with_config(self, config):
            observed["config"] = config
            return self

        def invoke(self, messages):
            observed["messages"] = messages
            return {
                "raw": SimpleNamespace(
                    usage_metadata={
                        "input_tokens": 11,
                        "output_tokens": 7,
                        "total_tokens": 18,
                    }
                ),
                "parsed": _page_ocr_result(1, "Extracted text"),
                "parsing_error": None,
            }

    class FakeLLM:
        def with_structured_output(self, schema, include_raw=False):
            observed["schema"] = schema
            observed["include_raw"] = include_raw
            return FakeRunnable()

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    monkeypatch.setattr(client, "_build_llm", lambda timeout_seconds: FakeLLM())

    usage_calls: list[tuple[str, dict[str, int]]] = []
    result = client.invoke(
        schema=PageOcrResult,
        messages=["message"],
        operation_name="OCR page 1",
        timeout_seconds=45.0,
        usage_notifier=lambda operation_name, usage: usage_calls.append((operation_name, usage)),
    )

    assert isinstance(result, PageOcrResult)
    assert observed["include_raw"] is True
    assert usage_calls == [
        (
            "OCR page 1",
            {
                "input_tokens": 11,
                "output_tokens": 7,
                "total_tokens": 18,
            },
        )
    ]


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
    assert "persisted as OCR JSON" in system_prompt
    assert "map direct evidence into a locked Excel workbook" in system_prompt
    assert "raw_text and source_snippets as compact evidence records" in system_prompt
    assert "Pull usable data out of charts and graphs" in system_prompt
    assert "saved to disk and later reused for workbook mapping and raw PDF ambiguity review" in prompt
    assert "do not save decorative or boilerplate text" in prompt
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


def test_structured_output_client_runs_tool_round_trip_before_final_parse(monkeypatch) -> None:
    settings = Settings.from_env()
    observed: dict[str, object] = {}

    query_tool = StructuredTool.from_function(
        name="query_transactions",
        description="Query transactions",
        func=lambda sql: '{"status":"ok","rows":[{"row_number":2,"amount":-95.7}]}',
    )

    class FakeToolRunnable:
        def __init__(self) -> None:
            self.calls = 0

        def with_config(self, config):
            observed["tool_config"] = config
            return self

        def invoke(self, messages):
            self.calls += 1
            observed.setdefault("tool_messages", []).append(messages)
            if self.calls == 1:
                return AIMessage(
                    content="Need the ledger rows.",
                    tool_calls=[
                        {
                            "name": "query_transactions",
                            "args": {
                                "sql": "SELECT row_number, amount FROM transactions_raw ORDER BY row_number LIMIT 1"
                            },
                            "id": "call-1",
                        }
                    ],
                    usage_metadata={"input_tokens": 10, "output_tokens": 5, "total_tokens": 15},
                )
            return AIMessage(
                content="Ready for final answer.",
                usage_metadata={"input_tokens": 7, "output_tokens": 3, "total_tokens": 10},
            )

    class FakeStructuredRunnable:
        def with_config(self, config):
            observed["structured_config"] = config
            return self

        def invoke(self, messages):
            observed["final_messages"] = messages
            return {
                "raw": SimpleNamespace(
                    usage_metadata={"input_tokens": 8, "output_tokens": 4, "total_tokens": 12}
                ),
                "parsed": _page_ocr_result(1, "final parsed result"),
                "parsing_error": None,
            }

    class FakeLLM:
        def bind_tools(self, tools):
            observed["tools"] = tools
            return FakeToolRunnable()

        def with_structured_output(self, schema, include_raw=False):
            observed["schema"] = schema
            observed["include_raw"] = include_raw
            return FakeStructuredRunnable()

    client = StructuredOutputClient(settings, model=settings.model_ocr)
    monkeypatch.setattr(client, "_build_llm", lambda timeout_seconds: FakeLLM())

    usage_calls: list[tuple[str, dict[str, int]]] = []
    result = client.invoke(
        schema=PageOcrResult,
        messages=[HumanMessage(content="Fill the sheet.")],
        operation_name="Workbook entry for Expenses",
        timeout_seconds=45.0,
        usage_notifier=lambda operation_name, usage: usage_calls.append((operation_name, usage)),
        tools=[query_tool],
    )

    assert isinstance(result, PageOcrResult)
    assert observed["tools"] == [query_tool]
    assert observed["include_raw"] is True
    final_messages = observed["final_messages"]
    assert any(isinstance(message, ToolMessage) for message in final_messages)
    assert any(
        isinstance(message, HumanMessage)
        and "return the final PageOcrResult structured output only" in str(message.content)
        for message in final_messages
    )
    assert usage_calls == [
        ("Workbook entry for Expenses", {"input_tokens": 10, "output_tokens": 5, "total_tokens": 15}),
        ("Workbook entry for Expenses", {"input_tokens": 7, "output_tokens": 3, "total_tokens": 10}),
        ("Workbook entry for Expenses", {"input_tokens": 8, "output_tokens": 4, "total_tokens": 12}),
    ]
