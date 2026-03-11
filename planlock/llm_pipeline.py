from __future__ import annotations

import base64
import json
from email.utils import parsedate_to_datetime
import re
import time
from collections.abc import Callable
from typing import Any, Protocol, TypeVar

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_openai import ChatOpenAI
from pydantic import BaseModel

from planlock.config import LLM_PROVIDER_OPENAI, Settings
from planlock.models import PageMappingResult, PageOcrResult
from planlock.pdf_renderer import RenderedPage
from planlock.template_schema import schema_reference_for_prompt


T = TypeVar("T")
RetryNotifier = Callable[[str, int, int, float, Exception], None]
THROTTLE_RESET_HEADER_NAMES = (
    "x-ratelimit-reset-requests",
    "x-ratelimit-reset-tokens",
)
PAGE_PROCESS_TIMEOUT_SCHEDULE_SECONDS = (60.0, 90.0, 120.0)
NON_RETRYABLE_STATUS_CODES = {400, 401, 403, 404, 422}
COMPOSITE_DURATION_PATTERN = re.compile(r"(\d+(?:\.\d+)?)(ms|s|m|h|d)")
SECONDS_PER_DURATION_UNIT = {
    "ms": 0.001,
    "s": 1.0,
    "m": 60.0,
    "h": 3600.0,
    "d": 86400.0,
}


OCR_SYSTEM_PROMPT = """
You are an OCR and layout-aware extraction engine for planner-facing financial plan PDFs.
Your output is used downstream to populate a locked Excel workbook for financial planners, so extract only page-supported facts that can be mapped or reviewed later.
Read the supplied page image and native PDF text together.

Rules:
- Extract only what is directly visible on the current page.
- Focus on planner-relevant content such as household names, dates, balances, income, expenses, savings targets, debts, account details, holdings, recommendations, assumptions, action items, and any current vs. proposed or target values.
- Preserve qualifiers such as current, suggested, target, required, annual, monthly, or one-time.
- Extract tables as tables when row and column structure is visible.
- Pull usable data out of charts and graphs, including titles, legends, labels, axes, time periods, categories, and plotted values or comparisons. Represent that information in figures and tables whenever possible.
- Capture recommendations and TODO items separately from numeric figures.
- If native text and the image disagree, trust the image.
- Do not infer hidden rows, spreadsheet values, or unstated household details.
- Return structured output only.
""".strip()


MAPPING_SYSTEM_PROMPT = """
You convert OCR output into a locked Excel workbook schema.

You will receive the current target page's OCR JSON first, followed by OCR JSON for the full document.

Rules:
- Map only values supported by the provided schema reference.
- Use the full-document OCR only as reference context for disambiguation, continued tables, or headings.
- Emit only records that are directly supported by the current target page.
- Emit literal values only for directly observed constants. If a target depends on extracted constants, populate the source constants and let the workbook compute the derived value with formulas.
- Never guess missing values.
- If a page has both current and suggested values and the workbook has one planning input, prefer the suggested or target value and mention that assumption in the comment.
- Leave unsupported items in unmapped_items.
- Use page_number on every mapped record, and keep it set to the current target page.
- Return structured output only.
""".strip()


def status_code(exc: Exception) -> int | None:
    for error in iter_exception_chain(exc):
        for value in (
            getattr(error, "status_code", None),
            getattr(error, "status", None),
            getattr(getattr(error, "response", None), "status_code", None),
            getattr(getattr(error, "response", None), "status", None),
        ):
            if isinstance(value, int):
                return value
            if isinstance(value, str) and value.isdigit():
                return int(value)
    return None


def iter_exception_chain(exc: Exception):
    seen: set[int] = set()
    current: Exception | None = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        yield current
        current = current.__cause__ or current.__context__


def error_text(exc: Exception) -> str:
    return " ".join(str(error).lower() for error in iter_exception_chain(exc))


def response_headers(exc: Exception) -> dict[str, str]:
    for error in iter_exception_chain(exc):
        for headers in (
            getattr(error, "headers", None),
            getattr(getattr(error, "response", None), "headers", None),
        ):
            if not hasattr(headers, "items"):
                continue
            return {
                str(key).strip().lower(): str(value).strip()
                for key, value in headers.items()
            }
    return {}


def is_quota_exhaustion_error(exc: Exception) -> bool:
    message = error_text(exc)
    return status_code(exc) == 429 and (
        "insufficient_quota" in message
        or "exceeded your current quota" in message
        or "billing details" in message
    )


def is_rate_limit_error(exc: Exception) -> bool:
    if is_quota_exhaustion_error(exc):
        return False
    message = error_text(exc)
    return status_code(exc) == 429 or "rate_limit_error" in message or "rate limit" in message


def is_throttle_error(exc: Exception) -> bool:
    if is_rate_limit_error(exc):
        return True
    message = error_text(exc)
    return status_code(exc) == 503 and (
        "slow down" in message
        or "reduce your request rate" in message
        or "temporarily overloaded" in message
    )


def is_timeout_error(exc: Exception) -> bool:
    message = error_text(exc)
    exception_name = type(exc).__name__.lower()
    timeout_markers = (
        "timed out",
        "timeout",
        "readtimeout",
        "writetimeout",
        "connecttimeout",
        "deadline exceeded",
        "apitimeouterror",
    )
    return isinstance(exc, TimeoutError) or any(
        marker in message or marker in exception_name for marker in timeout_markers
    )


def is_non_retryable_error(exc: Exception) -> bool:
    if is_quota_exhaustion_error(exc):
        return True
    return status_code(exc) in NON_RETRYABLE_STATUS_CODES


def retry_reason_for_error(exc: Exception) -> str:
    if is_throttle_error(exc):
        return "rate_limit"
    if is_timeout_error(exc):
        return "timeout"
    return "transient"


def parse_retry_after_seconds(value: str) -> float | None:
    try:
        retry_after = float(value)
    except (TypeError, ValueError):
        try:
            parsed = parsedate_to_datetime(value)
        except (TypeError, ValueError):
            return None
        retry_after = parsed.timestamp() - time.time()
    if retry_after > 0:
        return retry_after
    return None


def parse_duration_seconds(value: str) -> float | None:
    text = value.strip().lower().replace(" ", "")
    if not text:
        return None
    try:
        numeric = float(text)
    except ValueError:
        matches = list(COMPOSITE_DURATION_PATTERN.finditer(text))
        if not matches or "".join(match.group(0) for match in matches) != text:
            return None
        duration_seconds = 0.0
        for match in matches:
            duration_seconds += float(match.group(1)) * SECONDS_PER_DURATION_UNIT[match.group(2)]
        return duration_seconds if duration_seconds > 0 else None
    return numeric if numeric > 0 else None


def throttle_reset_seconds(exc: Exception) -> float | None:
    headers = response_headers(exc)
    candidates: list[float] = []
    retry_after = parse_retry_after_seconds(headers.get("retry-after", ""))
    if retry_after is not None:
        candidates.append(retry_after)
    for header_name in THROTTLE_RESET_HEADER_NAMES:
        reset_value = parse_duration_seconds(headers.get(header_name, ""))
        if reset_value is not None:
            candidates.append(reset_value)
    return max(candidates, default=None)


def backoff_seconds(settings: Settings, exc: Exception, attempt: int) -> float:
    exponential_backoff = min(
        settings.llm_retry_base_seconds * (2 ** (attempt - 1)),
        settings.llm_retry_max_seconds,
    )
    if not is_throttle_error(exc):
        return exponential_backoff

    retry_after = throttle_reset_seconds(exc)
    if retry_after is not None:
        return min(
            max(exponential_backoff, retry_after),
            settings.llm_retry_max_seconds,
        )
    return exponential_backoff


def rate_limit_max_attempts(settings: Settings, base_max_attempts: int) -> int:
    retry_attempts = 1
    retry_delay_seconds = min(
        settings.llm_retry_base_seconds,
        settings.llm_retry_max_seconds,
    )
    while retry_delay_seconds < settings.llm_retry_max_seconds:
        retry_attempts += 1
        retry_delay_seconds *= 2
    return max(base_max_attempts, retry_attempts + 1)


def page_process_timeout_seconds(attempt: int) -> float:
    safe_attempt = max(1, attempt)
    return PAGE_PROCESS_TIMEOUT_SCHEDULE_SECONDS[
        min(safe_attempt - 1, len(PAGE_PROCESS_TIMEOUT_SCHEDULE_SECONDS) - 1)
    ]


def attach_timeout_metadata(
    exc: Exception,
    *,
    current_timeout_seconds: float,
    next_timeout_seconds: float | None,
) -> None:
    try:
        setattr(exc, "planlock_timeout_seconds", current_timeout_seconds)
        setattr(exc, "planlock_next_timeout_seconds", next_timeout_seconds)
    except Exception:  # noqa: BLE001
        return


def timeout_seconds_from_error(exc: Exception, *, attr_name: str) -> float | None:
    value = getattr(exc, attr_name, None)
    if isinstance(value, (int, float)):
        return float(value)
    return None


def describe_retry_error(exc: Exception | None) -> str:
    if exc is None:
        return "unknown error"
    if not is_timeout_error(exc):
        return str(exc)

    timeout_seconds = timeout_seconds_from_error(exc, attr_name="planlock_timeout_seconds")
    timeout_copy = (
        f"request timed out after {timeout_seconds:.1f}s"
        if timeout_seconds is not None
        else "request timed out"
    )
    original = str(exc).strip()
    if not original:
        return timeout_copy
    if timeout_copy.lower() in original.lower():
        return original
    return f"{timeout_copy}: {original}"


def invoke_with_retries(
    settings: Settings,
    operation_name: str,
    invoke_fn: Callable[[float], T],
    retry_notifier: RetryNotifier | None = None,
) -> T:
    base_max_attempts = len(PAGE_PROCESS_TIMEOUT_SCHEDULE_SECONDS)
    rate_limit_attempts = rate_limit_max_attempts(settings, base_max_attempts)
    last_error: Exception | None = None
    max_attempts = base_max_attempts
    attempt = 0

    while True:
        attempt += 1
        request_timeout_seconds = page_process_timeout_seconds(attempt)
        settings.request_throttle.wait_for_availability()
        try:
            return invoke_fn(request_timeout_seconds)
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            if is_non_retryable_error(exc):
                max_attempts = attempt
                break
            if is_throttle_error(exc):
                max_attempts = rate_limit_attempts
            else:
                max_attempts = base_max_attempts
            if is_timeout_error(exc):
                attach_timeout_metadata(
                    exc,
                    current_timeout_seconds=request_timeout_seconds,
                    next_timeout_seconds=(
                        page_process_timeout_seconds(attempt + 1)
                        if attempt < max_attempts
                        else None
                    ),
                )
            if attempt >= max_attempts:
                break
            delay_seconds = 0.0 if is_timeout_error(exc) else backoff_seconds(settings, exc, attempt)
            if is_throttle_error(exc):
                settings.request_throttle.impose_cooldown(delay_seconds)
            if retry_notifier is not None:
                retry_notifier(operation_name, attempt + 1, max_attempts, delay_seconds, exc)
            if delay_seconds > 0:
                time.sleep(delay_seconds)
    raise RuntimeError(
        f"{operation_name} failed after {max_attempts} attempt(s): {describe_retry_error(last_error)}"
    ) from last_error


class StructuredOutputClient:
    def __init__(self, settings: Settings, *, model: str) -> None:
        self._settings = settings
        self._provider = settings.normalized_llm_provider()
        self._model = model
        self._tags = [
            "planlock",
            f"provider:{self._provider}",
            f"model:{self._model}",
        ]
        self._metadata = {
            "planlock_provider": self._provider,
            "planlock_model": self._model,
        }

        if self._provider == LLM_PROVIDER_OPENAI:
            if not settings.openai_api_key:
                raise ValueError("OPENAI_API_KEY is required to run PlanLock with provider 'openai'.")
        else:
            if not settings.anthropic_api_key:
                raise ValueError("ANTHROPIC_API_KEY is required to run PlanLock with provider 'anthropic'.")

    def _build_llm(self, *, timeout_seconds: float):
        if self._provider == LLM_PROVIDER_OPENAI:
            return ChatOpenAI(
                model=self._model,
                temperature=0,
                timeout=timeout_seconds,
                max_retries=0,
                api_key=self._settings.openai_api_key,
            )
        return ChatAnthropic(
            model=self._model,
            temperature=0,
            timeout=timeout_seconds,
            max_retries=0,
            anthropic_api_key=self._settings.anthropic_api_key,
        )

    def invoke(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str | None = None,
        timeout_seconds: float | None = None,
    ) -> BaseModel:
        llm = self._build_llm(timeout_seconds=timeout_seconds or self._settings.llm_timeout_seconds)
        runnable = llm.with_structured_output(schema).with_config(
            {
                "run_name": operation_name or f"PlanLock {schema.__name__}",
                "tags": self._tags,
                "metadata": self._metadata,
            }
        )
        return runnable.invoke(messages)


class ExtractionClient(Protocol):
    def ocr_page(
        self,
        page: RenderedPage,
        retry_notifier: RetryNotifier | None = None,
    ) -> PageOcrResult:
        ...

    def map_page(
        self,
        page: RenderedPage,
        ocr_result: PageOcrResult,
        document_ocr_results: list[PageOcrResult],
        retry_notifier: RetryNotifier | None = None,
    ) -> PageMappingResult:
        ...


class ProviderExtractionClient:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._ocr_llm = StructuredOutputClient(settings, model=settings.model_ocr)
        self._mapping_llm = StructuredOutputClient(settings, model=settings.model_mapping)

    def _image_payload(self, page: RenderedPage) -> dict[str, Any]:
        encoded = base64.b64encode(page.image_bytes).decode("utf-8")
        if self._settings.normalized_llm_provider() == LLM_PROVIDER_OPENAI:
            return {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{encoded}",
                },
            }
        return {
            "type": "image",
            "source_type": "base64",
            "mime_type": "image/png",
            "data": encoded,
        }

    def _invoke_with_retries(
        self,
        operation_name: str,
        invoke_fn: Callable[[float], T],
        retry_notifier: RetryNotifier | None = None,
    ) -> T:
        return invoke_with_retries(
            self._settings,
            operation_name,
            invoke_fn,
            retry_notifier=retry_notifier,
        )

    def _backoff_seconds(self, exc: Exception, attempt: int) -> float:
        return backoff_seconds(self._settings, exc, attempt)

    def ocr_page(
        self,
        page: RenderedPage,
        retry_notifier: RetryNotifier | None = None,
    ) -> PageOcrResult:
        prompt = (
            f"Page number: {page.page_number}\n"
            "This OCR output will be used to populate a locked planner workbook, so prioritize structured financial facts, recommendations, tables, and chart or graph values.\n"
            "Use the native text as a hint, but trust the image for layout, tables, charts, and graphs.\n\n"
            f"Native PDF text:\n{page.native_text}"
        )
        messages = [
            SystemMessage(content=OCR_SYSTEM_PROMPT),
            HumanMessage(
                content=[
                    {"type": "text", "text": prompt},
                    self._image_payload(page),
                ]
            ),
        ]
        return self._invoke_with_retries(
            operation_name=f"OCR page {page.page_number}",
            invoke_fn=lambda timeout_seconds: self._ocr_llm.invoke(
                schema=PageOcrResult,
                messages=messages,
                operation_name=f"OCR page {page.page_number}",
                timeout_seconds=timeout_seconds,
            ),
            retry_notifier=retry_notifier,
        )

    def map_page(
        self,
        page: RenderedPage,
        ocr_result: PageOcrResult,
        document_ocr_results: list[PageOcrResult],
        retry_notifier: RetryNotifier | None = None,
    ) -> PageMappingResult:
        prompt = self._build_mapping_prompt(page, ocr_result, document_ocr_results)
        messages = [
            SystemMessage(content=MAPPING_SYSTEM_PROMPT),
            HumanMessage(content=prompt),
        ]
        return self._invoke_with_retries(
            operation_name=f"Mapping page {page.page_number}",
            invoke_fn=lambda timeout_seconds: self._mapping_llm.invoke(
                schema=PageMappingResult,
                messages=messages,
                operation_name=f"Mapping page {page.page_number}",
                timeout_seconds=timeout_seconds,
            ),
            retry_notifier=retry_notifier,
        )

    @staticmethod
    def _build_mapping_prompt(
        page: RenderedPage,
        ocr_result: PageOcrResult,
        document_ocr_results: list[PageOcrResult],
    ) -> str:
        schema_reference = schema_reference_for_prompt()
        return (
            f"Target page for this mapping pass: {page.page_number}\n"
            "Use the current page OCR JSON as the primary source for emitted records. "
            "Use the full document OCR JSON only as reference context.\n\n"
            f"Workbook schema reference:\n{schema_reference}\n\n"
            "Current page OCR output JSON:\n"
            f"{json.dumps(ocr_result.model_dump(mode='json'), indent=2)}\n\n"
            "Full document OCR output JSON (reference context across all pages):\n"
            f"{json.dumps([result.model_dump(mode='json') for result in document_ocr_results], indent=2)}"
        )


OpenAICompatibleExtractionClient = ProviderExtractionClient
ClaudeExtractionClient = ProviderExtractionClient
