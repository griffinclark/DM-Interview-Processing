from __future__ import annotations

import base64
import json
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


OCR_SYSTEM_PROMPT = """
You are an OCR and layout-aware extraction engine for planner-facing financial plan PDFs.
Read the supplied page image and native PDF text together.

Rules:
- Extract only what is directly visible on the page.
- Preserve qualifiers such as current, suggested, target, required, annual, monthly, or one-time.
- Capture recommendations and TODO items separately from numeric figures.
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
    for value in (
        getattr(exc, "status_code", None),
        getattr(exc, "status", None),
        getattr(getattr(exc, "response", None), "status_code", None),
        getattr(getattr(exc, "response", None), "status", None),
    ):
        if isinstance(value, int):
            return value
        if isinstance(value, str) and value.isdigit():
            return int(value)
    return None


def is_rate_limit_error(exc: Exception) -> bool:
    error_text = str(exc).lower()
    return status_code(exc) == 429 or "rate_limit_error" in error_text or "rate limit" in error_text


def retry_after_seconds(exc: Exception) -> float | None:
    header_sources = [
        getattr(exc, "headers", None),
        getattr(getattr(exc, "response", None), "headers", None),
    ]
    for headers in header_sources:
        if headers is None:
            continue
        for key, value in headers.items():
            if str(key).lower() != "retry-after":
                continue
            try:
                retry_after = float(value)
            except (TypeError, ValueError):
                continue
            if retry_after > 0:
                return retry_after
    return None


def backoff_seconds(settings: Settings, exc: Exception, attempt: int) -> float:
    exponential_backoff = min(
        settings.llm_retry_base_seconds * (2 ** (attempt - 1)),
        settings.llm_retry_max_seconds,
    )
    if not is_rate_limit_error(exc):
        return exponential_backoff

    retry_after = retry_after_seconds(exc)
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


def invoke_with_retries(
    settings: Settings,
    operation_name: str,
    invoke_fn: Callable[[], T],
    retry_notifier: RetryNotifier | None = None,
) -> T:
    base_max_attempts = max(1, settings.llm_max_retries + 1)
    rate_limit_attempts = rate_limit_max_attempts(settings, base_max_attempts)
    last_error: Exception | None = None
    attempt = 0

    while True:
        attempt += 1
        try:
            return invoke_fn()
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            max_attempts = rate_limit_attempts if is_rate_limit_error(exc) else base_max_attempts
            if attempt >= max_attempts:
                break
            delay_seconds = backoff_seconds(settings, exc, attempt)
            if retry_notifier is not None:
                retry_notifier(operation_name, attempt + 1, max_attempts, delay_seconds, exc)
            time.sleep(delay_seconds)
    raise RuntimeError(
        f"{operation_name} failed after {max_attempts} attempt(s): {last_error}"
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
            self._llm = ChatOpenAI(
                model=self._model,
                temperature=0,
                max_tokens=4096,
                timeout=settings.llm_timeout_seconds,
                max_retries=0,
                api_key=settings.openai_api_key,
            )
        else:
            if not settings.anthropic_api_key:
                raise ValueError("ANTHROPIC_API_KEY is required to run PlanLock with provider 'anthropic'.")
            self._llm = ChatAnthropic(
                model=self._model,
                temperature=0,
                max_tokens=4096,
                timeout=settings.llm_timeout_seconds,
                max_retries=0,
                anthropic_api_key=settings.anthropic_api_key,
            )

    def invoke(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str | None = None,
    ) -> BaseModel:
        runnable = self._llm.with_structured_output(schema).with_config(
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
        invoke_fn: Callable[[], T],
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
            "Use the native text as a hint, but trust the image for layout and tables.\n\n"
            f"Native PDF text:\n{page.native_text[:12000]}"
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
            invoke_fn=lambda: self._ocr_llm.invoke(
                schema=PageOcrResult,
                messages=messages,
                operation_name=f"OCR page {page.page_number}",
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
            invoke_fn=lambda: self._mapping_llm.invoke(
                schema=PageMappingResult,
                messages=messages,
                operation_name=f"Mapping page {page.page_number}",
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
