from __future__ import annotations

import base64
import json
import time
from collections.abc import Callable
from typing import Protocol, TypeVar

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import HumanMessage, SystemMessage

from planlock.config import Settings
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

Rules:
- Map only values supported by the provided schema reference.
- Never guess missing values.
- If a page has both current and suggested values and the workbook has one planning input, prefer the suggested or target value and mention that assumption in the comment.
- Leave unsupported items in unmapped_items.
- Use page_number on every mapped record.
- Return structured output only.
""".strip()


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
        retry_notifier: RetryNotifier | None = None,
    ) -> PageMappingResult:
        ...


class ClaudeExtractionClient:
    def __init__(self, settings: Settings) -> None:
        if not settings.anthropic_api_key:
            raise ValueError("ANTHROPIC_API_KEY is required to run PlanLock.")
        self._settings = settings

        self._ocr_llm = ChatAnthropic(
            model=settings.model_ocr,
            temperature=0,
            max_tokens=4096,
            timeout=settings.llm_timeout_seconds,
            max_retries=0,
            anthropic_api_key=settings.anthropic_api_key,
        )
        self._mapping_llm = ChatAnthropic(
            model=settings.model_mapping,
            temperature=0,
            max_tokens=4096,
            timeout=settings.llm_timeout_seconds,
            max_retries=0,
            anthropic_api_key=settings.anthropic_api_key,
        )

    @staticmethod
    def _image_payload(page: RenderedPage) -> dict[str, str]:
        return {
            "type": "image",
            "source_type": "base64",
            "mime_type": "image/png",
            "data": base64.b64encode(page.image_bytes).decode("utf-8"),
        }

    def ocr_page(
        self,
        page: RenderedPage,
        retry_notifier: RetryNotifier | None = None,
    ) -> PageOcrResult:
        llm = self._ocr_llm.with_structured_output(PageOcrResult)
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
            invoke_fn=lambda: llm.invoke(messages),
            retry_notifier=retry_notifier,
        )

    def map_page(
        self,
        page: RenderedPage,
        ocr_result: PageOcrResult,
        retry_notifier: RetryNotifier | None = None,
    ) -> PageMappingResult:
        llm = self._mapping_llm.with_structured_output(PageMappingResult)
        schema_reference = schema_reference_for_prompt()
        prompt = (
            f"Page number: {page.page_number}\n\n"
            f"Workbook schema reference:\n{schema_reference}\n\n"
            "OCR output JSON:\n"
            f"{json.dumps(ocr_result.model_dump(mode='json'), indent=2)}"
        )
        messages = [
            SystemMessage(content=MAPPING_SYSTEM_PROMPT),
            HumanMessage(content=prompt),
        ]
        return self._invoke_with_retries(
            operation_name=f"Mapping page {page.page_number}",
            invoke_fn=lambda: llm.invoke(messages),
            retry_notifier=retry_notifier,
        )

    @staticmethod
    def _status_code(exc: Exception) -> int | None:
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

    @classmethod
    def _is_rate_limit_error(cls, exc: Exception) -> bool:
        error_text = str(exc).lower()
        return cls._status_code(exc) == 429 or "rate_limit_error" in error_text or "rate limit" in error_text

    @staticmethod
    def _retry_after_seconds(exc: Exception) -> float | None:
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

    def _backoff_seconds(self, exc: Exception, attempt: int) -> float:
        default_backoff = min(
            self._settings.llm_retry_base_seconds * (2 ** (attempt - 1)),
            self._settings.llm_retry_max_seconds,
        )
        if not self._is_rate_limit_error(exc):
            return default_backoff

        retry_after = self._retry_after_seconds(exc)
        if retry_after is not None:
            return max(default_backoff, retry_after)

        rate_limit_backoff = min(
            15.0 * (2 ** (attempt - 1)),
            max(self._settings.llm_retry_max_seconds, 90.0),
        )
        return max(default_backoff, rate_limit_backoff)

    def _invoke_with_retries(
        self,
        operation_name: str,
        invoke_fn: Callable[[], T],
        retry_notifier: RetryNotifier | None = None,
    ) -> T:
        base_max_attempts = max(1, self._settings.llm_max_retries + 1)
        rate_limit_max_attempts = max(base_max_attempts, 4)
        last_error: Exception | None = None
        attempt = 0

        while True:
            attempt += 1
            try:
                return invoke_fn()
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                max_attempts = rate_limit_max_attempts if self._is_rate_limit_error(exc) else base_max_attempts
                if attempt >= max_attempts:
                    break
                backoff_seconds = self._backoff_seconds(exc, attempt)
                if retry_notifier is not None:
                    retry_notifier(operation_name, attempt + 1, max_attempts, backoff_seconds, exc)
                time.sleep(backoff_seconds)
        raise RuntimeError(
            f"{operation_name} failed after {max_attempts} attempt(s): {last_error}"
        ) from last_error
