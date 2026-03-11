from __future__ import annotations

import base64
import json
from email.utils import parsedate_to_datetime
import re
import time
from collections.abc import Callable
from typing import Any, Protocol, TypeVar

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_core.tools import BaseTool
from langchain_openai import ChatOpenAI
from openai import OpenAI
from pydantic import BaseModel

from planlock import APP_NAME, APP_TRACE_SLUG
from planlock.config import LLM_PROVIDER_OPENAI, Settings
from planlock.models import PageMappingResult, PageOcrResult
from planlock.pdf_renderer import RenderedPage
from planlock.template_schema import schema_reference_for_prompt


T = TypeVar("T")
RetryNotifier = Callable[[str, int, int, float, Exception], None]
UsageNotifier = Callable[[str, dict[str, int]], None]
ProgressNotifier = Callable[[str, str], None]
THROTTLE_RESET_HEADER_NAMES = (
    "x-ratelimit-reset-requests",
    "x-ratelimit-reset-tokens",
)
PAGE_PROCESS_TIMEOUT_SCHEDULE_SECONDS = (120.0, 180.0, 240.0)
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
Your output is persisted as OCR JSON and later used in exactly two ways:
1. map direct evidence into a locked Excel workbook for financial planners
2. re-review saved OCR evidence when a planner-facing ambiguity needs to be resolved
Extract every page-supported token needed for those downstream tasks, but do not save content that cannot support a workbook write, recommendation review, or ambiguity resolution later.
Read the supplied page image and native PDF text together.

Rules:
- Extract only what is directly visible on the current page.
- Keep planner-relevant content such as household names, relationships, dates, balances, income, expenses, savings targets, debts, account details, holdings, recommendations, assumptions, action items, and any current vs. proposed or target values.
- Omit decorative or non-operative content such as repeated headers or footers, page numbers, branding, generic contact blocks, navigation copy, marketing filler, and boilerplate legal or compliance disclosures unless that text changes a financial fact, recommendation, assumption, or account constraint.
- When in doubt, keep the token if it could help defend a workbook write or resolve a later planner question.
- Preserve qualifiers such as current, suggested, target, required, annual, monthly, or one-time.
- Extract tables as tables when row and column structure is visible.
- Pull usable data out of charts and graphs, including titles, legends, labels, axes, time periods, categories, and plotted values or comparisons. Represent that information in figures and tables whenever possible.
- Capture recommendations and TODO items separately from numeric figures.
- Use raw_text and source_snippets as compact evidence records, not archival page dumps. Include enough surrounding text to justify extracted facts later, but do not copy irrelevant prose.
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


def token_usage_from_message(message: Any) -> dict[str, int] | None:
    usage_metadata = getattr(message, "usage_metadata", None)
    if isinstance(usage_metadata, dict):
        input_tokens = int(usage_metadata.get("input_tokens") or 0)
        output_tokens = int(usage_metadata.get("output_tokens") or 0)
        total_tokens = int(usage_metadata.get("total_tokens") or (input_tokens + output_tokens))
        return {
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": total_tokens,
        }

    response_metadata = getattr(message, "response_metadata", None)
    if not isinstance(response_metadata, dict):
        return None

    raw_usage = response_metadata.get("token_usage") or response_metadata.get("usage")
    if not isinstance(raw_usage, dict):
        return None

    input_tokens = int(raw_usage.get("prompt_tokens") or raw_usage.get("input_tokens") or 0)
    output_tokens = int(raw_usage.get("completion_tokens") or raw_usage.get("output_tokens") or 0)
    total_tokens = int(raw_usage.get("total_tokens") or (input_tokens + output_tokens))
    return {
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "total_tokens": total_tokens,
    }


def token_usage_from_response(response: Any) -> dict[str, int] | None:
    usage = getattr(response, "usage", None)
    if usage is None:
        return None

    input_tokens = int(getattr(usage, "input_tokens", 0) or 0)
    output_tokens = int(getattr(usage, "output_tokens", 0) or 0)
    total_tokens = int(getattr(usage, "total_tokens", 0) or (input_tokens + output_tokens))
    return {
        "input_tokens": input_tokens,
        "output_tokens": output_tokens,
        "total_tokens": total_tokens,
    }


def _is_sentence_boundary(text: str) -> bool:
    stripped = text.rstrip()
    if not stripped:
        return False
    return text.endswith("\n") or stripped.endswith((".", "!", "?"))


def _response_input_role(message: Any) -> str:
    if isinstance(message, SystemMessage):
        return "system"
    if isinstance(message, HumanMessage):
        return "user"

    message_type = str(getattr(message, "type", "")).lower()
    if message_type in {"system", "developer"}:
        return message_type
    if message_type in {"ai", "assistant"}:
        return "assistant"
    return "user"


def _response_input_content(content: Any) -> Any:
    if isinstance(content, str):
        return content

    if not isinstance(content, list):
        raise TypeError("Unsupported message content for OpenAI Responses streaming.")

    response_content: list[dict[str, Any]] = []
    for item in content:
        if not isinstance(item, dict):
            raise TypeError("Unsupported message content item for OpenAI Responses streaming.")

        item_type = str(item.get("type") or "").strip()
        if item_type == "text":
            response_content.append(
                {
                    "type": "input_text",
                    "text": str(item.get("text") or ""),
                }
            )
            continue
        if item_type == "image_url":
            image_url = item.get("image_url") or {}
            if not isinstance(image_url, dict):
                raise TypeError("OpenAI Responses streaming expected image_url payload to be a dict.")
            response_content.append(
                {
                    "type": "input_image",
                    "image_url": str(image_url.get("url") or ""),
                    "detail": str(image_url.get("detail") or "auto"),
                }
            )
            continue
        raise TypeError(f"Unsupported content type for OpenAI Responses streaming: {item_type or 'unknown'}")

    return response_content


def response_input_from_messages(messages: list[Any]) -> list[dict[str, Any]]:
    response_messages: list[dict[str, Any]] = []
    for message in messages:
        response_messages.append(
            {
                "role": _response_input_role(message),
                "content": _response_input_content(getattr(message, "content", message)),
            }
        )
    return response_messages


class ReasoningSummaryCoalescer:
    def __init__(
        self,
        *,
        operation_name: str,
        notifier: ProgressNotifier,
        min_emit_interval_seconds: float = 0.25,
    ) -> None:
        self._operation_name = operation_name
        self._notifier = notifier
        self._min_emit_interval_seconds = min_emit_interval_seconds
        self._summary_parts: dict[int, str] = {}
        self._last_emitted_text = ""
        self._last_emitted_at = 0.0

    def consume(self, event: Any) -> None:
        event_type = str(getattr(event, "type", ""))
        if event_type == "response.reasoning_summary_part.added":
            summary_index = int(getattr(event, "summary_index", 0))
            part = getattr(event, "part", None)
            initial_text = str(getattr(part, "text", "") or "")
            if initial_text:
                self._summary_parts[summary_index] = initial_text
                self._emit_if_ready()
            return

        if event_type == "response.reasoning_summary_text.delta":
            summary_index = int(getattr(event, "summary_index", 0))
            delta = str(getattr(event, "delta", "") or "")
            if delta:
                self._summary_parts[summary_index] = self._summary_parts.get(summary_index, "") + delta
                self._emit_if_ready()
            return

        if event_type == "response.reasoning_summary_text.done":
            summary_index = int(getattr(event, "summary_index", 0))
            self._summary_parts[summary_index] = str(getattr(event, "text", "") or "")
            self._emit_if_ready(force=True)

    def flush(self) -> None:
        self._emit_if_ready(force=True)

    def _visible_text(self) -> str:
        visible_parts = [
            text.strip()
            for _, text in sorted(self._summary_parts.items())
            if text and text.strip()
        ]
        return "\n\n".join(visible_parts)

    def _emit_if_ready(self, *, force: bool = False) -> None:
        visible_text = self._visible_text()
        if not visible_text or visible_text == self._last_emitted_text:
            return

        now = time.monotonic()
        if not force and (now - self._last_emitted_at) < self._min_emit_interval_seconds and not _is_sentence_boundary(visible_text):
            return

        self._notifier(self._operation_name, visible_text)
        self._last_emitted_text = visible_text
        self._last_emitted_at = now


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
        self._openai_client: OpenAI | None = None
        self._tags = [
            APP_TRACE_SLUG,
            f"provider:{self._provider}",
            f"model:{self._model}",
        ]
        self._metadata = {
            f"{APP_TRACE_SLUG}_provider": self._provider,
            f"{APP_TRACE_SLUG}_model": self._model,
        }

        if self._provider == LLM_PROVIDER_OPENAI:
            if not settings.openai_api_key:
                raise ValueError(f"OPENAI_API_KEY is required to run {APP_NAME} with provider 'openai'.")
            self._openai_client = OpenAI(api_key=settings.openai_api_key)
        else:
            if not settings.anthropic_api_key:
                raise ValueError(
                    f"ANTHROPIC_API_KEY is required to run {APP_NAME} with provider 'anthropic'."
                )

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
        usage_notifier: UsageNotifier | None = None,
        progress_notifier: ProgressNotifier | None = None,
        tools: list[BaseTool] | None = None,
    ) -> BaseModel:
        resolved_operation_name = operation_name or f"{APP_NAME} {schema.__name__}"
        resolved_timeout_seconds = timeout_seconds or self._settings.llm_timeout_seconds
        if tools:
            return self._invoke_with_tools(
                schema=schema,
                messages=messages,
                operation_name=resolved_operation_name,
                timeout_seconds=resolved_timeout_seconds,
                usage_notifier=usage_notifier,
                tools=tools,
            )
        if self._provider == LLM_PROVIDER_OPENAI:
            if progress_notifier is not None:
                return self._invoke_openai_responses_stream(
                    schema=schema,
                    messages=messages,
                    operation_name=resolved_operation_name,
                    timeout_seconds=resolved_timeout_seconds,
                    usage_notifier=usage_notifier,
                    progress_notifier=progress_notifier,
                )
            return self._invoke_openai_responses_parse(
                schema=schema,
                messages=messages,
                operation_name=resolved_operation_name,
                timeout_seconds=resolved_timeout_seconds,
                usage_notifier=usage_notifier,
            )

        return self._invoke_langchain_structured_output(
            schema=schema,
            messages=messages,
            operation_name=resolved_operation_name,
            timeout_seconds=resolved_timeout_seconds,
            usage_notifier=usage_notifier,
        )

    def _invoke_with_tools(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str,
        timeout_seconds: float,
        usage_notifier: UsageNotifier | None,
        tools: list[BaseTool],
        max_tool_round_trips: int = 6,
    ) -> BaseModel:
        llm = self._build_llm(timeout_seconds=timeout_seconds)
        tool_runnable = llm.bind_tools(tools).with_config(
            {
                "run_name": f"{operation_name} (tools)",
                "tags": self._tags,
                "metadata": self._metadata,
            }
        )
        tool_map = {tool.name: tool for tool in tools}
        conversation = list(messages)

        for _ in range(max_tool_round_trips):
            tool_response = tool_runnable.invoke(conversation)
            if usage_notifier is not None:
                token_usage = token_usage_from_message(tool_response)
                if token_usage is not None:
                    usage_notifier(operation_name, token_usage)

            if not isinstance(tool_response, AIMessage):
                raise TypeError("Tool-enabled runnable returned an unexpected payload.")

            conversation.append(tool_response)
            if not tool_response.tool_calls:
                return self._invoke_langchain_structured_output(
                    schema=schema,
                    messages=[
                        *conversation,
                        HumanMessage(
                            content=(
                                f"Using the tool results above, return the final {schema.__name__} "
                                "structured output only."
                            )
                        ),
                    ],
                    operation_name=operation_name,
                    timeout_seconds=timeout_seconds,
                    usage_notifier=usage_notifier,
                )

            for tool_call in tool_response.tool_calls:
                tool_name = str(tool_call.get("name", "")).strip()
                tool = tool_map.get(tool_name)
                if tool is None:
                    payload: str | dict[str, object] = {
                        "status": "error",
                        "error": f"Unknown tool '{tool_name}'.",
                    }
                else:
                    try:
                        payload = tool.invoke(tool_call.get("args", {}))
                    except Exception as exc:  # noqa: BLE001
                        payload = {
                            "status": "error",
                            "error": str(exc),
                        }
                conversation.append(
                    ToolMessage(
                        content=self._serialize_tool_payload(payload),
                        tool_call_id=str(tool_call.get("id", tool_name or "tool-call")),
                        name=tool_name,
                    )
                )

        raise RuntimeError(
            f"{operation_name} exceeded the maximum number of tool round-trips ({max_tool_round_trips})."
        )

    def _invoke_langchain_structured_output(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str,
        timeout_seconds: float,
        usage_notifier: UsageNotifier | None,
    ) -> BaseModel:
        llm = self._build_llm(timeout_seconds=timeout_seconds)
        runnable = llm.with_structured_output(schema, include_raw=True).with_config(
            {
                "run_name": operation_name,
                "tags": self._tags,
                "metadata": self._metadata,
            }
        )
        response = runnable.invoke(messages)
        if not isinstance(response, dict):
            raise TypeError("Structured output runnable returned an unexpected payload.")

        parsing_error = response.get("parsing_error")
        if parsing_error is not None:
            raise parsing_error

        raw_message = response.get("raw")
        if usage_notifier is not None and raw_message is not None:
            token_usage = token_usage_from_message(raw_message)
            if token_usage is not None:
                usage_notifier(operation_name, token_usage)

        parsed = response.get("parsed")
        if isinstance(parsed, schema):
            return parsed
        if parsed is None:
            raise ValueError("Structured output parsing returned no result.")
        return schema.model_validate(parsed)

    @staticmethod
    def _serialize_tool_payload(payload: str | dict[str, object] | object) -> str:
        if isinstance(payload, str):
            return payload
        return json.dumps(payload, ensure_ascii=True, default=str)

    def _invoke_openai_responses_parse(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str,
        timeout_seconds: float,
        usage_notifier: UsageNotifier | None,
    ) -> BaseModel:
        if self._openai_client is None:
            raise RuntimeError("OpenAI client is unavailable for Responses parsing.")

        response = self._openai_client.responses.parse(
            model=self._model,
            input=response_input_from_messages(messages),
            text_format=schema,
            temperature=0,
            metadata=self._metadata,
            timeout=timeout_seconds,
        )

        if usage_notifier is not None:
            token_usage = token_usage_from_response(response)
            if token_usage is not None:
                usage_notifier(operation_name, token_usage)

        parsed = response.output_parsed
        if isinstance(parsed, schema):
            return parsed
        if parsed is None:
            raise ValueError("Structured output parsing returned no result.")
        return schema.model_validate(parsed)

    def _invoke_openai_responses_stream(
        self,
        *,
        schema: type[BaseModel],
        messages: list[Any],
        operation_name: str,
        timeout_seconds: float,
        usage_notifier: UsageNotifier | None,
        progress_notifier: ProgressNotifier,
    ) -> BaseModel:
        if self._openai_client is None:
            raise RuntimeError("OpenAI client is unavailable for Responses streaming.")

        summary_coalescer = ReasoningSummaryCoalescer(
            operation_name=operation_name,
            notifier=progress_notifier,
        )
        with self._openai_client.responses.stream(
            model=self._model,
            input=response_input_from_messages(messages),
            text_format=schema,
            reasoning={"summary": "concise"},
            temperature=0,
            metadata=self._metadata,
            timeout=timeout_seconds,
        ) as stream:
            for event in stream:
                summary_coalescer.consume(event)
            response = stream.get_final_response()

        summary_coalescer.flush()
        if usage_notifier is not None:
            token_usage = token_usage_from_response(response)
            if token_usage is not None:
                usage_notifier(operation_name, token_usage)

        parsed = response.output_parsed
        if isinstance(parsed, schema):
            return parsed
        if parsed is None:
            raise ValueError("Structured output parsing returned no result.")
        return schema.model_validate(parsed)


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
            "This OCR JSON is saved to disk and later reused for workbook mapping and raw PDF ambiguity review.\n"
            "Capture every token needed for those steps, but do not save decorative or boilerplate text that cannot affect a workbook write, recommendation review, or planner decision.\n"
            "Prioritize structured financial facts, recommendations, tables, and chart or graph values.\n"
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
