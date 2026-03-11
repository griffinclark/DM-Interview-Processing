from __future__ import annotations

from planlock.config import Settings
from planlock.llm_pipeline import ClaudeExtractionClient


class FakeRateLimitError(Exception):
    status_code = 429


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
        assert "failed after 4 attempt(s)" in str(exc)
    else:
        raise AssertionError("Expected RuntimeError")

    assert attempts["count"] == 4
    assert sleep_calls == [15.0, 30.0, 60.0]
    assert notifications[0] == (2, 4, 15.0)
