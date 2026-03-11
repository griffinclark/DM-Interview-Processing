from __future__ import annotations

import time
from threading import Lock


class RequestThrottleCoordinator:
    def __init__(self) -> None:
        self._lock = Lock()
        self._cooldown_until_monotonic = 0.0

    def wait_for_availability(self) -> None:
        while True:
            with self._lock:
                remaining = self._cooldown_until_monotonic - time.monotonic()
            if remaining <= 0:
                return
            time.sleep(remaining)

    def impose_cooldown(self, delay_seconds: float) -> None:
        if delay_seconds <= 0:
            return
        with self._lock:
            self._cooldown_until_monotonic = max(
                self._cooldown_until_monotonic,
                time.monotonic() + delay_seconds,
            )
