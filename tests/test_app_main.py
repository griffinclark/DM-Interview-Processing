from __future__ import annotations

from types import SimpleNamespace

import planlock.streamlit_app as app

from planlock.config import Settings
from planlock.models import RunEvent, Stage


class _Block:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def container(self):
        return self

    def empty(self):
        return None


class _FakeStreamlit:
    def __init__(self) -> None:
        self.session_state = {}
        self.markdown_calls: list[str] = []
        self.empty_calls = 0

    def markdown(self, *args, **kwargs) -> None:
        if args:
            self.markdown_calls.append(args[0])
        return None

    def empty(self) -> _Block:
        self.empty_calls += 1
        return _Block()

    def columns(self, *args, **kwargs) -> list[_Block]:
        return [_Block(), _Block()]

    def error(self, *args, **kwargs) -> None:
        return None


class _UploadedFile:
    name = "sample.pdf"

    def getvalue(self) -> bytes:
        return b"%PDF-1.4"


def test_main_passes_settings_to_init_and_work_area(monkeypatch) -> None:
    observed: dict[str, object] = {}
    base_settings = SimpleNamespace(validate_template_lock=lambda: "locked-template-sha")
    runtime_settings = SimpleNamespace(ocr_parallel_workers=3, max_pages=40)

    monkeypatch.setattr(app, "st", _FakeStreamlit())
    monkeypatch.setattr(app, "inject_styles", lambda: None)
    monkeypatch.setattr(app, "mount_live_countdown_bridge", lambda: observed.setdefault("countdown_bridge", True))
    monkeypatch.setattr(app.Settings, "from_env", classmethod(lambda cls: base_settings))
    monkeypatch.setattr(app, "apply_runtime_settings", lambda settings: runtime_settings)
    monkeypatch.setattr(app, "render_masthead", lambda: False)
    monkeypatch.setattr(app, "render_status", lambda placeholder: None)
    monkeypatch.setattr(app, "render_stage_progress", lambda: None)
    monkeypatch.setattr(app, "render_logs", lambda placeholder: None)

    def fake_init_state(settings) -> None:
        observed["init_state"] = settings

    def fake_render_work_area(work_placeholder, settings):
        observed["render_work_area"] = settings
        return None, False

    monkeypatch.setattr(app, "init_state", fake_init_state)
    monkeypatch.setattr(app, "render_work_area", fake_render_work_area)

    app.main()

    assert observed["countdown_bridge"] is True
    assert observed["init_state"] is base_settings
    assert observed["render_work_area"] is runtime_settings


def test_append_event_ignores_heartbeat_updates_in_logs(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Reviewing page 1/1 in lane 1.",
            stage_completed=0,
            stage_total=1,
            page_number=1,
            pipe_number=1,
            pipe_total=1,
            phase="start",
        )
    )

    prior_status = fake_st.session_state["last_status"]
    prior_log_count = len(fake_st.session_state["logs"])

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Document review in progress.",
            stage_completed=0,
            stage_total=1,
            pipe_total=1,
            phase="heartbeat",
        )
    )

    assert fake_st.session_state["last_status"] == prior_status
    assert len(fake_st.session_state["logs"]) == prior_log_count


def test_build_timing_markup_adds_live_countdown_target() -> None:
    markup = app.build_timing_markup("12.0s", countdown_target_ms=123456)

    assert 'class="ocr-live-countdown"' in markup
    assert 'data-countdown-target-ms="123456"' in markup
    assert ">12.0s<" in markup


def test_build_timing_markup_adds_live_elapsed_target() -> None:
    markup = app.build_timing_markup("3.2s", elapsed_started_at_ms=654321)

    assert 'class="ocr-live-elapsed"' in markup
    assert 'data-elapsed-started-at-ms="654321"' in markup
    assert ">3.2s<" in markup


def test_build_timing_markup_supports_handoff_animation_offset() -> None:
    markup = app.build_timing_markup(
        "Retrying",
        extra_classes=("ocr-live-handoff",),
        animation_offset_ms=180,
    )

    assert 'class="ocr-live-handoff"' in markup
    assert 'style="animation-delay: -180ms;"' in markup
    assert ">Retrying<" in markup


def test_append_event_marks_finished_cooldown_once_timer_expires(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    timestamps = iter([100.0, 101.2])
    monkeypatch.setattr(app.time, "time", lambda: next(timestamps))

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Cooling down before retrying page 1/1 in lane 1.",
            stage_completed=0,
            stage_total=1,
            page_number=1,
            page_total=1,
            pipe_number=1,
            pipe_total=1,
            attempt_number=2,
            max_attempts=3,
            retry_delay_seconds=1.0,
            retry_reason="rate_limit",
            phase="retry",
        )
    )

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    assert pipe["cooldown_finished_at_ms"] is None

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Document review in progress.",
            stage_completed=0,
            stage_total=1,
            pipe_total=1,
            phase="heartbeat",
        )
    )

    assert pipe["cooldown_finished_at_ms"] == pipe["retry_until_ms"]


def test_append_event_reuses_lane_without_batch_reset(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Reviewing page 1/2 in lane 1.",
            stage_completed=0,
            stage_total=2,
            page_number=1,
            page_total=2,
            pipe_number=1,
            pipe_total=1,
            phase="start",
        )
    )
    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Finished reviewing page 1/2 in lane 1.",
            stage_completed=1,
            stage_total=2,
            page_number=1,
            page_total=2,
            pipe_number=1,
            pipe_total=1,
            phase="complete",
        )
    )
    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Reviewing page 2/2 in lane 1.",
            stage_completed=1,
            stage_total=2,
            page_number=2,
            page_total=2,
            pipe_number=1,
            pipe_total=1,
            phase="start",
        )
    )

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    assert pipe["status"] == "running"
    assert pipe["page_number"] == 2
    assert pipe["last_completed_page"] == 1


def test_render_ocr_parallel_uses_live_card_and_consumes_flash(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "complete"
    pipe["last_completed_page"] = 1
    pipe["completed_at_ms"] = 123456
    pipe["flash_complete"] = True
    fake_st.session_state["ocr_pipeline"]["page_total"] = 1
    fake_st.session_state["ocr_pipeline"]["completed_pages"] = 1

    monkeypatch.setattr(app.time, "time", lambda: 123.456)

    app.render_ocr_parallel()

    assert 'class="section-card live-card"' in fake_st.markdown_calls[-1]
    assert "ocr-row inactive just-completed" in fake_st.markdown_calls[-1]
    assert pipe["flash_complete"] is False

    app.render_ocr_parallel()

    assert "ocr-row inactive just-completed" not in fake_st.markdown_calls[-1]


def test_main_skips_ui_rerender_on_heartbeat(monkeypatch) -> None:
    observed = {
        "status_calls": 0,
        "progress_calls": 0,
        "log_calls": 0,
        "work_calls": 0,
    }
    fake_st = _FakeStreamlit()
    base_settings = Settings.from_env()
    uploaded_file = _UploadedFile()

    class _FakeJobRunner:
        def __init__(self, settings) -> None:
            self.settings = settings

        def run(self, pdf_bytes: bytes, original_filename: str):
            yield RunEvent(
                stage=Stage.OCR,
                message="Reviewing page 1/1 in lane 1.",
                stage_completed=0,
                stage_total=1,
                page_number=1,
                page_total=1,
                pipe_number=1,
                pipe_total=1,
                phase="start",
            )
            yield RunEvent(
                stage=Stage.OCR,
                message="Document review in progress.",
                stage_completed=0,
                stage_total=1,
                pipe_total=1,
                phase="heartbeat",
            )
            yield RunEvent(
                stage=Stage.OCR,
                message="Finished reviewing page 1/1 in lane 1.",
                stage_completed=1,
                stage_total=1,
                page_number=1,
                page_total=1,
                pipe_number=1,
                pipe_total=1,
                phase="complete",
            )

    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "inject_styles", lambda: None)
    monkeypatch.setattr(app, "mount_live_countdown_bridge", lambda: None)
    monkeypatch.setattr(app.Settings, "from_env", classmethod(lambda cls: base_settings))
    monkeypatch.setattr(app, "apply_runtime_settings", lambda settings: settings)
    monkeypatch.setattr(app, "render_masthead", lambda: False)
    monkeypatch.setattr(app, "render_pdf_previews", lambda uploaded_bytes, max_pages: {})
    monkeypatch.setattr(app, "JobRunner", _FakeJobRunner)

    def fake_render_status(placeholder) -> None:
        observed["status_calls"] += 1

    def fake_render_stage_progress() -> None:
        observed["progress_calls"] += 1

    def fake_render_logs(placeholder) -> None:
        observed["log_calls"] += 1

    def fake_render_work_area(work_placeholder, settings):
        observed["work_calls"] += 1
        if observed["work_calls"] == 1:
            return uploaded_file, True
        return None, False

    monkeypatch.setattr(app, "render_status", fake_render_status)
    monkeypatch.setattr(app, "render_stage_progress", fake_render_stage_progress)
    monkeypatch.setattr(app, "render_logs", fake_render_logs)
    monkeypatch.setattr(app, "render_work_area", fake_render_work_area)

    app.main()

    assert observed == {
        "status_calls": 5,
        "progress_calls": 5,
        "log_calls": 5,
        "work_calls": 5,
    }
