from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace

import planlock.streamlit_app as app

from planlock.config import Settings
from planlock.models import (
    AgentQuestion,
    CellAssignment,
    QuestionOption,
    ReviewReport,
    RunEvent,
    Severity,
    Stage,
    ValueKind,
)


class _Block:
    def __init__(self, owner=None):
        self.owner = owner

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def container(self):
        return self

    def empty(self):
        return None

    def form_submit_button(self, label="Submit", **kwargs) -> bool:
        if self.owner is None:
            return False
        return self.owner.form_submit_button(label, **kwargs)

    def button(self, label="Submit", **kwargs) -> bool:
        if self.owner is None:
            return False
        return self.owner.button(label, **kwargs)


class _Placeholder:
    def __init__(self) -> None:
        self.markdown_calls: list[str] = []

    def markdown(self, *args, **kwargs) -> None:
        if args:
            self.markdown_calls.append(args[0])
        return None


class _FakeStreamlit:
    def __init__(self) -> None:
        self.session_state = {}
        self.markdown_calls: list[str] = []
        self.empty_calls = 0
        self.dataframe_rows = None
        self.radio_value = None
        self.selectbox_value = None
        self.text_input_value = ""
        self.submit_responses: dict[str, bool] = {}
        self.button_responses: dict[str, bool] = {}

    def markdown(self, *args, **kwargs) -> None:
        if args:
            self.markdown_calls.append(args[0])
        return None

    def caption(self, *args, **kwargs) -> None:
        if args:
            self.markdown_calls.append(args[0])
        return None

    def empty(self) -> _Block:
        self.empty_calls += 1
        return _Block(self)

    def columns(self, *args, **kwargs) -> list[_Block]:
        if not args:
            count = 2
        else:
            spec = args[0]
            count = spec if isinstance(spec, int) else len(spec)
        return [_Block(self) for _ in range(count)]

    def form(self, *args, **kwargs) -> _Block:
        return _Block(self)

    def radio(self, label, options, index=0, **kwargs):
        if self.radio_value is not None:
            return self.radio_value
        if index is None:
            return None
        return options[index]

    def text_input(self, *args, **kwargs) -> str:
        return self.text_input_value

    def number_input(self, label, value=None, **kwargs):
        return value

    def selectbox(self, label, options, index=0, **kwargs):
        if self.selectbox_value is not None:
            return self.selectbox_value
        if index is None:
            return None
        return options[index]

    def form_submit_button(self, label="Submit", **kwargs) -> bool:
        return self.submit_responses.get(label, False)

    def button(self, label, key=None, **kwargs) -> bool:
        lookup_key = key if key is not None else label
        return self.button_responses.get(lookup_key, False)

    def expander(self, *args, **kwargs) -> _Block:
        return _Block(self)

    def download_button(self, *args, **kwargs) -> None:
        return None

    def dataframe(self, data, **kwargs) -> None:
        self.dataframe_rows = data

    def json(self, *args, **kwargs) -> None:
        return None

    def error(self, *args, **kwargs) -> None:
        return None

    def rerun(self) -> None:
        return None


class _HtmlStreamlit(_FakeStreamlit):
    def __init__(self) -> None:
        super().__init__()
        self.html_calls: list[str] = []

    def html(self, *args, **kwargs) -> None:
        if args:
            self.html_calls.append(args[0])
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


def test_build_provider_selector_markup_uses_external_logo_images() -> None:
    markup = app.build_provider_selector_markup("openai")

    assert "provider-logo-image" in markup
    assert 'src="https://us1.discourse-cdn.com/openai1/original/4X/3/2/1/321a1ba297482d3d4060d114860de1aa5610f8a9.png"' in markup
    assert (
        'src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRlh7pFp_23fKHEyUCA6K6V44mSYZNcboaY9A&amp;s"'
        in markup
    )
    assert "provider-card is-selected" in markup
    assert "provider-card is-disabled" in markup
    assert ">Active<" in markup
    assert ">Unavailable<" in markup
    assert "<svg" not in markup


def test_render_work_area_copy_only_mentions_parallel_lanes(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    observed: dict[str, object] = {}

    def fake_render_upload_panel(*, context_key, title, copy, cache_path):
        observed["context_key"] = context_key
        observed["title"] = title
        observed["copy"] = copy
        observed["cache_path"] = cache_path
        return None, None

    monkeypatch.setattr(app, "render_upload_panel", fake_render_upload_panel)

    app.render_work_area(_Block(fake_st), Settings.from_env())

    assert observed["context_key"] == "primary"
    assert observed["title"] == "Upload PDF"
    assert "change the number of parallel lanes" in str(observed["copy"])
    assert "provider credentials" in str(observed["copy"])
    assert "change the provider" not in str(observed["copy"])
    assert "page limit" not in str(observed["copy"])
    assert "retry timing" not in str(observed["copy"])
    assert "models" not in str(observed["copy"])


def test_apply_runtime_settings_forces_openai_provider(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    fake_st.session_state["runtime_settings"] = app.build_runtime_settings_state(Settings.from_env())
    fake_st.session_state["runtime_settings"]["llm_provider"] = "anthropic"
    monkeypatch.setattr(app, "st", fake_st)

    settings = app.apply_runtime_settings(Settings.from_env())

    assert settings.llm_provider == "openai"
    assert settings.model_ocr == "gpt-5.2"
    assert fake_st.session_state["runtime_settings"]["llm_provider"] == "openai"


def test_run_scoped_widget_key_suffixes_duplicate_widget_ids(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    assert app.run_scoped_widget_key("pdf_upload_retry") == "pdf_upload_retry"
    assert app.run_scoped_widget_key("pdf_upload_retry") == "pdf_upload_retry__2"
    assert app.run_scoped_widget_key("pdf_upload_retry") == "pdf_upload_retry__3"


def test_render_result_stringifies_mixed_assignment_values(monkeypatch, tmp_path: Path) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)

    workbook_path = tmp_path / "filled.xlsx"
    workbook_path.write_bytes(b"workbook")

    report = ReviewReport(
        job_id="job-123",
        template_sha256="template-sha",
        success=True,
        mapped_assignments=[
            CellAssignment(
                sheet_name="Data Input",
                cell="B2",
                semantic_key="profile.client_1.first_name",
                value=42,
                value_kind=ValueKind.NUMBER,
            ),
            CellAssignment(
                sheet_name="Expenses",
                cell="B10",
                semantic_key="expense.travel.label",
                value="Planner Import",
                value_kind=ValueKind.STRING,
            ),
        ],
    )

    app.render_result(
        app.ImportArtifacts(
            success=True,
            job_id="job-123",
            job_dir=tmp_path,
            output_workbook_path=workbook_path,
            review_report=report,
        )
    )

    assert fake_st.dataframe_rows is not None
    assert fake_st.dataframe_rows[0]["value"] == "42"
    assert fake_st.dataframe_rows[1]["value"] == "Planner Import"


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


def test_append_event_keeps_retry_detail_off_top_status(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    app.append_event(
        RunEvent(
            stage=Stage.OCR,
            message="Restarting lane 1 for page 1/23 (pass 2/3) after 2.0s.",
            detail_message="OCR page 1: 1 validation error for PageOcrResult",
            severity=Severity.WARNING,
            stage_completed=0,
            stage_total=23,
            page_number=1,
            page_total=23,
            pipe_number=1,
            pipe_total=1,
            attempt_number=2,
            max_attempts=3,
            retry_delay_seconds=2.0,
            retry_reason="transient",
            phase="retry",
        )
    )

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    assert fake_st.session_state["last_status"]["message"] == "Restarting lane 1 for page 1/23 (pass 2/3) after 2.0s."
    assert pipe["last_error"] == "OCR page 1: 1 validation error for PageOcrResult"


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


def test_stage_index_collapses_financial_calculations_into_workbook_step() -> None:
    assert app.stage_index(Stage.OCR) == 1
    assert app.stage_index(Stage.DATA_ENTRY) == 2
    assert app.stage_index(Stage.FINANCIAL_CALCULATIONS) == 2


def test_render_status_uses_two_step_workflow_for_financial_calculations(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    app.append_event(
        RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Formula review complete.",
            stage_completed=2,
            stage_total=3,
        )
    )

    placeholder = _Placeholder()
    app.render_status(placeholder)

    assert "Stage 2 of 2" in placeholder.markdown_calls[-1]
    assert "Workbook entry" in placeholder.markdown_calls[-1]
    assert "Formula review complete." in placeholder.markdown_calls[-1]


def test_render_status_uses_throttle_card_for_rate_limits(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["workbook_retry"] = {
        "sheet_name": "Data Input",
        "retry_until_ms": 135000,
        "attempt_number": 2,
        "max_attempts": 7,
        "retry_reason": "rate_limit",
    }
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "OpenAI rate limit hit while filling Data Input. Resuming pass 2/7 in 10.0s.",
        "severity": Severity.WARNING,
    }

    placeholder = _Placeholder()
    app.render_status(placeholder)

    markup = placeholder.markdown_calls[-1]
    assert 'class="status-card throttle"' in markup
    assert "Rate limit cooldown" in markup
    assert "Data Input is paused" in markup
    assert 'data-countdown-target-ms="135000"' in markup


def test_render_stage_progress_collapses_to_two_steps(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["active_stage"] = Stage.FINANCIAL_CALCULATIONS.value
    fake_st.session_state["stage_progress"][Stage.OCR.value] = (3, 3)
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (7, 7)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (1, 3)

    app.render_stage_progress()

    markup = fake_st.markdown_calls[-1]
    assert markup.count("stage-chip-number") == 2
    assert "Step 3" not in markup
    assert "Workbook entry" in markup


def test_render_stage_focus_uses_shared_workbook_stage_ui(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.FINANCIAL_CALCULATIONS.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (7, 7)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (1, 3)
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "Running workbook checks.",
        "severity": Severity.INFO,
    }

    app.render_stage_focus()

    markup = fake_st.markdown_calls[-1]
    assert "Workbook entry" in markup
    assert "LangGraph sheet entry" in markup
    assert "Workbook validation" in markup
    assert "Running workbook checks." in markup


def test_render_workbook_stage_shows_rate_limit_banner(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["workbook_retry"] = {
        "sheet_name": "Data Input",
        "retry_until_ms": 135000,
        "retry_delay_seconds": 10.0,
        "attempt_number": 2,
        "max_attempts": 7,
        "retry_reason": "rate_limit",
    }

    app.render_workbook_stage()

    markup = fake_st.markdown_calls[-1]
    assert "OpenAI rate limit detected" in markup
    assert "Data Input is paused." in markup
    assert "Retry in" in markup
    assert 'data-countdown-target-ms="135000"' in markup


def test_render_entry_question_supports_figure_it_out(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    fake_st.submit_responses = {
        "Submit answer and continue": False,
        "Figure it out": True,
    }
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: fake_st.markdown_calls.append("WORKBOOK_STAGE"))

    answer, source, submitted = app.render_entry_question(
        SimpleNamespace(
            job_id="job-1",
            entry_state_path=None,
            pending_question=AgentQuestion(
                id="data-input-name",
                sheet_name="Data Input",
                prompt="Which first name should populate client 1?",
                rationale="The OCR shows two possible names for client 1.",
                affected_targets=["profile.client_1.first_name"],
                options=[
                    QuestionOption(label="Taylor", value="Taylor", description="Matches the recommended planner value."),
                    QuestionOption(label="Tyler", value="Tyler", description="Appears in a secondary note."),
                ],
            ),
        )
    )

    assert submitted is True
    assert answer == ""
    assert source == "agent"
    assert any("Question for you" in call for call in fake_st.markdown_calls)
    assert any("Figure it out" in call for call in fake_st.markdown_calls)


def test_render_entry_question_prefers_custom_answer(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    fake_st.text_input_value = "Jordan"
    fake_st.submit_responses = {
        "Submit answer and continue": True,
        "Figure it out": False,
    }
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: None)

    answer, source, submitted = app.render_entry_question(
        SimpleNamespace(
            job_id="job-2",
            entry_state_path=None,
            pending_question=AgentQuestion(
                id="data-input-name",
                sheet_name="Data Input",
                prompt="Which first name should populate client 1?",
                rationale="The OCR shows two possible names for client 1.",
                affected_targets=["profile.client_1.first_name"],
                options=[QuestionOption(label="Taylor", value="Taylor")],
            ),
        )
    )

    assert submitted is True
    assert answer == "Jordan"
    assert source == "free_text"


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


def test_render_ocr_parallel_inlines_lane_warning_on_card(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "retrying"
    pipe["page_number"] = 1
    pipe["attempt_number"] = 2
    pipe["max_attempts"] = 3
    pipe["retry_delay_seconds"] = 2.0
    pipe["last_error"] = "OCR page 1: 1 validation error for PageOcrResult"
    fake_st.session_state["ocr_pipeline"]["page_total"] = 23

    app.render_ocr_parallel()

    markup = fake_st.markdown_calls[-1]
    assert "ocr-row has-warning" in markup
    assert "Lane warning" in markup
    assert "Alert" in markup
    assert "OCR page 1: 1 validation error for PageOcrResult" in markup


def test_render_ocr_parallel_shows_paused_rate_limit_state(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "retrying"
    pipe["page_number"] = 4
    pipe["attempt_number"] = 2
    pipe["max_attempts"] = 4
    pipe["retry_reason"] = "rate_limit"
    pipe["retry_until_ms"] = 135000
    fake_st.session_state["ocr_pipeline"]["page_total"] = 8
    fake_st.session_state["ocr_pipeline"]["completed_pages"] = 3

    monkeypatch.setattr(app.time, "time", lambda: 123.0)

    app.render_ocr_parallel()

    markup = fake_st.markdown_calls[-1]
    assert "Paused on page 4" in markup
    assert "OpenAI rate limit hit." in markup
    assert "Retry in" in markup
    assert 'class="ocr-dot paused"' in markup
    assert 'data-countdown-target-ms="135000"' in markup


def test_render_ocr_parallel_renders_preview_as_background_image(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "running"
    pipe["page_number"] = 5
    pipe["started_at_ms"] = 1000
    fake_st.session_state["ocr_pipeline"]["page_total"] = 23
    fake_st.session_state["page_previews"] = {
        5: "data:image/png;base64,preview-image",
    }

    monkeypatch.setattr(app.time, "time", lambda: 2.0)

    app.render_ocr_parallel()

    markup = fake_st.markdown_calls[-1]
    assert "background-image: url('data:image/png;base64,preview-image');" in markup
    assert 'role="img"' in markup
    assert "Page 5 in progress" in markup


def test_render_ocr_parallel_prefers_raw_html_renderer_when_available(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "running"
    pipe["page_number"] = 5
    pipe["started_at_ms"] = 1000
    fake_st.session_state["ocr_pipeline"]["page_total"] = 23
    fake_st.session_state["page_previews"] = {
        5: "data:image/png;base64,preview-image",
    }

    monkeypatch.setattr(app.time, "time", lambda: 2.0)

    app.render_ocr_parallel()

    assert fake_st.markdown_calls == []
    assert fake_st.html_calls
    assert "ocr-preview-shell" in fake_st.html_calls[-1]
    assert "background-image: url('data:image/png;base64,preview-image');" in fake_st.html_calls[-1]


def test_reset_run_state_clears_lane_warnings_for_new_document(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    pipe = fake_st.session_state["ocr_pipeline"]["pipes"][0]
    pipe["status"] = "failed"
    pipe["page_number"] = 3
    pipe["last_error"] = "Lane failed on previous document."

    app.reset_run_state(2, "next.pdf", {})

    assert fake_st.session_state["source_filename"] == "next.pdf"
    assert len(fake_st.session_state["ocr_pipeline"]["pipes"]) == 2
    assert all(pipe["status"] == "idle" for pipe in fake_st.session_state["ocr_pipeline"]["pipes"])
    assert all(pipe["last_error"] is None for pipe in fake_st.session_state["ocr_pipeline"]["pipes"])


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


def test_main_can_start_from_cache(monkeypatch, tmp_path: Path) -> None:
    observed: dict[str, object] = {}
    fake_st = _FakeStreamlit()
    base_settings = replace(Settings.from_env(), debug_cache_path=tmp_path / "debug_cache.txt")

    class _FakeJobRunner:
        def __init__(self, settings) -> None:
            observed["runner_settings"] = settings

        def load_phase_one_cache(self, path):
            observed["load_cache_path"] = path
            return SimpleNamespace(
                source_filename="cached.pdf",
                ocr_results=[SimpleNamespace(page_number=1), SimpleNamespace(page_number=2)],
            )

        def start_job_from_cache(self, path, *, cache=None):
            observed["start_cache_path"] = path
            observed["start_cache_filename"] = None if cache is None else cache.source_filename
            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message="Beginning workbook entry.",
                stage_completed=0,
                stage_total=1,
            )

    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "inject_styles", lambda: None)
    monkeypatch.setattr(app, "mount_live_countdown_bridge", lambda: None)
    monkeypatch.setattr(app.Settings, "from_env", classmethod(lambda cls: base_settings))
    monkeypatch.setattr(app, "apply_runtime_settings", lambda settings: settings)
    monkeypatch.setattr(app, "render_masthead", lambda: False)
    monkeypatch.setattr(app, "render_status", lambda placeholder: None)
    monkeypatch.setattr(app, "render_stage_progress", lambda: None)
    monkeypatch.setattr(app, "render_logs", lambda placeholder: None)
    monkeypatch.setattr(app, "JobRunner", _FakeJobRunner)

    def fake_render_work_area(work_placeholder, settings):
        observed["work_calls"] = int(observed.get("work_calls", 0)) + 1
        if observed["work_calls"] == 1:
            return None, "cache", None
        return None, None, None

    monkeypatch.setattr(app, "render_work_area", fake_render_work_area)
    monkeypatch.setattr(
        app,
        "reset_cached_run_state",
        lambda pipe_total, source_filename, page_total: observed.setdefault(
            "reset_cached_run_state",
            (pipe_total, source_filename, page_total),
        ),
    )

    app.main()

    assert observed["runner_settings"] is base_settings
    assert observed["load_cache_path"] == base_settings.debug_cache_path
    assert observed["start_cache_path"] == base_settings.debug_cache_path
    assert observed["start_cache_filename"] == "cached.pdf"
    assert observed["reset_cached_run_state"] == (base_settings.ocr_parallel_workers, "cached.pdf", 2)
