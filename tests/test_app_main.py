from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import planlock.streamlit_app as app

from planlock.config import Settings
from planlock.models import (
    AgentQuestion,
    CellAssignment,
    EntrySessionState,
    ImportArtifacts,
    QuestionOption,
    ReviewReport,
    RunEvent,
    SheetEntrySummary,
    Severity,
    Stage,
    ValueKind,
)
from planlock.template_entry_agent import save_entry_state


class _Block:
    def __init__(self, owner=None):
        self.owner = owner
        self.was_cleared = False
        self.empty_count = 0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def container(self):
        return self

    def empty(self):
        self.was_cleared = True
        self.empty_count += 1
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
        self.empty_blocks: list[_Block] = []
        self.button_calls: list[dict[str, object]] = []
        self.dataframe_rows = None
        self.radio_value = None
        self.selectbox_value = None
        self.text_input_value = ""
        self.submit_responses: dict[str, bool] = {}
        self.button_responses: dict[str, bool] = {}
        self.uploaded_file_value = None
        self.dialog_calls: list[dict[str, object]] = []
        self.rerun_called = False

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
        block = _Block(self)
        self.empty_blocks.append(block)
        return block

    def columns(self, *args, **kwargs) -> list[_Block]:
        if not args:
            count = 2
        else:
            spec = args[0]
            count = spec if isinstance(spec, int) else len(spec)
        return [_Block(self) for _ in range(count)]

    def form(self, *args, **kwargs) -> _Block:
        return _Block(self)

    def file_uploader(self, *args, **kwargs):
        return self.uploaded_file_value

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
        self.button_calls.append({"label": label, "key": key, **kwargs})
        lookup_key = key if key is not None else label
        return self.button_responses.get(lookup_key, False)

    def dialog(self, title, **kwargs):
        def decorator(fn):
            def wrapped(*args, **inner_kwargs):
                self.dialog_calls.append({"title": title, **kwargs})
                return fn(*args, **inner_kwargs)

            return wrapped

        return decorator

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
        self.rerun_called = True
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


def _persist_entry_state(tmp_path: Path) -> Path:
    state_path = tmp_path / "entry_state.json"
    save_entry_state(
        state_path,
        EntrySessionState(
            job_id="job-123",
            template_sha256="template-sha",
            workbook_path=tmp_path / "output.xlsx",
            ocr_results_path=tmp_path / "ocr_results.json",
            current_sheet_index=1,
            sheet_order=["Data Input", "Expenses", "Retirement Accounts"],
            sheet_summaries=[
                SheetEntrySummary(
                    sheet_name="Data Input",
                    status="completed",
                    mapped_count=2,
                    touched_cells=["C6", "D6"],
                    message="Committed household names into the workbook.",
                )
            ],
            mapped_assignments=[
                CellAssignment(
                    sheet_name="Data Input",
                    cell="C6",
                    value="Alicia",
                    value_kind=ValueKind.STRING,
                    semantic_key="profile.client_1.first_name",
                    source_pages=[1],
                ),
                CellAssignment(
                    sheet_name="Data Input",
                    cell="D6",
                    value="Smith",
                    value_kind=ValueKind.STRING,
                    semantic_key="profile.client_1.last_name",
                    source_pages=[1],
                ),
            ],
        ),
    )
    return state_path


def test_inject_styles_forces_white_text_inside_primary_buttons(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)

    app.inject_styles()

    assert fake_st.html_calls
    markup = fake_st.html_calls[-1]
    assert '.stButton > button[kind="primary"] *,' in markup
    assert '.stButton > button[kind="primary"] [data-testid="stMarkdownContainer"] *,' in markup
    assert '.stFormSubmitButton > button [data-testid="stMarkdownContainer"] * {' in markup
    assert '-webkit-text-fill-color: #fffaf1 !important;' in markup


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


def test_append_event_tracks_langgraph_sheet_trace(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    monkeypatch.setattr(app.time, "time", lambda: 100.0)

    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="LangGraph is working on Data Input.",
            sheet_name="Data Input",
            detail_message="Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
            stage_completed=0,
            stage_total=7,
            phase="start",
        )
    )

    trace = fake_st.session_state["agent_trace"]
    assert trace["status"] == "running"
    assert trace["current_sheet"] == "Data Input"
    assert "Reviewing OCR evidence" in trace["message"]
    assert trace["started_at_ms"] == 100000
    assert trace["recent_events"][-1]["label"] == "Now working"


def test_append_event_tracks_live_reasoning_summary(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    observed_times = iter([100.0, 101.0])
    monkeypatch.setattr(app.time, "time", lambda: next(observed_times))

    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="LangGraph is working on Data Input.",
            sheet_name="Data Input",
            detail_message="Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
            stage_completed=0,
            stage_total=7,
            phase="start",
        )
    )
    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Workbook entry in progress.",
            sheet_name="Data Input",
            progress_message="Reviewing household profile evidence.",
            stage_completed=0,
            stage_total=7,
            phase="heartbeat",
        )
    )

    trace = fake_st.session_state["agent_trace"]
    assert trace["live_summary"] == "Reviewing household profile evidence."
    assert trace["live_summary_updated_at_ms"] == 101000


def test_append_event_clears_live_reasoning_summary_on_retry(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    observed_times = iter([100.0, 101.0, 102.0, 103.0])
    monkeypatch.setattr(app.time, "time", lambda: next(observed_times))

    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="LangGraph is working on Data Input.",
            sheet_name="Data Input",
            detail_message="Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
            stage_completed=0,
            stage_total=7,
            phase="start",
        )
    )
    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Workbook entry in progress.",
            sheet_name="Data Input",
            progress_message="Reviewing household profile evidence.",
            stage_completed=0,
            stage_total=7,
            phase="heartbeat",
        )
    )
    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="OpenAI rate limit hit while filling Data Input. Resuming pass 2/7 in 10.0s.",
            detail_message="Workbook entry for Data Input: HTTP 429",
            sheet_name="Data Input",
            retry_delay_seconds=10.0,
            stage_completed=0,
            stage_total=7,
            phase="retry",
            severity=Severity.WARNING,
        )
    )

    trace = fake_st.session_state["agent_trace"]
    assert trace["live_summary"] is None
    assert trace["live_summary_updated_at_ms"] is None


def test_append_event_preserves_workbook_failure_detail(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    app.append_event(
        RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Workbook entry failed on Data Input.",
            detail_message="Workbook entry for Data Input: request timed out after 120.0s",
            severity=Severity.ERROR,
            stage_completed=0,
            stage_total=7,
            sheet_name="Data Input",
            phase="failed",
        )
    )

    assert fake_st.session_state["last_status"]["detail_message"] == (
        "Workbook entry for Data Input: request timed out after 120.0s"
    )
    assert fake_st.session_state["workbook_retry"]["phase"] == "failed"
    assert fake_st.session_state["workbook_retry"]["sheet_name"] == "Data Input"


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

    def fake_render_upload_panel(*, context_key, title, copy):
        observed["context_key"] = context_key
        observed["title"] = title
        observed["copy"] = copy
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


def test_render_work_area_places_question_handoff_before_stage_focus(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["is_running"] = True
    fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": "question-signature",
        "sheet_name": "Expenses",
        "prompt": "Please provide the monthly totals by expense line item.",
        "progress_text": "2/7 completed • 5 remaining",
        "pdf_rereviewed": True,
    }

    monkeypatch.setattr(app, "render_stage_focus", lambda: fake_st.markdown_calls.append("WORKBOOK_STAGE"))

    result = app.render_work_area(_Block(fake_st), Settings.from_env())

    assert result == (None, None, None)
    handoff_index = next(index for index, call in enumerate(fake_st.markdown_calls) if "Answer logged" in call)
    stage_index = fake_st.markdown_calls.index("WORKBOOK_STAGE")
    assert handoff_index < stage_index
    assert fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] is None


def test_render_work_area_clears_placeholder_before_repeat_renders(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["is_running"] = True

    monkeypatch.setattr(app, "render_stage_focus", lambda: None)

    work_placeholder = _Block(fake_st)
    app.render_work_area(work_placeholder, Settings.from_env())
    app.render_work_area(work_placeholder, Settings.from_env())

    assert work_placeholder.empty_count == 2


def test_render_upload_panel_replaces_controls_immediately_on_build_click(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    fake_st.uploaded_file_value = _UploadedFile()
    fake_st.button_responses = {"run_import_primary": True}
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    uploaded_file, run_mode = app.render_upload_panel(
        context_key="primary",
        title="Upload PDF",
        copy="Start a new workbook build.",
    )

    assert uploaded_file is fake_st.uploaded_file_value
    assert run_mode == "upload"
    assert fake_st.empty_blocks[-1].was_cleared is True
    assert "Starting workbook build" in fake_st.markdown_calls[-1]
    assert [call["label"] for call in fake_st.button_calls] == ["Build workbook"]


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


def test_render_status_uses_timeout_card_for_workbook_timeouts(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["workbook_retry"] = {
        "sheet_name": "Data Input",
        "message": "Data Input exceeded the 120s processing limit. Retrying immediately with 180s timeout (pass 2/3).",
        "detail_message": "Workbook entry for Data Input: request timed out after 120.0s",
        "attempt_number": 2,
        "max_attempts": 3,
        "retry_delay_seconds": 0.0,
        "retry_reason": "timeout",
        "phase": "retry",
        "severity": Severity.WARNING,
    }
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "Data Input exceeded the 120s processing limit. Retrying immediately with 180s timeout (pass 2/3).",
        "severity": Severity.WARNING,
        "detail_message": "Workbook entry for Data Input: request timed out after 120.0s",
    }

    placeholder = _Placeholder()
    app.render_status(placeholder)

    markup = placeholder.markdown_calls[-1]
    assert 'class="status-card warning"' in markup
    assert "Processing timeout" in markup
    assert "Data Input is retrying" in markup
    assert "request timed out after 120.0s" in markup
    assert "Next pass" in markup
    assert "Immediate" in markup


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


def test_render_shell_chrome_keeps_chrome_in_immersive_workbook_mode(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    observed = {"masthead": 0}
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["is_running"] = True

    def fake_render_masthead() -> bool:
        observed["masthead"] += 1
        return False

    monkeypatch.setattr(app, "render_masthead", fake_render_masthead)

    chrome_placeholder = _Block()
    settings_clicked = app.render_shell_chrome(chrome_placeholder)

    assert settings_clicked is False
    assert chrome_placeholder.was_cleared is True
    assert chrome_placeholder.empty_count == 1
    assert observed == {"masthead": 1}


def test_render_shell_chrome_clears_placeholder_before_repeat_renders(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    observed = {"masthead": 0}
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    def fake_render_masthead() -> bool:
        observed["masthead"] += 1
        return False

    monkeypatch.setattr(app, "render_masthead", fake_render_masthead)

    chrome_placeholder = _Block()
    app.render_shell_chrome(chrome_placeholder)
    app.render_shell_chrome(chrome_placeholder)

    assert chrome_placeholder.empty_count == 2
    assert observed == {"masthead": 2}


def test_render_shell_chrome_run_scopes_settings_button_key_on_repeat_render(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())

    chrome_placeholder = _Block()
    app.render_shell_chrome(chrome_placeholder)
    app.render_shell_chrome(chrome_placeholder)

    assert [call["key"] for call in fake_st.button_calls] == [
        "open_run_settings",
        "open_run_settings__2",
    ]


def test_render_stage_focus_uses_shared_workbook_stage_ui(monkeypatch, tmp_path: Path) -> None:
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
    fake_st.session_state["agent_trace"] = {
        "status": "validating",
        "current_sheet": None,
        "message": "Formula review is running now.",
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "recent_events": [
            {"label": "Validation", "message": "Running workbook checks.", "tone": "success"},
        ],
    }
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=_persist_entry_state(tmp_path),
    )

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_stage_focus()

    markup = fake_st.markdown_calls[-1]
    assert "immersive-workbook-shell" in markup
    assert "Final review" in markup
    assert "Running workbook checks." in markup
    assert "Status updates" in markup
    assert "Checking the completed workbook" in markup
    assert "Final review roadmap" in markup
    assert "Committed household names into the workbook." in markup


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
    assert "Pass 2/7" in markup
    assert 'data-countdown-target-ms="135000"' in markup


def test_render_workbook_stage_shows_timeout_banner(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["workbook_retry"] = {
        "sheet_name": "Data Input",
        "message": "Data Input exceeded the 120s processing limit. Retrying immediately with 180s timeout (pass 2/3).",
        "detail_message": "Workbook entry for Data Input: request timed out after 120.0s",
        "retry_delay_seconds": 0.0,
        "attempt_number": 2,
        "max_attempts": 3,
        "retry_reason": "timeout",
        "phase": "retry",
        "severity": Severity.WARNING,
    }

    app.render_workbook_stage()

    markup = fake_st.markdown_calls[-1]
    assert "Processing timeout" in markup
    assert "Data Input is retrying" in markup
    assert "request timed out after 120.0s" in markup
    assert "Immediate" in markup


def test_render_workbook_stage_shows_setup_shell_before_first_agent_response(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "Preparing workbook entry.",
        "severity": Severity.INFO,
    }

    app.render_workbook_stage()

    markup = fake_st.markdown_calls[-1]
    assert "workbook-setup-shell" in markup
    assert "Getting the workbook ready" in markup
    assert "Waiting for the first workbook update" in markup
    assert "sample.pdf" in markup
    assert 'class="workbook-setup-step-list"' in markup
    assert "Coming up" not in markup


def test_render_workbook_stage_setup_shell_prefers_raw_html_renderer(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "Preparing workbook entry.",
        "severity": Severity.INFO,
    }

    app.render_workbook_stage()

    assert fake_st.html_calls
    markup = fake_st.html_calls[-1]
    assert "workbook-setup-shell" in markup
    assert "Preparing the live workbook view" in markup
    assert "Waiting for the first workbook update" in markup
    assert 'class="workbook-setup-step-list"' in markup
    assert fake_st.markdown_calls == []


def test_render_work_area_failure_copy_includes_detail(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = False
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "Run failed: Workbook entry failed on Data Input.",
        "severity": Severity.ERROR,
        "detail_message": "Workbook entry for Data Input: request timed out after 120.0s",
    }
    observed: dict[str, object] = {}

    def fake_render_upload_panel(*, context_key, title, copy):
        observed["context_key"] = context_key
        observed["title"] = title
        observed["copy"] = copy
        return None, None

    monkeypatch.setattr(app, "render_upload_panel", fake_render_upload_panel)

    app.render_work_area(_Block(fake_st), Settings.from_env())

    assert observed["context_key"] == "retry"
    assert observed["title"] == "Run failed"
    assert observed["copy"] == (
        "Run failed: Workbook entry failed on Data Input. Details: "
        "Workbook entry for Data Input: request timed out after 120.0s"
    )


def test_render_workbook_stage_prefers_raw_html_renderer(monkeypatch, tmp_path: Path) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (3, 6)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (0, 3)
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Data Input",
        "message": "Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
        "token_count": 1824,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "live_summary": "Reviewing the current workbook context before writing cells.",
        "live_summary_updated_at_ms": 120500,
        "last_rendered_summary": None,
        "recent_events": [
            {
                "label": "Sheet live",
                "message": "Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
                "tone": "info",
            }
        ],
    }
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=_persist_entry_state(tmp_path),
    )

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_workbook_stage()

    assert fake_st.html_calls
    assert "immersive-workbook-shell" in fake_st.html_calls[-1]
    assert "Status updates" in fake_st.html_calls[-1]
    assert "Section roadmap" in fake_st.html_calls[-1]
    assert "Committed household names into the workbook." in fake_st.html_calls[-1]
    assert "Fill workbook" in fake_st.html_calls[-1]
    assert 'class="agent-window-status running"' in fake_st.html_calls[-1]
    assert "Reviewing Data Input" in fake_st.html_calls[-1]
    assert "Reasoning summary" in fake_st.html_calls[-1]
    assert "Reviewing the current workbook context before writing cells." in fake_st.html_calls[-1]
    assert "Alicia" not in fake_st.html_calls[-1]
    assert 'class="sheet-desk-summary"' not in fake_st.html_calls[-1]
    assert 'data-elapsed-started-at-ms="120000"' in fake_st.html_calls[-1]
    assert fake_st.markdown_calls == []


def test_render_workbook_stage_falls_back_to_markdown_without_html_renderer(monkeypatch, tmp_path: Path) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (3, 6)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (0, 3)
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Data Input",
        "message": "Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
        "token_count": 1824,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "live_summary": "Reviewing the current workbook context before writing cells.",
        "live_summary_updated_at_ms": 120500,
        "last_rendered_summary": None,
        "recent_events": [],
    }
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=_persist_entry_state(tmp_path),
    )

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_workbook_stage()

    assert fake_st.markdown_calls
    assert "Reasoning summary" in fake_st.markdown_calls[-1]
    assert "Reviewing the current workbook context before writing cells." in fake_st.markdown_calls[-1]


def test_render_workbook_stage_roadmap_includes_all_sheets(monkeypatch, tmp_path: Path) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (1, 6)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (0, 3)
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Expenses",
        "message": "Reviewing the source material before adding workbook entries.",
        "token_count": 0,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "recent_events": [],
    }

    state_path = tmp_path / "entry_state.json"
    save_entry_state(
        state_path,
        EntrySessionState(
            job_id="job-123",
            template_sha256="template-sha",
            workbook_path=tmp_path / "output.xlsx",
            ocr_results_path=tmp_path / "ocr_results.json",
            current_sheet_index=1,
            sheet_order=[
                "Data Input",
                "Expenses",
                "Retirement Accounts",
                "Taxable Accounts",
                "Education Accounts",
                "Net Worth",
            ],
            sheet_summaries=[
                SheetEntrySummary(sheet_name="Data Input", status="completed", mapped_count=2),
            ],
        ),
    )
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=state_path,
    )

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_workbook_stage()

    markup = fake_st.html_calls[-1]
    for sheet_name in [
        "Transactions Raw",
        "Data Input",
        "Expenses",
        "Retirement Accounts",
        "Taxable Accounts",
        "Education Accounts",
        "Net Worth",
    ]:
        assert sheet_name in markup
    assert markup.count('class="sheet-queue-item') == 7
    assert markup.index("Transactions Raw") < markup.index("Data Input")
    assert "Autofilled" in markup
    assert "not generated by the agent" in markup


def test_render_workbook_stage_only_shows_roadmap_without_cell_preview(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (1, 6)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (0, 3)
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Expenses",
        "message": "Reviewing the source material before adding workbook entries.",
        "token_count": 0,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "recent_events": [],
    }
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=_persist_entry_state(tmp_path),
    )

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_workbook_stage()

    markup = fake_st.html_calls[-1]
    assert "Section roadmap" in markup
    assert "Entries will appear here" not in markup
    assert "Saved entries will appear here as soon as they are ready." not in markup
    assert "Cell B10" not in markup


def test_render_workbook_stage_hides_stale_question_during_resume_handoff(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (1, 6)
    fake_st.session_state["stage_progress"][Stage.FINANCIAL_CALCULATIONS.value] = (0, 3)
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Expenses",
        "message": "Letting the agent resolve the outstanding section and continue workbook entry.",
        "token_count": 0,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "recent_events": [],
    }
    question = AgentQuestion(
        id="expenses-monthly-totals",
        sheet_name="Expenses",
        prompt="Please provide the monthly totals by expense line item.",
        rationale="The transactions tab still needs category cleanup.",
    )
    state_path = tmp_path / "entry_state.json"
    save_entry_state(
        state_path,
        EntrySessionState(
            job_id="job-123",
            template_sha256="template-sha",
            workbook_path=tmp_path / "output.xlsx",
            ocr_results_path=tmp_path / "ocr_results.json",
            current_sheet_index=1,
            sheet_order=["Data Input", "Expenses", "Retirement Accounts"],
            pending_question=question,
            sheet_summaries=[
                SheetEntrySummary(
                    sheet_name="Data Input",
                    status="completed",
                    mapped_count=2,
                    message="Committed household names into the workbook.",
                ),
                SheetEntrySummary(
                    sheet_name="Expenses",
                    status="needs_input",
                    message=question.prompt,
                ),
            ],
        ),
    )
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=state_path,
        pending_question=question,
    )
    fake_st.session_state[app.QUESTION_ACTIVE_RESUME_STATE_KEY] = "job-123"

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    app.render_workbook_stage()

    markup = fake_st.html_calls[-1]
    assert "Needs your input" not in markup
    assert question.prompt not in markup
    assert "Waiting for your answer before workbook entry can continue." not in markup
    assert "Section roadmap" in markup
    assert "Expenses is active now. 1/6 sections are complete so far." in markup


def test_render_workbook_stage_uses_neutral_waiting_copy_for_pending_question(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = False
    fake_st.session_state["source_filename"] = "sample.pdf"
    fake_st.session_state["stage_progress"][Stage.DATA_ENTRY.value] = (1, 6)
    fake_st.session_state["agent_trace"] = {
        "status": "needs_input",
        "current_sheet": "Expenses",
        "message": "Workbook entry is waiting for planner confirmation.",
        "token_count": 0,
        "started_at_ms": None,
        "retry_until_ms": None,
        "live_summary": None,
        "live_summary_updated_at_ms": None,
        "last_rendered_summary": None,
        "recent_events": [],
    }
    question = AgentQuestion(
        id="expenses-monthly-totals",
        sheet_name="Expenses",
        prompt="Please provide the monthly totals by expense line item.",
        rationale="The transactions tab still needs category cleanup.",
    )
    state_path = tmp_path / "entry_state.json"
    save_entry_state(
        state_path,
        EntrySessionState(
            job_id="job-123",
            template_sha256="template-sha",
            workbook_path=tmp_path / "output.xlsx",
            ocr_results_path=tmp_path / "ocr_results.json",
            current_sheet_index=1,
            sheet_order=["Data Input", "Expenses", "Retirement Accounts"],
            pending_question=question,
            sheet_summaries=[
                SheetEntrySummary(
                    sheet_name="Data Input",
                    status="completed",
                    mapped_count=2,
                    message="Committed household names into the workbook.",
                ),
                SheetEntrySummary(
                    sheet_name="Expenses",
                    status="needs_input",
                    message=question.prompt,
                ),
            ],
        ),
    )
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=state_path,
        pending_question=question,
    )

    app.render_workbook_stage()

    markup = fake_st.html_calls[-1]
    assert "Needs your input" not in markup
    assert question.prompt not in markup
    assert "Waiting" in markup
    assert "Workbook entry is waiting for planner confirmation." in markup


def test_render_masthead_prefers_raw_html_renderer_when_available(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "LangGraph is working on Data Input.",
        "severity": Severity.INFO,
    }
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Data Input",
        "message": "LangGraph is working on Data Input.",
        "token_count": 0,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "live_summary": None,
        "live_summary_updated_at_ms": None,
        "last_rendered_summary": None,
        "recent_events": [],
    }

    monkeypatch.setattr(app.time, "time", lambda: 121.0)
    settings_clicked = app.render_masthead()

    assert settings_clicked is False
    assert fake_st.html_calls
    assert "HollyPlanner" in fake_st.html_calls[-1]
    assert 'class="taskbar-shell"' in fake_st.html_calls[-1]
    assert "Turn planner PDFs into locked workbooks" in fake_st.html_calls[-1]
    assert "Stage 2 of 2" in fake_st.html_calls[-1]
    assert "LangGraph is working on Data Input." in fake_st.html_calls[-1]
    assert "In progress" in fake_st.html_calls[-1]
    assert 'data-elapsed-started-at-ms="120000"' in fake_st.html_calls[-1]


def test_render_masthead_falls_back_to_markdown_without_html_renderer(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["last_status"] = {
        "index": 2,
        "stage": "Workbook entry",
        "message": "LangGraph is working on Data Input.",
        "severity": Severity.INFO,
    }

    settings_clicked = app.render_masthead()

    assert settings_clicked is False
    assert fake_st.markdown_calls
    assert "HollyPlanner" in fake_st.markdown_calls[0]
    assert 'class="taskbar-shell"' in fake_st.markdown_calls[0]
    assert "Turn planner PDFs into locked workbooks" in fake_st.markdown_calls[0]
    assert "Stage 2 of 2" in fake_st.markdown_calls[0]
    assert "LangGraph is working on Data Input." in fake_st.markdown_calls[0]


def test_render_masthead_compacts_rate_limit_status_into_taskbar(monkeypatch) -> None:
    fake_st = _HtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["run_started"] = True
    fake_st.session_state["is_running"] = True
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
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

    app.render_masthead()

    markup = fake_st.html_calls[-1]
    assert 'class="taskbar-status throttle"' in markup
    assert "Cooldown" in markup
    assert "Pass 2/7" in markup
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

    assert (answer, source, submitted) == (None, None, False)
    assert fake_st.dialog_calls == []
    assert fake_st.rerun_called is True
    assert any("Planner decision required" in call for call in fake_st.markdown_calls)
    assert not any("Affected targets" in call for call in fake_st.markdown_calls)

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
    assert any("Answer logged" in call for call in fake_st.markdown_calls)


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

    assert (answer, source, submitted) == (None, None, False)
    assert fake_st.dialog_calls == []
    assert fake_st.rerun_called is True

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
    assert any("Answer logged" in call for call in fake_st.markdown_calls)


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
        "status_calls": 0,
        "progress_calls": 0,
        "log_calls": 5,
        "work_calls": 5,
    }


def test_consume_runner_events_rerenders_for_live_summary_heartbeat(monkeypatch) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    app.init_state(Settings.from_env())
    fake_st.session_state["is_running"] = True
    fake_st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    fake_st.session_state["agent_trace"] = {
        "status": "running",
        "current_sheet": "Data Input",
        "message": "Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
        "token_count": 0,
        "started_at_ms": 120000,
        "retry_until_ms": None,
        "live_summary": None,
        "live_summary_updated_at_ms": None,
        "last_rendered_summary": None,
        "recent_events": [],
    }

    observed = {"chrome": 0, "logs": 0, "work": 0}

    monkeypatch.setattr(app, "render_shell_chrome", lambda placeholder: observed.__setitem__("chrome", observed["chrome"] + 1))
    monkeypatch.setattr(app, "render_logs", lambda placeholder: observed.__setitem__("logs", observed["logs"] + 1))
    monkeypatch.setattr(app, "render_work_area", lambda placeholder, settings: observed.__setitem__("work", observed["work"] + 1))

    app.consume_runner_events(
        events=[
            RunEvent(
                stage=Stage.DATA_ENTRY,
                message="Workbook entry in progress.",
                sheet_name="Data Input",
                progress_message="Reviewing household profile evidence.",
                stage_completed=0,
                stage_total=7,
                phase="heartbeat",
            )
        ],
        settings=Settings.from_env(),
        chrome_placeholder=object(),
        log_placeholder=object(),
        work_placeholder=object(),
    )

    assert observed == {"chrome": 2, "logs": 2, "work": 2}
    assert fake_st.session_state["agent_trace"]["live_summary"] == "Reviewing household profile evidence."
def test_main_queues_question_resume_before_running_job(monkeypatch, tmp_path: Path) -> None:
    observed: dict[str, object] = {"resume_calls": 0}
    fake_st = _FakeStreamlit()
    base_settings = Settings.from_env()
    fake_st.session_state["current_job_id"] = "job-42"
    pending_question = AgentQuestion(
        id="expenses-monthly-totals",
        sheet_name="Expenses",
        prompt="Please provide the monthly totals by expense line item.",
        rationale="The transactions tab still needs category cleanup.",
    )
    fake_st.session_state["result"] = ImportArtifacts(
        success=False,
        job_id="job-42",
        job_dir=tmp_path,
        pending_question=pending_question,
    )
    fake_st.session_state["agent_trace"] = {
        "status": "needs_input",
        "current_sheet": "Expenses",
        "message": pending_question.prompt,
        "token_count": 0,
        "started_at_ms": None,
        "retry_until_ms": None,
        "recent_events": [],
    }

    class _FakeJobRunner:
        def __init__(self, settings) -> None:
            observed["runner_settings"] = settings

        def resume_job(self, job_id, answer, *, source):
            observed["resume_calls"] = int(observed["resume_calls"]) + 1
            observed["resume_args"] = (job_id, answer, source)
            return []

    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "inject_styles", lambda: None)
    monkeypatch.setattr(app, "mount_live_countdown_bridge", lambda: None)
    monkeypatch.setattr(app.Settings, "from_env", classmethod(lambda cls: base_settings))
    monkeypatch.setattr(app, "apply_runtime_settings", lambda settings: settings)
    monkeypatch.setattr(app, "render_shell_chrome", lambda placeholder: False)
    monkeypatch.setattr(app, "render_logs", lambda placeholder: None)
    monkeypatch.setattr(app, "JobRunner", _FakeJobRunner)
    monkeypatch.setattr(app, "render_work_area", lambda work_placeholder, settings: (None, None, ("", "agent")))

    app.main()

    assert fake_st.rerun_called is True
    assert observed["resume_calls"] == 0
    assert fake_st.session_state[app.QUESTION_PENDING_RESUME_STATE_KEY] == {
        "job_id": "job-42",
        "answer": "",
        "source": "agent",
    }
    assert fake_st.session_state[app.QUESTION_ACTIVE_RESUME_STATE_KEY] == "job-42"
    assert fake_st.session_state["is_running"] is True
    assert fake_st.session_state["active_stage"] == Stage.DATA_ENTRY.value
    assert fake_st.session_state["last_status"]["message"] == "Resuming workbook entry."
    assert fake_st.session_state["result"].pending_question is None
    assert fake_st.session_state["agent_trace"]["status"] == "running"
    assert fake_st.session_state["agent_trace"]["current_sheet"] == "Expenses"


def test_main_resumes_question_after_follow_up_rerun(monkeypatch) -> None:
    observed: dict[str, object] = {}
    fake_st = _FakeStreamlit()
    base_settings = Settings.from_env()
    fake_st.session_state["is_running"] = True
    fake_st.session_state[app.QUESTION_PENDING_RESUME_STATE_KEY] = {
        "job_id": "job-42",
        "answer": "",
        "source": "agent",
    }

    class _FakeJobRunner:
        def __init__(self, settings) -> None:
            observed["runner_settings"] = settings

        def resume_job(self, job_id, answer, *, source):
            observed["resume_args"] = (job_id, answer, source)
            return []

    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "inject_styles", lambda: None)
    monkeypatch.setattr(app, "mount_live_countdown_bridge", lambda: None)
    monkeypatch.setattr(app.Settings, "from_env", classmethod(lambda cls: base_settings))
    monkeypatch.setattr(app, "apply_runtime_settings", lambda settings: settings)
    monkeypatch.setattr(app, "render_shell_chrome", lambda placeholder: False)
    monkeypatch.setattr(app, "render_logs", lambda placeholder: None)
    monkeypatch.setattr(app, "JobRunner", _FakeJobRunner)
    monkeypatch.setattr(app, "render_work_area", lambda work_placeholder, settings: (None, None, None))
    monkeypatch.setattr(
        app,
        "consume_runner_events",
        lambda **kwargs: observed.setdefault(
            "consume_args",
            {
                "events": kwargs["events"],
                "settings": kwargs["settings"],
            },
        ),
    )

    app.main()

    assert observed["runner_settings"] is base_settings
    assert observed["resume_args"] == ("job-42", "", "agent")
    assert observed["consume_args"]["events"] == []
    assert observed["consume_args"]["settings"] is base_settings
    assert fake_st.session_state.get(app.QUESTION_PENDING_RESUME_STATE_KEY) is None
