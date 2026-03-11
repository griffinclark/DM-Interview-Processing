from __future__ import annotations

from pathlib import Path

import planlock.streamlit_app as app

from planlock.models import AgentQuestion, EntrySessionState, ImportArtifacts, QuestionOption, SheetEntrySummary
from planlock.template_entry_agent import save_entry_state


class _FormBlock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def form_submit_button(self, *args, **kwargs) -> bool:
        return False


class _ActionFormBlock(_FormBlock):
    def __init__(self, owner) -> None:
        self.owner = owner

    def form_submit_button(self, label="Submit", **kwargs) -> bool:
        return self.owner.submit_responses.get(label, False)


class _QuestionStreamlit:
    def __init__(self) -> None:
        self.session_state = {}
        self.markdown_calls: list[str] = []
        self.rerun_calls: list[dict[str, object]] = []

    def markdown(self, *args, **kwargs) -> None:
        if args:
            self.markdown_calls.append(args[0])

    def form(self, *args, **kwargs) -> _FormBlock:
        return _FormBlock()

    def columns(self, *args, **kwargs) -> list[_FormBlock]:
        return [_FormBlock(), _FormBlock()]

    def radio(self, label, options, index=0, **kwargs):
        return options[index]

    def text_input(self, *args, **kwargs) -> str:
        return ""

    def rerun(self, *args, **kwargs) -> None:
        self.rerun_calls.append(kwargs)
        return None


class _QuestionHtmlStreamlit(_QuestionStreamlit):
    def __init__(self) -> None:
        super().__init__()
        self.html_calls: list[str] = []

    def html(self, *args, **kwargs) -> None:
        if args:
            self.html_calls.append(args[0])


class _StatefulQuestionStreamlit(_QuestionHtmlStreamlit):
    def __init__(self) -> None:
        super().__init__()
        self.radio_calls: list[dict[str, object]] = []
        self.text_input_calls: list[dict[str, object]] = []

    def radio(self, label, options, index=0, **kwargs):
        key = kwargs.get("key")
        value = self.session_state.get(key, options[index])
        if key is not None:
            self.session_state[key] = value
        self.radio_calls.append({"key": key, "value": value, "options": list(options)})
        return value

    def text_input(self, *args, **kwargs) -> str:
        key = kwargs.get("key")
        value = str(self.session_state.get(key, ""))
        if key is not None:
            self.session_state[key] = value
        self.text_input_calls.append({"key": key, "value": value})
        return value


class _SubmittingQuestionStreamlit(_StatefulQuestionStreamlit):
    def __init__(self, submit_responses: dict[str, bool]) -> None:
        super().__init__()
        self.submit_responses = submit_responses

    def columns(self, *args, **kwargs) -> list[_ActionFormBlock]:
        return [_ActionFormBlock(self), _ActionFormBlock(self)]


def _entry_state(tmp_path: Path) -> EntrySessionState:
    return EntrySessionState(
        job_id="job-123",
        template_sha256="template-sha",
        workbook_path=tmp_path / "output.xlsx",
        ocr_results_path=tmp_path / "ocr_results.json",
        current_sheet_index=2,
        sheet_order=["Data Input", "Net Worth", "Expenses"],
        sheet_summaries=[
            SheetEntrySummary(
                sheet_name="Expenses",
                status="needs_input",
                message="Still ambiguous after re-checking the document.",
            )
        ],
    )


def test_render_entry_question_shows_raw_pdf_rereview_chip(monkeypatch, tmp_path: Path) -> None:
    fake_st = _QuestionStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: None)

    state_path = tmp_path / "entry_state.json"
    save_entry_state(state_path, _entry_state(tmp_path))

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=state_path,
        pending_question=AgentQuestion(
            id="expenses-travel-target",
            sheet_name="Expenses",
            prompt="Which travel target should be used?",
            rationale="The raw OCR still does not clearly distinguish the travel target.",
            affected_targets=["expenses.travel.monthly_amount"],
            pdf_rereviewed=True,
        ),
    )

    answer, source, submitted = app.render_entry_question(result)

    assert (answer, source, submitted) == (None, None, False)
    markup = "\n".join(fake_st.markdown_calls)
    assert "Raw PDF re-reviewed" in markup
    assert "Which travel target should be used?" in markup
    assert "Affected targets" not in markup
    assert "Last write summary" not in markup


def test_render_entry_question_places_live_prompt_before_workbook_stage(monkeypatch, tmp_path: Path) -> None:
    fake_st = _QuestionStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: fake_st.markdown_calls.append("WORKBOOK_STAGE"))

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use monthly or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
        ),
    )

    answer, source, submitted = app.render_entry_question(result)

    assert (answer, source, submitted) == (None, None, False)
    question_index = next(
        index for index, call in enumerate(fake_st.markdown_calls) if "Planner decision required" in call
    )
    stage_index = fake_st.markdown_calls.index("WORKBOOK_STAGE")
    assert question_index < stage_index


def test_render_entry_question_prefers_html_renderer_with_inline_markup(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _QuestionHtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use paycheck, monthly, or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
            pdf_rereviewed=True,
        ),
    )

    app.render_entry_question_form(result, {"progress_text": "0/7 completed • 7 remaining"})

    assert fake_st.html_calls
    markup = fake_st.html_calls[-1]
    assert 'class="question-shell is-inline is-live"' in markup
    assert 'class="question-panel"' in markup
    assert "Planner decision required" in markup
    assert "Data Input" in markup
    assert "0/7 completed • 7 remaining" in markup
    assert "Choose the best supported answer below" in markup
    assert "Affected targets" not in markup
    assert "Last write summary" not in markup
    assert "Current sheet" not in markup


def test_render_entry_question_form_clears_inline_inputs_between_questions(monkeypatch, tmp_path: Path) -> None:
    fake_st = _StatefulQuestionStreamlit()
    monkeypatch.setattr(app, "st", fake_st)

    stale_widget_keys = app.entry_question_widget_keys("job-123", "prior-question")
    fake_st.session_state[app.QUESTION_WIDGET_STATE_KEY] = {"job-123": "prior-question"}
    fake_st.session_state[stale_widget_keys["options"]] = "Quarterly"
    fake_st.session_state[stale_widget_keys["free_text"]] = "Legacy answer"

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use monthly or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
            options=[
                QuestionOption(label="Monthly", value="monthly"),
                QuestionOption(label="Annual", value="annual"),
            ],
        ),
    )

    app.render_entry_question_form(result, {"progress_text": "0/7 completed • 7 remaining"})

    current_widget_keys = app.entry_question_widget_keys("job-123", "income-cadence")
    assert fake_st.session_state[app.QUESTION_WIDGET_STATE_KEY]["job-123"] == app.entry_question_signature(
        result.pending_question
    )
    assert fake_st.session_state[current_widget_keys["options"]] == "Monthly"
    assert fake_st.session_state[current_widget_keys["free_text"]] == ""
    assert fake_st.radio_calls[-1]["value"] == "Monthly"
    assert fake_st.text_input_calls[-1]["value"] == ""


def test_render_entry_question_inline_submit_queues_answer_and_forces_app_rerun(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _SubmittingQuestionStreamlit({"Submit answer and continue": True})
    monkeypatch.setattr(app, "st", fake_st)

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use monthly or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
            options=[
                QuestionOption(label="Monthly", value="monthly"),
                QuestionOption(label="Annual", value="annual"),
            ],
        ),
    )

    answer, source, submitted = app.render_entry_question_form(result, {"progress_text": "0/7 completed • 7 remaining"})

    assert (answer, source, submitted) == (None, None, False)
    assert fake_st.session_state[app.QUESTION_SUBMISSION_STATE_KEY] == {
        "job_id": "job-123",
        "question_signature": app.entry_question_signature(result.pending_question),
        "answer": "monthly",
        "source": "option",
    }
    assert fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] == {
        "job_id": "job-123",
        "question_signature": app.entry_question_signature(result.pending_question),
        "sheet_name": "Data Input",
        "prompt": "Should the income values use monthly or annual amounts?",
        "progress_text": "0/7 completed • 7 remaining",
        "pdf_rereviewed": False,
    }
    assert fake_st.rerun_calls == [{"scope": "app"}]


def test_render_entry_question_inline_delegate_queues_agent_choice_and_forces_app_rerun(
    monkeypatch, tmp_path: Path
) -> None:
    fake_st = _SubmittingQuestionStreamlit({"Figure it out": True})
    monkeypatch.setattr(app, "st", fake_st)

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use monthly or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
            options=[
                QuestionOption(label="Monthly", value="monthly"),
                QuestionOption(label="Annual", value="annual"),
            ],
        ),
    )

    answer, source, submitted = app.render_entry_question_form(result, {"progress_text": "0/7 completed • 7 remaining"})

    assert (answer, source, submitted) == (None, None, False)
    assert fake_st.session_state[app.QUESTION_SUBMISSION_STATE_KEY] == {
        "job_id": "job-123",
        "question_signature": app.entry_question_signature(result.pending_question),
        "answer": "",
        "source": "agent",
    }
    assert fake_st.rerun_calls == [{"scope": "app"}]


def test_render_entry_question_handoff_renders_exit_markup_once(monkeypatch) -> None:
    fake_st = _QuestionHtmlStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": "question-signature",
        "sheet_name": "Expenses",
        "prompt": "Please provide the monthly totals by expense line item.",
        "progress_text": "2/7 completed • 5 remaining",
        "pdf_rereviewed": True,
    }

    app.render_entry_question_handoff(job_id="job-123", clear_after_render=True)

    assert fake_st.html_calls
    markup = fake_st.html_calls[-1]
    assert 'class="question-shell is-inline is-exiting"' in markup
    assert "Answer logged" in markup
    assert "Workbook entry is resuming for this section now." in markup
    assert fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] is None


def test_render_entry_question_places_handoff_before_workbook_stage(monkeypatch, tmp_path: Path) -> None:
    fake_st = _QuestionStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: fake_st.markdown_calls.append("WORKBOOK_STAGE"))

    question = AgentQuestion(
        id="income-cadence",
        sheet_name="Data Input",
        prompt="Should the income values use monthly or annual amounts?",
        rationale="The source document is ambiguous.",
        affected_targets=["income.client_1.base_salary_gross"],
    )
    question_signature = app.entry_question_signature(question)
    fake_st.session_state[app.QUESTION_SUBMISSION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": question_signature,
        "answer": "",
        "source": "agent",
    }
    fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": question_signature,
        "sheet_name": question.sheet_name,
        "prompt": question.prompt,
        "progress_text": "0/7 completed • 7 remaining",
        "pdf_rereviewed": False,
    }

    answer, source, submitted = app.render_entry_question(
        ImportArtifacts(
            success=False,
            job_id="job-123",
            job_dir=tmp_path,
            entry_state_path=None,
            pending_question=question,
        )
    )

    assert (answer, source, submitted) == ("", "agent", True)
    question_index = next(index for index, call in enumerate(fake_st.markdown_calls) if "Answer logged" in call)
    stage_index = fake_st.markdown_calls.index("WORKBOOK_STAGE")
    assert question_index < stage_index


def test_render_entry_question_ignores_submission_from_prior_question(monkeypatch, tmp_path: Path) -> None:
    fake_st = _QuestionStreamlit()
    monkeypatch.setattr(app, "st", fake_st)
    monkeypatch.setattr(app, "render_workbook_stage", lambda: None)

    prior_question = AgentQuestion(
        id="prior-question",
        sheet_name="Data Input",
        prompt="Use monthly amounts?",
        rationale="The source document is ambiguous.",
        affected_targets=["income.client_1.base_salary_gross"],
    )
    fake_st.session_state[app.QUESTION_SUBMISSION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": app.entry_question_signature(prior_question),
        "answer": "Legacy answer",
        "source": "free_text",
    }
    fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] = {
        "job_id": "job-123",
        "question_signature": app.entry_question_signature(prior_question),
        "sheet_name": "Data Input",
        "prompt": "Use monthly amounts?",
        "progress_text": "0/7 completed • 7 remaining",
        "pdf_rereviewed": False,
    }

    result = ImportArtifacts(
        success=False,
        job_id="job-123",
        job_dir=tmp_path,
        entry_state_path=None,
        pending_question=AgentQuestion(
            id="income-cadence",
            sheet_name="Data Input",
            prompt="Should the income values use monthly or annual amounts?",
            rationale="The source document is ambiguous.",
            affected_targets=["income.client_1.base_salary_gross"],
        ),
    )

    answer, source, submitted = app.render_entry_question(result)

    assert (answer, source, submitted) == (None, None, False)
    assert app.QUESTION_SUBMISSION_STATE_KEY not in fake_st.session_state
    assert fake_st.session_state[app.QUESTION_TRANSITION_STATE_KEY] is None
