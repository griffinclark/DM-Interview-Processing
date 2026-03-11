from __future__ import annotations

from pathlib import Path

import planlock.streamlit_app as app

from planlock.models import AgentQuestion, EntrySessionState, ImportArtifacts, SheetEntrySummary
from planlock.template_entry_agent import save_entry_state


class _FormBlock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def form_submit_button(self, *args, **kwargs) -> bool:
        return False


class _QuestionStreamlit:
    def __init__(self) -> None:
        self.session_state = {}
        self.markdown_calls: list[str] = []

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
    assert "Raw PDF re-reviewed" in "\n".join(fake_st.markdown_calls)
