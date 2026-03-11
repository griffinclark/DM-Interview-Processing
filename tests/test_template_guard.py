from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from planlock.config import Settings
from planlock.models import CellAssignment, ValueKind
from planlock.template_guard import check_for_drift
from planlock.template_schema import ALLOWED_WRITE_CELLS_BY_SHEET
from planlock.workbook_writer import apply_assignments_to_workbook


def _copy_template(tmp_path: Path) -> Path:
    settings = Settings.from_env()
    target = tmp_path / "template-copy.xlsx"
    target.write_bytes(settings.template_path.read_bytes())
    return target


def test_allowed_input_change_does_not_trigger_drift(tmp_path: Path) -> None:
    settings = Settings.from_env()
    workbook_path = _copy_template(tmp_path)
    apply_assignments_to_workbook(
        workbook_path,
        [
            CellAssignment(
                sheet_name="Data Input",
                cell="C6",
                value="Taylor",
                value_kind=ValueKind.STRING,
                semantic_key="profile.client_1.first_name",
                source_pages=[1],
            )
        ],
    )

    drift = check_for_drift(settings.template_path, workbook_path, ALLOWED_WRITE_CELLS_BY_SHEET)

    assert drift.passed, drift.violations


def test_sheet_rename_triggers_drift_failure(tmp_path: Path) -> None:
    settings = Settings.from_env()
    workbook_path = _copy_template(tmp_path)
    workbook = load_workbook(workbook_path)
    workbook["Expenses"].title = "Expenses Changed"
    workbook.save(workbook_path)

    drift = check_for_drift(settings.template_path, workbook_path, ALLOWED_WRITE_CELLS_BY_SHEET)

    assert not drift.passed
    assert any("sheet names" in violation.lower() for violation in drift.violations)
