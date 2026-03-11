from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
import tempfile

from openpyxl import load_workbook

from planlock.models import CalculationValidationResult


def _count_formula_cells(path: Path) -> int:
    workbook = load_workbook(path)
    total = 0
    for sheet in workbook.worksheets:
        for cell in sheet._cells.values():
            if isinstance(cell.value, str) and cell.value.startswith("="):
                total += 1
    return total


def run_calculation_validation(template_path: Path, workbook_path: Path) -> CalculationValidationResult:
    template_formula_cells = _count_formula_cells(template_path)
    recalc_attempted = False
    recalc_completed = False
    warnings: list[str] = []

    soffice = shutil.which("soffice")
    if soffice:
        recalc_attempted = True
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            scratch_input = temp_dir / workbook_path.name
            shutil.copy2(workbook_path, scratch_input)
            result = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "xlsx",
                    "--outdir",
                    str(temp_dir),
                    str(scratch_input),
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            candidate = temp_dir / workbook_path.name
            if result.returncode == 0 and candidate.exists():
                recalc_completed = True
            else:
                warnings.append("LibreOffice recalculation did not complete; Excel recalc-on-open will be required.")
    else:
        warnings.append("LibreOffice not available; workbook will rely on Excel recalc-on-open.")

    output_formula_cells = _count_formula_cells(workbook_path)
    passed = output_formula_cells >= template_formula_cells
    if not passed:
        warnings.append("Formula cell count dropped below template baseline.")

    return CalculationValidationResult(
        passed=passed,
        recalc_attempted=recalc_attempted,
        recalc_completed=recalc_completed,
        formula_cells_template=template_formula_cells,
        formula_cells_output=output_formula_cells,
        warnings=warnings,
    )
