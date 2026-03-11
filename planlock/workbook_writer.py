from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
import shutil
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import to_excel

from planlock.config import Settings
from planlock.models import (
    AccountCandidate,
    CanonicalPlanDocument,
    CellAssignment,
    ImportWarning,
    Severity,
    Stage,
    ValueKind,
)
from planlock.template_schema import (
    DATA_INPUT_ACCOUNT_ROWS,
    EDUCATION_BLOCKS,
    EXPENSE_ROW_BLOCKS,
    FIELD_TARGETS,
    NET_WORTH_ASSET_ROWS,
    NET_WORTH_LIABILITY_ROWS,
    PORTFOLIO_BLOCKS,
    RETIREMENT_BLOCKS,
    TAXABLE_BLOCKS,
    is_allowed_formula,
    is_allowed_write,
)


SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", SPREADSHEET_NS)
ET.register_namespace("r", REL_NS)


LIABILITY_ACCOUNT_HINTS = (
    "credit",
    "debt",
    "heloc",
    "liability",
    "loan",
    "mortgage",
    "payable",
)


def copy_locked_template(settings: Settings, job_dir: Path) -> Path:
    output_path = job_dir / settings.workbook_output_name
    shutil.copy2(settings.template_path, output_path)
    return output_path


def _coerce_value(value: object, value_kind: ValueKind) -> object:
    if value is None:
        return None
    if value_kind == ValueKind.DATE and isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
        return value
    return value


def _tag(name: str) -> str:
    return f"{{{SPREADSHEET_NS}}}{name}"


def _sheet_paths(workbook_path: Path) -> dict[str, str]:
    ns = {"a": SPREADSHEET_NS, "r": REL_NS}
    with ZipFile(workbook_path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root}
    mapping: dict[str, str] = {}
    for sheet in workbook_root.find("a:sheets", ns):
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        mapping[sheet.attrib["name"]] = "xl/" + rel_map[rel_id]
    return mapping


def _find_or_create_row(sheet_root: ET.Element, row_number: int) -> ET.Element:
    sheet_data = sheet_root.find(_tag("sheetData"))
    assert sheet_data is not None
    for row in sheet_data.findall(_tag("row")):
        if int(row.attrib["r"]) == row_number:
            return row
    row = ET.Element(_tag("row"), {"r": str(row_number)})
    sheet_data.append(row)
    return row


def _find_or_create_cell(row_element: ET.Element, cell_ref: str) -> ET.Element:
    for cell in row_element.findall(_tag("c")):
        if cell.attrib.get("r") == cell_ref:
            return cell
    cell = ET.Element(_tag("c"), {"r": cell_ref})
    row_element.append(cell)
    return cell


def _clear_cell_payload(cell_element: ET.Element) -> None:
    for child in list(cell_element):
        if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
            cell_element.remove(child)


def _set_string(cell_element: ET.Element, value: str) -> None:
    cell_element.set("t", "inlineStr")
    is_element = ET.SubElement(cell_element, _tag("is"))
    text_element = ET.SubElement(is_element, _tag("t"))
    if value != value.strip() or "\n" in value:
        text_element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_element.text = value


def _set_number(cell_element: ET.Element, value: int | float) -> None:
    cell_element.attrib.pop("t", None)
    v = ET.SubElement(cell_element, _tag("v"))
    v.text = str(value)


def _set_boolean(cell_element: ET.Element, value: bool) -> None:
    cell_element.set("t", "b")
    v = ET.SubElement(cell_element, _tag("v"))
    v.text = "1" if value else "0"


def _set_formula(cell_element: ET.Element, value: str) -> None:
    cell_element.attrib.pop("t", None)
    formula = ET.SubElement(cell_element, _tag("f"))
    formula.text = value[1:] if value.startswith("=") else value
    ET.SubElement(cell_element, _tag("v"))


def _set_date(cell_element: ET.Element, value: object) -> None:
    coerced = _coerce_value(value, ValueKind.DATE)
    if isinstance(coerced, datetime):
        serial = to_excel(coerced)
    elif isinstance(coerced, date):
        serial = to_excel(datetime.combine(coerced, datetime.min.time()))
    else:
        _set_string(cell_element, str(value))
        return
    _set_number(cell_element, serial)


def _apply_to_sheet_xml(xml_bytes: bytes, assignments: list[CellAssignment]) -> bytes:
    sheet_root = ET.fromstring(xml_bytes)
    for assignment in assignments:
        row_number, _ = coordinate_to_tuple(assignment.cell)
        row_element = _find_or_create_row(sheet_root, row_number)
        cell_element = _find_or_create_cell(row_element, assignment.cell)
        _clear_cell_payload(cell_element)
        if assignment.value is None:
            cell_element.attrib.pop("t", None)
            continue
        if assignment.value_kind == ValueKind.STRING:
            _set_string(cell_element, str(assignment.value))
        elif assignment.value_kind == ValueKind.NUMBER:
            _set_number(cell_element, assignment.value)  # type: ignore[arg-type]
        elif assignment.value_kind == ValueKind.BOOLEAN:
            _set_boolean(cell_element, bool(assignment.value))
        elif assignment.value_kind == ValueKind.FORMULA:
            _set_formula(cell_element, str(assignment.value))
        elif assignment.value_kind == ValueKind.DATE:
            _set_date(cell_element, assignment.value)
        else:
            _set_string(cell_element, str(assignment.value))
    return ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)


def _update_calc_properties(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    calc_pr = root.find(_tag("calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(root, _tag("calcPr"))
    calc_pr.set("calcMode", "auto")
    calc_pr.set("fullCalcOnLoad", "1")
    calc_pr.set("forceFullCalc", "1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _expense_assignments(document: CanonicalPlanDocument) -> list[CellAssignment]:
    assignments: list[CellAssignment] = []
    for expense in document.expenses:
        start_row, _ = EXPENSE_ROW_BLOCKS[expense.category]
        source_pages = [expense.page_number] if expense.page_number else []
        assignments.append(
            CellAssignment(
                sheet_name="Expenses",
                cell=f"B{start_row}",
                value=expense.label,
                value_kind=ValueKind.STRING,
                semantic_key=f"expense.{expense.category}.label",
                source_pages=source_pages,
                comment=expense.comment,
            )
        )
        if expense.monthly_amount is not None:
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"D{start_row}",
                    value=expense.monthly_amount,
                    value_kind=ValueKind.NUMBER,
                    semantic_key=f"expense.{expense.category}.monthly",
                    source_pages=source_pages,
                    comment=expense.comment,
                )
            )
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"C{start_row}",
                    value=f'=IF(D{start_row}="","",D{start_row}*12)',
                    value_kind=ValueKind.FORMULA,
                    semantic_key=f"expense.{expense.category}.yearly_formula",
                    source_pages=source_pages,
                    comment="Derived from imported monthly amount.",
                )
            )
        elif expense.yearly_amount is not None:
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"C{start_row}",
                    value=expense.yearly_amount,
                    value_kind=ValueKind.NUMBER,
                    semantic_key=f"expense.{expense.category}.yearly",
                    source_pages=source_pages,
                    comment=expense.comment,
                )
            )
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"D{start_row}",
                    value=f'=IF(C{start_row}="","",C{start_row}/12)',
                    value_kind=ValueKind.FORMULA,
                    semantic_key=f"expense.{expense.category}.monthly_formula",
                    source_pages=source_pages,
                    comment="Derived from imported yearly amount.",
                )
            )
        if expense.comment:
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"F{start_row}",
                    value=expense.comment,
                    value_kind=ValueKind.STRING,
                    semantic_key=f"expense.{expense.category}.note",
                    source_pages=source_pages,
                    comment=expense.comment,
                )
            )
        if expense.discretionary is not None:
            assignments.append(
                CellAssignment(
                    sheet_name="Expenses",
                    cell=f"G{start_row}",
                    value=expense.discretionary,
                    value_kind=ValueKind.BOOLEAN,
                    semantic_key=f"expense.{expense.category}.discretionary",
                    source_pages=source_pages,
                    comment=expense.comment,
                )
            )
    return assignments


def _clean_account_text(value: str | None) -> str | None:
    cleaned = (value or "").strip()
    return cleaned or None


def _net_worth_section_for_account(account: AccountCandidate) -> str:
    if account.net_worth_section in {"asset", "liability"}:
        return account.net_worth_section
    if isinstance(account.balance, (int, float)) and account.balance < 0:
        return "liability"

    text_parts = [
        account.account_type,
        account.owner_name,
        account.account_identifier,
        account.institution,
        account.notes,
        account.source_excerpt,
    ]
    normalized = " ".join(part.strip().lower() for part in text_parts if isinstance(part, str) and part.strip())
    if any(keyword in normalized for keyword in LIABILITY_ACCOUNT_HINTS):
        return "liability"
    return "asset"


def _net_worth_label_for_account(account: AccountCandidate) -> str:
    parts: list[str] = []
    seen: set[str] = set()
    for raw_value in [
        account.institution,
        account.account_type,
        account.owner_name,
        account.account_identifier,
    ]:
        cleaned = _clean_account_text(raw_value)
        if cleaned is None:
            continue
        signature = cleaned.casefold()
        if signature in seen:
            continue
        seen.add(signature)
        parts.append(cleaned)

    if parts:
        return " - ".join(parts)

    fallback = _clean_account_text(account.notes) or _clean_account_text(account.source_excerpt)
    if fallback is not None:
        return fallback
    return "Imported account"


def _account_assignments(document: CanonicalPlanDocument) -> tuple[list[CellAssignment], list[ImportWarning]]:
    assignments: list[CellAssignment] = []
    warnings: list[ImportWarning] = []
    for row, account in zip(DATA_INPUT_ACCOUNT_ROWS, document.accounts):
        source_pages = [account.page_number] if account.page_number else []
        mapping = {
            f"B{row}": account.account_type,
            f"C{row}": account.owner_name,
            f"D{row}": account.account_identifier,
            f"E{row}": account.apy,
            f"F{row}": account.institution,
            f"G{row}": account.balance,
            f"H{row}": account.monthly_contribution,
            f"I{row}": account.last_updated,
            f"J{row}": account.notes,
        }
        for cell, value in mapping.items():
            if value is None:
                continue
            assignments.append(
                CellAssignment(
                    sheet_name="Data Input",
                    cell=cell,
                    value=value,
                    value_kind=ValueKind.NUMBER if isinstance(value, (int, float)) else ValueKind.STRING,
                    semantic_key=f"account.{row}.{cell}",
                    source_pages=source_pages,
                    comment=account.notes,
                )
            )

    net_worth_row_index = {"asset": 0, "liability": 0}
    net_worth_overflow: set[str] = set()
    net_worth_rows_by_section = {
        "asset": NET_WORTH_ASSET_ROWS,
        "liability": NET_WORTH_LIABILITY_ROWS,
    }
    for account in document.accounts:
        if account.balance is None:
            continue

        section = _net_worth_section_for_account(account)
        rows = net_worth_rows_by_section[section]
        row_index = net_worth_row_index[section]
        if row_index >= len(rows):
            net_worth_overflow.add(section)
            continue

        row = rows[row_index]
        net_worth_row_index[section] += 1
        source_pages = [account.page_number] if account.page_number else []
        balance = abs(account.balance) if section == "liability" else account.balance
        label = _net_worth_label_for_account(account)
        assignments.extend(
            [
                CellAssignment(
                    sheet_name="Net Worth",
                    cell=f"B{row}",
                    value=label,
                    value_kind=ValueKind.STRING,
                    semantic_key=f"net_worth.{section}.{row}.label",
                    source_pages=source_pages,
                    comment=account.notes,
                ),
                CellAssignment(
                    sheet_name="Net Worth",
                    cell=f"C{row}",
                    value=balance,
                    value_kind=ValueKind.NUMBER,
                    semantic_key=f"net_worth.{section}.{row}.balance",
                    source_pages=source_pages,
                    comment=account.notes,
                ),
            ]
        )

    if len(document.accounts) > len(DATA_INPUT_ACCOUNT_ROWS):
        warnings.append(
            ImportWarning(
                code="account_capacity_exceeded",
                message="Not all net-worth accounts fit into the locked workbook rows.",
                severity=Severity.WARNING,
                stage=Stage.DATA_ENTRY,
            )
        )
    for section in sorted(net_worth_overflow):
        warnings.append(
            ImportWarning(
                code="net_worth_capacity_exceeded",
                message=f"Not all {section} accounts fit into the Net Worth sheet.",
                severity=Severity.WARNING,
                stage=Stage.DATA_ENTRY,
            )
        )
    return assignments, warnings


def _portfolio_blocks_by_key() -> dict[tuple[str, str], list]:
    grouped: dict[tuple[str, str], list] = defaultdict(list)
    for block in PORTFOLIO_BLOCKS:
        grouped[(block.sheet_name, block.owner_section)].append(block)
    return grouped


def _holding_assignments(document: CanonicalPlanDocument) -> tuple[list[CellAssignment], list[ImportWarning]]:
    assignments: list[CellAssignment] = []
    warnings: list[ImportWarning] = []
    blocks = _portfolio_blocks_by_key()
    grouped_holdings: dict[tuple[str, str, str], list] = defaultdict(list)
    for holding in document.holdings:
        grouped_holdings[(holding.sheet_name, holding.owner_section, holding.account_name)].append(holding)

    for key, holdings in grouped_holdings.items():
        sheet_name, owner_section, account_name = key
        available_blocks = blocks.get((sheet_name, owner_section), [])
        if not available_blocks:
            warnings.append(
                ImportWarning(
                    code="missing_portfolio_block",
                    message=f"No template block available for {sheet_name} / {owner_section}.",
                    severity=Severity.WARNING,
                    stage=Stage.DATA_ENTRY,
                )
            )
            continue

        block = available_blocks.pop(0)
        assignments.append(
            CellAssignment(
                sheet_name=sheet_name,
                cell=f"B{block.account_row}",
                value=account_name,
                value_kind=ValueKind.STRING,
                semantic_key=f"{sheet_name}.{owner_section}.{account_name}.header",
                source_pages=sorted({holding.page_number for holding in holdings if holding.page_number}),
            )
        )
        for row, holding in zip(range(block.holding_start_row, block.holding_end_row + 1), holdings):
            source_pages = [holding.page_number] if holding.page_number else []
            assignments.extend(
                [
                    CellAssignment(
                        sheet_name=sheet_name,
                        cell=f"B{row}",
                        value=holding.holding_name,
                        value_kind=ValueKind.STRING,
                        semantic_key=f"{sheet_name}.{account_name}.{row}.holding_name",
                        source_pages=source_pages,
                        comment=holding.comment,
                    ),
                    CellAssignment(
                        sheet_name=sheet_name,
                        cell=f"C{row}",
                        value=holding.symbol,
                        value_kind=ValueKind.STRING,
                        semantic_key=f"{sheet_name}.{account_name}.{row}.symbol",
                        source_pages=source_pages,
                        comment=holding.comment,
                    ),
                    CellAssignment(
                        sheet_name=sheet_name,
                        cell=f"D{row}",
                        value=holding.category,
                        value_kind=ValueKind.STRING,
                        semantic_key=f"{sheet_name}.{account_name}.{row}.category",
                        source_pages=source_pages,
                        comment=holding.comment,
                    ),
                ]
            )
            numeric_pairs = {
                f"E{row}": holding.expense_ratio,
                f"F{row}": holding.yield_pct,
                f"G{row}": holding.one_year_return_pct,
                f"H{row}": holding.five_year_return_pct,
                f"I{row}": holding.shares,
                f"J{row}": holding.price,
            }
            for cell, value in numeric_pairs.items():
                if value is not None:
                    assignments.append(
                        CellAssignment(
                            sheet_name=sheet_name,
                            cell=cell,
                            value=value,
                            value_kind=ValueKind.NUMBER,
                            semantic_key=f"{sheet_name}.{account_name}.{row}.{cell}",
                            source_pages=source_pages,
                            comment=holding.comment,
                        )
                    )

            assignments.append(
                CellAssignment(
                    sheet_name=sheet_name,
                    cell=f"K{row}",
                    value=f'=IF(AND(I{row}<>"",J{row}<>""),I{row}*J{row},"")',
                    value_kind=ValueKind.FORMULA,
                    semantic_key=f"{sheet_name}.{account_name}.{row}.value_formula",
                    source_pages=source_pages,
                    comment="Derived from shares x price.",
                )
            )

            if sheet_name in {"Taxable Accounts", "Education Accounts"}:
                if holding.purchase_price is not None:
                    assignments.append(
                        CellAssignment(
                            sheet_name=sheet_name,
                            cell=f"L{row}",
                            value=holding.purchase_price,
                            value_kind=ValueKind.NUMBER,
                            semantic_key=f"{sheet_name}.{account_name}.{row}.purchase_price",
                            source_pages=source_pages,
                            comment=holding.comment,
                        )
                    )
                assignments.append(
                    CellAssignment(
                        sheet_name=sheet_name,
                        cell=f"M{row}",
                        value=f'=IF(AND(I{row}<>"",J{row}<>"",L{row}<>""),I{row}*(J{row}-L{row}),"")',
                        value_kind=ValueKind.FORMULA,
                        semantic_key=f"{sheet_name}.{account_name}.{row}.gain_formula",
                        source_pages=source_pages,
                        comment="Derived unrealized gain/loss.",
                    )
                )

        if len(holdings) > (block.holding_end_row - block.holding_start_row + 1):
            warnings.append(
                ImportWarning(
                    code="holding_capacity_exceeded",
                    message=f"Not all holdings fit in {sheet_name} block {account_name}.",
                    severity=Severity.WARNING,
                    stage=Stage.DATA_ENTRY,
                )
            )

    return assignments, warnings


def build_assignments(document: CanonicalPlanDocument) -> tuple[list[CellAssignment], list[ImportWarning]]:
    assignments: list[CellAssignment] = []
    warnings: list[ImportWarning] = []

    for key, candidate in document.fields.items():
        sheet_name, cell = FIELD_TARGETS[key]
        assignments.append(
            CellAssignment(
                sheet_name=sheet_name,
                cell=cell,
                value=candidate.value,
                value_kind=candidate.value_kind,
                semantic_key=key,
                source_pages=[candidate.page_number] if candidate.page_number else [],
                comment=candidate.comment,
            )
        )

    assignments.extend(_expense_assignments(document))
    account_assignments, account_warnings = _account_assignments(document)
    assignments.extend(account_assignments)
    warnings.extend(account_warnings)
    holding_assignments, holding_warnings = _holding_assignments(document)
    assignments.extend(holding_assignments)
    warnings.extend(holding_warnings)

    return assignments, warnings


def apply_assignments_to_workbook(
    workbook_path: Path,
    assignments: list[CellAssignment],
) -> None:
    sheet_paths = _sheet_paths(workbook_path)
    assignments_by_sheet: dict[str, list[CellAssignment]] = defaultdict(list)
    for assignment in assignments:
        if not is_allowed_write(assignment.sheet_name, assignment.cell):
            raise ValueError(
                f"Attempted write outside whitelist: {assignment.sheet_name}!{assignment.cell}"
            )
        if assignment.value_kind == ValueKind.FORMULA and not is_allowed_formula(
            assignment.sheet_name, assignment.cell
        ):
            raise ValueError(
                f"Attempted formula write outside allowed formula cells: "
                f"{assignment.sheet_name}!{assignment.cell}"
            )
        assignments_by_sheet[assignment.sheet_name].append(assignment)

    with ZipFile(workbook_path, "r") as source_archive:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_handle:
            temp_path = Path(temp_handle.name)

        with ZipFile(temp_path, "w") as target_archive:
            for member in source_archive.infolist():
                data = source_archive.read(member.filename)
                if member.filename == "xl/workbook.xml":
                    data = _update_calc_properties(data)
                else:
                    for sheet_name, sheet_assignments in assignments_by_sheet.items():
                        sheet_path = sheet_paths[sheet_name]
                        if member.filename == sheet_path:
                            data = _apply_to_sheet_xml(data, sheet_assignments)
                            break
                target_archive.writestr(member, data)

    shutil.move(str(temp_path), workbook_path)
