from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

from planlock.transactions_query import (
    TransactionQueryService,
    build_query_transactions_tool,
    has_transaction_data,
)


def _workbook_with_transactions(tmp_path: Path) -> Path:
    workbook_path = tmp_path / "transactions.xlsx"
    workbook = Workbook()
    data_input = workbook.active
    data_input.title = "Data Input"
    transactions = workbook.create_sheet("Transactions Raw")
    transactions.append(
        [
            "account",
            "date_posted",
            "amount",
            "merchant",
            "description",
            "default category",
        ]
    )
    transactions.append(
        [
            "Amex Joint",
            "2024-05-02",
            -95.70,
            "Preschool Smiles",
            "BT*PRESCHOOL SMILES EDEN PRAIRIE        MN",
            "Aftercare/Childcare/Tuition",
        ]
    )
    transactions.append(
        [
            "Amex Joint",
            "2024-10-19",
            -136.0,
            "TAP Air Portugal",
            "TAP AIR PORTUGAL    LISBOA              PR",
            "Airfare/Transportation",
        ]
    )
    workbook.save(workbook_path)
    return workbook_path


def test_transaction_query_service_reads_transactions_as_typed_rows(tmp_path: Path) -> None:
    workbook_path = _workbook_with_transactions(tmp_path)
    service = TransactionQueryService(workbook_path=workbook_path)

    result = service.query(
        """
        SELECT row_number, account, amount, default_category
        FROM transactions_raw
        ORDER BY row_number
        """
    )

    assert result["status"] == "ok"
    assert result["row_count"] == 2
    assert result["rows"][0]["row_number"] == 2
    assert result["rows"][0]["account"] == "Amex Joint"
    assert result["rows"][0]["amount"] == -95.7
    assert result["rows"][1]["default_category"] == "Airfare/Transportation"


def test_transaction_query_service_exposes_row_and_column_provenance(tmp_path: Path) -> None:
    workbook_path = _workbook_with_transactions(tmp_path)
    service = TransactionQueryService(workbook_path=workbook_path)

    result = service.query(
        """
        SELECT row_number, column_letter, column_name, value_text
        FROM transactions_raw_cells
        WHERE row_number = 2
        ORDER BY column_letter
        """
    )

    assert result["status"] == "ok"
    assert result["row_count"] == 6
    assert result["rows"][0] == {
        "row_number": 2,
        "column_letter": "A",
        "column_name": "account",
        "value_text": "Amex Joint",
    }


def test_query_transactions_tool_returns_error_payload_for_non_read_only_sql(tmp_path: Path) -> None:
    workbook_path = _workbook_with_transactions(tmp_path)
    tool = build_query_transactions_tool(workbook_path)

    payload = json.loads(tool.invoke({"sql": "DELETE FROM transactions_raw"}))

    assert payload["status"] == "error"
    assert "read-only" in payload["error"]
    assert has_transaction_data(workbook_path) is True
