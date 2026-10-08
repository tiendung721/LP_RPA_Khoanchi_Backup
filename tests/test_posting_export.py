from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.database import Database
from app.repositories.excel_run_repository import ExcelRunRepository
from app.repositories.expense_posting_repository import ExpensePostingRepository
from app.services.excel.models import (
    ExcelRunStatus,
    PostingItemStatus,
    PostingPlan,
    PostingResult,
    ResolutionAction,
    WorkbookFingerprint,
)
from app.services.excel.posting_export import (
    HEADERS,
    PostingExportService,
    build_posting_export_snapshot,
)


def _source(index: int, *, document: str, invoice: str, amount: int) -> dict:
    return {
        "source_item_index": index,
        "source_document_id": document,
        "source_document_name": f"{document}.pdf",
        "invoice_no": invoice,
        "invoice_date": "2026-10-08",
        "invoice_container_count": 2 if document == "A" else 1,
        "container_count_basis": "EXPLICIT",
        "container": f"CONT{index}",
        "bl": "BL-01",
        "input_sqt": None,
        "vessel_voyage_raw": "PROSPER 2625S",
        "vessel_name": "PROSPER",
        "voyage_no": "2625S",
        "fee": "VTN",
        "rule": "ST",
        "carrier": "VT ABC",
        "amount": amount,
    }


def test_posting_export_retains_all_invoice_rows_and_counts_grouped_write_once(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    runs = ExcelRunRepository(database)
    postings = ExpensePostingRepository(database)
    run = runs.create_run(operation="EXPENSE_POSTING")
    bk = tmp_path / "BK.xlsx"
    plan = PostingPlan(
        batch_id=1,
        batch_path=tmp_path / "ready.json",
        batch_hash="hash",
        target_path=bk,
        target_fingerprint=WorkbookFingerprint(0, 0, "hash"),
        items=[],
        conflicts=[],
        sheet_candidates=[],
        run_id=run.id,
        source_rows=[
            _source(0, document="A", invoice="INV-A", amount=100),
            _source(1, document="B", invoice="=1+1", amount=400),
            _source(2, document="A", invoice="INV-A", amount=200),
        ],
    )
    actions = [
        {
            "source_indices": [0, 2], "status": PostingItemStatus.POSTED,
            "action": ResolutionAction.OVERWRITE, "amount_write": True,
            "invoice_write": True, "carrier_write": False,
            "fee_selected": "VTN", "value_after": 300,
            "invoice_selected": "INV-A", "invoice_value_after": "INV-A",
            "carrier_effective": "VT ABC", "source_sqt": 701,
            "sheet_name": "T07 26", "target_row": 2, "target_cell": "L2",
        },
        {
            "source_indices": [1], "status": PostingItemStatus.USER_SKIPPED,
            "action": ResolutionAction.SKIP, "amount_write": False,
            "invoice_write": False, "carrier_write": False,
            "fee_selected": "VTN", "value_after": None,
            "sheet_name": "T07 26", "target_row": None, "target_cell": None,
        },
    ]
    result = PostingResult(
        status=ExcelRunStatus.SUCCEEDED, target_path=bk,
        target_sheets=("T07 26",), run_id=run.id,
    )
    snapshot = build_posting_export_snapshot(plan, actions, result)
    assert snapshot["invoice_count"] == 2
    assert snapshot["row_count"] == 3
    assert snapshot["written_total"] == "300"
    assert [row["amount_written_this_run"] for row in snapshot["rows"]] == [300, None, None]
    assert [row["status"] for row in snapshot["rows"]] == ["Đã ghi tiền", "Đã ghi tiền", "Bỏ qua"]
    assert [row["source_row"] for row in snapshot["rows"]] == [1, 3, 2]
    postings.save_export_snapshot(run.id, snapshot)
    runs.finish_run(run.id, status="SUCCEEDED")

    # The report remains available from persisted data in a fresh repository.
    reloaded = ExpensePostingRepository(Database(tmp_path / "state.db"))
    exporter = PostingExportService(ExcelRunRepository(reloaded.database), reloaded)
    preview = exporter.latest_preview()
    assert preview is not None
    assert preview.invoice_count == 2
    target = tmp_path / preview.filename
    exported = exporter.export(run.id, target)
    assert exported.written_total == Decimal(300)
    assert exporter.latest_preview().last_export_path == target
    original_bytes = target.read_bytes()
    with pytest.raises(FileExistsError):
        exporter.export(run.id, target)
    assert target.read_bytes() == original_bytes

    book = load_workbook(target, data_only=False)
    try:
        assert book.sheetnames == ["Chi phí nhập BK"]
        sheet = book.active
        assert tuple(cell.value for cell in sheet[7]) == HEADERS
        assert sheet.max_row == 10
        assert [sheet.cell(row, 2).value for row in range(8, 11)] == ["A.pdf", "A.pdf", "B.pdf"]
        assert [sheet.cell(row, 22).value for row in range(8, 11)] == [300, None, None]
        assert sheet.cell(10, 5).value == "=1+1"
        assert sheet.cell(10, 5).data_type == "s"
        assert sheet.cell(8, 6).number_format == "dd/mm/yyyy"
        assert sheet.cell(10, 30).value == "Bỏ qua"
    finally:
        book.close()


def test_posting_export_add_uses_adjustment_not_final_bk_cell(tmp_path: Path) -> None:
    plan = PostingPlan(
        batch_id=None, batch_path=tmp_path / "ready.json", batch_hash="hash",
        target_path=tmp_path / "BK.xlsx",
        target_fingerprint=WorkbookFingerprint(0, 0, "hash"),
        items=[], conflicts=[], sheet_candidates=[],
        source_rows=[_source(0, document="A", invoice="INV-A", amount=-200)],
    )
    result = PostingResult(status=ExcelRunStatus.SUCCEEDED, target_path=plan.target_path)
    snapshot = build_posting_export_snapshot(
        plan,
        [{
            "source_indices": [0], "status": PostingItemStatus.POSTED,
            "action": ResolutionAction.ADD, "amount_write": True,
            "value_before": 1000, "value_after": 800,
            "sheet_name": "T07 26", "target_row": 2, "target_cell": "L2",
        }],
        result,
    )
    assert snapshot["rows"][0]["value_after"] == 800
    assert snapshot["rows"][0]["amount_written_this_run"] == -200
    assert snapshot["written_total"] == "-200"
