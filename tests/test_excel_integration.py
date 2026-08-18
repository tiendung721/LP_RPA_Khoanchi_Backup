from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Iterable

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from app.database import Database
from app.repositories.expense_posting_repository import ExpensePostingRepository
from app.services.excel.daily_sync import (
    SOURCE_HEADER_ALIASES,
    DailySyncError,
    DailySyncService,
    parse_sqt,
)
from app.services.excel.models import (
    ConflictType,
    ExcelRunStatus,
    PostingItemStatus,
    ResolutionAction,
    TargetCellKind,
)
from app.services.excel.posting import (
    FEE_HEADER_ALIASES,
    ExpensePostingError,
    ExpensePostingService,
)
from app.services.excel.review import CorrectionRequiredError, SourceDataChangedError
from app.services.excel.workbook import WorkbookChangedError


SYNC_HEADERS = [
    aliases[0] for aliases in SOURCE_HEADER_ALIASES.values()
]

POSTING_BASE_HEADERS = (
    "SQT PM",
    "Ngày Đóng",
    "Số Container",
    "Loại hàng",
    "Tên tàu",
    "Người nhận",
)
POSTING_FEE_COLUMNS = {
    fee: column
    for column, fee in enumerate(FEE_HEADER_ALIASES, len(POSTING_BASE_HEADERS) + 1)
}


class _ImmediateStabilityChecker:
    def wait(self, _path: str | Path) -> None:
        return None


class _ReadyProvider:
    def __init__(self, path: Path) -> None:
        self.path = path

    def get_latest_ready_json_path(self) -> Path:
        return self.path

    def get_ready_json_path(self, _batch_id: int) -> Path:
        return self.path


class _PostedIndexRepository:
    def __init__(self, indices: set[int]) -> None:
        self.indices = indices

    def successful_source_indices(self, _batch_hash: str) -> set[int]:
        return set(self.indices)


class _RunHistoryRepository:
    def __init__(self) -> None:
        self.created: list[dict[str, Any]] = []
        self.updated: list[dict[str, Any]] = []
        self.finished: list[dict[str, Any]] = []

    def create_run(self, **values: Any) -> Any:
        self.created.append(dict(values))
        return SimpleNamespace(id=len(self.created))

    def update_run(self, run_id: int, **values: Any) -> None:
        self.updated.append({"run_id": run_id, **values})

    def finish_run(self, run_id: int, **values: Any) -> None:
        self.finished.append({"run_id": run_id, **values})

    def get_latest_sync_sheet(self) -> None:
        return None


class _PostingHistoryRepository:
    def __init__(self) -> None:
        self.items: list[dict[str, Any]] = []
        self.metadata: dict[str, Any] = {}

    def successful_source_indices(self, _batch_hash: str) -> set[int]:
        return set()

    def create_items(self, items: Iterable[dict[str, Any]], **values: Any) -> list[Any]:
        self.items.extend(dict(item) for item in items)
        self.metadata = dict(values)
        return []


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _sync_row(
    sqt: Any,
    container: str = "DRYU3026167",
    *,
    closing_date: str = "2026-07-28",
    weight: int = 20,
    cargo_type: str = "Gạo",
    closing_place: str = "Kho A",
    transport: str = "Xe A",
) -> list[Any]:
    return [
        sqt,
        closing_date,
        container,
        weight,
        cargo_type,
        closing_place,
        "Tàu A",
        "2026-07-30",
        "2026-08-02",
        "Công ty B",
        "VTB",
        transport,
    ]


def _save_daily(
    path: Path,
    rows: Iterable[Iterable[Any]],
    *,
    month: int = 7,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = f"Tháng {month}"
    sheet.append(SYNC_HEADERS)
    for row in rows:
        sheet.append(list(row))
    workbook.save(path)
    workbook.close()


def _populate_target_sheet(
    sheet: Any,
    rows: Iterable[Iterable[Any]],
) -> None:
    for column, header in enumerate(SYNC_HEADERS[:11], 1):
        sheet.cell(1, column).value = header
    sheet["L1"] = "Chi phí L"
    sheet["M1"] = "Chi phí M"
    sheet["N1"] = "Chi phí N"
    sheet["O1"] = "Chi phí O"
    sheet["P1"] = SYNC_HEADERS[11]
    sheet["Q1"] = "Hóa đơn"
    for row_number, source_row in enumerate(rows, 2):
        source_values = list(source_row)
        for column, value in enumerate(source_values[:11], 1):
            sheet.cell(row_number, column).value = value
        sheet.cell(row_number, 16).value = source_values[11]


def _save_target(
    path: Path,
    rows: Iterable[Iterable[Any]],
    *,
    month: int,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = f"T{month:02d} 26"
    _populate_target_sheet(sheet, rows)
    workbook.save(path)
    workbook.close()


def _save_ready(path: Path, rows: Iterable[Iterable[Any]]) -> None:
    normalized_rows = []
    for row in rows:
        values = list(row)
        if len(values) == 5:
            values = [*values[:4], None, None, values[4]]
        normalized_rows.append(values)
    path.write_text(
        json.dumps({"v": 1, "d": normalized_rows}, ensure_ascii=False),
        encoding="utf-8",
    )


def _add_posting_row(
    sheet: Any,
    row: int,
    container: str,
    *,
    sqt: Any = 700,
    closing_date: Any = "2026-07-28",
    cargo_type: Any = "Gạo",
) -> None:
    sheet.cell(row, 1).value = sqt
    sheet.cell(row, 2).value = closing_date
    sheet.cell(row, 3).value = container
    sheet.cell(row, 4).value = cargo_type
    sheet.cell(row, 5).value = "Tàu A"
    sheet.cell(row, 6).value = "Công ty B"


def _new_posting_sheet(workbook: Workbook, name: str) -> Any:
    sheet = workbook.active
    if sheet.max_row == 1 and sheet.max_column == 1 and sheet["A1"].value is None:
        sheet.title = name
    else:
        sheet = workbook.create_sheet(name)
    for column, title in enumerate(POSTING_BASE_HEADERS, 1):
        sheet.cell(1, column).value = title
    for fee, column in POSTING_FEE_COLUMNS.items():
        sheet.cell(1, column).value = FEE_HEADER_ALIASES[fee][0]
    sheet.cell(1, len(POSTING_BASE_HEADERS) + len(POSTING_FEE_COLUMNS) + 1).value = (
        "GHI CHÚ"
    )
    sheet.cell(1, len(POSTING_BASE_HEADERS) + len(POSTING_FEE_COLUMNS) + 2).value = (
        "Hóa đơn"
    )
    return sheet


POSTING_INVOICE_LAYOUT: dict[str, tuple[int, int | None]] = {
    "CB": (7, 8),
    "CBDH": (9, 10),
    "VTN": (11, 12),
    "NV": (13, 14),
    "HH": (15, 16),
    "NH": (17, 18),
    "HV": (19, 20),
    "VSDL": (21, 22),
    "LC": (23, 24),
    "QT": (25, 26),
    "LL": (27, None),
    "SC": (28, 29),
}

FULL_POSTING_LAYOUT: dict[str, tuple[int, int | None]] = {
    "CB": (13, 14),
    "CBDH": (17, 18),
    "NV": (19, 20),
    "HH": (21, 22),
    "VTN": (24, 25),
    "NH": (26, 27),
    "HV": (28, 29),
    "VSDL": (30, 31),
    "LC": (32, 33),
    "QT": (34, 35),
    "LL": (36, None),
    "SC": (37, 38),
}


def _new_full_posting_sheet(workbook: Workbook, name: str) -> Any:
    sheet = workbook.active
    if sheet.max_row == 1 and sheet.max_column == 1 and sheet["A1"].value is None:
        sheet.title = name
    else:
        sheet = workbook.create_sheet(name)
    for column, header in enumerate(SYNC_HEADERS[:11], 1):
        sheet.cell(1, column).value = header
    sheet.cell(1, 16).value = SYNC_HEADERS[11]
    for fee, (amount_column, invoice_column) in FULL_POSTING_LAYOUT.items():
        sheet.cell(1, amount_column).value = FEE_HEADER_ALIASES[fee][0]
        if invoice_column is not None:
            sheet.cell(1, invoice_column).value = "Số HĐ"
    sheet.cell(1, 39).value = "GHI CHÚ"
    return sheet


def _add_full_plan_row(
    sheet: Any,
    row: int,
    *,
    sqt: int,
    container: str,
    closing_date: str,
) -> list[Any]:
    values = _sync_row(sqt, container, closing_date=closing_date)
    for column, value in enumerate(values[:11], 1):
        sheet.cell(row, column).value = value
    sheet.cell(row, 16).value = values[11]
    return values


def _new_posting_sheet_with_invoices(workbook: Workbook, name: str) -> Any:
    sheet = workbook.active
    sheet.title = name
    for column, title in enumerate(POSTING_BASE_HEADERS, 1):
        sheet.cell(1, column).value = title
    for fee, (amount_column, invoice_column) in POSTING_INVOICE_LAYOUT.items():
        sheet.cell(1, amount_column).value = FEE_HEADER_ALIASES[fee][0]
        if invoice_column is not None:
            sheet.cell(1, invoice_column).value = (
                "Hóa đơn cước biển" if fee == "CB" else "Số HĐ"
            )
    sheet.cell(1, 30).value = "GHI CHÚ"
    return sheet


def test_posting_reader_accepts_v1_review_fields_without_using_them_for_amount(
    tmp_path: Path,
) -> None:
    service = ExpensePostingService(provider=object(), bk_path=tmp_path / "BK.xlsx")
    raw = json.dumps(
        {
            "v": 1,
            "d": [
                [
                    "GAOU2619968",
                    None,
                    "CBDH",
                    "CV",
                    "000130/HD",
                    "Vận tải ABC",
                    2_484_000,
                ]
            ],
        },
        ensure_ascii=False,
    ).encode("utf-8")

    parsed = service._validate_json(raw)
    assert parsed[0]["source_document_id"].startswith("LEGACY_BATCH_")
    assert parsed == [
        {
            "source_item_index": 0,
            "source_document_id": parsed[0]["source_document_id"],
            "source_document_name": "Dữ liệu bóc tách cũ",
            "container": "GAOU2619968",
            "bl": None,
            "fee": "CBDH",
            "rule": "CV",
            "invoice_no": "000130/HD",
            "invoice_date": None,
            "carrier": "Vận tải ABC",
            "amount": 2_484_000,
        }
    ]


def test_bang_ke_posts_across_sheets_migrates_columns_and_adds_negative(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ket_qua_boc_tach_bang_ke_fixture.json"
    target = tmp_path / "BK.xlsx"

    def source_row(
        container: str,
        fee: str,
        amount: int,
        vessel: str,
        voyage: str,
    ) -> dict[str, Any]:
        return {
            "container": container,
            "bl": "BL-01",
            "vessel_voyage_raw": f"{vessel} V.{voyage}",
            "vessel_name": vessel,
            "voyage_no": voyage,
            "invoice_container_count": None,
            "container_count_basis": "UNKNOWN",
            "fee": fee,
            "rule": "GV",
            "invoice_no": None,
            "invoice_date": None,
            "carrier": None,
            "amount": amount,
        }

    ready.write_text(
        json.dumps(
            {
                "v": 2,
                "d": [
                    source_row("DRYU3045911", "VSDL", -282_000, "PRIME", "2606S"),
                    source_row("DRYU3045911", "NH", -409_300, "PRIME", "2606S"),
                    source_row("DRYU3045911", "HV", -560_000, "PRIME", "2606S"),
                    source_row("DRYU3045911", "VAT", 80_000, "PRIME", "2606S"),
                    source_row("MSCU1234567", "GH", 150_000, "PROSPER", "2625S"),
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    workbook = Workbook()
    january = _new_full_posting_sheet(workbook, "T01 26")
    _add_full_plan_row(
        january,
        2,
        sqt=101,
        container="DRYU3045911",
        closing_date="2026-01-10",
    )
    january.cell(2, 7).value = "PRIME\nVoyage 2606S"
    january.cell(2, FULL_POSTING_LAYOUT["VSDL"][0]).value = 1_000_000
    january.cell(2, FULL_POSTING_LAYOUT["NH"][0]).value = 1_000_000
    january.cell(2, FULL_POSTING_LAYOUT["HV"][0]).value = 1_000_000
    april = _new_full_posting_sheet(workbook, "T04 26")
    _add_full_plan_row(
        april,
        2,
        sqt=402,
        container="MSCU1234567",
        closing_date="2026-04-10",
    )
    april.cell(2, 7).value = "PROSPER - Chuyến số 2625S"
    workbook.save(target)
    workbook.close()

    class BangKeProvider(_ReadyProvider):
        repository = SimpleNamespace(
            get_by_id=lambda _batch_id: SimpleNamespace(source_kind="BANG_KE")
        )

    service = ExpensePostingService(
        BangKeProvider(ready),
        bk_path=target,
        temp_dir=tmp_path / "Temp",
        backup_dir=tmp_path / "Backup",
    )
    plan = service.analyze(batch_id=1)

    assert plan.source_kind == "BANG_KE"
    assert plan.selected_sheet is None
    assert plan.target_sheets == {"T01 26", "T04 26"}
    negative_conflicts = [
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.NEGATIVE_ADJUSTMENT
    ]
    assert {conflict.fee: conflict.amount for conflict in negative_conflicts} == {
        "VSDL": -282_000,
        "NH": -409_300,
        "HV": -560_000,
    }
    assert {
        conflict.fee: conflict.details["value_after"]
        for conflict in negative_conflicts
    } == {"VSDL": 718_000, "NH": 590_700, "HV": 440_000}

    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {"action": ResolutionAction.ADD}
            for conflict in negative_conflicts
        },
    )
    assert not refined.conflicts, [
        (conflict.fee, conflict.sheet_name, conflict.message)
        for conflict in refined.conflicts
    ]
    result = service.apply(refined, {})

    assert result.target_sheets == ("T01 26", "T04 26")
    assert result.backup_path is not None
    workbook = load_workbook(target, data_only=False)
    try:
        for sheet_name in ("T01 26", "T04 26"):
            sheet = workbook[sheet_name]
            assert sheet.cell(1, 39).value == "THUẾ GTGT"
            assert sheet.cell(1, 40).value == "GIA HẠN"
        assert workbook["T01 26"].cell(2, 30).value == 718_000
        assert workbook["T01 26"].cell(2, 26).value == 590_700
        assert workbook["T01 26"].cell(2, 28).value == 440_000
        assert workbook["T01 26"].cell(2, 39).value == 80_000
        assert workbook["T04 26"].cell(2, 40).value == 150_000
    finally:
        workbook.close()


def _posting_service(
    ready: Path,
    target: Path,
    runtime_dir: Path,
    *,
    posting_repository: Any | None = None,
    run_repository: Any | None = None,
    clock: Any | None = None,
) -> ExpensePostingService:
    return ExpensePostingService(
        _ReadyProvider(ready),
        bk_path=target,
        temp_dir=runtime_dir / "Temp",
        backup_dir=runtime_dir / "Backup",
        posting_repository=posting_repository,
        run_repository=run_repository,
        clock=clock,
    )


def _sync_service(
    daily: Path,
    target: Path,
    runtime_dir: Path,
    *,
    run_repository: Any | None = None,
) -> DailySyncService:
    return DailySyncService(
        daily_path=daily,
        bk_path=target,
        temp_dir=runtime_dir / "Temp",
        backup_dir=runtime_dir / "Backup",
        stability_checker=_ImmediateStabilityChecker(),
        run_repository=run_repository,
    )


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (700, 700),
        (700.0, 700),
        (" 7 0 0 ", 700),
        (True, None),
        (0, None),
        (-1, None),
        (700.5, None),
        ("700A", None),
        ("", None),
        (None, None),
    ],
)
def test_parse_sqt_accepts_only_positive_integers(
    raw: Any,
    expected: int | None,
) -> None:
    assert parse_sqt(raw) == expected


def test_daily_sync_compares_full_sheet_preserves_order_and_is_idempotent(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    old_row = _sync_row(699, "OLDU0000001")
    duplicate = _sync_row(701, "NEWC0000003")
    _save_daily(
        daily,
        [
            old_row,
            _sync_row(700, "NEWA0000001"),
            _sync_row(700, "NEWB0000002"),
            duplicate,
            duplicate,
            _sync_row("not-an-sqt", "BADS0000000"),
        ],
    )
    _save_target(target, [old_row], month=7)
    workbook = load_workbook(target)
    sheet = workbook["T07 26"]
    sheet["L2"] = 125_000
    sheet["M2"] = "=1+1"
    sheet["N2"] = "do not touch"
    sheet["O2"] = 0
    sheet["Q2"] = "INV-001"
    workbook.save(target)
    workbook.close()
    source_before = _sha256(daily)
    target_before = _sha256(target)

    service = _sync_service(daily, target, runtime_dir)
    plan = service.analyze(source_sheet_name="Tháng 7")

    assert plan.selected_month == 7
    assert [row.sqt for row in plan.rows] == [700, 700, 701, 701]
    assert plan.conflicts == []
    assert plan.insert_count == 4
    assert plan.update_count == 0
    assert plan.unchanged_count == 1
    assert plan.invalid_count == 1
    assert plan.requires_user_input

    result = service.apply(plan, {})

    assert result.status is ExcelRunStatus.SUCCEEDED
    assert result.added_rows == 4
    assert result.inserted_rows == 4
    assert result.updated_rows == 0
    assert result.skipped_rows == 1
    assert result.backup_path is not None
    assert result.backup_path.is_file()
    assert result.backup_path.name == "BK 2026_latest.xlsx"
    assert _sha256(result.backup_path) == target_before
    assert _sha256(daily) == source_before
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        assert [sheet.cell(row, 1).value for row in range(2, 7)] == [
            699,
            700,
            700,
            701,
            701,
        ]
        assert [sheet.cell(row, 3).value for row in range(3, 7)] == [
            "NEWA0000001",
            "NEWB0000002",
            "NEWC0000003",
            "NEWC0000003",
        ]
        assert sheet["F3"].value == "Kho A"
        assert sheet["P3"].value == "Xe A"
        assert [sheet.cell(2, column).value for column in range(12, 16)] == [
            125_000,
            "=1+1",
            "do not touch",
            0,
        ]
        assert sheet["Q2"].value == "INV-001"
    finally:
        workbook.close()

    target_after_first_apply = _sha256(target)
    backup_count = len(list((runtime_dir / "Backup").iterdir()))
    second_plan = service.analyze()
    second_result = service.apply(second_plan, {})

    assert not second_plan.has_changes
    assert second_result.status is ExcelRunStatus.NO_CHANGES
    assert second_result.backup_path is None
    assert second_plan.invalid_count == 1
    assert _sha256(target) == target_after_first_apply
    assert len(list((runtime_dir / "Backup").iterdir())) == backup_count


def test_daily_sync_updates_existing_rows_and_preserves_bk_only_columns(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    updated = _sync_row(
        699,
        "NEWU0000001",
        weight=28,
        closing_place="Kho mới",
        transport="Xe mới",
    )
    inserted = _sync_row(700, "NEWA0000002")
    _save_daily(daily, [updated, inserted])
    _save_target(
        target,
        [
            _sync_row(698, "TARGET000001"),
            _sync_row(699, "OLDU0000001"),
        ],
        month=7,
    )
    workbook = load_workbook(target)
    sheet = workbook["T07 26"]
    sheet["L3"] = 125_000
    sheet["M3"] = "=1+1"
    sheet["N3"] = "không được đổi"
    sheet["O3"] = 0
    sheet["Q3"] = "INV-001"
    workbook.save(target)
    workbook.close()

    service = _sync_service(daily, target, runtime_dir)
    plan = service.analyze(source_sheet_name="Tháng 7")

    assert plan.update_count == 1
    assert plan.insert_count == 1
    assert plan.unchanged_count == 0
    assert plan.target_only_count == 1

    result = service.apply(plan, {})

    assert result.updated_rows == 1
    assert result.inserted_rows == 1
    assert result.target_only_rows == 1
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        assert sheet["A2"].value == 698
        assert sheet["C2"].value == "TARGET000001"
        assert sheet["C3"].value == "NEWU0000001"
        assert sheet["D3"].value == 28
        assert sheet["F3"].value == "Kho mới"
        assert sheet["P3"].value == "Xe mới"
        assert [sheet.cell(3, column).value for column in range(12, 16)] == [
            125_000,
            "=1+1",
            "không được đổi",
            0,
        ]
        assert sheet["Q3"].value == "INV-001"
        assert sheet["A4"].value == 700
        assert sheet["C4"].value == "NEWA0000002"
    finally:
        workbook.close()


def test_daily_sync_blocks_mismatched_duplicate_sqt_without_backup(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_daily(
        daily,
        [
            _sync_row(700, "NEWA0000001"),
            _sync_row(700, "NEWB0000002"),
        ],
    )
    _save_target(target, [_sync_row(700, "NEWA0000001")], month=7)
    service = _sync_service(daily, target, runtime_dir)

    plan = service.analyze(source_sheet_name="Tháng 7")

    assert plan.conflict_count == 1
    assert plan.conflicts[0].conflict_type is (
        ConflictType.SYNC_GROUP_COUNT_MISMATCH
    )
    with pytest.raises(DailySyncError, match="số dòng"):
        service.apply(plan, {})
    assert not (runtime_dir / "Backup").exists()


def test_daily_sync_aborts_when_source_changes_after_analysis(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_daily(daily, [_sync_row(700)])
    _save_target(target, [_sync_row(699)], month=7)
    service = _sync_service(daily, target, runtime_dir)
    plan = service.analyze(source_sheet_name="Tháng 7")

    _save_daily(daily, [_sync_row(700, "CHANGED00001")])

    with pytest.raises(WorkbookChangedError):
        service.apply(plan, {})
    assert not (runtime_dir / "Backup").exists()


def test_daily_sync_accepts_vt_bo_as_target_transport_header(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_daily(daily, [_sync_row(700)], month=7)
    _save_target(target, [_sync_row(699)], month=7)

    workbook = load_workbook(target)
    try:
        workbook["T07 26"]["P1"] = "VT bộ"
        workbook.save(target)
    finally:
        workbook.close()

    service = _sync_service(daily, target, runtime_dir)
    plan = service.analyze(source_sheet_name="Tháng 7")

    assert plan.selected_sheet == "T07 26"
    assert [row.sqt for row in plan.rows] == [700]
    service.cancel(plan)


def test_daily_sync_lists_month_sheets_and_analyzes_only_selected_source(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày.xlsx"
    target = tmp_path / "BK.xlsx"
    runtime_dir = tmp_path / "Excel"

    workbook = Workbook()
    july = workbook.active
    july.title = "Tháng 7"
    july.append(SYNC_HEADERS)
    july.append(_sync_row(700, "JULY0000001"))
    august = workbook.create_sheet("Tháng 8")
    august.append(["Header sai"])
    august.append([800])
    workbook.create_sheet("Ghi chú")
    workbook.save(daily)
    workbook.close()

    workbook = Workbook()
    july_target = workbook.active
    july_target.title = "T07 26"
    _populate_target_sheet(july_target, [_sync_row(699, "OLDU0000001")])
    august_target = workbook.create_sheet("T08 26")
    _populate_target_sheet(august_target, [_sync_row(799, "OLDU0000002")])
    workbook.save(target)
    workbook.close()

    service = _sync_service(daily, target, runtime_dir)

    assert [
        (candidate.month, candidate.sheet_name)
        for candidate in service.source_sheet_candidates()
    ] == [(7, "Tháng 7"), (8, "Tháng 8")]

    progress: list[str] = []
    plan = service.analyze(
        source_sheet_name="Tháng 7",
        progress_callback=progress.append,
    )

    assert plan.selected_month == 7
    assert plan.selected_sheet == "T07 26"
    assert [row.sqt for row in plan.rows] == [700]
    assert [candidate.source_sheet for candidate in plan.month_candidates] == [
        "Tháng 7"
    ]
    assert any("Tháng 7" in message for message in progress)
    assert not any("Tháng 8" in message for message in progress)
    service.cancel(plan)


def test_daily_sync_rejects_source_sheet_not_in_daily_workbook(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày.xlsx"
    target = tmp_path / "BK.xlsx"
    _save_daily(daily, [_sync_row(700)], month=7)
    _save_target(target, [_sync_row(699)], month=7)
    service = _sync_service(daily, target, tmp_path / "Excel")

    with pytest.raises(DailySyncError, match="Không tìm thấy sheet nguồn"):
        service.analyze(source_sheet_name="Tháng 8")


def test_daily_sync_creates_new_month_from_previous_nonempty_template(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK Tổng hợp 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_daily(daily, [_sync_row(700)], month=7)
    _save_target(target, [_sync_row(699, "OLDU0000001")], month=6)
    workbook = load_workbook(target)
    template = workbook["T06 26"]
    template.freeze_panes = "A2"
    template.page_setup.orientation = "landscape"
    template.auto_filter.ref = "A1:Q2"
    template.merge_cells("R1:S1")
    template["R1"] = "Thông tin mẫu"
    template["C2"].fill = PatternFill("solid", fgColor="FFF2CC")
    template["C2"].comment = Comment("old row", "test")
    validation = DataValidation(type="list", formula1='"A,B"')
    template.add_data_validation(validation)
    validation.add("C2:C100")
    template.conditional_formatting.add(
        "A2:A100",
        CellIsRule(operator="greaterThan", formula=["0"]),
    )
    workbook.save(target)
    workbook.close()

    service = _sync_service(daily, target, runtime_dir)
    result = service.apply(service.analyze(), {})

    assert result.status is ExcelRunStatus.SUCCEEDED
    assert result.sheet_name == "T07 26"
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook.sheetnames == ["T06 26", "T07 26"]
        template = workbook["T06 26"]
        created = workbook["T07 26"]
        assert template["A2"].value == 699
        assert template["C2"].value == "OLDU0000001"
        assert created["A2"].value == 700
        assert created["C2"].value == "DRYU3026167"
        assert created["C2"].comment is None
        assert created["C2"].fill.fgColor.rgb == template["C2"].fill.fgColor.rgb
        assert created.freeze_panes == "A2"
        assert created.page_setup.orientation == "landscape"
        assert created.auto_filter.ref == "A1:Q2"
        assert "R1:S1" in {str(item) for item in created.merged_cells.ranges}
        assert created["R1"].value == "Thông tin mẫu"
        assert len(created.data_validations.dataValidation) == 1
        assert len(created.conditional_formatting) == 1
        assert created["L2"].value is None
        assert created["Q2"].value is None
    finally:
        workbook.close()


def test_daily_sync_aborts_before_backup_when_bk_changed_after_analyze(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_daily(daily, [_sync_row(700)])
    _save_target(target, [_sync_row(699)], month=7)
    service = _sync_service(daily, target, runtime_dir)
    plan = service.analyze()

    workbook = load_workbook(target)
    workbook["T07 26"]["L2"] = "external change"
    workbook.save(target)
    workbook.close()
    changed_hash = _sha256(target)

    with pytest.raises(WorkbookChangedError):
        service.apply(plan, {})

    assert _sha256(target) == changed_hash
    assert not (runtime_dir / "Backup").exists()


def test_posting_carries_previous_month_plan_once_and_reuses_target_row(
    tmp_path: Path,
) -> None:
    first_ready = tmp_path / "first.json"
    second_ready = tmp_path / "second.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(first_ready, [[container, None, "VTN", "ST", "INV-1", None, 100]])
    _save_ready(second_ready, [[container, None, "HH", "ST", "INV-2", None, 250]])

    workbook = Workbook()
    july = _new_full_posting_sheet(workbook, "T07 26")
    _add_full_plan_row(
        july,
        2,
        sqt=708,
        container="MSCU1234567",
        closing_date="2026-07-10",
    )
    june = _new_full_posting_sheet(workbook, "T06 26")
    source_values = _add_full_plan_row(
        june,
        2,
        sqt=619,
        container=container,
        closing_date="2026-06-10",
    )
    june.cell(2, FULL_POSTING_LAYOUT["CB"][0]).value = 999
    april = _new_full_posting_sheet(workbook, "T04 26")
    _add_full_plan_row(
        april,
        2,
        sqt=401,
        container=container,
        closing_date="2026-04-10",
    )
    workbook.save(target)
    workbook.close()

    database = Database(tmp_path / "app_state.db")
    postings = ExpensePostingRepository(database)
    first_service = _posting_service(
        first_ready,
        target,
        tmp_path / "First",
        posting_repository=postings,
    )
    first_plan = first_service.analyze(sheet_name="T07 26")

    assert not first_plan.conflicts
    assert first_plan.items[0].selected_source_sheet == "T06 26"
    assert first_plan.items[0].carry_forward_required
    assert first_plan.items[0].target_row == 3
    first_result = first_service.apply(first_plan, {})
    assert first_result.backup_path is not None

    workbook = load_workbook(target, data_only=False)
    try:
        july = workbook["T07 26"]
        copied = [july.cell(3, column).value for column in range(1, 12)]
        assert copied == source_values[:11]
        assert july.cell(3, 16).value == source_values[11]
        assert july.cell(3, FULL_POSTING_LAYOUT["CB"][0]).value is None
        assert july.cell(3, FULL_POSTING_LAYOUT["VTN"][0]).value == 100
    finally:
        workbook.close()

    second_service = _posting_service(
        second_ready,
        target,
        tmp_path / "Second",
        posting_repository=postings,
    )
    second_plan = second_service.analyze(sheet_name="T07 26")
    assert not second_plan.conflicts
    assert second_plan.items[0].selected_source_sheet == "T06 26"
    assert not second_plan.items[0].carry_forward_required
    assert second_plan.items[0].target_row == 3
    second_service.apply(second_plan, {})

    workbook = load_workbook(target, data_only=False)
    try:
        july = workbook["T07 26"]
        assert july.cell(3, FULL_POSTING_LAYOUT["VTN"][0]).value == 100
        assert july.cell(3, FULL_POSTING_LAYOUT["HH"][0]).value == 250
        assert july.cell(4, 1).value is None
    finally:
        workbook.close()
        database.close()


def test_posting_allocates_two_documents_to_two_months_in_one_atomic_apply(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready-v3.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime = tmp_path / "Excel"
    rows: list[dict[str, Any]] = []
    july_containers = [f"JULY{i:07d}" for i in range(1, 6)]
    june_containers = [f"JUNE{i:07d}" for i in range(1, 11)]
    for index, container in enumerate(july_containers, start=1):
        rows.append(
            {
                "source_document_id": "DOC_001",
                "source_document_name": "Hoa_don_A.pdf",
                "container": container,
                "bl": None,
                "vessel_voyage_raw": None,
                "vessel_name": None,
                "voyage_no": None,
                "invoice_container_count": None,
                "container_count_basis": "UNKNOWN",
                "fee": "VTN",
                "rule": "ST",
                "invoice_no": "INV-JULY",
                "invoice_date": "2026-07-15",
                "carrier": None,
                "amount": 1000 + index,
            }
        )
    for index, container in enumerate(june_containers, start=1):
        rows.append(
            {
                "source_document_id": "DOC_002",
                "source_document_name": "Hoa_don_B.pdf",
                "container": container,
                "bl": None,
                "vessel_voyage_raw": None,
                "vessel_name": None,
                "voyage_no": None,
                "invoice_container_count": None,
                "container_count_basis": "UNKNOWN",
                "fee": "HH",
                "rule": "ST",
                "invoice_no": "INV-JUNE",
                "invoice_date": "2026-06-15",
                "carrier": None,
                "amount": 2000 + index,
            }
        )
    ready.write_text(
        json.dumps({"v": 3, "d": rows}, ensure_ascii=False), encoding="utf-8"
    )
    workbook = Workbook()
    july = _new_full_posting_sheet(workbook, "T07 26")
    for row_number, container in enumerate(july_containers, start=2):
        _add_full_plan_row(
            july, row_number, sqt=700 + row_number, container=container,
            closing_date="2026-07-10",
        )
    june = _new_full_posting_sheet(workbook, "T06 26")
    for row_number, container in enumerate(june_containers, start=2):
        _add_full_plan_row(
            june, row_number, sqt=600 + row_number, container=container,
            closing_date="2026-06-10",
        )
    workbook.save(target)
    workbook.close()
    service = _posting_service(ready, target, runtime)

    unassigned = service.analyze()
    assert len(unassigned.source_groups) == 2
    assert {group.suggested_sheet for group in unassigned.source_groups} == {
        "T06 26", "T07 26"
    }
    assert all(group.target_sheet is None for group in unassigned.source_groups)
    assignments = {
        group.group_id: group.suggested_sheet
        for group in unassigned.source_groups
        if group.suggested_sheet is not None
    }
    plan = service.analyze(group_target_sheets=assignments)

    assert not plan.conflicts
    assert plan.target_sheets == {"T06 26", "T07 26"}
    result = service.apply(plan, {})
    assert result.target_sheets == ("T06 26", "T07 26")
    assert result.posted_source_items == 15
    assert len(list((runtime / "Backup").glob("*.xlsx"))) == 1
    workbook = load_workbook(target, data_only=False)
    try:
        assert [
            workbook["T07 26"].cell(row, FULL_POSTING_LAYOUT["VTN"][0]).value
            for row in range(2, 7)
        ] == [1001, 1002, 1003, 1004, 1005]
        assert [
            workbook["T06 26"].cell(row, FULL_POSTING_LAYOUT["HH"][0]).value
            for row in range(2, 12)
        ] == list(range(2001, 2011))
    finally:
        workbook.close()

def test_posting_can_split_one_same_month_document_by_invoice(tmp_path: Path) -> None:
    ready = tmp_path / "same-month.json"
    target = tmp_path / "BK.xlsx"
    base = {
        "source_document_id": "DOC_001",
        "source_document_name": "Hai_hoa_don.pdf",
        "bl": None,
        "vessel_voyage_raw": None,
        "vessel_name": None,
        "voyage_no": None,
        "invoice_container_count": None,
        "container_count_basis": "UNKNOWN",
        "fee": "VTN",
        "rule": "ST",
        "invoice_date": "2026-07-15",
        "carrier": None,
    }
    ready.write_text(
        json.dumps(
            {
                "v": 3,
                "d": [
                    base | {"container": "SPLT0000001", "invoice_no": "A", "amount": 1},
                    base | {"container": "SPLT0000002", "invoice_no": "B", "amount": 2},
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    workbook = Workbook()
    sheet = _new_full_posting_sheet(workbook, "T07 26")
    _add_full_plan_row(sheet, 2, sqt=701, container="SPLT0000001", closing_date="2026-07-01")
    _add_full_plan_row(sheet, 3, sqt=702, container="SPLT0000002", closing_date="2026-07-02")
    workbook.save(target)
    workbook.close()
    service = _posting_service(ready, target, tmp_path / "Runtime")

    grouped = service.analyze()
    assert len(grouped.source_groups) == 1
    assert grouped.source_groups[0].can_split_by_invoice
    split = service.analyze(split_document_ids=["DOC_001"])
    assert len(split.source_groups) == 2
    assert {group.invoice_no for group in split.source_groups} == {"A", "B"}
    assert all(not group.can_split_by_invoice for group in split.source_groups)
    service.cancel(split)


def test_posting_three_month_window_handles_year_boundary(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2027.xlsx"
    container = "DRYU3026167"
    _save_ready(ready, [[container, None, "VTN", "ST", 100]])
    workbook = Workbook()
    january = _new_full_posting_sheet(workbook, "T01 27")
    _add_full_plan_row(
        january,
        2,
        sqt=1001,
        container="MSCU1234567",
        closing_date="2027-01-10",
    )
    _new_full_posting_sheet(workbook, "T12 26")
    november = _new_full_posting_sheet(workbook, "T11 26")
    _add_full_plan_row(
        november,
        2,
        sqt=1101,
        container=container,
        closing_date="2026-11-10",
    )
    october = _new_full_posting_sheet(workbook, "T10 26")
    _add_full_plan_row(
        october,
        2,
        sqt=1001,
        container=container,
        closing_date="2026-10-10",
    )
    workbook.save(target)
    workbook.close()

    plan = _posting_service(ready, target, tmp_path / "Excel").analyze(
        sheet_name="T01 27"
    )

    assert not plan.conflicts
    assert plan.items[0].selected_source_sheet == "T11 26"
    assert plan.items[0].source_sqt == 1101


def test_posting_duplicate_container_across_months_requires_source_sheet(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(ready, [[container, None, "VTN", "ST", 100]])
    workbook = Workbook()
    july = _new_full_posting_sheet(workbook, "T07 26")
    _add_full_plan_row(
        july,
        2,
        sqt=708,
        container=container,
        closing_date="2026-07-10",
    )
    june = _new_full_posting_sheet(workbook, "T06 26")
    _add_full_plan_row(
        june,
        2,
        sqt=619,
        container=container,
        closing_date="2026-06-10",
    )
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_CONTAINER_MATCH
    )
    assert [candidate.source_sheet for candidate in conflict.row_candidates] == [
        "T07 26",
        "T06 26",
    ]
    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_ROW",
                "selected_source_sheet": "T06 26",
                "selected_row": 2,
            }
        },
    )
    assert not refined.conflicts
    assert refined.items[0].selected_source_sheet == "T06 26"
    assert refined.items[0].source_sqt == 619
    assert refined.items[0].carry_forward_required


def test_posting_keeps_normalized_duplicate_expenses_separate(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(
        ready,
        [
            [" dryu-302 6167 ", None, "VTN", "ST", 100],
            ["DRYU3026167", "BL-01", "VTN", "ST", 250],
            [None, "BL-ONLY-01", "CB", "HD", 500],
            [None, "BL-ONLY-02", "CB", "HD", 600],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    unselected = service.analyze()
    assert unselected.selected_sheet is None
    assert [item.target_sheet for item in unselected.sheet_candidates] == ["T07 26"]
    plan = service.analyze(sheet_name="T07 26")

    assert plan.selected_sheet == "T07 26"
    assert len(plan.items) == 4
    assert plan.items[0].container == "DRYU3026167"
    assert plan.items[0].source_indices == [0]
    assert plan.items[0].amount == 100
    assert plan.items[1].source_indices == [1]
    assert plan.items[1].amount == 250
    assert plan.items[2].source_indices == [2]
    assert plan.items[3].source_indices == [3]
    assert [conflict.conflict_type for conflict in plan.conflicts] == [
        ConflictType.BL_ONLY_NO_CONTAINER,
        ConflictType.BL_ONLY_NO_CONTAINER,
        ConflictType.MULTIPLE_EXPENSE_SAME_CELL,
    ]
    same_cell = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL
    )
    refined = service.refine(
        plan,
        {
            same_cell.conflict_id: {
                "action": "SELECT_SOURCE_ITEM",
                "selected_source_item_index": 0,
            }
        },
    )

    assert len(refined.items) == 4
    assert refined.items[0].source_indices == [0]
    assert refined.items[0].amount == 100
    assert refined.items[1].status is PostingItemStatus.USER_SKIPPED
    assert [
        conflict.conflict_type for conflict in refined.conflicts
    ] == [
        ConflictType.BL_ONLY_NO_CONTAINER,
        ConflictType.BL_ONLY_NO_CONTAINER,
    ]

    result = service.apply(refined, {})

    assert result.status is ExcelRunStatus.SUCCEEDED
    assert result.posted_source_items == 1
    assert result.written_cells == 1
    assert result.skipped_source_items == 3
    assert result.backup_path is not None
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook["T07 26"].cell(
            2, POSTING_FEE_COLUMNS["VTN"]
        ).value == 100
    finally:
        workbook.close()


def test_posting_keeps_repeated_container_separate_for_different_sqt_choices(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [
            [container, "BL-01", "VTN", "ST", 100],
            [container, "BL-02", "VTN", "ST", 250],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, container, sqt=701)
    _add_posting_row(sheet, 3, container, sqt=702)
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")
    conflicts = [
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_CONTAINER_MATCH
    ]

    assert len(plan.items) == 2
    assert [item.source_indices for item in plan.items] == [[0], [1]]
    assert [conflict.amount for conflict in conflicts] == [100, 250]
    assert len(conflicts) == 2

    refined = service.refine(
        plan,
        {
            conflicts[0].conflict_id: {
                "action": "SELECT_ROW",
                "selected_row": 2,
            },
            conflicts[1].conflict_id: {
                "action": "SELECT_ROW",
                "selected_row": 3,
            },
        },
    )

    assert not refined.conflicts
    assert [(item.target_row, item.amount) for item in refined.items] == [
        (2, 100),
        (3, 250),
    ]

    result = service.apply(refined, {})

    assert result.posted_source_items == 2
    assert result.written_cells == 2
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        column = POSTING_FEE_COLUMNS["VTN"]
        assert sheet.cell(2, column).value == 100
        assert sheet.cell(3, column).value == 250
    finally:
        workbook.close()


def test_posting_never_sums_repeated_expenses_for_same_target_cell(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [
            [container, "BL-01", "VTN", "ST", 100],
            [container, "BL-02", "VTN", "ST", 250],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, container, sqt=701)
    _add_posting_row(sheet, 3, container, sqt=702)
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")
    conflicts = [
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_CONTAINER_MATCH
    ]
    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_ROW",
                "selected_row": 2,
            }
            for conflict in conflicts
        },
    )

    same_cell = next(
        conflict
        for conflict in refined.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL
    )
    refined = service.refine(
        refined,
        {
            same_cell.conflict_id: {
                "action": "SELECT_SOURCE_ITEM",
                "selected_source_item_index": 1,
            }
        },
    )
    assert not refined.conflicts
    assert len(refined.items) == 2

    result = service.apply(refined, {})

    assert result.posted_source_items == 1
    assert result.skipped_source_items == 1
    assert result.written_cells == 1
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        column = POSTING_FEE_COLUMNS["VTN"]
        assert sheet.cell(2, column).value == 250
        assert sheet.cell(3, column).value is None
    finally:
        workbook.close()


def test_posting_requires_confirmation_for_repeated_container_with_one_bk_row(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [
            [container, "BL-01", "VTN", "ST", 100],
            [container, "BL-02", "VTN", "ST", 250],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, container, sqt=701)
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL
    )

    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_SOURCE_ITEM",
                "selected_source_item_index": 0,
            },
        },
    )

    assert not refined.conflicts
    result = service.apply(refined, {})

    assert result.posted_source_items == 1
    assert result.skipped_source_items == 1
    assert result.written_cells == 1
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook["T07 26"].cell(
            2, POSTING_FEE_COLUMNS["VTN"]
        ).value == 100
    finally:
        workbook.close()


def test_posting_cell_conflicts_never_offer_add(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    containers = [
        "DRYU3026167",
        "MSCU1234567",
        "OOLU7654321",
        "TGHU2345678",
        "CMAU3456789",
        "TEMU4567890",
    ]
    amounts = [100, 200, 300, 100, 400, 500]
    _save_ready(
        ready,
        [
            [container, None, "CB", "ST", amount]
            for container, amount in zip(containers, amounts, strict=True)
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    for row, container in enumerate(containers, 2):
        _add_posting_row(sheet, row, container, sqt=698 + row)
    cb_column = POSTING_FEE_COLUMNS["CB"]
    sheet.cell(2, cb_column).value = None
    sheet.cell(3, cb_column).value = 0
    sheet.cell(4, cb_column).value = 300
    sheet.cell(5, cb_column).value = 50
    sheet.cell(6, cb_column).value = "=SUM(A1:A2)"
    sheet.cell(7, cb_column).value = "đã nhập tay"
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")
    by_container = {item.container: item for item in plan.items}

    assert by_container[containers[0]].cell_state.kind is TargetCellKind.EMPTY
    assert by_container[containers[1]].cell_state.kind is TargetCellKind.ZERO
    assert by_container[containers[2]].status is PostingItemStatus.ALREADY_EXISTS
    conflicts = {
        conflict.conflict_type: conflict for conflict in plan.conflicts
    }
    assert set(conflicts) == {
        ConflictType.TARGET_CELL_OCCUPIED,
        ConflictType.TARGET_CELL_FORMULA,
        ConflictType.TARGET_CELL_TEXT,
    }
    assert conflicts[ConflictType.TARGET_CELL_OCCUPIED].default_action is None
    assert conflicts[ConflictType.TARGET_CELL_OCCUPIED].allowed_actions == (
        ResolutionAction.KEEP_EXISTING,
        ResolutionAction.OVERWRITE,
    )
    assert all(
        conflict.default_action is None
        and ResolutionAction.SKIP not in conflict.allowed_actions
        for conflict in conflicts.values()
    )
    assert ResolutionAction.ADD not in conflicts[
        ConflictType.TARGET_CELL_OCCUPIED
    ].allowed_actions
    with pytest.raises(CorrectionRequiredError):
        service.apply(plan, {})
    resolutions = {
        conflicts[ConflictType.TARGET_CELL_OCCUPIED].conflict_id: {
            "action": "OVERWRITE"
        },
        conflicts[ConflictType.TARGET_CELL_FORMULA].conflict_id: {
            "action": "OVERWRITE"
        },
        conflicts[ConflictType.TARGET_CELL_TEXT].conflict_id: {
            "action": "KEEP_EXISTING"
        },
    }

    result = service.apply(plan, resolutions)

    assert result.posted_source_items == 4
    assert result.written_cells == 4
    assert result.already_existing_items == 1
    assert result.skipped_source_items == 1
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        assert [sheet.cell(row, cb_column).value for row in range(2, 8)] == [
            100,
            200,
            300,
            100,
            400,
            "đã nhập tay",
        ]
    finally:
        workbook.close()


def test_posting_uses_primary_row_beside_ron_and_reports_unsafe_matches(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    primary_and_ron = "DRYU3026167"
    ambiguous = "MSCU1234567"
    missing = "OOLU7654321"
    _save_ready(
        ready,
        [
            [primary_and_ron, None, "NV", "ST", 100],
            [ambiguous, None, "NV", "ST", 200],
            [missing, None, "NV", "ST", 300],
            [None, "BL-ONLY", "CB", "HD", 400],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, primary_and_ron, sqt=700)
    _add_posting_row(
        sheet,
        3,
        primary_and_ron,
        sqt=None,
        closing_date=None,
        cargo_type="Ron",
    )
    _add_posting_row(sheet, 4, ambiguous, sqt=701)
    _add_posting_row(sheet, 5, ambiguous, sqt=702)
    workbook.save(target)
    workbook.close()

    plan = _posting_service(ready, target, runtime_dir).analyze(
        sheet_name="T07 26"
    )
    by_container = {item.container: item for item in plan.items if item.container}

    assert by_container[primary_and_ron].target_row is None
    assert by_container[ambiguous].target_row is None
    assert by_container[missing].target_row is None
    conflict_types = {
        conflict.conflict_type for conflict in plan.conflicts
    }
    multiple = [
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_CONTAINER_MATCH
    ]
    assert len(multiple) == 2
    assert ConflictType.CONTAINER_NOT_FOUND in conflict_types
    assert ConflictType.BL_ONLY_NO_CONTAINER in conflict_types


def test_posting_unknown_fee_and_invoice_header_are_never_auto_mapped(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(
        ready,
        [
            ["DRYU3026167", None, "CXD", "ST", 100],
            ["MSCU1234567", None, "CB", "ST", 200],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    _add_posting_row(sheet, 3, "MSCU1234567")
    sheet.cell(1, POSTING_FEE_COLUMNS["CB"]).value = "Hóa đơn"
    workbook.save(target)
    workbook.close()

    plan = _posting_service(ready, target, runtime_dir).analyze(
        sheet_name="T07 26"
    )
    conflicts = {
        conflict.conflict_type: conflict for conflict in plan.conflicts
    }

    assert ConflictType.UNKNOWN_FEE_CODE in conflicts
    assert conflicts[ConflictType.UNKNOWN_FEE_CODE].fee == "CXD"
    assert set(conflicts[ConflictType.UNKNOWN_FEE_CODE].details["fees"]) == set(
        FEE_HEADER_ALIASES
    )
    assert ConflictType.FEE_COLUMN_MISSING in conflicts
    assert conflicts[ConflictType.FEE_COLUMN_MISSING].fee == "CB"
    assert conflicts[ConflictType.FEE_COLUMN_MISSING].target_column is None


def test_posting_rejects_invoice_unit_price_alias_after_notes(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [["DRYU3026167", None, "CBDH", "ST", 200]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(1, POSTING_FEE_COLUMNS["CBDH"]).value = None
    notes_column = len(POSTING_BASE_HEADERS) + len(POSTING_FEE_COLUMNS) + 1
    sheet.cell(1, notes_column + 1).value = "ĐƠN GIÁ"
    workbook.save(target)
    workbook.close()

    plan = _posting_service(ready, target, tmp_path / "Excel").analyze(
        sheet_name="T07 26"
    )

    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.FEE_COLUMN_MISSING
    )
    assert conflict.fee == "CBDH"
    assert conflict.target_column is None


def test_posting_partial_replay_sums_only_unposted_source_rows(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(
        ready,
        [
            ["DRYU3026167", None, "VTN", "ST", 100],
            ["DRYU3026167", None, "VTN", "ST", 250],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    workbook.save(target)
    workbook.close()
    repository = _PostedIndexRepository({0})

    service = _posting_service(
        ready,
        target,
        runtime_dir,
        posting_repository=repository,
    )
    first_plan = service.analyze(sheet_name="T07 26")

    assert first_plan.already_posted_indices == {0}
    assert len(first_plan.previously_posted_items) == 1
    assert first_plan.requires_user_input
    plan = service.analyze(
        sheet_name="T07 26",
        repost_source_indices=[],
    )
    assert len(plan.items) == 1
    assert plan.items[0].source_indices == [1]
    assert plan.items[0].amount == 250

    result = service.apply(plan, {})

    assert result.posted_source_items == 1
    assert result.written_cells == 1
    workbook = load_workbook(target)
    try:
        assert workbook["T07 26"].cell(
            2, POSTING_FEE_COLUMNS["VTN"]
        ).value == 250
    finally:
        workbook.close()


def test_posting_aborts_without_backup_when_target_changes_after_analyze(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(ready, [["DRYU3026167", None, "CB", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    workbook.save(target)
    workbook.close()
    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")

    workbook = load_workbook(target)
    workbook["T07 26"]["A30"] = "external change"
    workbook.save(target)
    workbook.close()
    changed_hash = _sha256(target)

    with pytest.raises(WorkbookChangedError):
        service.apply(plan, {})

    assert _sha256(target) == changed_hash
    assert not (runtime_dir / "Backup").exists()


def test_refine_cxd_surfaces_target_cell_conflict_before_apply(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [["DRYU3026167", None, "CXD", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(2, POSTING_FEE_COLUMNS["VTN"]).value = 50
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.UNKNOWN_FEE_CODE
    )

    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_FEE",
                "selected_fee": "VTN",
            }
        },
    )

    assert refined.items[0].selected_fee == "VTN"
    assert [
        item.conflict_type for item in refined.conflicts
    ] == [ConflictType.TARGET_CELL_OCCUPIED]


def test_refine_compares_each_json_member_instead_of_bundle_hash(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [["DRYU3026167", None, "CXD", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.UNKNOWN_FEE_CODE
    )
    assert plan.source_members
    assert plan.source_members[0]["sha256"] == _sha256(ready)

    # A real posting bundle may combine several JSON members, so this digest
    # intentionally differs from every member's raw file hash.
    plan.batch_hash = hashlib.sha256(b"combined posting bundle").hexdigest()

    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_FEE",
                "selected_fee": "VTN",
            }
        },
    )

    assert refined.items[0].selected_fee == "VTN"


def test_refine_reports_when_a_json_member_really_changed(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [["DRYU3026167", None, "CXD", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.UNKNOWN_FEE_CODE
    )
    _save_ready(ready, [["DRYU3026167", None, "CXD", "ST", 250]])

    with pytest.raises(SourceDataChangedError):
        service.refine(
            plan,
            {
                conflict.conflict_id: {
                    "action": "SELECT_FEE",
                    "selected_fee": "VTN",
                }
            },
        )


def test_refine_manual_row_surfaces_formula_conflict(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [[None, "BL-01", "CB", "HD", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(2, POSTING_FEE_COLUMNS["CB"]).value = "=10+20"
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.BL_ONLY_NO_CONTAINER
    )
    assert [candidate.row for candidate in conflict.row_candidates] == [2]

    refined = service.refine(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_ROW",
                "selected_row": 2,
            }
        },
    )

    assert [
        item.conflict_type for item in refined.conflicts
    ] == [ConflictType.TARGET_CELL_FORMULA]


def test_apply_rejects_direct_fee_selection_until_plan_is_refined(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(ready, [["DRYU3026167", None, "CXD", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(2, POSTING_FEE_COLUMNS["VTN"]).value = 50
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, runtime_dir)
    plan = service.analyze(sheet_name="T07 26")
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.UNKNOWN_FEE_CODE
    )
    before = _sha256(target)

    with pytest.raises(ExpensePostingError, match="refine"):
        service.apply(
            plan,
            {
                conflict.conflict_id: {
                    "action": "SELECT_FEE",
                    "selected_fee": "VTN",
                }
            },
        )

    assert _sha256(target) == before
    assert not (runtime_dir / "Backup").exists()


@pytest.mark.parametrize(
    ("existing", "action", "conflict_type"),
    [
        (50, ResolutionAction.KEEP_EXISTING, ConflictType.TARGET_CELL_OCCUPIED),
        (
            "=10+20",
            ResolutionAction.KEEP_FORMULA,
            ConflictType.TARGET_CELL_FORMULA,
        ),
    ],
)
def test_posting_history_preserves_keep_action_and_cell_value(
    tmp_path: Path,
    existing: Any,
    action: ResolutionAction,
    conflict_type: ConflictType,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    _save_ready(ready, [["DRYU3026167", None, "CB", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(2, POSTING_FEE_COLUMNS["CB"]).value = existing
    workbook.save(target)
    workbook.close()
    runs = _RunHistoryRepository()
    postings = _PostingHistoryRepository()
    service = _posting_service(
        ready,
        target,
        runtime_dir,
        posting_repository=postings,
        run_repository=runs,
    )
    plan = service.analyze(batch_id=71, sheet_name="T07 26")
    conflict = next(
        item for item in plan.conflicts if item.conflict_type is conflict_type
    )

    result = service.apply(
        plan,
        {conflict.conflict_id: {"action": action.value}},
    )

    assert result.status is ExcelRunStatus.NO_CHANGES
    assert len(postings.items) == 1
    assert postings.items[0]["action"] is action
    assert postings.items[0]["value_before"] == existing
    assert postings.items[0]["value_after"] == existing
    assert postings.metadata["batch_id"] == 71
    assert not (runtime_dir / "Backup").exists()


def test_posting_writes_invoice_next_to_each_supported_fee_and_skips_ll(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    runtime_dir = tmp_path / "Excel"
    rows: list[list[Any]] = []
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    for offset, fee in enumerate(POSTING_INVOICE_LAYOUT, 2):
        container = f"TEST{offset:07d}"
        _add_posting_row(sheet, offset, container, sqt=700 + offset)
        rows.append(
            [
                container,
                None,
                fee,
                "HD" if fee == "CB" else "ST",
                f"INV-{fee}",
                None,
                offset * 100,
            ]
        )
    workbook.save(target)
    workbook.close()
    _save_ready(ready, rows)

    runs = _RunHistoryRepository()
    postings = _PostingHistoryRepository()
    service = _posting_service(
        ready,
        target,
        runtime_dir,
        run_repository=runs,
        posting_repository=postings,
    )
    plan = service.analyze(batch_id=71, sheet_name="T07 26")
    assert not plan.conflicts

    result = service.apply(plan, {})

    assert result.written_cells == len(POSTING_INVOICE_LAYOUT)
    assert result.invoice_written_cells == len(POSTING_INVOICE_LAYOUT) - 1
    cb_history = next(
        item for item in postings.items if item["fee_selected"] == "CB"
    )
    assert cb_history["invoice_no"] == "INV-CB"
    assert cb_history["invoice_selected"] == "INV-CB"
    assert cb_history["invoice_target_cell"] == "H2"
    assert cb_history["invoice_value_after"] == "INV-CB"
    assert cb_history["invoice_action"] is ResolutionAction.OVERWRITE
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        for offset, fee in enumerate(POSTING_INVOICE_LAYOUT, 2):
            amount_column, invoice_column = POSTING_INVOICE_LAYOUT[fee]
            assert sheet.cell(offset, amount_column).value == offset * 100
            if fee == "LL":
                assert invoice_column is None
            else:
                assert sheet.cell(offset, int(invoice_column)).value == f"INV-{fee}"
    finally:
        workbook.close()


@pytest.mark.parametrize("existing_invoice", ["INV-OLD", "=1+1"])
def test_invoice_conflict_is_independent_from_amount_conflict(
    tmp_path: Path,
    existing_invoice: str,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [[container, None, "CB", "HD", "INV-NEW", None, 100]],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    _add_posting_row(sheet, 2, container)
    amount_column, invoice_column = POSTING_INVOICE_LAYOUT["CB"]
    sheet.cell(2, amount_column).value = 50
    sheet.cell(2, int(invoice_column)).value = existing_invoice
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    amount_conflict = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.TARGET_CELL_OCCUPIED
    )
    invoice_conflict = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.INVOICE_VALUE_CONFLICT
    )
    assert invoice_conflict.allowed_actions == (
        ResolutionAction.KEEP_EXISTING,
        ResolutionAction.OVERWRITE,
    )

    result = service.apply(
        plan,
        {
            amount_conflict.conflict_id: {"action": "KEEP_EXISTING"},
            invoice_conflict.conflict_id: {"action": "OVERWRITE"},
        },
    )

    assert result.status is ExcelRunStatus.SUCCEEDED
    assert result.written_cells == 0
    assert result.invoice_written_cells == 1
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        assert sheet.cell(2, amount_column).value == 50
        assert sheet.cell(2, int(invoice_column)).value == "INV-NEW"
    finally:
        workbook.close()


def test_same_cell_expenses_require_one_json_line_before_invoice_write(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [
            [container, None, "VTN", "ST", "INV-A", None, 100],
            [container, None, "VTN", "ST", "INV-B", None, 200],
        ],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    _add_posting_row(sheet, 2, container)
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    same_cell = next(
        conflict
        for conflict in plan.conflicts
        if conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL
    )
    selected = service.refine(
        plan,
        {
            same_cell.conflict_id: {
                "action": "SELECT_SOURCE_ITEM",
                "selected_source_item_index": 1,
            }
        },
    )
    assert not selected.conflicts
    result = service.apply(selected, {})

    assert result.written_cells == 1
    assert result.invoice_written_cells == 1
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook["T07 26"]
        amount_column, invoice_column = POSTING_INVOICE_LAYOUT["VTN"]
        assert sheet.cell(2, amount_column).value == 200
        assert sheet.cell(2, int(invoice_column)).value == "INV-B"
    finally:
        workbook.close()


def test_blank_invoice_preserves_existing_value_and_missing_header_is_explicit(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(ready, [[container, None, "CBDH", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    _add_posting_row(sheet, 2, container)
    amount_column, invoice_column = POSTING_INVOICE_LAYOUT["CBDH"]
    sheet.cell(2, int(invoice_column)).value = "INV-OLD"
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    result = service.apply(service.analyze(sheet_name="T07 26"), {})
    assert result.invoice_written_cells == 0
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook["T07 26"].cell(2, int(invoice_column)).value == "INV-OLD"
    finally:
        workbook.close()

    missing_ready = tmp_path / "missing-ready.json"
    missing_target = tmp_path / "missing-BK.xlsx"
    _save_ready(
        missing_ready,
        [[container, None, "CBDH", "ST", "INV-NEW", None, 200]],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    _add_posting_row(sheet, 2, container)
    sheet.cell(1, int(invoice_column)).value = "Không phải hóa đơn"
    workbook.save(missing_target)
    workbook.close()
    missing_service = _posting_service(
        missing_ready, missing_target, tmp_path / "MissingExcel"
    )
    missing_plan = missing_service.analyze(sheet_name="T07 26")
    missing_conflict = next(
        conflict
        for conflict in missing_plan.conflicts
        if conflict.conflict_type is ConflictType.INVOICE_COLUMN_MISSING
    )
    assert missing_conflict.default_action is ResolutionAction.SKIP_INVOICE
    missing_result = missing_service.apply(missing_plan, {})
    assert missing_result.written_cells == 1
    assert missing_result.invoice_written_cells == 0


@pytest.mark.parametrize(
    ("existing_invoice", "incoming_invoice", "expected_invoice", "written"),
    [
        (790, "790", 790, 0),
        (None, "000790", "000790", 1),
    ],
)
def test_invoice_comparison_is_textual_and_leading_zero_is_preserved(
    tmp_path: Path,
    existing_invoice: Any,
    incoming_invoice: str,
    expected_invoice: Any,
    written: int,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    container = "DRYU3026167"
    _save_ready(
        ready,
        [[container, None, "HH", "ST", incoming_invoice, None, 100]],
    )
    workbook = Workbook()
    sheet = _new_posting_sheet_with_invoices(workbook, "T07 26")
    _add_posting_row(sheet, 2, container)
    _, invoice_column = POSTING_INVOICE_LAYOUT["HH"]
    sheet.cell(2, int(invoice_column)).value = existing_invoice
    workbook.save(target)
    workbook.close()

    service = _posting_service(ready, target, tmp_path / "Excel")
    plan = service.analyze(sheet_name="T07 26")
    assert not any(
        conflict.conflict_type is ConflictType.INVOICE_VALUE_CONFLICT
        for conflict in plan.conflicts
    )
    result = service.apply(plan, {})
    assert result.invoice_written_cells == written
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook["T07 26"].cell(2, int(invoice_column)).value == expected_invoice
    finally:
        workbook.close()


def test_daily_sync_ignores_filename_year_and_requires_target_sheet_when_ambiguous(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "bao cao noi bo 2025-2027.xlsx"
    target = tmp_path / "so cai tong hop.xlsx"
    _save_daily(daily, [_sync_row(700)])
    workbook = Workbook()
    first = workbook.active
    first.title = "T07 25"
    _populate_target_sheet(first, [_sync_row(700, "OLDU0000001")])
    second = workbook.create_sheet("T07 26")
    _populate_target_sheet(second, [_sync_row(699, "OLDU0000002")])
    workbook.save(target)
    workbook.close()

    service = _sync_service(daily, target, tmp_path / "Excel")
    plan = service.analyze(source_sheet_name="Tháng 7")

    assert plan.source_year is None
    assert plan.selected_sheet is None
    assert {item.target_sheet for item in plan.month_candidates} == {
        "T07 25",
        "T07 26",
    }
    assert {
        item.target_sheet: item.new_row_count
        for item in plan.month_candidates
    } == {"T07 25": 0, "T07 26": 1}
    conflict = next(
        item
        for item in plan.conflicts
        if item.conflict_type is ConflictType.TARGET_MONTH_AMBIGUOUS
    )
    result = service.apply(
        plan,
        {
            conflict.conflict_id: {
                "action": "SELECT_MONTH",
                "selected_month": 7,
                "selected_sheet": "T07 26",
            }
        },
    )

    assert result.sheet_name == "T07 26"
    workbook = load_workbook(target)
    try:
        assert workbook["T07 25"]["A3"].value is None
        assert workbook["T07 26"]["A3"].value == 700
    finally:
        workbook.close()


def test_daily_sync_uses_the_only_existing_sheet_for_source_month(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "du lieu tuy chon.xlsx"
    target = tmp_path / "bk nhieu nam.xlsx"
    _save_daily(daily, [_sync_row(700)])
    workbook = Workbook()
    july = workbook.active
    july.title = "T07 25"
    _populate_target_sheet(july, [_sync_row(699, "OLDU0000001")])
    june = workbook.create_sheet("T06 26")
    _populate_target_sheet(june, [_sync_row(699, "OLDU0000002")])
    workbook.save(target)
    workbook.close()

    service = _sync_service(daily, target, tmp_path / "Excel")
    plan = service.analyze()

    assert [item.target_sheet for item in plan.month_candidates] == ["T07 25"]
    assert plan.selected_sheet == "T07 25"
    assert not any(
        item.conflict_type is ConflictType.TARGET_MONTH_AMBIGUOUS
        for item in plan.conflicts
    )

    result = service.apply(plan, {})

    assert result.sheet_name == "T07 25"
    workbook = load_workbook(target)
    try:
        assert workbook["T07 25"]["A3"].value == 700
        assert "T07 26" not in workbook.sheetnames
    finally:
        workbook.close()


@pytest.mark.parametrize("existing_header", [None, "Data cập nhật", "Date cập nhật"])
def test_posting_writes_or_reuses_update_date_column(
    tmp_path: Path,
    existing_header: str | None,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK tuy chon.xlsx"
    fixed = datetime(2026, 7, 29, 15, 4, 5)
    _save_ready(ready, [["DRYU3026167", None, "NV", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    last_header = len(POSTING_BASE_HEADERS) + len(POSTING_FEE_COLUMNS) + 2
    sheet.cell(1, last_header).fill = PatternFill("solid", fgColor="FFF2CC")
    update_column = last_header + 1
    if existing_header is not None:
        sheet.cell(1, update_column).value = existing_header
    workbook.save(target)
    workbook.close()

    service = _posting_service(
        ready,
        target,
        tmp_path / "Excel",
        clock=lambda: fixed,
    )
    result = service.apply(
        service.analyze(sheet_name="T07 26"),
        {},
    )

    assert result.written_cells == 1
    workbook = load_workbook(target)
    try:
        sheet = workbook["T07 26"]
        assert sheet.cell(1, update_column).value == (
            existing_header or "Date cập nhật"
        )
        assert sheet.cell(2, update_column).value == fixed
        assert sheet.cell(2, update_column).number_format == "dd/mm/yyyy hh:mm:ss"
        if existing_header is None:
            assert (
                sheet.cell(1, update_column).fill.fgColor.rgb
                == sheet.cell(1, last_header).fill.fgColor.rgb
            )
    finally:
        workbook.close()


def test_posting_same_value_only_updates_date_when_explicitly_reposted(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK.xlsx"
    fixed = datetime(2026, 7, 29, 16, 30, 0)
    _save_ready(ready, [["DRYU3026167", None, "NV", "ST", 100]])
    workbook = Workbook()
    sheet = _new_posting_sheet(workbook, "T07 26")
    _add_posting_row(sheet, 2, "DRYU3026167")
    sheet.cell(2, POSTING_FEE_COLUMNS["NV"]).value = 100
    workbook.save(target)
    workbook.close()

    normal = _posting_service(
        ready,
        target,
        tmp_path / "Normal",
        clock=lambda: fixed,
    )
    normal_result = normal.apply(normal.analyze(sheet_name="T07 26"), {})
    assert normal_result.written_cells == 0
    assert normal_result.backup_path is None
    workbook = load_workbook(target)
    try:
        headers = [
            workbook["T07 26"].cell(1, column).value
            for column in range(1, workbook["T07 26"].max_column + 1)
        ]
        assert "Date cập nhật" not in headers
    finally:
        workbook.close()

    repost = _posting_service(
        ready,
        target,
        tmp_path / "Repost",
        posting_repository=_PostedIndexRepository({0}),
        clock=lambda: fixed,
    )
    repost_plan = repost.analyze(
        sheet_name="T07 26",
        repost_source_indices=[0],
    )
    repost_result = repost.apply(repost_plan, {})

    assert repost_plan.items[0].force_repost
    assert repost_result.written_cells == 1
    workbook = load_workbook(target)
    try:
        sheet = workbook["T07 26"]
        date_column = next(
            column
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(1, column).value == "Date cập nhật"
        )
        assert sheet.cell(2, date_column).value == fixed
    finally:
        workbook.close()


def test_posting_sheet_cancel_is_audited_as_cancelled(
    tmp_path: Path,
) -> None:
    ready = tmp_path / "ready.json"
    target = tmp_path / "BK 2026.xlsx"
    _save_ready(ready, [[None, "BL-ONLY", "CB", "ST", 100]])
    workbook = Workbook()
    _new_posting_sheet(workbook, "T07 26")
    _new_posting_sheet(workbook, "T08 26")
    workbook.save(target)
    workbook.close()
    runs = _RunHistoryRepository()
    service = _posting_service(
        ready,
        target,
        tmp_path / "Excel",
        run_repository=runs,
    )
    plan = service.analyze()
    assert plan.selected_sheet is None
    assert not plan.conflicts
    service.cancel(plan)

    assert runs.finished[-1]["status"] is ExcelRunStatus.CANCELLED


def test_daily_month_cancel_is_audited_as_cancelled(
    tmp_path: Path,
) -> None:
    daily = tmp_path / "Hàng ngày 2026.xlsx"
    target = tmp_path / "BK 2026.xlsx"
    workbook = Workbook()
    for month, container in ((7, "DRYU3026167"), (8, "MSCU1234567")):
        sheet = workbook.active if month == 7 else workbook.create_sheet()
        sheet.title = f"Tháng {month}"
        sheet.append(SYNC_HEADERS)
        sheet.append(_sync_row(700, container))
    workbook.save(daily)
    workbook.close()
    workbook = Workbook()
    for month, container in ((7, "DRYU3026167"), (8, "MSCU1234567")):
        sheet = workbook.active if month == 7 else workbook.create_sheet()
        sheet.title = f"T{month:02d} 26"
        _populate_target_sheet(sheet, [_sync_row(699, container)])
    workbook.save(target)
    workbook.close()
    runs = _RunHistoryRepository()
    service = _sync_service(
        daily,
        target,
        tmp_path / "Excel",
        run_repository=runs,
    )
    plan = service.analyze()
    assert plan.source_target_sheets == {
        "Tháng 7": "T07 26",
        "Tháng 8": "T08 26",
    }
    service.cancel(plan)

    assert runs.finished[-1]["status"] is ExcelRunStatus.CANCELLED


def test_daily_sync_applies_two_selected_months_with_one_backup(tmp_path: Path) -> None:
    daily = tmp_path / "Hàng ngày.xlsx"
    target = tmp_path / "BK.xlsx"
    runtime = tmp_path / "Excel"
    workbook = Workbook()
    june = workbook.active
    june.title = "Tháng 6"
    june.append(SYNC_HEADERS)
    june.append(_sync_row(601, "JUNE0000001", closing_date="2026-06-10"))
    july = workbook.create_sheet("Tháng 7")
    july.append(SYNC_HEADERS)
    july.append(_sync_row(701, "JULY0000001", closing_date="2026-07-10"))
    workbook.save(daily)
    workbook.close()
    workbook = Workbook()
    june_target = workbook.active
    june_target.title = "T06 26"
    _populate_target_sheet(june_target, [_sync_row(600, "OLDJUNE0001")])
    july_target = workbook.create_sheet("T07 26")
    _populate_target_sheet(july_target, [_sync_row(700, "OLDJULY0001")])
    workbook.save(target)
    workbook.close()
    service = _sync_service(daily, target, runtime)

    plan = service.analyze(source_sheet_names=["Tháng 6", "Tháng 7"])

    assert plan.source_target_sheets == {
        "Tháng 6": "T06 26",
        "Tháng 7": "T07 26",
    }
    result = service.apply(plan, {})
    assert result.target_sheets == ("T06 26", "T07 26")
    assert result.inserted_rows == 2
    assert len(list((runtime / "Backup").glob("*.xlsx"))) == 1
    workbook = load_workbook(target, data_only=False)
    try:
        assert workbook["T06 26"]["A3"].value == 601
        assert workbook["T07 26"]["A3"].value == 701
    finally:
        workbook.close()
