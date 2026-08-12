from __future__ import annotations

from pathlib import Path

import pytest

from app.database import Database
from app.config import AppSettings
from app.models import DataRow
from app.repositories.batch_repository import BatchRepository
from app.services.batch_service import BatchService
from app.services.reviewed_batch_provider import ReviewedBatchProvider
from app.services.excel.posting import ExpensePostingService
from app.sea_freight import (
    BkContainerSnapshot,
    ContainerRecord,
    GroupStatus,
    SeaFreightReconciliationService,
    SeaFreightRepository,
    iso6346_check_digit,
    normalize_match_key,
    validate_vessel_voyage,
)
from app.ui.sea_freight_center import SeaFreightReconciliationDialog


def _container(serial: int) -> str:
    body = f"TSTU{serial:06d}"
    return body + str(iso6346_check_digit(body))


def _snapshot(tmp_path: Path, count: int = 6, fingerprint: str = "fp-1") -> BkContainerSnapshot:
    return BkContainerSnapshot(
        bk_path=str((tmp_path / "BK.xlsx").resolve()),
        bk_sheet="T07 26",
        vessel_voyage_raw="PROSPER 2625S",
        vessel_key="PROSPER",
        voyage_key="2625S",
        combined_key="PROSPER2625S",
        workbook_fingerprint=fingerprint,
        snapshot_hash=f"snapshot-{count}-{fingerprint}",
        containers=tuple(
            ContainerRecord(_container(index), "T07 26", index + 10, index, None)
            for index in range(1, count + 1)
        ),
    )


class _Matcher:
    def __init__(self, snapshot: BkContainerSnapshot) -> None:
        self.current = snapshot

    def snapshot(self, *_args, **_kwargs) -> BkContainerSnapshot:
        return self.current

    def sheet_names(self, _path) -> tuple[str, ...]:
        return ("T07 26",)


def _row(count: int, amount: int, invoice: str, carrier: str = "HÃNG TÀU") -> DataRow:
    return DataRow(
        cont=None,
        bl=f"BL-{invoice}",
        fee="CB",
        rule="HD",
        amount=amount,
        invoice_no=invoice,
        carrier=carrier,
        vessel_voyage_raw="PROSPER 2625S",
        vessel_name="PROSPER",
        voyage_no="2625S",
        invoice_container_count=count,
        container_count_basis="EXPLICIT",
        invoice_date="2026-07-15",
    )


def test_vessel_voyage_normalization_keeps_voyage_letters() -> None:
    assert normalize_match_key("  prosper   2625s ") == "PROSPER2625S"
    assert validate_vessel_voyage(
        "PHUC KHANH V.424S", "PHUC KHANH", "V.424S"
    ) == ("PHUCKHANH", "V424S", "PHUCKHANHV424S")
    with pytest.raises(ValueError):
        validate_vessel_voyage("TP 16/5", "TP", "16/5")


def test_accumulates_pending_then_ready_and_allocates_exact_total(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    matcher = _Matcher(_snapshot(tmp_path))
    service = SeaFreightReconciliationService(repository, matcher=matcher)

    first = service.attach(
        _row(4, 32_000_000, "INV-001"),
        bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="source-a",
    )
    assert first.status is GroupStatus.PENDING
    assert first.missing_count == 2

    ready = service.attach(
        _row(2, 16_000_003, "INV-002"),
        bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=1, source_sha256="source-b",
    )
    assert ready.status is GroupStatus.READY
    rows = service.prepare_allocation(ready.id)
    assert len(rows) == 6
    assert sum(int(row.amount or 0) for row in rows) == 48_000_003
    assert [row.amount for row in rows[:5]] == [8_000_000] * 5
    assert rows[-1].amount == 8_000_003
    assert rows[0].invoice_no == "INV-001, INV-002"
    assert rows[0].bl == "BL-INV-001, BL-INV-002"
    database.close()


def test_multiple_carriers_are_merged_and_snapshot_change_blocks_confirmation(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    matcher = _Matcher(_snapshot(tmp_path))
    service = SeaFreightReconciliationService(repository, matcher=matcher)
    service.attach(
        _row(4, 10, "A"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="a",
    )
    conflict = service.attach(
        _row(2, 20, "B", "KHÁC"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=1, source_sha256="b",
    )
    assert conflict.status is GroupStatus.READY
    assert service.allocation_rows(conflict.id)[0].carrier == "HÃNG TÀU / KHÁC"
    matcher.current = _snapshot(tmp_path, fingerprint="fp-2")
    with pytest.raises(RuntimeError, match="BK đã thay đổi"):
        service.prepare_allocation(conflict.id)
    assert repository.get_group(conflict.id).status is GroupStatus.NEEDS_RECHECK
    database.close()


def test_unknown_count_basis_is_not_accepted(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    service = SeaFreightReconciliationService(
        SeaFreightRepository(database), matcher=_Matcher(_snapshot(tmp_path))
    )
    row = _row(1, 1, "A")
    row.container_count_basis = "UNKNOWN"
    with pytest.raises(ValueError, match="Căn cứ"):
        service.inspect(row, bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26")
    database.close()


def test_ready_group_creates_only_one_system_batch(tmp_path: Path) -> None:
    settings = AppSettings(data_root=tmp_path, output_dir=tmp_path / "Output")
    database = Database(settings.paths.database_path)
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.attach(
        _row(2, 101, "INV"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="source",
    )
    batch_service = BatchService(
        settings.paths, BatchRepository(database)
    )
    first = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    repository.mark_allocated(group.id, first.metadata.id)
    second = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    assert first.metadata.id == second.metadata.id
    assert first.metadata.source_kind == "SEA_FREIGHT_RECONCILIATION"
    assert first.metadata.reconciliation_group_id == group.id
    assert [row.amount for row in first.document.rows] == [50, 51]
    database.close()


def test_managed_source_row_is_excluded_without_blocking_independent_rows() -> None:
    rows = [
        {"source_item_index": 0, "container": None},
        {"source_item_index": 1, "container": "TSTU0000010"},
    ]
    grouped = ExpensePostingService._group_rows(
        rows, already_posted=set(), managed_source_indices={0}
    )
    assert grouped == [[rows[1]]]


def test_supplement_after_allocation_or_posting_is_not_auto_merged(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.attach(
        _row(2, 100, "A"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="a",
    )
    repository.mark_allocated(group.id)
    with pytest.raises(RuntimeError, match="đã hoàn tất"):
        service.attach(
            _row(1, 10, "B"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
            source_batch_id=None, source_item_index=1, source_sha256="b",
        )
    repository.mark_posted(group.id, None)
    # Posted revisions remain inspectable in the read-only dossier view.
    assert len(service.allocation_rows(group.id)) == 2
    with pytest.raises(RuntimeError, match="đã hoàn tất"):
        service.attach(
            _row(1, 10, "C"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
            source_batch_id=None, source_item_index=2, source_sha256="c",
        )
    database.close()


def test_open_period_save_manual_and_import_matching_supplement(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=5))
    )
    group = service.open_or_create(
        _row(2, 20_000_000, "INV-001"),
        bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=None, source_item_index=0, source_sha256="first",
    )
    assert group.reconciliation_month == 7
    assert group.reconciliation_year == 2026
    contributions = repository.list_contributions(group.id)
    saved = service.save_group(
        group.id,
        vessel_voyage_raw="PROSPER 2625S",
        vessel_name="PROSPER",
        voyage_no="2625S",
        invoices=[
            {
                "id": contributions[0].id,
                "invoice_no": "INV-001",
                "invoice_date": "2026-07-15",
                "bl": "BL-1",
                "invoice_container_count": 2,
                "amount": 20_000_000,
                "carrier": "HÃNG A",
                "source_kind": "INITIAL",
                "source_sha256": "first",
            },
            {
                "invoice_no": "INV-002",
                "invoice_container_count": 1,
                "amount": 10_000_000,
                "carrier": "HÃNG B",
                "source_kind": "MANUAL",
            },
        ],
    )
    assert saved.status is GroupStatus.PENDING
    supplement = _row(2, 15_000_001, "INV-003", "HÃNG C")
    other = _row(1, 1, "INV-X")
    other.vessel_voyage_raw = "OTHER 999S"
    other.vessel_name = "OTHER"
    other.voyage_no = "999S"
    imported = service.import_supplement(
        group.id, [supplement, other, supplement],
        source_batch_id=None, source_sha256="supplement",
    )
    assert imported.added_count == 1
    assert imported.duplicate_invoices == ("INV-003",)
    assert imported.skipped_count == 1
    ready = repository.get_group(group.id)
    assert ready is not None and ready.status is GroupStatus.READY
    rows = service.prepare_allocation(group.id)
    assert rows[-1].amount == 9_000_001
    assert rows[0].invoice_no == "INV-001, INV-002, INV-003"
    assert rows[0].carrier == "HÃNG A / HÃNG B / HÃNG C"
    assert rows[0].invoice_date is None
    database.close()


def test_duplicate_invoice_number_is_rejected_on_save(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.open_or_create(
        _row(1, 100, "INV"), bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=None, source_item_index=0, source_sha256="a",
    )
    original = repository.list_contributions(group.id)[0]
    with pytest.raises(RuntimeError, match="bị trùng"):
        service.save_group(
            group.id,
            vessel_voyage_raw="PROSPER 2625S", vessel_name="PROSPER", voyage_no="2625S",
            invoices=[
                {"id": original.id, "invoice_no": "INV", "invoice_container_count": 1, "amount": 100},
                {"invoice_no": " inv ", "invoice_container_count": 1, "amount": 100},
            ],
        )
    assert len(repository.list_contributions(group.id)) == 1
    database.close()


def test_reconciliation_dialog_has_only_the_six_confirmed_actions(qtbot, tmp_path: Path) -> None:
    settings = AppSettings(data_root=tmp_path, output_dir=tmp_path / "Output")
    database = Database(settings.paths.database_path)
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.open_or_create(
        _row(2, 101, "INV"), bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=None, source_item_index=0, source_sha256="source",
    )
    batch_service = BatchService(settings.paths, BatchRepository(database))
    dialog = SeaFreightReconciliationDialog(
        service,
        batch_service=batch_service,
        open_assistant=lambda: None,
        group_id=group.id,
    )
    qtbot.addWidget(dialog)
    dialog.show()
    try:
        assert [
            dialog.assistant_button.text(), dialog.add_button.text(),
            dialog.delete_button.text(), dialog.save_button.text(),
            dialog.confirm_button.text(), dialog.close_button.text(),
        ] == [
            "Mở Trợ lý bóc tách", "Thêm dòng HĐ", "Xóa dòng",
            "Lưu", "Xác nhận", "Đóng",
        ]
        assert dialog.confirm_button.isEnabled()
        assert dialog.invoice_table.rowCount() == 1
        assert dialog.container_table.rowCount() == 2
        assert dialog.result_table.rowCount() == 2
        assert dialog.status_label.text().startswith("Đủ 2/2 cont")
    finally:
        dialog.close()
        database.close()


def test_cancelled_group_releases_invoice_for_a_new_reconciliation_and_keeps_history(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    source = _row(2, 100, "INV-RETRY")
    old_group = service.open_or_create(
        source, bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=None, source_item_index=0, source_sha256="same-file",
    )
    old_history = repository.list_contribution_history(old_group.id)
    assert len(old_history) == 1
    repository.cancel_group(old_group.id)

    assert repository.list_contributions(old_group.id) == []
    retained = repository.list_contribution_history(old_group.id)
    assert len(retained) == 1
    assert retained[0].status == "REMOVED"
    assert retained[0].removed_at is not None
    assert repository.get_group(old_group.id).status is GroupStatus.CANCELLED

    new_group = service.open_or_create(
        source, bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=None, source_item_index=0, source_sha256="same-file",
    )
    assert new_group.id != old_group.id
    assert new_group.status is GroupStatus.READY
    assert [item.invoice_no for item in repository.list_contributions(new_group.id)] == [
        "INV-RETRY"
    ]
    old_events = database.query_all(
        "SELECT event_type FROM sea_freight_reconciliation_events WHERE group_id = ? ORDER BY id",
        (old_group.id,),
    )
    assert [str(row["event_type"]) for row in old_events][-1] == "GROUP_CANCELLED"
    database.close()


def test_cancelling_allocated_group_archives_old_result_batch(tmp_path: Path) -> None:
    settings = AppSettings(data_root=tmp_path, output_dir=tmp_path / "Output")
    database = Database(settings.paths.database_path)
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.attach(
        _row(2, 101, "INV"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="source",
    )
    batch_service = BatchService(settings.paths, BatchRepository(database))
    result = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    repository.mark_allocated(group.id, result.metadata.id)

    repository.cancel_group(group.id)

    archived = batch_service.get_batch(result.metadata.id)
    assert archived is not None
    assert archived.status.value == "ARCHIVED"
    assert repository.list_contributions(group.id) == []
    assert repository.list_contribution_history(group.id)[0].status == "REMOVED"
    database.close()


def test_invoice_fingerprint_includes_vessel_voyage_and_container_count() -> None:
    base = _row(2, 100, "INV")
    fingerprint = SeaFreightReconciliationService.contribution_fingerprint(
        base, source_sha256="same-source"
    )
    different_count = base.copy_with(invoice_container_count=3)
    different_voyage = base.copy_with(
        vessel_voyage_raw="PROSPER 2626S", voyage_no="2626S"
    )

    assert SeaFreightReconciliationService.contribution_fingerprint(
        different_count, source_sha256="same-source"
    ) != fingerprint
    assert SeaFreightReconciliationService.contribution_fingerprint(
        different_voyage, source_sha256="same-source"
    ) != fingerprint


def test_completed_group_can_create_immutable_revision(tmp_path: Path) -> None:
    settings = AppSettings(data_root=tmp_path, output_dir=tmp_path / "Output")
    database = Database(settings.paths.database_path)
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = service.attach(
        _row(2, 101, "INV-REV"), bk_path=tmp_path / "BK.xlsx", bk_sheet="T07 26",
        source_batch_id=None, source_item_index=0, source_sha256="source",
    )
    batch_service = BatchService(settings.paths, BatchRepository(database))
    result = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    repository.mark_allocated(group.id, result.metadata.id)

    revised = service.create_revision(group.id)

    old = repository.get_group(group.id)
    assert old is not None and not old.is_current
    assert old.status is GroupStatus.ALLOCATED
    assert revised.is_current and revised.revision_no == 2
    assert revised.supersedes_group_id == group.id
    assert revised.status is GroupStatus.READY
    assert [item.invoice_no for item in repository.list_contributions(revised.id)] == [
        "INV-REV"
    ]
    archived = batch_service.get_batch(result.metadata.id)
    assert archived is not None and archived.status.value == "ARCHIVED"
    database.close()


def test_posting_bundle_keeps_normal_rows_and_replaces_managed_sea_freight(
    tmp_path: Path,
) -> None:
    settings = AppSettings(
        data_root=tmp_path,
        output_dir=tmp_path / "Output",
        bk_workbook_path=str(tmp_path / "BK.xlsx"),
    )
    database = Database(settings.paths.database_path)
    batch_repository = BatchRepository(database)
    batch_service = BatchService(settings.paths, batch_repository)
    source_path = settings.output_dir / "incoming.json"
    source_path.parent.mkdir(parents=True, exist_ok=True)
    source_path.write_text(
        '{"v":2,"d":['
        '{"container":"TSTU0000017","bl":null,"vessel_voyage_raw":null,'
        '"vessel_name":null,"voyage_no":null,"invoice_container_count":null,'
        '"container_count_basis":"UNKNOWN","fee":"VTN","rule":"ST",'
        '"invoice_no":"NORMAL","invoice_date":"2026-07-15",'
        '"carrier":"VT","amount":50},'
        '{"container":null,"bl":"BL-SEA","vessel_voyage_raw":"PROSPER 2625S",'
        '"vessel_name":"PROSPER","voyage_no":"2625S","invoice_container_count":2,'
        '"container_count_basis":"EXPLICIT","fee":"CB","rule":"HD",'
        '"invoice_no":"SEA","invoice_date":"2026-07-15",'
        '"carrier":"HÃNG TÀU","amount":100}]}'
        , encoding="utf-8"
    )
    received = batch_service.receive_file(source_path)
    sea_repository = SeaFreightRepository(database)
    sea_service = SeaFreightReconciliationService(
        sea_repository, matcher=_Matcher(_snapshot(tmp_path, count=2))
    )
    group = sea_service.open_or_create(
        _row(2, 100, "SEA"), bk_path=tmp_path / "BK.xlsx", month=7, year=2026,
        source_batch_id=received.batch.id, source_item_index=1,
        source_sha256=received.batch.sha256,
    )
    child = batch_service.create_reconciliation_batch(
        group.id, sea_service.prepare_allocation(group.id)
    )
    sea_repository.mark_allocated(group.id, child.metadata.id)
    assert received.review is not None
    batch_service.confirm_batch(received.batch.id, received.review.document)

    posting = ExpensePostingService(
        ReviewedBatchProvider(batch_service),
        settings,
        sea_freight_repository=sea_repository,
        sea_freight_service=sea_service,
    )
    bundle = posting._posting_bundle(received.batch.id)

    assert len(bundle["rows"]) == 3
    assert bundle["original_source_count"] == 1
    assert bundle["reconciliation_source_count"] == 2
    assert [row["container"] for row in bundle["rows"]][1:] == [
        record.container for record in _snapshot(tmp_path, count=2).containers
    ]
    assert all(
        row.get("invoice_no") != "SEA" or row["container"]
        for row in bundle["rows"]
    )
    database.close()
