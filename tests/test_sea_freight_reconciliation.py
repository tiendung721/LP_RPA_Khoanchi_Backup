from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from app.database import Database
from app.config import AppSettings
from app.models import DataRow
from app.repositories.batch_repository import BatchRepository
from app.services.batch_service import BatchDataError, BatchService
from app.services.reviewed_batch_provider import ReviewedBatchProvider
from app.services.excel.posting import ExpensePostingService
from app.sea_freight import (
    BkContainerSnapshot,
    BkVesselMatcher,
    ContainerRecord,
    GroupStatus,
    SeaFreightReconciliationService,
    SeaFreightRepository,
    VesselVoyageResolutionKind,
    VesselVoyageNotFoundError,
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

    assert validate_vessel_voyage(
        "AI ĐỌC SAI", "PHUC KHANH", "V.424S"
    ) == ("PHUCKHANH", "V424S", "PHUCKHANHV424S")


def test_vessel_not_found_does_not_create_empty_group(tmp_path: Path) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    matcher = _Matcher(_snapshot(tmp_path, count=0))
    service = SeaFreightReconciliationService(repository, matcher=matcher)

    with pytest.raises(VesselVoyageNotFoundError) as caught:
        service.open_or_create(
            _row(1, 100, "INV-NOT-FOUND"),
            bk_path=tmp_path / "BK.xlsx",
            month=7,
            year=2026,
            source_batch_id=None,
            source_item_index=0,
            source_sha256="missing",
        )

    assert caught.value.vessel_voyage == "PROSPER 2625S"
    assert caught.value.bk_sheet == "T07 26"
    assert repository.list_groups(include_closed=True) == []
    database.close()


def test_matcher_normalizes_exact_text_and_ranks_bk_suggestions(tmp_path: Path) -> None:
    bk_path = tmp_path / "BK.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "T07 26"
    sheet.append(
        [
            "SQT PM",
            "Ngày Đóng",
            "Số Container",
            "Số tấn",
            "Loại hàng",
            "Nơi đóng",
            "Tên tàu",
            "Ngày chạy",
            "Dự kiến giao",
            "Người nhận",
            "VT biển",
            "Vận chuyển",
        ]
    )
    sheet.append(
        [1, None, _container(1), None, None, None, "PHÚC KHANH / V.424S"]
    )
    sheet.append(
        [2, None, _container(2), None, None, None, "NEW VISION 2610S"]
    )
    sheet.append(
        [3, None, _container(3), None, None, None, "NEW VISION 2610S"]
    )
    sheet.append(
        [4, None, _container(4), None, None, None, "NEW VISION 2611N"]
    )
    workbook.save(bk_path)
    workbook.close()
    matcher = BkVesselMatcher()

    exact = matcher.snapshot(
        bk_path,
        "T07 26",
        vessel_voyage_raw="AI ĐỌC KHÁC",
        vessel_name="phuc khanh",
        voyage_no="V 424S",
    )
    suggestions = matcher.suggestions(
        bk_path,
        "T07 26",
        vessel_name="NEW VISON",
        voyage_no="2610S",
    )

    assert exact.container_count == 1
    assert exact.vessel_voyage_raw == "PHÚC KHANH / V.424S"
    assert exact.canonical_vessel_name == "PHÚC KHANH"
    assert exact.canonical_voyage_no == "V.424S"
    assert suggestions[0].vessel_voyage == "NEW VISION 2610S"
    assert suggestions[0].container_count == 2


@pytest.mark.parametrize(
    ("bk_value", "input_voyage", "canonical_voyage"),
    [
        ("VIETSUN DYNAMIC V.1228S", "1228S", "V.1228S"),
        ("BIEN DONG MARINER MB2627S", "2627S", "MB2627S"),
        ("BIEN DONG STAR BS2627S", "2627S", "BS2627S"),
        ("VIMC PIONEER VPN2619S", "2619S", "VPN2619S"),
        ("TEST SHIP ZX123S", "123S", "ZX123S"),
        ("VIETSUN DYNAMIC 1228S", "V.1228S", "1228S"),
    ],
)
def test_matcher_auto_resolves_unique_missing_voyage_prefix(
    tmp_path: Path,
    bk_value: str,
    input_voyage: str,
    canonical_voyage: str,
) -> None:
    vessel_name = bk_value[: -len(canonical_voyage)].strip()
    path = _write_vessel_bk(tmp_path, [(bk_value, _container(1))])

    snapshot = BkVesselMatcher().snapshot(
        path,
        "T08 26",
        vessel_voyage_raw=f"{vessel_name} {input_voyage}",
        vessel_name=vessel_name,
        voyage_no=input_voyage,
    )

    assert snapshot.resolution_kind is VesselVoyageResolutionKind.AUTO_ALIAS
    assert snapshot.canonical_voyage_no == canonical_voyage
    assert snapshot.vessel_voyage_raw == bk_value
    assert snapshot.container_count == 1


def test_exact_voyage_wins_when_prefixed_alias_also_exists(tmp_path: Path) -> None:
    path = _write_vessel_bk(
        tmp_path,
        [
            ("VIETSUN DYNAMIC 1228S", _container(1)),
            ("VIETSUN DYNAMIC V.1228S", _container(2)),
        ],
    )

    snapshot = BkVesselMatcher().snapshot(
        path,
        "T08 26",
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )

    assert snapshot.resolution_kind is VesselVoyageResolutionKind.EXACT
    assert snapshot.vessel_voyage_raw == "VIETSUN DYNAMIC 1228S"
    assert [item.container for item in snapshot.containers] == [_container(1)]


def test_matcher_does_not_guess_when_prefix_alias_is_ambiguous(tmp_path: Path) -> None:
    path = _write_vessel_bk(
        tmp_path,
        [
            ("VIETSUN DYNAMIC V.1228S", _container(1)),
            ("VIETSUN DYNAMIC MB1228S", _container(2)),
        ],
    )

    snapshot = BkVesselMatcher().snapshot(
        path,
        "T08 26",
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )

    assert snapshot.resolution_kind is VesselVoyageResolutionKind.AMBIGUOUS
    assert snapshot.container_count == 0
    assert {item.voyage_no for item in snapshot.alias_candidates} == {
        "V.1228S",
        "MB1228S",
    }
    assert all(item.selectable for item in snapshot.alias_candidates)


def test_service_returns_selectable_candidates_for_ambiguous_alias(
    tmp_path: Path,
) -> None:
    path = _write_vessel_bk(
        tmp_path,
        [
            ("VIETSUN DYNAMIC V.1228S", _container(1)),
            ("VIETSUN DYNAMIC MB1228S", _container(2)),
        ],
    )
    database = Database(tmp_path / "ambiguous-state.db")
    repository = SeaFreightRepository(database)
    row = _row(1, 6_850_000, "INV-AMBIGUOUS").copy_with(
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )

    with pytest.raises(VesselVoyageNotFoundError) as caught:
        SeaFreightReconciliationService(repository).open_or_create(
            row,
            bk_path=path,
            month=8,
            year=2026,
            source_batch_id=None,
            source_item_index=0,
            source_sha256="ambiguous-source",
        )

    assert caught.value.ambiguous is True
    assert {item.voyage_no for item in caught.value.suggestions} == {
        "V.1228S",
        "MB1228S",
    }
    assert repository.list_groups(include_closed=True) == []
    database.close()


@pytest.mark.parametrize(
    ("vessel_name", "voyage_no"),
    [
        ("VIETSUN DYNAMC", "1228S"),
        ("VIETSUN DYNAMIC", "1228N"),
        ("VIETSUN DYNAMIC", "MB1228S"),
    ],
)
def test_matcher_never_auto_aliases_typo_suffix_or_different_prefix(
    tmp_path: Path,
    vessel_name: str,
    voyage_no: str,
) -> None:
    path = _write_vessel_bk(
        tmp_path, [("VIETSUN DYNAMIC V.1228S", _container(1))]
    )

    snapshot = BkVesselMatcher().snapshot(
        path,
        "T08 26",
        vessel_voyage_raw=f"{vessel_name} {voyage_no}",
        vessel_name=vessel_name,
        voyage_no=voyage_no,
    )

    assert snapshot.resolution_kind is VesselVoyageResolutionKind.NOT_FOUND
    assert snapshot.container_count == 0


def test_alias_candidate_without_valid_container_is_not_applied(tmp_path: Path) -> None:
    path = _write_vessel_bk(
        tmp_path, [("VIETSUN DYNAMIC V.1228S", "NOT-A-CONTAINER")]
    )

    snapshot = BkVesselMatcher().snapshot(
        path,
        "T08 26",
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )

    assert snapshot.resolution_kind is VesselVoyageResolutionKind.NOT_FOUND
    assert snapshot.alias_candidates == ()


def test_service_persists_canonical_identity_and_keeps_ai_raw(tmp_path: Path) -> None:
    path = _write_vessel_bk(
        tmp_path,
        [
            ("VIETSUN DYNAMIC V.1228S", _container(1)),
            ("VIETSUN DYNAMIC V.1228S", _container(2)),
        ],
    )
    database = Database(tmp_path / "alias-state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(repository)
    first = _row(1, 6_850_000, "INV-ALIAS-1").copy_with(
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )
    second = _row(1, 6_850_000, "INV-ALIAS-2").copy_with(
        vessel_voyage_raw="VIETSUN DYNAMIC V.1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="V.1228S",
    )

    outcome = service.open_or_create_many(
        [(0, first), (1, second)],
        bk_path=path,
        month=8,
        year=2026,
        source_batch_id=None,
        source_sha256="alias-source",
    )

    assert outcome.resolution_kind is VesselVoyageResolutionKind.AUTO_ALIAS
    assert outcome.canonical_voyage_no == "V.1228S"
    assert outcome.group.voyage_key == "V1228S"
    contributions = repository.list_contributions(outcome.group.id)
    assert {item.voyage_no for item in contributions} == {"V.1228S"}
    assert contributions[0].vessel_voyage_raw == "VIETSUN DYNAMIC 1228S"

    supplement = _row(1, 6_850_000, "INV-ALIAS-3").copy_with(
        vessel_voyage_raw="VIETSUN DYNAMIC 1228S",
        vessel_name="VIETSUN DYNAMIC",
        voyage_no="1228S",
    )
    imported = service.import_supplement(
        outcome.group.id,
        [supplement],
        source_batch_id=None,
        source_sha256="alias-supplement",
    )
    assert imported.added_count == 1
    assert repository.list_contributions(outcome.group.id)[-1].voyage_no == "V.1228S"
    database.close()


def test_custom_gpt_prompts_require_complete_voyage_prefixes() -> None:
    root = Path(__file__).resolve().parents[1]
    for filename in (
        "gpt_custom_instructions_chi_tiet.txt",
        "gpt_custom_instructions_ngan.txt",
    ):
        instructions = (root / filename).read_text(encoding="utf-8")
        assert "V.1228S" in instructions
        assert "MB2627S" in instructions
        assert "Không tự thêm tiền tố" in instructions

    short = (root / "gpt_custom_instructions_ngan.txt").read_text(encoding="utf-8")
    assert len(short) < 8_000


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


def test_open_many_invoices_is_atomic_ready_and_preserves_document_sources(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=10))
    )
    first = _row(4, 32_000_000, "INV-004")
    first.source_document_id = "DOC_001"
    first.source_document_name = "Hoa_don_A.pdf"
    second = _row(6, 48_000_000, "INV-006")
    second.source_document_id = "DOC_002"
    second.source_document_name = "Hoa_don_B.pdf"

    outcome = service.open_or_create_many(
        [(0, first), (1, second)],
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=None,
        source_sha256="batch-sha",
    )

    group = outcome.group
    assert group.status is GroupStatus.READY
    assert outcome.added_source_indices == (0, 1)
    contributions = repository.list_contributions(group.id)
    assert [(item.invoice_no, item.source_document_id) for item in contributions] == [
        ("INV-004", "DOC_001"),
        ("INV-006", "DOC_002"),
    ]
    assert sum(int(item.invoice_container_count) for item in contributions) == 10
    database.close()


def test_cross_batch_duplicate_links_existing_posted_group_without_changing_totals(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batch_repository = BatchRepository(database)
    old_batch = batch_repository.create_batch(
        source_filename="old.json",
        source_output_path=tmp_path / "old.json",
        original_archive_path=tmp_path / "old.json",
        working_path=tmp_path / "old.json",
        sha256="a" * 64,
    )


def _write_vessel_bk(tmp_path: Path, vessels: list[tuple[str, str]]) -> Path:
    path = tmp_path / "BK-alias.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "T08 26"
    sheet.append(
        [
            "SQT PM",
            "Ngày Đóng",
            "Số Container",
            "Số tấn",
            "Loại hàng",
            "Nơi đóng",
            "Tên tàu",
            "Ngày chạy",
            "Dự kiến giao",
            "Người nhận",
            "VT biển",
            "Vận chuyển",
        ]
    )
    for index, (vessel, container) in enumerate(vessels, start=1):
        sheet.append([index, None, container, None, None, None, vessel])
    workbook.save(path)
    workbook.close()
    return path
    new_batch = batch_repository.create_batch(
        source_filename="new.json",
        source_output_path=tmp_path / "new.json",
        original_archive_path=tmp_path / "new.json",
        working_path=tmp_path / "new.json",
        sha256="b" * 64,
    )
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=3))
    )
    original = _row(3, 20_550_000, "00006504", carrier="VIETSUN")
    original.source_document_name = "C26TVS-00006504.pdf"
    group = service.open_or_create(
        original,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=old_batch.id,
        source_item_index=0,
        source_sha256=old_batch.sha256,
    )
    repository.mark_allocated(group.id)
    repository.mark_posted(group.id, None)
    duplicate = original.copy_with()
    duplicate.source_document_name = "C26TVS-00006504 (1).pdf"

    matches = service.sync_batch_history(
        [(3, duplicate)],
        source_batch_id=new_batch.id,
        source_sha256=new_batch.sha256,
    )
    outcome = service.open_or_create_many(
        [(3, duplicate)],
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=new_batch.id,
        source_sha256=new_batch.sha256,
    )

    assert matches[3].kind.value == "EXACT"
    assert matches[3].group is not None and matches[3].group.id == group.id
    assert outcome.group.id == group.id
    assert outcome.duplicate_source_indices == (3,)
    assert not outcome.requires_revision
    assert len(repository.list_groups(include_closed=True)) == 1
    assert len(repository.list_contributions(group.id)) == 1
    assert repository.get_group(group.id).total_amount == 20_550_000
    history = repository.list_contribution_history(group.id)
    assert [item.status for item in history] == ["ACTIVE", "DUPLICATE"]
    assert repository.managed_source_indices(new_batch.id) == {3}
    assert repository.groups_for_source_batch(new_batch.id)[3].id == group.id
    assert [item.id for item in repository.posting_groups_for_source_batch(new_batch.id)] == [
        group.id
    ]

    revised = service.create_revision(group.id)
    assert [item.status for item in repository.list_contribution_history(revised.id)] == [
        "ACTIVE",
        "DUPLICATE",
    ]
    assert repository.get_group(revised.id).total_amount == 20_550_000
    assert repository.groups_for_source_batch(new_batch.id)[3].id == revised.id
    database.close()


def test_same_invoice_with_changed_business_data_is_conflict_not_link(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batch_repository = BatchRepository(database)
    old_batch = batch_repository.create_batch(
        source_filename="old.json",
        source_output_path=tmp_path / "old.json",
        original_archive_path=tmp_path / "old.json",
        working_path=tmp_path / "old.json",
        sha256="c" * 64,
    )
    new_batch = batch_repository.create_batch(
        source_filename="new.json",
        source_output_path=tmp_path / "new.json",
        original_archive_path=tmp_path / "new.json",
        working_path=tmp_path / "new.json",
        sha256="d" * 64,
    )
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=1))
    )
    original = _row(1, 6_850_000, "INV-CONFLICT")
    group = service.open_or_create(
        original,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=old_batch.id,
        source_item_index=0,
        source_sha256=old_batch.sha256,
    )
    repository.mark_allocated(group.id)
    changed = original.copy_with(amount=7_000_000)

    matches = service.sync_batch_history(
        [(0, changed)],
        source_batch_id=new_batch.id,
        source_sha256=new_batch.sha256,
    )

    assert matches[0].kind.value == "CONFLICT"
    assert matches[0].differing_fields == ("số tiền",)
    assert repository.managed_source_indices(new_batch.id) == set()
    assert repository.groups_for_source_batch(new_batch.id) == {}
    outcome = service.open_or_create_many(
        [(0, changed)],
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=new_batch.id,
        source_sha256=new_batch.sha256,
    )
    assert outcome.group.id == group.id
    assert outcome.requires_revision
    assert outcome.duplicate_source_indices == ()
    database.close()


def test_new_invoice_for_completed_voyage_returns_existing_group_for_rerun(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=1))
    )
    group = service.open_or_create(
        _row(1, 6_850_000, "INV-OLD"),
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=None,
        source_item_index=0,
        source_sha256="old",
    )
    repository.mark_allocated(group.id)

    outcome = service.open_or_create_many(
        [(1, _row(1, 7_000_000, "INV-NEW"))],
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=None,
        source_sha256="new",
    )

    assert outcome.group.id == group.id
    assert outcome.requires_revision
    assert [item.invoice_no for item in repository.list_contributions(group.id)] == [
        "INV-OLD"
    ]
    assert len(repository.list_groups(include_closed=True)) == 1
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


def test_missing_reconciliation_json_is_recreated_from_current_group(
    tmp_path: Path,
) -> None:
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
    first = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    repository.mark_allocated(group.id, first.metadata.id)
    result_path = first.metadata.ready_path
    assert result_path is not None
    result_path.unlink()

    with pytest.raises(BatchDataError):
        batch_service.load_batch(first.metadata.id)

    repaired = batch_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )

    assert repaired.metadata.id == first.metadata.id
    assert repaired.metadata.status.value == "READY"
    assert repaired.metadata.last_error is None
    assert result_path.is_file()
    assert [row.amount for row in repaired.document.rows] == [50, 51]
    database.close()


def test_restart_preserves_reconciliation_json(tmp_path: Path) -> None:
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
    batch_repository = BatchRepository(database)
    first_service = BatchService(settings.paths, batch_repository)
    result = first_service.create_reconciliation_batch(
        group.id, service.prepare_allocation(group.id)
    )
    result_path = result.metadata.ready_path
    assert result_path is not None and result_path.is_file()

    restarted = BatchService(settings.paths, batch_repository)

    assert result_path.is_file()
    assert restarted.load_batch(result.metadata.id).document.to_dict() == (
        result.document.to_dict()
    )
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
        assert dialog.group_value.isReadOnly()
        assert dialog.group_value.text().startswith("PROSPER 2625S – Lần 1 –")
        assert not hasattr(dialog, "group_combo")
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


def test_deleting_batch_row_removes_invoice_and_reindexes_remaining_source(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batch = BatchRepository(database).create_batch(
        source_filename="source.json",
        source_output_path=tmp_path / "source.json",
        original_archive_path=tmp_path / "source.json",
        working_path=tmp_path / "source.json",
        sha256="e" * 64,
    )
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=3))
    )
    outcome = service.open_or_create_many(
        [(1, _row(1, 6_000_000, "INV-REMOVE")), (2, _row(2, 12_000_000, "INV-KEEP"))],
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=batch.id,
        source_sha256=batch.sha256,
    )

    service.apply_batch_row_deletions(
        source_batch_id=batch.id,
        deleted_source_indices={0, 1},
        remaining_source_indices={2: 1},
    )

    remaining = repository.list_contributions(outcome.group.id)
    assert [(item.invoice_no, item.source_item_index) for item in remaining] == [
        ("INV-KEEP", 1)
    ]
    assert repository.get_group(outcome.group.id).status is GroupStatus.PENDING
    assert [
        item.status for item in repository.list_contribution_history(outcome.group.id)
    ] == ["REMOVED", "ACTIVE"]
    database.close()


def test_deleting_last_invoice_cancels_profile_and_allows_reconciliation_again(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batch_repository = BatchRepository(database)
    first_batch = batch_repository.create_batch(
        source_filename="first.json",
        source_output_path=tmp_path / "first.json",
        original_archive_path=tmp_path / "first.json",
        working_path=tmp_path / "first.json",
        sha256="f" * 64,
    )
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=1))
    )
    invoice = _row(1, 6_500_000, "INV-REUSE")
    original = service.open_or_create(
        invoice,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=first_batch.id,
        source_item_index=0,
        source_sha256=first_batch.sha256,
    )

    service.apply_batch_row_deletions(
        source_batch_id=first_batch.id,
        deleted_source_indices={0},
        remaining_source_indices={},
    )

    assert repository.get_group(original.id).status is GroupStatus.CANCELLED
    assert repository.list_contributions(original.id) == []
    second_batch = batch_repository.create_batch(
        source_filename="second.json",
        source_output_path=tmp_path / "second.json",
        original_archive_path=tmp_path / "second.json",
        working_path=tmp_path / "second.json",
        sha256="1" * 64,
    )
    repeated = service.open_or_create(
        invoice,
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=second_batch.id,
        source_item_index=0,
        source_sha256=second_batch.sha256,
    )
    assert repeated.id != original.id
    assert repeated.status is GroupStatus.READY
    database.close()


def test_deleting_invoice_from_completed_profile_preserves_old_revision(
    tmp_path: Path,
) -> None:
    database = Database(tmp_path / "state.db")
    batch = BatchRepository(database).create_batch(
        source_filename="posted.json",
        source_output_path=tmp_path / "posted.json",
        original_archive_path=tmp_path / "posted.json",
        working_path=tmp_path / "posted.json",
        sha256="2" * 64,
    )
    repository = SeaFreightRepository(database)
    service = SeaFreightReconciliationService(
        repository, matcher=_Matcher(_snapshot(tmp_path, count=1))
    )
    original = service.open_or_create(
        _row(1, 7_000_000, "INV-COMPLETE"),
        bk_path=tmp_path / "BK.xlsx",
        month=7,
        year=2026,
        source_batch_id=batch.id,
        source_item_index=0,
        source_sha256=batch.sha256,
    )
    repository.mark_allocated(original.id)
    repository.mark_posted(original.id, None)

    service.apply_batch_row_deletions(
        source_batch_id=batch.id,
        deleted_source_indices={0},
        remaining_source_indices={},
    )

    revisions = repository.list_revisions(original.id)
    assert len(revisions) == 2
    current = next(item for item in revisions if item.is_current)
    old = next(item for item in revisions if not item.is_current)
    assert current.status is GroupStatus.CANCELLED
    assert old.id == original.id
    assert old.status is GroupStatus.POSTED
    assert [item.status for item in repository.list_contribution_history(old.id)] == [
        "ACTIVE"
    ]
    database.close()


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
    child_path = child.metadata.ready_path
    assert child_path is not None
    child_path.unlink()
    assert received.review is not None
    batch_service.confirm_batch(received.batch.id, received.review.document)

    posting = ExpensePostingService(
        ReviewedBatchProvider(batch_service),
        settings,
        sea_freight_repository=sea_repository,
        sea_freight_service=sea_service,
    )
    bundle = posting._posting_bundle(received.batch.id)

    assert child_path.is_file()
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
