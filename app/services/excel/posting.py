"""Two-phase posting of reviewed expense JSON into the BK workbook."""

from __future__ import annotations

import copy
import hashlib
import json
import re
import unicodedata
from collections import defaultdict
from collections.abc import Callable, Iterable, Mapping, Sequence
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl.utils import get_column_letter

from app.constants import FEE_CODES, RULE_CODES
from app.services.json_codec import JsonCodec, JsonCodecError
from app.services.validation_service import normalize_bl, normalize_container

from .headers import HeaderResolution, HeaderResolutionError, HeaderResolver, normalize_header
from .carrier import (
    BK_CARRIER_HEADER_ALIASES,
    DAILY_MANAGED_CARRIER_GROUPS,
    carrier_group_for_fee,
    carrier_key,
    carrier_text,
    detail_rows_from_actions,
    join_carriers,
    read_bk_detail_rows,
    split_carriers,
    unique_carriers,
    upsert_bk_detail_rows,
)
from .daily_sync import (
    SOURCE_HEADER_ALIASES,
    SYNC_FIELDS,
    TARGET_EXPECTED_COLUMNS,
    DailySyncService,
)
from .models import (
    ConflictType,
    ExcelOperation,
    ExcelRunStatus,
    MonthCandidate,
    PostingConflict,
    PostingItem,
    PostingItemStatus,
    PostingPlan,
    PostingResolution,
    PostingResult,
    PostingSourceGroup,
    ResolutionAction,
    RowCandidate,
    TargetCellKind,
    TargetCellState,
    resolution_map,
)
from .outcomes import posting_outcomes
from .resolvers import MonthSheetService, YearResolver
from .review import (
    CorrectionRequiredError,
    SourceDataChangedError,
    validate_conflict_resolutions,
)
from .workbook import (
    ExcelBackupService,
    ExcelLockService,
    WorkbookGateway,
    ensure_supported_workbook,
)


ProgressCallback = Callable[[str], None] | None

BASE_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "sqt": ("SQT PM", "SQT", "Số thứ tự PM"),
    "closing_date": ("Ngày Đóng", "Ngày đóng hàng"),
    "container": ("Số Container", "Container", "Số cont"),
    "cargo_type": ("Loại hàng", "Tên hàng"),
    "vessel": ("Tên tàu", "Tàu"),
    "recipient": ("Người nhận", "Khách hàng"),
    "notes": ("Ghi chú", "Ghi Chú"),
    "carrier_sea": BK_CARRIER_HEADER_ALIASES["SEA"],
    "carrier_road": BK_CARRIER_HEADER_ALIASES["ROAD"],
    "carrier_nam": BK_CARRIER_HEADER_ALIASES["NAM"],
}

FEE_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "CB": ("Cước biển",),
    "CBDH": ("Cước bộ đóng hàng", "ĐƠN GIÁ"),
    "VTN": ("Cước VTN",),
    "NV": ("Nâng vỏ",),
    "HH": ("Hạ Hàng",),
    "NH": ("Nâng Hàng",),
    "HV": ("Hạ vỏ",),
    "VSDL": ("VS D/O LỆNH", "VS DO LỆNH", "VS D/O"),
    "LC": ("Lưu cont",),
    "QT": ("Quá tải",),
    "LL": ("LÀM LỆNH",),
    "SC": ("SỬA CHỮA",),
    "GH": ("GIA HẠN",),
}

FEE_INVOICE_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "CB": ("Hóa đơn cước biển",),
}

INVOICE_HEADER_NAMES = frozenset(
    normalize_header(value)
    for value in (
        "Hóa đơn",
        "Số hóa đơn",
        "Ngày hóa đơn",
        "Tiền hóa đơn",
        "VAT",
        "Thuế GTGT",
        "Số HĐ",
        "Hóa đơn cước biển",
        "HD",
        "HĐ",
        "hoá đon",
    )
)
INVOICE_NUMBER_HEADER_NAMES = frozenset(
    normalize_header(value)
    for value in (
        "Hóa đơn",
        "Số hóa đơn",
        "Số HĐ",
        "Hóa đơn cước biển",
        "HD",
        "HĐ",
        "hoá đon",
    )
)
UPDATE_HEADER_NAMES = frozenset(
    normalize_header(value) for value in ("Date cập nhật", "Data cập nhật")
)
UPDATE_HEADER_CANONICAL = "Date cập nhật"
UPDATE_NUMBER_FORMAT = "dd/mm/yyyy hh:mm:ss"


class ExpensePostingError(RuntimeError):
    pass


def _stable_id(*parts: Any) -> str:
    payload = json.dumps(parts, ensure_ascii=False, default=str, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:24]


def _invoice_text(value: Any) -> str | None:
    if value in (None, ""):
        return None
    return str(value).strip() or None


def _invoice_key(value: Any) -> str | None:
    text = _invoice_text(value)
    return " ".join(text.casefold().split()) if text is not None else None


def _unique_invoice_values(values: Sequence[Any]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = _invoice_text(value)
        key = _invoice_key(text)
        if text is None or key is None or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def _period_offset(month: int, year: int, offset: int) -> tuple[int, int]:
    absolute = year * 12 + month - 1 - offset
    target_year, zero_based_month = divmod(absolute, 12)
    return zero_based_month + 1, target_year


def _progress(callback: ProgressCallback, message: str) -> None:
    if callback is not None:
        callback(message)


def classify_target_cell(cell: Any, amount: int) -> TargetCellState:
    value = cell.value
    coordinate = cell.coordinate
    data_type = getattr(cell, "data_type", None)
    if data_type == "f" or (isinstance(value, str) and value.startswith("=")):
        kind = TargetCellKind.FORMULA
    elif value in (None, ""):
        kind = TargetCellKind.EMPTY
    elif isinstance(value, bool):
        kind = TargetCellKind.TEXT
    elif isinstance(value, (int, float)):
        if value == amount:
            kind = TargetCellKind.SAME_VALUE
        elif value == 0:
            kind = TargetCellKind.ZERO
        else:
            kind = TargetCellKind.NUMBER
    else:
        kind = TargetCellKind.TEXT
    return TargetCellState(kind, value, coordinate, data_type)


class ExpensePostingService:
    def __init__(
        self,
        provider: Any | None = None,
        settings: Any | None = None,
        *,
        reviewed_batch_provider: Any | None = None,
        bk_path: str | Path | None = None,
        temp_dir: str | Path | None = None,
        backup_dir: str | Path | None = None,
        gateway: WorkbookGateway | None = None,
        lock_service: ExcelLockService | None = None,
        header_resolver: HeaderResolver | None = None,
        month_service: MonthSheetService | None = None,
        year_resolver: YearResolver | None = None,
        run_repository: Any | None = None,
        posting_repository: Any | None = None,
        sea_freight_repository: Any | None = None,
        sea_freight_service: Any | None = None,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.provider = reviewed_batch_provider or provider
        if self.provider is None:
            raise TypeError("ExpensePostingService cần ReviewedBatchProvider.")
        paths = getattr(settings, "paths", None)
        self.bk_path = Path(
            bk_path or getattr(settings, "bk_workbook_path", "")
        )
        system_dir = getattr(paths, "system_dir", Path("Output") / "_system")
        self.temp_dir = Path(
            temp_dir
            or getattr(paths, "excel_temp_dir", system_dir / "Excel" / "Temp")
        )
        self.backup_dir = Path(
            backup_dir
            or getattr(paths, "excel_backup_dir", system_dir / "Excel" / "Backup")
        )
        self.gateway = gateway or WorkbookGateway()
        self.lock_service = lock_service or ExcelLockService()
        self.headers = header_resolver or HeaderResolver()
        self.months = month_service or MonthSheetService()
        self.years = year_resolver or YearResolver()
        self.run_repository = run_repository
        self.posting_repository = posting_repository
        self.sea_freight_repository = sea_freight_repository
        self.sea_freight_service = sea_freight_service
        self.clock = clock or (lambda: datetime.now().astimezone().replace(tzinfo=None))
        # Working copy nằm cạnh BK để atomic replace không bao giờ cross-volume.
        self.backups = ExcelBackupService(self.backup_dir)

    def _ordered_sheet_names(self, names: Iterable[str]) -> list[str]:
        """Sắp sheet tháng theo năm/tháng để history và apply có thứ tự ổn định."""

        def key(name: str) -> tuple[int, int, str]:
            parsed = self.months.parse_target_sheet(name)
            if parsed is None:
                return (9999, 99, name)
            month, year = parsed
            return (year, month, name)

        return sorted({str(name) for name in names}, key=key)

    def analyze(
        self,
        batch_id: int | None = None,
        sheet_name: str | None = None,
        repost_source_indices: Sequence[int] | None = None,
        group_target_sheets: Mapping[str, str] | None = None,
        split_document_ids: Sequence[str] | None = None,
        progress_callback: ProgressCallback = None,
    ) -> PostingPlan:
        target = ensure_supported_workbook(self.bk_path)
        if not target.is_file():
            raise ExpensePostingError(f"Không tìm thấy file BK: {target}")
        bundle = self._posting_bundle(batch_id)
        batch_path = bundle["batch_path"]
        raw = bundle["raw"]
        batch_hash = bundle["bundle_hash"]
        resolved_batch_id = bundle["batch_id"]
        document_rows = bundle["rows"]
        source_members = bundle["source_members"]
        reconciliation_members = bundle["reconciliation_members"]
        source_kind = str(bundle.get("source_kind") or "ASSISTANT")
        already_posted = self._successful_bundle_indices(document_rows)
        repost_indices = set(repost_source_indices or ())
        invalid_reposts = repost_indices.difference(already_posted)
        if invalid_reposts:
            raise ExpensePostingError(
                "Danh sách khoản nhập lại không còn hợp lệ; vui lòng đọc lại."
            )
        repost_selection_done = repost_source_indices is not None
        previously_posted = self._previously_posted_items(
            batch_hash, document_rows, already_posted,
        )
        run_id = self._create_run(batch_path, target)
        try:
            _progress(progress_callback, "Đang đọc cấu trúc file BK…")
            self.lock_service.ensure_readable(target)
            fingerprint = self.gateway.fingerprint(target)
            workbook = self.gateway.load(target, read_only=False)
            try:
                sheet_names = [
                    name
                    for name in workbook.sheetnames
                    if self.months.parse_target_sheet(name) is not None
                ]
                if not sheet_names:
                    raise ExpensePostingError("File BK không có sheet tháng TMM YY.")
                groups = self._group_rows(
                    document_rows,
                    already_posted,
                    repost_indices,
                    set(),
                )
                candidates = [
                    MonthCandidate(
                        month=parsed[0],
                        year=parsed[1],
                        source_sheet="",
                        target_sheet=name,
                    )
                    for name in sheet_names
                    if (parsed := self.months.parse_target_sheet(name)) is not None
                ]
                selected = self._choose_sheet(sheet_name, candidates, groups, None)
                items = self._items_from_groups(groups, repost_indices)
                conflicts: list[PostingConflict] = []
                source_groups: list[PostingSourceGroup] = []
                if source_kind == "BANG_KE":
                    self._ensure_bang_ke_columns(workbook)
                    items, item_conflicts = self._analyze_bang_ke_items(
                        workbook,
                        items,
                        batch_hash=batch_hash,
                    )
                    conflicts.extend(item_conflicts)
                    selected = None
                else:
                    source_groups = self._build_posting_source_groups(
                        document_rows,
                        sheet_names,
                        group_target_sheets=group_target_sheets,
                        legacy_sheet=selected,
                        split_document_ids=set(split_document_ids or ()),
                    )
                    target_by_source_index = {
                        source_index: group.target_sheet
                        for group in source_groups
                        for source_index in group.source_item_indices
                    }
                    for item in items:
                        targets = {
                            target_by_source_index.get(source_index)
                            for source_index in item.source_indices
                        }
                        if len(targets) != 1:
                            raise ExpensePostingError(
                                "Một khoản chi đang thuộc nhiều nhóm phân bổ sheet."
                            )
                        item.sheet_name = next(iter(targets))
                    items, conflicts = self._analyze_assigned_items(
                        workbook, items, batch_hash=batch_hash
                    )
                    assigned = {
                        group.target_sheet
                        for group in source_groups
                        if group.target_sheet is not None
                    }
                    selected = next(iter(assigned)) if len(assigned) == 1 else None
            finally:
                workbook.close()

            plan = PostingPlan(
                batch_id=resolved_batch_id,
                batch_path=batch_path,
                batch_hash=batch_hash,
                target_path=target,
                target_fingerprint=fingerprint,
                items=items,
                conflicts=conflicts,
                sheet_candidates=candidates,
                selected_sheet=selected,
                source_item_count=len(document_rows),
                already_posted_indices=already_posted,
                previously_posted_items=previously_posted,
                repost_source_indices=repost_indices,
                repost_selection_done=repost_selection_done,
                run_id=run_id,
                source_members=source_members,
                reconciliation_members=reconciliation_members,
                original_source_count=int(bundle["original_source_count"]),
                reconciliation_source_count=int(bundle["reconciliation_source_count"]),
                confirmation_required=True,
                source_kind=source_kind,
                source_groups=source_groups,
                split_document_ids=set(split_document_ids or ()),
                target_sheets={
                    item.sheet_name for item in items if item.sheet_name is not None
                },
            )
            status = (
                ExcelRunStatus.WAITING_USER
                if plan.requires_user_input
                else ExcelRunStatus.NO_CHANGES
                if not items
                else ExcelRunStatus.ANALYZING
            )
            self._update_run(
                run_id,
                status=status,
                source_fingerprint={
                    "size": sum(int(member["size"]) for member in source_members),
                    "mtime_ns": batch_path.stat().st_mtime_ns,
                    "sha256": batch_hash,
                },
                target_fingerprint_before=fingerprint,
                total_items=len(document_rows),
                conflict_count=len(conflicts),
                sheet_name=", ".join(self._ordered_sheet_names(plan.target_sheets)) or selected,
            )
            return plan
        except Exception as exc:
            self._finish_failed(run_id, exc)
            raise

    def apply(
        self,
        plan: PostingPlan,
        resolutions: Mapping[str, Any] | Sequence[PostingResolution] | None = None,
        progress_callback: ProgressCallback = None,
    ) -> PostingResult:
        resolved = resolution_map(resolutions)
        target_value_conflicts = [
            conflict
            for conflict in plan.conflicts
            if conflict.conflict_type
            in {
                ConflictType.TARGET_CELL_OCCUPIED,
                ConflictType.TARGET_CELL_FORMULA,
                ConflictType.TARGET_CELL_TEXT,
            }
        ]
        issues = validate_conflict_resolutions(target_value_conflicts, resolved)
        if issues:
            raise CorrectionRequiredError(issues)
        if self._has_selector_resolution(plan, resolved):
            # SELECT_SHEET/SELECT_ROW/SELECT_FEE changes the target cell itself. Calling
            # apply directly would bypass the mandatory second analysis of
            # that cell and could silently keep an occupied/formula/text value.
            raise ExpensePostingError(
                "Lựa chọn sheet, dòng hoặc mã phí phải được phân tích lại bằng "
                "ExpensePostingService.refine() trước khi apply."
            )
        if plan.source_kind == "BANG_KE":
            return self._apply_bang_ke(
                plan,
                resolved,
                progress_callback=progress_callback,
            )
        if plan.source_groups:
            return self._apply_multi_sheet(
                plan,
                resolved,
                progress_callback=progress_callback,
            )
        try:
            self._check_batch_resolution(plan, resolved)
            selected_sheet = self._selected_sheet(plan, resolved)
            if selected_sheet is None and plan.items:
                raise ExpensePostingError(
                    "Chưa chọn sheet tháng cần nhập khoản chi."
                )
            self.gateway.assert_unchanged(
                plan.target_path, plan.target_fingerprint, label="File BK"
            )
            self._assert_source_members_unchanged(plan)

            workbook = self.gateway.load(plan.target_path, read_only=False)
            try:
                if (
                    selected_sheet is not None
                    and selected_sheet not in workbook.sheetnames
                ):
                    raise ExpensePostingError(
                        f"Không tìm thấy sheet {selected_sheet}."
                    )
                base_items = copy.deepcopy(plan.items)
                if (
                    selected_sheet is not None
                    and plan.selected_sheet != selected_sheet
                ):
                    base_items, dynamic_conflicts = self._analyze_items(
                        workbook,
                        selected_sheet,
                        base_items,
                        batch_hash=plan.batch_hash,
                    )
                else:
                    dynamic_conflicts = []
                actions, history = self._resolve_apply_actions(
                    workbook[selected_sheet] if selected_sheet else None,
                    base_items,
                    [*plan.conflicts, *dynamic_conflicts],
                    resolved,
                    batch_hash=plan.batch_hash,
                )
            finally:
                workbook.close()
        except Exception as exc:
            self._finish_failed(plan.run_id, exc)
            raise

        write_actions = [
            action
            for action in actions
            if action.get("amount_write")
            or action.get("invoice_write")
            or action.get("carrier_write")
        ]
        detail_actions = [
            action
            for action in actions
            if action.get("carrier_group") in {"HP", "NAM"}
            and action.get("status")
            in {PostingItemStatus.POSTED, PostingItemStatus.ALREADY_EXISTS}
        ]
        already_actions = [
            action
            for action in actions
            if action["status"] is PostingItemStatus.ALREADY_EXISTS
        ]
        skipped_actions = [
            action
            for action in actions
            if action["status"]
            in {
                PostingItemStatus.USER_SKIPPED,
                PostingItemStatus.NOT_MATCHED,
                PostingItemStatus.UNRESOLVED,
            }
        ]
        working_path: Path | None = None
        backup_path: Path | None = None
        after = plan.target_fingerprint
        update_timestamp: datetime | None = None
        update_column: int | None = None
        workbook_replaced = False
        self._update_run(plan.run_id, status=ExcelRunStatus.APPLYING)
        try:
            if write_actions or detail_actions:
                with self.lock_service.acquire(plan.target_path):
                    pass
                self.gateway.assert_unchanged(
                    plan.target_path,
                    plan.target_fingerprint,
                    label="File BK",
                )
                working_path = self.backups.create_working_copy(
                    plan.target_path, run_id=plan.run_id
                )
                write_book = self.gateway.load(working_path, read_only=False)
                try:
                    worksheet = write_book[selected_sheet]
                    carried_plan_rows = self._write_carried_plan_rows(
                        worksheet, write_actions
                    )
                    base = self._resolve_base_headers(worksheet)
                    fee_columns = self._resolve_fee_columns(worksheet, base)
                    invoice_columns = self._resolve_invoice_columns(
                        worksheet, base, fee_columns
                    )
                    carrier_columns = self._resolve_carrier_columns(base)
                    update_column = self._ensure_update_column(worksheet)
                    update_timestamp = self.clock().replace(
                        tzinfo=None,
                        microsecond=0,
                    )
                    updated_rows: set[int] = set()
                    for action in write_actions:
                        column = int(action["target_column"])
                        fee = str(action["fee_selected"])
                        if fee_columns.get(fee) != column:
                            raise ExpensePostingError(
                                f"Cột phí {fee} không còn khớp header."
                            )
                        if action.get("amount_write"):
                            worksheet.cell(
                                int(action["target_row"]), column
                            ).value = action["value_after"]
                        if action.get("invoice_write"):
                            invoice_column = int(action["invoice_target_column"])
                            if invoice_columns.get(fee) != invoice_column:
                                raise ExpensePostingError(
                                    f"Cột Số HĐ của phí {fee} không còn khớp header."
                                )
                            worksheet.cell(
                                int(action["target_row"]), invoice_column
                            ).value = action["invoice_value_after"]
                        if action.get("carrier_write"):
                            carrier_group = str(action["carrier_group"])
                            carrier_column = int(action["carrier_target_column"])
                            if carrier_columns.get(carrier_group) != carrier_column:
                                raise ExpensePostingError(
                                    f"Cột bên vận tải {carrier_group} không còn khớp header."
                                )
                            worksheet.cell(
                                int(action["target_row"]), carrier_column
                            ).value = action["carrier_value_after"]
                        updated_rows.add(int(action["target_row"]))
                    for target_row in updated_rows:
                        update_cell = worksheet.cell(target_row, update_column)
                        update_cell.value = update_timestamp
                        update_cell.number_format = UPDATE_NUMBER_FORMAT
                    if carried_plan_rows:
                        from .payment_sync import (
                            find_summary_start,
                            refresh_bk_summary_formulas,
                        )

                        if find_summary_start(worksheet) is not None:
                            refresh_bk_summary_formulas(worksheet)
                    detail_rows = detail_rows_from_actions(
                        detail_actions,
                        batch_id=plan.batch_id,
                        batch_hash=plan.batch_hash,
                        sheet_name=str(selected_sheet),
                        updated_at=update_timestamp,
                    )
                    upsert_bk_detail_rows(write_book, detail_rows)
                    self.gateway.save(write_book, working_path)
                finally:
                    write_book.close()
                self._verify_posting(
                    working_path,
                    selected_sheet,
                    write_actions,
                    detail_rows=detail_rows,
                    update_column=update_column,
                    update_timestamp=update_timestamp,
                )
                backup_path = self.backups.create_backup(plan.target_path)
                after = self.gateway.atomic_replace(
                    working_path,
                    plan.target_path,
                    expected=plan.target_fingerprint,
                )
                working_path = None
                workbook_replaced = True
                self._record_carry_forwards(plan, write_actions)

            self._record_history(plan, history)
            self._mark_completed_reconciliations(plan)
            status = (
                ExcelRunStatus.SUCCEEDED
                if write_actions or detail_actions
                else ExcelRunStatus.NO_CHANGES
            )
            posted_count = sum(
                len(action["source_indices"]) for action in write_actions
            )
            already_count = sum(
                len(action["source_indices"]) for action in already_actions
            )
            skipped_count = sum(
                len(action["source_indices"]) for action in skipped_actions
            )
            amount_written_count = sum(
                bool(action.get("amount_write")) for action in write_actions
            )
            invoice_written_count = sum(
                bool(action.get("invoice_write")) for action in write_actions
            )
            carrier_written_count = sum(
                bool(action.get("carrier_write")) for action in write_actions
            )
            carrier_appended_count = sum(
                bool(action.get("carrier_write"))
                and action.get("carrier_action") is ResolutionAction.APPEND_CARRIER
                for action in write_actions
            )
            carrier_kept_count = sum(
                action.get("carrier_action") is ResolutionAction.KEEP_EXISTING
                for action in actions
                if action.get("carrier_group") is not None
            )
            unmapped_count = sum(
                len(action.get("source_indices", ()))
                for action in detail_actions
                if not action.get("carrier_effective")
            )
            result = PostingResult(
                status=status,
                target_path=plan.target_path,
                sheet_name=selected_sheet,
                posted_source_items=posted_count,
                written_cells=amount_written_count,
                invoice_written_cells=invoice_written_count,
                carrier_written_cells=carrier_written_count,
                carrier_appended_cells=carrier_appended_count,
                carrier_kept_cells=carrier_kept_count,
                unmapped_carrier_items=unmapped_count,
                skipped_source_items=skipped_count,
                already_existing_items=already_count,
                conflict_count=len(plan.conflicts),
                backup_path=backup_path,
                fingerprint_before=plan.target_fingerprint,
                fingerprint_after=after,
                run_id=plan.run_id,
                message=(
                    f"Đã nhập {posted_count} khoản vào {amount_written_count} ô tiền; "
                    f"cập nhật {invoice_written_count} ô Số HĐ và "
                    f"{carrier_written_count} ô bên vận tải. "
                    f"VT: ghi đè {carrier_written_count - carrier_appended_count}, "
                    f"ghi thêm {carrier_appended_count}, giữ nguyên {carrier_kept_count}, "
                    f"chưa xác định {unmapped_count}."
                    if write_actions or detail_actions
                    else "Không có ô cần ghi."
                ),
                item_outcomes=posting_outcomes(actions),
            )
            self._finish_result(result)
            return result
        except Exception as exc:
            if workbook_replaced and backup_path is not None:
                try:
                    rollback_path = self.backups.create_working_copy(
                        backup_path,
                        run_id=f"{plan.run_id}-rollback",
                    )
                    try:
                        self.gateway.atomic_replace(
                            rollback_path,
                            plan.target_path,
                            expected=after,
                        )
                    finally:
                        rollback_path.unlink(missing_ok=True)
                except Exception as rollback_error:
                    exc.add_note(f"Khôi phục BK thất bại: {rollback_error}")
            self._finish_failed(plan.run_id, exc)
            raise
        finally:
            if working_path is not None and working_path.exists():
                working_path.unlink()

    def refine(
        self,
        plan: PostingPlan,
        resolutions: Mapping[str, Any] | Sequence[PostingResolution] | None,
        progress_callback: ProgressCallback = None,
    ) -> PostingPlan:
        try:
            if plan.source_groups and plan.source_kind != "BANG_KE":
                return self._refine_multi_sheet(
                    plan,
                    resolutions,
                    progress_callback=progress_callback,
                )
            return self._refine_plan(
                plan,
                resolutions,
                progress_callback=progress_callback,
            )
        except Exception as exc:
            self._finish_failed(plan.run_id, exc)
            raise

    def _refine_plan(
        self,
        plan: PostingPlan,
        resolutions: Mapping[str, Any] | Sequence[PostingResolution] | None,
        progress_callback: ProgressCallback = None,
    ) -> PostingPlan:
        """Áp dụng lựa chọn dòng/mã phí rồi phân tích lại ô đích.

        Pha này chỉ đọc workbook. Nó bảo đảm một lựa chọn ``SELECT_ROW`` hoặc
        ``SELECT_FEE`` không thể đi thẳng tới apply trước khi các xung đột
        number/formula/text phát sinh được đưa lại cho dialog tổng hợp.
        """

        resolved = resolution_map(resolutions)
        applied_conflict_ids: set[str] = set()
        self._check_batch_resolution(plan, resolved)
        if plan.source_kind == "BANG_KE":
            return self._refine_bang_ke(
                plan,
                resolved,
                progress_callback=progress_callback,
            )
        selected_sheet = self._selected_sheet(plan, resolved)
        if selected_sheet is None:
            raise ExpensePostingError("Chưa chọn sheet tháng cần nhập khoản chi.")
        self.gateway.assert_unchanged(
            plan.target_path, plan.target_fingerprint, label="File BK"
        )
        self._assert_source_members_unchanged(plan)

        items = copy.deepcopy(plan.items)
        for conflict in plan.conflicts:
            value = resolved.get(conflict.conflict_id)
            if conflict.conflict_type in {
                ConflictType.MULTIPLE_SOURCE_CARRIERS,
                ConflictType.CARRIER_VALUE_CONFLICT,
                ConflictType.CARRIER_COLUMN_MISSING,
            }:
                if value is None:
                    continue
                action = self._action(value)
                if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                indexes = tuple(
                    int(index)
                    for index in conflict.details.get(
                        "item_indexes",
                        (() if conflict.item_index is None else (conflict.item_index,)),
                    )
                )
                selected_carrier = None
                if action is ResolutionAction.SELECT_CARRIER:
                    selected_carrier = self._resolution_attr(value, "selected_carrier")
                    valid = {
                        carrier_key(candidate)
                        for candidate in conflict.details.get("carrier_candidates", ())
                    }
                    if carrier_key(selected_carrier) not in valid:
                        raise ExpensePostingError("Bên vận tải được chọn không hợp lệ.")
                ambiguous_invoices = {
                    _invoice_key(invoice)
                    for invoice in conflict.details.get(
                        "ambiguous_invoice_carriers", {}
                    )
                }
                for index in indexes:
                    belongs_to_ambiguous_invoice = bool(
                        ambiguous_invoices.intersection(
                            _invoice_key(invoice)
                            for invoice in items[index].invoice_candidates
                        )
                    )
                    items[index].selected_carrier = (
                        str(selected_carrier)
                        if selected_carrier is not None
                        and (
                            not ambiguous_invoices
                            or belongs_to_ambiguous_invoice
                        )
                        else None
                    )
                    items[index].carrier_action = (
                        ResolutionAction.OVERWRITE
                        if action is ResolutionAction.SELECT_CARRIER
                        else action
                    )
                applied_conflict_ids.add(conflict.conflict_id)
                continue
            if conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL:
                if value is None:
                    continue
                action = self._action(value)
                if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                if action is not ResolutionAction.SELECT_SOURCE_ITEM:
                    raise ExpensePostingError("Chưa chọn dòng JSON cần ghi vào ô phí.")
                selected_source_index = self._resolution_attr(
                    value, "selected_source_item_index"
                )
                valid_indexes = {
                    int(option["source_item_index"])
                    for option in conflict.details.get("source_item_options", ())
                }
                if (
                    selected_source_index is None
                    or int(selected_source_index) not in valid_indexes
                ):
                    raise ExpensePostingError("Dòng JSON được chọn không hợp lệ.")
                selected_source_index = int(selected_source_index)
                for grouped_index in conflict.details.get("item_indexes", ()):
                    grouped_item = items[int(grouped_index)]
                    if selected_source_index in grouped_item.source_indices:
                        grouped_item.status = PostingItemStatus.PLANNED
                        grouped_item.action = None
                    else:
                        grouped_item.status = PostingItemStatus.USER_SKIPPED
                        grouped_item.action = ResolutionAction.SKIP
                applied_conflict_ids.add(conflict.conflict_id)
                continue
            if conflict.item_index is None:
                continue
            if value is None:
                continue
            action = self._action(value)
            item = items[conflict.item_index]
            invoice_conflict = conflict.conflict_type in {
                ConflictType.INVOICE_COLUMN_MISSING,
                ConflictType.INVOICE_VALUE_CONFLICT,
                ConflictType.MULTIPLE_SOURCE_INVOICES,
            }
            if invoice_conflict:
                if action is ResolutionAction.SELECT_INVOICE:
                    selected_invoice = self._resolution_attr(
                        value, "selected_invoice"
                    )
                    if _invoice_key(selected_invoice) not in {
                        _invoice_key(candidate)
                        for candidate in item.invoice_candidates
                    }:
                        raise ExpensePostingError("Số HĐ được chọn không hợp lệ.")
                    item.selected_invoice = str(selected_invoice)
                    item.invoice_action = None
                elif action in {
                    ResolutionAction.SKIP_INVOICE,
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.OVERWRITE,
                }:
                    item.invoice_action = action
                elif action in {
                    ResolutionAction.CANCEL,
                    ResolutionAction.CANCEL_ALL,
                }:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                applied_conflict_ids.add(conflict.conflict_id)
                continue
            if action is ResolutionAction.SELECT_FEE:
                selected_fee = self._resolution_attr(value, "selected_fee")
                if selected_fee not in FEE_HEADER_ALIASES:
                    raise ExpensePostingError("Mã phí được chọn không hợp lệ.")
                item.selected_fee = str(selected_fee)
                item.selected_invoice = None
                item.invoice_action = None
            elif action is ResolutionAction.SELECT_ROW:
                selected_row = self._resolution_attr(value, "selected_row")
                selected_source_sheet = self._resolution_attr(
                    value, "selected_source_sheet"
                )
                if selected_row is None:
                    raise ExpensePostingError("Chưa chọn dòng BK.")
                item.selected_source_row = int(selected_row)
                item.selected_source_sheet = str(
                    selected_source_sheet or selected_sheet
                )
                item.target_row = None
                item.invoice_action = None
            elif action in {
                ResolutionAction.OVERWRITE,
            }:
                item.action = action
            elif action in {
                ResolutionAction.SKIP,
                ResolutionAction.KEEP_EXISTING,
                ResolutionAction.KEEP_FORMULA,
            }:
                item.action = action
                item.status = PostingItemStatus.USER_SKIPPED
            elif action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
            applied_conflict_ids.add(conflict.conflict_id)

        _progress(progress_callback, "Đang kiểm tra lại dòng và ô đã chọn…")
        workbook = self.gateway.load(plan.target_path, read_only=False)
        try:
            if selected_sheet not in workbook.sheetnames:
                raise ExpensePostingError(f"Không tìm thấy sheet {selected_sheet}.")
            items, conflicts = self._analyze_items(
                workbook,
                selected_sheet,
                items,
                batch_hash=plan.batch_hash,
            )
        finally:
            workbook.close()
        # Chỉ loại conflict mà quyết định đã thực sự được tiêu thụ trong vòng
        # hiện tại. Không lọc bằng toàn bộ resolution ledger của phiên.
        conflicts = [
            conflict
            for conflict in conflicts
            if conflict.conflict_id not in applied_conflict_ids
        ]
        refined = PostingPlan(
            batch_id=plan.batch_id,
            batch_path=plan.batch_path,
            batch_hash=plan.batch_hash,
            target_path=plan.target_path,
            target_fingerprint=plan.target_fingerprint,
            items=items,
            conflicts=conflicts,
            sheet_candidates=plan.sheet_candidates,
            selected_sheet=selected_sheet,
            source_item_count=plan.source_item_count,
            already_posted_indices=plan.already_posted_indices,
            previously_posted_items=plan.previously_posted_items,
            repost_source_indices=plan.repost_source_indices,
            repost_selection_done=plan.repost_selection_done,
            run_id=plan.run_id,
            source_members=plan.source_members,
            reconciliation_members=plan.reconciliation_members,
            original_source_count=plan.original_source_count,
            reconciliation_source_count=plan.reconciliation_source_count,
            confirmation_required=plan.confirmation_required,
            confirmation_done=plan.confirmation_done,
        )
        self._update_run(
            plan.run_id,
            status=(
                ExcelRunStatus.WAITING_USER
                if conflicts
                else ExcelRunStatus.ANALYZING
            ),
            sheet_name=selected_sheet,
            conflict_count=len(conflicts),
        )
        return refined

    def _refine_multi_sheet(
        self,
        plan: PostingPlan,
        resolutions: Mapping[str, Any] | Sequence[PostingResolution] | None,
        *,
        progress_callback: ProgressCallback = None,
    ) -> PostingPlan:
        resolved = resolution_map(resolutions)
        self._check_batch_resolution(plan, resolved)
        self.gateway.assert_unchanged(
            plan.target_path, plan.target_fingerprint, label="File BK"
        )
        self._assert_source_members_unchanged(plan)
        combined_items: list[PostingItem] = []
        combined_conflicts: list[PostingConflict] = []
        ordered_sheets = self._ordered_sheet_names(plan.target_sheets)
        for position, sheet_name in enumerate(ordered_sheets, start=1):
            _progress(
                progress_callback,
                f"Đang phân tích lại {sheet_name} ({position}/{len(ordered_sheets)})…",
            )
            indexes = [
                index
                for index, item in enumerate(plan.items)
                if item.sheet_name == sheet_name
            ]
            local = copy.copy(plan)
            local.items = [copy.deepcopy(plan.items[index]) for index in indexes]
            local.conflicts = self._localized_conflicts(plan.conflicts, indexes)
            local.selected_sheet = sheet_name
            local.target_sheets = {sheet_name}
            local.source_groups = []
            local.confirmation_required = False
            local.confirmation_done = True
            refined = self._refine_plan(
                local,
                resolved,
                progress_callback=progress_callback,
            )
            offset = len(combined_items)
            for conflict in refined.conflicts:
                if conflict.item_index is not None:
                    conflict.item_index += offset
                if "item_indexes" in conflict.details:
                    conflict.details["item_indexes"] = [
                        int(index) + offset
                        for index in conflict.details["item_indexes"]
                    ]
            combined_items.extend(refined.items)
            combined_conflicts.extend(refined.conflicts)
        refined_plan = copy.copy(plan)
        refined_plan.items = combined_items
        refined_plan.conflicts = combined_conflicts
        refined_plan.selected_sheet = (
            next(iter(plan.target_sheets)) if len(plan.target_sheets) == 1 else None
        )
        refined_plan.target_sheets = {
            item.sheet_name for item in combined_items if item.sheet_name is not None
        }
        self._update_run(
            plan.run_id,
            status=(
                ExcelRunStatus.WAITING_USER
                if combined_conflicts
                else ExcelRunStatus.ANALYZING
            ),
            sheet_name=", ".join(sorted(refined_plan.target_sheets)) or None,
            conflict_count=len(combined_conflicts),
        )
        return refined_plan

    @staticmethod
    def _bang_ke_item_key(item: PostingItem) -> tuple[str, str]:
        return (
            item.container or "",
            ExpensePostingService._vessel_voyage_key(
                item.vessel_name,
                item.voyage_no,
                item.vessel_voyage_raw
                if not (item.vessel_name or item.voyage_no)
                else None,
            ),
        )

    def _refine_bang_ke(
        self,
        plan: PostingPlan,
        resolved: Mapping[str, Any],
        *,
        progress_callback: ProgressCallback = None,
    ) -> PostingPlan:
        self.gateway.assert_unchanged(
            plan.target_path, plan.target_fingerprint, label="File BK"
        )
        self._assert_source_members_unchanged(plan)
        items = copy.deepcopy(plan.items)
        selected_by_key: dict[tuple[str, str], tuple[str, int]] = {}
        applied_conflict_ids: set[str] = set()
        for conflict in plan.conflicts:
            value = resolved.get(conflict.conflict_id)
            if value is None:
                continue
            action = self._action(value)
            if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                raise ExpensePostingError("Người dùng đã hủy nhập bảng kê.")
            if conflict.item_index is None:
                continue
            item = items[conflict.item_index]
            if action is ResolutionAction.SELECT_ROW:
                selected_row = self._resolution_attr(value, "selected_row")
                selected_sheet = self._resolution_attr(
                    value, "selected_source_sheet"
                ) or self._resolution_attr(value, "selected_sheet")
                if selected_row is None or not selected_sheet:
                    raise ExpensePostingError("Chưa chọn đủ sheet và dòng BK.")
                valid = {
                    (candidate.source_sheet, candidate.row)
                    for candidate in conflict.row_candidates
                }
                choice = (str(selected_sheet), int(selected_row))
                if choice not in valid:
                    raise ExpensePostingError("Dòng BK được chọn không hợp lệ.")
                selected_by_key[self._bang_ke_item_key(item)] = choice
            elif action is ResolutionAction.SKIP:
                item.action = ResolutionAction.SKIP
                item.status = PostingItemStatus.USER_SKIPPED
            elif action in {
                ResolutionAction.ADD,
                ResolutionAction.OVERWRITE,
                ResolutionAction.KEEP_EXISTING,
                ResolutionAction.KEEP_FORMULA,
            }:
                item.action = action
                if action in {
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.KEEP_FORMULA,
                }:
                    item.status = PostingItemStatus.USER_SKIPPED
            applied_conflict_ids.add(conflict.conflict_id)

        for item in items:
            choice = selected_by_key.get(self._bang_ke_item_key(item))
            if choice is not None:
                item.selected_source_sheet, item.selected_source_row = choice
                item.status = PostingItemStatus.PLANNED
                item.action = None

        _progress(progress_callback, "Đang kiểm tra lại các dòng BK đã chọn…")
        workbook = self.gateway.load(plan.target_path, read_only=False)
        try:
            self._ensure_bang_ke_columns(workbook)
            items, conflicts = self._analyze_bang_ke_items(
                workbook, items, batch_hash=plan.batch_hash
            )
        finally:
            workbook.close()
        conflicts = [
            conflict
            for conflict in conflicts
            if conflict.conflict_id not in applied_conflict_ids
        ]
        refined = PostingPlan(
            batch_id=plan.batch_id,
            batch_path=plan.batch_path,
            batch_hash=plan.batch_hash,
            target_path=plan.target_path,
            target_fingerprint=plan.target_fingerprint,
            items=items,
            conflicts=conflicts,
            sheet_candidates=plan.sheet_candidates,
            selected_sheet=None,
            source_item_count=plan.source_item_count,
            already_posted_indices=plan.already_posted_indices,
            previously_posted_items=plan.previously_posted_items,
            repost_source_indices=plan.repost_source_indices,
            repost_selection_done=plan.repost_selection_done,
            run_id=plan.run_id,
            source_members=plan.source_members,
            reconciliation_members=plan.reconciliation_members,
            original_source_count=plan.original_source_count,
            reconciliation_source_count=plan.reconciliation_source_count,
            confirmation_required=plan.confirmation_required,
            confirmation_done=plan.confirmation_done,
            source_kind="BANG_KE",
            target_sheets={
                item.sheet_name for item in items if item.sheet_name is not None
            },
        )
        self._update_run(
            plan.run_id,
            status=(
                ExcelRunStatus.WAITING_USER
                if conflicts
                else ExcelRunStatus.ANALYZING
            ),
            sheet_name=", ".join(sorted(refined.target_sheets)) or None,
            conflict_count=len(conflicts),
        )
        return refined

    def _apply_bang_ke(
        self,
        plan: PostingPlan,
        resolved: Mapping[str, Any],
        *,
        progress_callback: ProgressCallback = None,
    ) -> PostingResult:
        self._check_batch_resolution(plan, resolved)
        if plan.conflicts:
            raise ExpensePostingError(
                "Batch bảng kê vẫn còn xung đột; hãy xử lý và phân tích lại trước khi ghi."
            )
        self.gateway.assert_unchanged(
            plan.target_path, plan.target_fingerprint, label="File BK"
        )
        self._assert_source_members_unchanged(plan)
        actions: list[dict[str, Any]] = []
        history: list[dict[str, Any]] = []
        workbook = self.gateway.load(plan.target_path, read_only=False)
        try:
            self._ensure_bang_ke_columns(workbook)
            for item in plan.items:
                status = item.status
                before = item.current_value
                after = before
                action = item.action or ResolutionAction.SKIP
                amount_write = False
                if status is PostingItemStatus.PLANNED:
                    if item.sheet_name is None or item.target_row is None:
                        status = PostingItemStatus.NOT_MATCHED
                    elif item.target_column is None:
                        status = PostingItemStatus.UNRESOLVED
                    else:
                        cell = workbook[item.sheet_name].cell(
                            item.target_row, item.target_column
                        )
                        before = cell.value
                        state = classify_target_cell(cell, item.amount)
                        if item.amount < 0:
                            if action is not ResolutionAction.ADD:
                                raise ExpensePostingError(
                                    "Khoản điều chỉnh giảm chưa được xác nhận bằng thao tác ADD."
                                )
                            if state.kind not in {
                                TargetCellKind.NUMBER,
                                TargetCellKind.ZERO,
                                TargetCellKind.SAME_VALUE,
                            }:
                                raise ExpensePostingError(
                                    f"Ô {item.sheet_name}!{cell.coordinate} không còn là số."
                                )
                            after = before + item.amount
                            amount_write = True
                            status = PostingItemStatus.POSTED
                        elif state.kind is TargetCellKind.SAME_VALUE and not item.force_repost:
                            action = ResolutionAction.KEEP_EXISTING
                            status = PostingItemStatus.ALREADY_EXISTS
                        elif state.kind in {TargetCellKind.EMPTY, TargetCellKind.ZERO} or (
                            state.kind is TargetCellKind.SAME_VALUE and item.force_repost
                        ):
                            action = ResolutionAction.OVERWRITE
                            after = item.amount
                            amount_write = True
                            status = PostingItemStatus.POSTED
                        elif action is ResolutionAction.OVERWRITE:
                            after = item.amount
                            amount_write = True
                            status = PostingItemStatus.POSTED
                        else:
                            status = PostingItemStatus.USER_SKIPPED
                record = self._action_record(
                    item,
                    item.selected_fee,
                    item.target_row,
                    item.target_column,
                    before,
                    after,
                    action,
                    status,
                )
                record.update(
                    {
                        "amount_write": amount_write,
                        "invoice_write": False,
                        "carrier_write": False,
                    }
                )
                actions.append(record)
                history.extend(self._history_rows(record, item))
        finally:
            workbook.close()

        write_actions = [action for action in actions if action["amount_write"]]
        working_path: Path | None = None
        backup_path: Path | None = None
        after_fingerprint = plan.target_fingerprint
        replaced = False
        self._update_run(plan.run_id, status=ExcelRunStatus.APPLYING)
        try:
            with self.lock_service.acquire(plan.target_path):
                pass
            self.gateway.assert_unchanged(
                plan.target_path, plan.target_fingerprint, label="File BK"
            )
            working_path = self.backups.create_working_copy(
                plan.target_path, run_id=plan.run_id
            )
            write_book = self.gateway.load(working_path, read_only=False)
            update_meta: dict[str, tuple[int, datetime]] = {}
            try:
                self._ensure_bang_ke_columns(write_book)
                grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
                for action in write_actions:
                    grouped[str(action["sheet_name"])].append(action)
                for sheet_name, sheet_actions in grouped.items():
                    worksheet = write_book[sheet_name]
                    base = self._resolve_base_headers(worksheet)
                    fee_columns = self._resolve_fee_columns(worksheet, base)
                    update_column = self._ensure_update_column(worksheet)
                    timestamp = self.clock().replace(tzinfo=None, microsecond=0)
                    update_meta[sheet_name] = (update_column, timestamp)
                    updated_rows: set[int] = set()
                    for action in sheet_actions:
                        fee = str(action["fee_selected"])
                        column = int(action["target_column"])
                        if fee_columns.get(fee) != column:
                            raise ExpensePostingError(
                                f"Cột phí {fee} trên {sheet_name} không còn khớp header."
                            )
                        row = int(action["target_row"])
                        worksheet.cell(row, column).value = action["value_after"]
                        updated_rows.add(row)
                    for row in updated_rows:
                        cell = worksheet.cell(row, update_column)
                        cell.value = timestamp
                        cell.number_format = UPDATE_NUMBER_FORMAT
                self.gateway.save(write_book, working_path)
            finally:
                write_book.close()

            for sheet_name, sheet_actions in grouped.items():
                update_column, timestamp = update_meta[sheet_name]
                self._verify_posting(
                    working_path,
                    sheet_name,
                    sheet_actions,
                    update_column=update_column,
                    update_timestamp=timestamp,
                )
            backup_path = self.backups.create_backup(plan.target_path)
            after_fingerprint = self.gateway.atomic_replace(
                working_path,
                plan.target_path,
                expected=plan.target_fingerprint,
            )
            working_path = None
            replaced = True
            self._record_history(plan, history)
            posted = sum(len(action["source_indices"]) for action in write_actions)
            skipped = sum(
                len(action["source_indices"])
                for action in actions
                if action["status"]
                in {
                    PostingItemStatus.USER_SKIPPED,
                    PostingItemStatus.NOT_MATCHED,
                    PostingItemStatus.UNRESOLVED,
                }
            )
            already = sum(
                len(action["source_indices"])
                for action in actions
                if action["status"] is PostingItemStatus.ALREADY_EXISTS
            )
            target_sheets = tuple(
                sorted({str(action["sheet_name"]) for action in write_actions})
            )
            result = PostingResult(
                status=(
                    ExcelRunStatus.SUCCEEDED
                    if write_actions
                    else ExcelRunStatus.NO_CHANGES
                ),
                target_path=plan.target_path,
                sheet_name=", ".join(target_sheets) or None,
                target_sheets=target_sheets,
                posted_source_items=posted,
                written_cells=len(write_actions),
                skipped_source_items=skipped,
                already_existing_items=already,
                conflict_count=0,
                backup_path=backup_path,
                fingerprint_before=plan.target_fingerprint,
                fingerprint_after=after_fingerprint,
                run_id=plan.run_id,
                message=(
                    f"Đã nhập {posted} khoản bảng kê vào {len(target_sheets)} sheet."
                    if write_actions
                    else "Không có khoản bảng kê cần ghi; cấu trúc cột đã được chuẩn hóa."
                ),
                item_outcomes=posting_outcomes(actions),
            )
            self._finish_result(result)
            return result
        except Exception as exc:
            if replaced and backup_path is not None:
                try:
                    rollback = self.backups.create_working_copy(
                        backup_path, run_id=f"{plan.run_id}-rollback"
                    )
                    try:
                        self.gateway.atomic_replace(
                            rollback,
                            plan.target_path,
                            expected=after_fingerprint,
                        )
                    finally:
                        rollback.unlink(missing_ok=True)
                except Exception as rollback_error:
                    exc.add_note(f"Khôi phục BK thất bại: {rollback_error}")
            self._finish_failed(plan.run_id, exc)
            raise
        finally:
            if working_path is not None and working_path.exists():
                working_path.unlink()

    @staticmethod
    def _localized_conflicts(
        conflicts: Sequence[PostingConflict], indexes: Sequence[int]
    ) -> list[PostingConflict]:
        index_map = {global_index: local_index for local_index, global_index in enumerate(indexes)}
        result: list[PostingConflict] = []
        for conflict in conflicts:
            if conflict.item_index is None or conflict.item_index not in index_map:
                continue
            item = copy.deepcopy(conflict)
            item.item_index = index_map[int(conflict.item_index)]
            if "item_indexes" in item.details:
                item.details["item_indexes"] = [
                    index_map[int(index)]
                    for index in item.details["item_indexes"]
                    if int(index) in index_map
                ]
            result.append(item)
        return result

    def _apply_multi_sheet(
        self,
        plan: PostingPlan,
        resolved: Mapping[str, Any],
        *,
        progress_callback: ProgressCallback = None,
    ) -> PostingResult:
        self._check_batch_resolution(plan, resolved)
        if any(item.sheet_name is None for item in plan.items):
            raise ExpensePostingError("Chưa phân bổ đủ sheet cho mọi khoản chi.")
        self.gateway.assert_unchanged(
            plan.target_path, plan.target_fingerprint, label="File BK"
        )
        self._assert_source_members_unchanged(plan)
        actions: list[dict[str, Any]] = []
        history: list[dict[str, Any]] = []
        workbook = self.gateway.load(plan.target_path, read_only=False)
        try:
            for sheet_name in self._ordered_sheet_names(plan.target_sheets):
                if sheet_name not in workbook.sheetnames:
                    raise ExpensePostingError(f"Không tìm thấy sheet {sheet_name}.")
                indexes = [
                    index
                    for index, item in enumerate(plan.items)
                    if item.sheet_name == sheet_name
                ]
                local_items = [copy.deepcopy(plan.items[index]) for index in indexes]
                local_conflicts = self._localized_conflicts(plan.conflicts, indexes)
                sheet_actions, sheet_history = self._resolve_apply_actions(
                    workbook[sheet_name],
                    local_items,
                    local_conflicts,
                    resolved,
                    batch_hash=plan.batch_hash,
                )
                actions.extend(sheet_actions)
                history.extend(sheet_history)
        finally:
            workbook.close()

        write_actions = [
            action
            for action in actions
            if action.get("amount_write")
            or action.get("invoice_write")
            or action.get("carrier_write")
        ]
        detail_actions = [
            action
            for action in actions
            if action.get("carrier_group") in {"HP", "NAM"}
            and action.get("status")
            in {PostingItemStatus.POSTED, PostingItemStatus.ALREADY_EXISTS}
        ]
        already_actions = [
            action
            for action in actions
            if action["status"] is PostingItemStatus.ALREADY_EXISTS
        ]
        skipped_actions = [
            action
            for action in actions
            if action["status"]
            in {
                PostingItemStatus.USER_SKIPPED,
                PostingItemStatus.NOT_MATCHED,
                PostingItemStatus.UNRESOLVED,
            }
        ]
        working_path: Path | None = None
        backup_path: Path | None = None
        after = plan.target_fingerprint
        replaced = False
        self._update_run(plan.run_id, status=ExcelRunStatus.APPLYING)
        try:
            detail_rows: list[dict[str, Any]] = []
            update_meta: dict[str, tuple[int, datetime]] = {}
            grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
            for action in write_actions:
                grouped[str(action["sheet_name"])].append(action)
            grouped_details: dict[str, list[dict[str, Any]]] = defaultdict(list)
            for action in detail_actions:
                grouped_details[str(action["sheet_name"])].append(action)
            if write_actions or detail_actions:
                with self.lock_service.acquire(plan.target_path):
                    pass
                self.gateway.assert_unchanged(
                    plan.target_path, plan.target_fingerprint, label="File BK"
                )
                working_path = self.backups.create_working_copy(
                    plan.target_path, run_id=plan.run_id
                )
                write_book = self.gateway.load(working_path, read_only=False)
                try:
                    for sheet_name, sheet_actions in grouped.items():
                        _progress(
                            progress_callback,
                            f"Đang ghi {sheet_name} ({len(update_meta) + 1}/{len(grouped)})…",
                        )
                        worksheet = write_book[sheet_name]
                        carried = self._write_carried_plan_rows(worksheet, sheet_actions)
                        base = self._resolve_base_headers(worksheet)
                        fee_columns = self._resolve_fee_columns(worksheet, base)
                        invoice_columns = self._resolve_invoice_columns(
                            worksheet, base, fee_columns
                        )
                        carrier_columns = self._resolve_carrier_columns(base)
                        update_column = self._ensure_update_column(worksheet)
                        timestamp = self.clock().replace(tzinfo=None, microsecond=0)
                        update_meta[sheet_name] = (update_column, timestamp)
                        updated_rows: set[int] = set()
                        for action in sheet_actions:
                            fee = str(action["fee_selected"])
                            row = int(action["target_row"])
                            column = int(action["target_column"])
                            if fee_columns.get(fee) != column:
                                raise ExpensePostingError(
                                    f"Cột phí {fee} trên {sheet_name} không còn khớp header."
                                )
                            if action.get("amount_write"):
                                worksheet.cell(row, column).value = action["value_after"]
                            if action.get("invoice_write"):
                                invoice_column = int(action["invoice_target_column"])
                                if invoice_columns.get(fee) != invoice_column:
                                    raise ExpensePostingError(
                                        f"Cột Số HĐ của phí {fee} trên {sheet_name} không còn khớp header."
                                    )
                                worksheet.cell(row, invoice_column).value = action[
                                    "invoice_value_after"
                                ]
                            if action.get("carrier_write"):
                                carrier_group = str(action["carrier_group"])
                                carrier_column = int(action["carrier_target_column"])
                                if carrier_columns.get(carrier_group) != carrier_column:
                                    raise ExpensePostingError(
                                        f"Cột bên vận tải {carrier_group} trên {sheet_name} không còn khớp header."
                                    )
                                worksheet.cell(row, carrier_column).value = action[
                                    "carrier_value_after"
                                ]
                            updated_rows.add(row)
                        for row in updated_rows:
                            cell = worksheet.cell(row, update_column)
                            cell.value = timestamp
                            cell.number_format = UPDATE_NUMBER_FORMAT
                        if carried:
                            from .payment_sync import find_summary_start, refresh_bk_summary_formulas

                            if find_summary_start(worksheet) is not None:
                                refresh_bk_summary_formulas(worksheet)
                    for sheet_name, sheet_detail_actions in grouped_details.items():
                        timestamp = update_meta.get(
                            sheet_name,
                            (0, self.clock().replace(tzinfo=None, microsecond=0)),
                        )[1]
                        detail_rows.extend(
                            detail_rows_from_actions(
                                sheet_detail_actions,
                                batch_id=plan.batch_id,
                                batch_hash=plan.batch_hash,
                                sheet_name=sheet_name,
                                updated_at=timestamp,
                            )
                        )
                    upsert_bk_detail_rows(write_book, detail_rows)
                    self.gateway.save(write_book, working_path)
                finally:
                    write_book.close()
                for sheet_name in set(grouped) | set(grouped_details):
                    sheet_actions = grouped.get(sheet_name, [])
                    update_column, timestamp = update_meta.get(
                        sheet_name,
                        (1, self.clock().replace(tzinfo=None, microsecond=0)),
                    )
                    self._verify_posting(
                        working_path,
                        sheet_name,
                        sheet_actions,
                        detail_rows=[
                            row
                            for row in detail_rows
                            if str(row.get("sheet_name") or "") == sheet_name
                        ],
                        update_column=update_column,
                        update_timestamp=timestamp,
                    )
                backup_path = self.backups.create_backup(plan.target_path)
                after = self.gateway.atomic_replace(
                    working_path,
                    plan.target_path,
                    expected=plan.target_fingerprint,
                )
                working_path = None
                replaced = True
                self._record_carry_forwards(plan, write_actions)

            self._record_history(plan, history)
            self._mark_completed_reconciliations(plan)
            posted_count = sum(len(action["source_indices"]) for action in write_actions)
            skipped_count = sum(len(action["source_indices"]) for action in skipped_actions)
            already_count = sum(len(action["source_indices"]) for action in already_actions)
            target_sheets = tuple(self._ordered_sheet_names(plan.target_sheets))
            result = PostingResult(
                status=(
                    ExcelRunStatus.SUCCEEDED
                    if write_actions or detail_actions
                    else ExcelRunStatus.NO_CHANGES
                ),
                target_path=plan.target_path,
                sheet_name=", ".join(target_sheets) or None,
                target_sheets=target_sheets,
                posted_source_items=posted_count,
                written_cells=sum(bool(action.get("amount_write")) for action in write_actions),
                invoice_written_cells=sum(bool(action.get("invoice_write")) for action in write_actions),
                carrier_written_cells=sum(bool(action.get("carrier_write")) for action in write_actions),
                skipped_source_items=skipped_count,
                already_existing_items=already_count,
                conflict_count=len(plan.conflicts),
                backup_path=backup_path,
                fingerprint_before=plan.target_fingerprint,
                fingerprint_after=after,
                run_id=plan.run_id,
                message=(
                    f"Đã nhập {posted_count} khoản vào {len(target_sheets)} sheet BK."
                    if write_actions or detail_actions
                    else "Không có ô cần ghi."
                ),
                item_outcomes=posting_outcomes(actions),
            )
            self._finish_result(result)
            return result
        except Exception as exc:
            if replaced and backup_path is not None:
                try:
                    rollback = self.backups.create_working_copy(
                        backup_path, run_id=f"{plan.run_id}-rollback"
                    )
                    try:
                        self.gateway.atomic_replace(
                            rollback, plan.target_path, expected=after
                        )
                    finally:
                        rollback.unlink(missing_ok=True)
                except Exception as rollback_error:
                    exc.add_note(f"Khôi phục BK thất bại: {rollback_error}")
            self._finish_failed(plan.run_id, exc)
            raise
        finally:
            if working_path is not None and working_path.exists():
                working_path.unlink()

    def _validate_json(
        self, raw: bytes, *, allow_negative: bool = False
    ) -> list[dict[str, Any]]:
        try:
            document = JsonCodec().loads(raw)
        except JsonCodecError as exc:
            raise ExpensePostingError(
                f"JSON đã xác nhận không đọc được: {exc}"
            ) from exc
        result: list[dict[str, Any]] = []
        for index, row in enumerate(document.rows):
            container = row.cont
            bl = row.bl
            fee = row.fee
            rule = row.rule
            invoice_no = row.invoice_no
            carrier = row.carrier
            amount = row.amount
            if container is not None and not isinstance(container, str):
                raise ExpensePostingError(f"Container dòng {index + 1} không hợp lệ.")
            if bl is not None and not isinstance(bl, str):
                raise ExpensePostingError(f"B/L dòng {index + 1} không hợp lệ.")
            if invoice_no is not None and not isinstance(invoice_no, str):
                raise ExpensePostingError(f"Số HĐ dòng {index + 1} không hợp lệ.")
            if carrier is not None and not isinstance(carrier, str):
                raise ExpensePostingError(
                    f"Bên vận tải dòng {index + 1} không hợp lệ."
                )
            if not isinstance(fee, str) or fee.strip().upper() not in FEE_CODES:
                raise ExpensePostingError(f"Mã phí dòng {index + 1} không hợp lệ.")
            if rule is not None and (
                not isinstance(rule, str) or rule.strip().upper() not in RULE_CODES
            ):
                raise ExpensePostingError(f"Rule dòng {index + 1} không hợp lệ.")
            if (
                isinstance(amount, bool)
                or type(amount) is not int
                or (amount < 0 and not allow_negative)
            ):
                raise ExpensePostingError(f"Số tiền dòng {index + 1} không hợp lệ.")
            payload = {
                    "source_item_index": index,
                    "source_document_id": row.source_document_id,
                    "source_document_name": row.source_document_name,
                    "container": normalize_container(container),
                    "bl": normalize_bl(bl),
                    "fee": fee.strip().upper(),
                    "rule": rule.strip().upper() if isinstance(rule, str) else None,
                    "invoice_no": invoice_no,
                    "invoice_date": row.invoice_date,
                    "carrier": carrier,
                    "amount": amount,
            }
            if any((row.vessel_voyage_raw, row.vessel_name, row.voyage_no)):
                payload.update(
                    {
                        "vessel_voyage_raw": row.vessel_voyage_raw,
                        "vessel_name": row.vessel_name,
                        "voyage_no": row.voyage_no,
                    }
                )
            result.append(payload)
        return result

    def _posting_bundle(self, batch_id: int | None) -> dict[str, Any]:
        """Ghép file nguồn với các batch con đối soát thành một nguồn logic."""

        batch_path = self._ready_path(batch_id)
        if batch_path is None:
            raise ExpensePostingError(
                "Không có JSON hiện hành đã xác nhận để nhập khoản chi."
            )
        batch_path = Path(batch_path)
        raw = batch_path.read_bytes()
        resolved_batch_id = batch_id or self._batch_id_for_path(batch_path)
        source_hash = hashlib.sha256(raw).hexdigest()
        repository = getattr(self.provider, "repository", None)
        metadata = (
            repository.get_by_id(resolved_batch_id)
            if repository is not None and resolved_batch_id is not None
            else None
        )
        source_kind = str(getattr(metadata, "source_kind", "ASSISTANT"))
        source_rows = self._validate_json(
            raw, allow_negative=source_kind == "BANG_KE"
        )
        members: list[dict[str, Any]] = [
            {
                "batch_id": resolved_batch_id,
                "path": batch_path,
                "sha256": source_hash,
                "size": len(raw),
                "kind": source_kind,
            }
        ]

        # Fallback tương thích provider độc lập và trường hợp caller chủ động
        # yêu cầu chính batch kết quả đối soát.
        if (
            resolved_batch_id is None
            or self.sea_freight_repository is None
            or source_kind == "SEA_FREIGHT_RECONCILIATION"
        ):
            rows = []
            group_id = self._reconciliation_group_id(resolved_batch_id)
            if group_id is not None and self.sea_freight_service is not None:
                group = self.sea_freight_repository.get_group(group_id)
                if group is not None and str(group.status.value) == "ALLOCATED":
                    self.sea_freight_service.prepare_allocation(group_id)
            for index, row in enumerate(source_rows):
                rows.append(
                    {
                        **row,
                        "source_item_index": index,
                        "origin_batch_id": resolved_batch_id,
                        "origin_batch_hash": source_hash,
                        "origin_source_item_index": index,
                        "reconciliation_group_id": group_id,
                    }
                )
            return {
                "batch_id": resolved_batch_id,
                "batch_path": batch_path,
                "raw": raw,
                "bundle_hash": source_hash,
                "rows": rows,
                "source_members": members,
                "reconciliation_members": (
                    [{"group_id": group_id, "batch_id": resolved_batch_id, "row_count": len(rows)}]
                    if group_id is not None
                    else []
                ),
                "original_source_count": 0 if group_id is not None else len(rows),
                "reconciliation_source_count": len(rows) if group_id is not None else 0,
                "source_kind": source_kind,
            }

        managed = self.sea_freight_repository.groups_for_source_batch(
            resolved_batch_id
        )
        incomplete = [
            group
            for group in managed.values()
            if str(group.status.value) not in {"ALLOCATED", "POSTED"}
        ]
        if incomplete:
            raise ExpensePostingError(
                f"Còn {len(incomplete)} hồ sơ cước biển chưa xác nhận; "
                "hãy hoàn tất đối soát trước khi nhập BK."
            )

        combined: list[dict[str, Any]] = []
        for source_index, row in enumerate(source_rows):
            if source_index in managed:
                continue
            combined.append(
                {
                    **row,
                    "origin_batch_id": resolved_batch_id,
                    "origin_batch_hash": source_hash,
                    "origin_source_item_index": source_index,
                    "reconciliation_group_id": None,
                }
            )
        original_count = len(combined)
        reconciliation_members: list[dict[str, Any]] = []
        for group in self.sea_freight_repository.posting_groups_for_source_batch(
            resolved_batch_id
        ):
            status = str(group.status.value)
            if status not in {"ALLOCATED", "POSTED"}:
                raise ExpensePostingError(
                    f"Hồ sơ đối soát lần {group.revision_no} chưa được xác nhận."
                )
            allocation_rows = None
            if self.sea_freight_service is not None:
                allocation_rows = (
                    self.sea_freight_service.prepare_allocation(group.id)
                    if status == "ALLOCATED"
                    else self.sea_freight_service.allocation_rows(group.id)
                )
            child = None
            batch_service = getattr(self.provider, "batch_service", None)
            if batch_service is not None and allocation_rows is not None:
                # JSON đối soát là đầu ra có thể tái tạo. Luôn dựng từ dữ liệu
                # hồ sơ hiện tại thay vì coi đường dẫn batch lịch sử là đầu vào.
                child = batch_service.create_reconciliation_batch(
                    group.id, allocation_rows
                ).metadata
            elif group.generated_batch_id is not None and repository is not None:
                child = repository.get_by_id(int(group.generated_batch_id))
            if child is None:
                raise ExpensePostingError(
                    f"Hồ sơ đối soát #{group.id} chưa tạo được kết quả hiện tại."
                )
            child_batch_id = int(child.id)
            child_path = getattr(child, "ready_path", None) if child is not None else None
            if child_path is None or not Path(child_path).is_file():
                raise ExpensePostingError(
                    f"Thiếu JSON kết quả của hồ sơ đối soát #{group.id}."
                )
            child_path = Path(child_path)
            child_raw = child_path.read_bytes()
            child_hash = hashlib.sha256(child_raw).hexdigest()
            child_rows = self._validate_json(child_raw)
            if len(child_rows) != group.bk_container_count:
                raise ExpensePostingError(
                    f"Batch kết quả hồ sơ #{group.id} không còn đủ số cont."
                )
            members.append(
                {
                    "batch_id": child_batch_id,
                    "path": child_path,
                    "sha256": child_hash,
                    "size": len(child_raw),
                    "kind": "SEA_FREIGHT_RECONCILIATION",
                }
            )
            reconciliation_members.append(
                {
                    "group_id": group.id,
                    "batch_id": child_batch_id,
                    "row_count": len(child_rows),
                }
            )
            for child_index, row in enumerate(child_rows):
                combined.append(
                    {
                        **row,
                        "origin_batch_id": child_batch_id,
                        "origin_batch_hash": child_hash,
                        "origin_source_item_index": child_index,
                        "reconciliation_group_id": group.id,
                    }
                )

        for index, row in enumerate(combined):
            row["source_item_index"] = index
        digest_payload = [
            (member["batch_id"], member["sha256"])
            for member in members
        ]
        bundle_hash = hashlib.sha256(
            json.dumps(digest_payload, ensure_ascii=True, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        return {
            "batch_id": resolved_batch_id,
            "batch_path": batch_path,
            "raw": raw,
            "bundle_hash": bundle_hash,
            "rows": combined,
            "source_members": members,
            "reconciliation_members": reconciliation_members,
            "original_source_count": original_count,
            "reconciliation_source_count": len(combined) - original_count,
            "source_kind": source_kind,
        }

    def _successful_bundle_indices(
        self, rows: Sequence[Mapping[str, Any]]
    ) -> set[int]:
        if self.posting_repository is None:
            return set()
        cache: dict[int, set[int]] = {}
        hash_cache: dict[str, set[int]] = {}
        result: set[int] = set()
        lookup = getattr(
            self.posting_repository, "successful_source_indices_for_batch", None
        )
        for row in rows:
            origin_batch = row.get("origin_batch_id")
            if origin_batch is None:
                origin_hash = str(row.get("origin_batch_hash") or "")
                if origin_hash:
                    hash_cache.setdefault(
                        origin_hash, self._successful_indices(origin_hash)
                    )
                    if int(row["origin_source_item_index"]) in hash_cache[origin_hash]:
                        result.add(int(row["source_item_index"]))
                continue
            if not callable(lookup):
                continue
            origin_batch = int(origin_batch)
            cache.setdefault(origin_batch, set(lookup(origin_batch)))
            if int(row["origin_source_item_index"]) in cache[origin_batch]:
                result.add(int(row["source_item_index"]))
        return result

    @staticmethod
    def _assert_source_members_unchanged(plan: PostingPlan) -> None:
        members = plan.source_members
        if not members:
            if hashlib.sha256(plan.batch_path.read_bytes()).hexdigest() != plan.batch_hash:
                raise SourceDataChangedError(
                    "JSON đã xác nhận",
                    "JSON đã xác nhận đã thay đổi sau khi phân tích.",
                )
            return
        for member in members:
            path = Path(member["path"])
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != member["sha256"]:
                raise SourceDataChangedError(
                    "JSON nguồn của gói nhập BK",
                    "Một JSON nguồn của gói nhập BK đã thay đổi sau khi phân tích.",
                )

    def cancel(self, plan: PostingPlan) -> None:
        """Đánh dấu lần nhập đang chờ lựa chọn là đã hủy, không sửa BK."""

        if self.run_repository is not None and plan.run_id is not None:
            self.run_repository.finish_run(
                plan.run_id,
                status=ExcelRunStatus.CANCELLED,
                sheet_name=(
                    ", ".join(self._ordered_sheet_names(plan.target_sheets))
                    or plan.selected_sheet
                ),
                total_items=plan.source_item_count,
                changed_items=0,
                skipped_items=0,
                conflict_count=len(plan.conflicts),
            )

    @staticmethod
    def _group_rows(
        rows: Sequence[dict[str, Any]],
        already_posted: set[int],
        repost_source_indices: set[int] | None = None,
        managed_source_indices: set[int] | None = None,
    ) -> list[list[dict[str, Any]]]:
        repost = repost_source_indices or set()
        managed = managed_source_indices or set()
        # Mỗi dòng hóa đơn phải còn độc lập cho tới khi xác định được dòng BK.
        # Cùng container/mã phí chưa đủ để kết luận các khoản thuộc cùng một SQT.
        return [
            [row]
            for row in rows
            if row["source_item_index"] not in managed
            and (
                row["source_item_index"] not in already_posted
                or row["source_item_index"] in repost
            )
        ]

    def _reconciliation_group_id(self, batch_id: int | None) -> int | None:
        if self.sea_freight_repository is None or batch_id is None:
            return None
        row = self.sea_freight_repository.database.query_one(
            "SELECT reconciliation_group_id FROM batches WHERE id = ?",
            (batch_id,),
        )
        if row is None or row["reconciliation_group_id"] is None:
            return None
        return int(row["reconciliation_group_id"])

    def _build_posting_source_groups(
        self,
        rows: Sequence[dict[str, Any]],
        sheet_names: Sequence[str],
        *,
        group_target_sheets: Mapping[str, str] | None,
        legacy_sheet: str | None,
        split_document_ids: set[str],
    ) -> list[PostingSourceGroup]:
        """Tạo các đơn vị phân bổ theo chứng từ, tách theo HĐ khi khác tháng."""

        available = set(sheet_names)
        grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        document_rows: dict[str, list[dict[str, Any]]] = defaultdict(list)
        reconciliation_rows: dict[int, list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            reconciliation_group_id = row.get("reconciliation_group_id")
            if reconciliation_group_id is not None:
                reconciliation_rows[int(reconciliation_group_id)].append(row)
            else:
                document_rows[str(row["source_document_id"])].append(row)

        unknown_split_documents = split_document_ids.difference(document_rows)
        if unknown_split_documents:
            raise ExpensePostingError(
                "Danh sách tài liệu cần tách theo hóa đơn không còn tương thích."
            )

        for document_id, members in document_rows.items():
            periods = {
                str(value)[:7]
                for value in (row.get("invoice_date") for row in members)
                if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)
            }
            invoice_keys = {
                _invoice_key(row.get("invoice_no"))
                for row in members
                if _invoice_key(row.get("invoice_no"))
            }
            split_by_invoice = len(periods) > 1 or (
                document_id in split_document_ids and len(invoice_keys) > 1
            )
            if split_by_invoice:
                for row in members:
                    invoice_key = _invoice_key(row.get("invoice_no")) or "__NO_INVOICE__"
                    grouped[(document_id, invoice_key)].append(row)
            else:
                grouped[(document_id, "")].extend(members)

        result: list[PostingSourceGroup] = []
        for (document_id, invoice_key), members in grouped.items():
            invoice_values = _unique_invoice_values(
                [row.get("invoice_no") for row in members]
            )
            invoice_no = (
                invoice_values[0]
                if len(invoice_values) == 1
                else ", ".join(invoice_values) or None
            )
            group_id = _stable_id("posting-source", document_id, invoice_key)
            result.append(
                self._posting_source_group(
                    group_id,
                    members,
                    available,
                    invoice_no=invoice_no,
                    target_locked=False,
                    locked_target=None,
                    group_target_sheets=group_target_sheets,
                    legacy_sheet=legacy_sheet,
                    can_split_by_invoice=(not invoice_key and len(invoice_values) > 1),
                )
            )

        for reconciliation_group_id, members in reconciliation_rows.items():
            reconciliation = (
                self.sea_freight_repository.get_group(reconciliation_group_id)
                if self.sea_freight_repository is not None
                else None
            )
            if reconciliation is None:
                raise ExpensePostingError(
                    f"Không tìm thấy hồ sơ đối soát #{reconciliation_group_id}."
                )
            group_id = f"SEA_RECON_{reconciliation_group_id}_R{reconciliation.revision_no}"
            result.append(
                self._posting_source_group(
                    group_id,
                    members,
                    available,
                    invoice_no=", ".join(
                        _unique_invoice_values(
                            [row.get("invoice_no") for row in members]
                        )
                    ) or None,
                    target_locked=True,
                    locked_target=reconciliation.bk_sheet,
                    group_target_sheets=group_target_sheets,
                    legacy_sheet=None,
                    reconciliation_group_id=reconciliation_group_id,
                    can_split_by_invoice=False,
                )
            )

        assigned_indices = sorted(
            source_index
            for group in result
            for source_index in group.source_item_indices
        )
        expected_indices = sorted(int(row["source_item_index"]) for row in rows)
        if assigned_indices != expected_indices:
            raise ExpensePostingError(
                "Không thể phân bổ duy nhất mọi khoản chi vào nhóm chứng từ."
            )
        known_group_ids = {group.group_id for group in result}
        unknown = set(group_target_sheets or {}).difference(known_group_ids)
        if unknown:
            raise ExpensePostingError("Ánh xạ nhóm chứng từ không còn tương thích.")
        return result

    def _posting_source_group(
        self,
        group_id: str,
        members: Sequence[dict[str, Any]],
        available: set[str],
        *,
        invoice_no: str | None,
        target_locked: bool,
        locked_target: str | None,
        group_target_sheets: Mapping[str, str] | None,
        legacy_sheet: str | None,
        reconciliation_group_id: int | None = None,
        can_split_by_invoice: bool = False,
    ) -> PostingSourceGroup:
        dates = sorted(
            {
                str(row["invoice_date"])
                for row in members
                if row.get("invoice_date") not in (None, "")
            }
        )
        periods = {
            (int(value[5:7]), int(value[:4]))
            for value in dates
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)
        }
        suggested = None
        if len(periods) == 1:
            month, year = next(iter(periods))
            candidate = f"T{month:02d} {year % 100:02d}"
            if candidate in available:
                suggested = candidate
        requested = (group_target_sheets or {}).get(group_id)
        target = locked_target if target_locked else requested or legacy_sheet
        if target_locked and requested not in (None, locked_target):
            raise ExpensePostingError(
                f"Hồ sơ cước biển {group_id} bị khóa tại sheet {locked_target}."
            )
        if target is not None and target not in available:
            raise ExpensePostingError(f"Sheet {target!r} không tồn tại trong BK.")
        return PostingSourceGroup(
            group_id=group_id,
            source_document_id=str(members[0]["source_document_id"]),
            source_document_name=str(members[0]["source_document_name"]),
            invoice_no=invoice_no,
            source_item_indices=[int(row["source_item_index"]) for row in members],
            invoice_dates=dates,
            suggested_sheet=suggested,
            target_sheet=target,
            reconciliation_group_id=reconciliation_group_id,
            target_locked=target_locked,
            can_split_by_invoice=can_split_by_invoice,
        )

    def _analyze_assigned_items(
        self,
        workbook: Any,
        items: list[PostingItem],
        *,
        batch_hash: str,
    ) -> tuple[list[PostingItem], list[PostingConflict]]:
        assigned: dict[str, list[PostingItem]] = defaultdict(list)
        unassigned: list[PostingItem] = []
        for item in items:
            if item.sheet_name is None:
                unassigned.append(item)
            else:
                assigned[item.sheet_name].append(item)
        analyzed: list[PostingItem] = []
        conflicts: list[PostingConflict] = []
        for sheet_name, sheet_items in assigned.items():
            offset = len(analyzed)
            sheet_items, sheet_conflicts = self._analyze_items(
                workbook, sheet_name, sheet_items, batch_hash=batch_hash
            )
            for conflict in sheet_conflicts:
                if conflict.item_index is not None:
                    conflict.item_index += offset
                if "item_indexes" in conflict.details:
                    conflict.details["item_indexes"] = [
                        int(index) + offset
                        for index in conflict.details["item_indexes"]
                    ]
            analyzed.extend(sheet_items)
            conflicts.extend(sheet_conflicts)
        analyzed.extend(unassigned)
        return analyzed, conflicts

    @staticmethod
    def _items_from_groups(
        groups: Sequence[Sequence[dict[str, Any]]],
        repost_source_indices: set[int] | None = None,
    ) -> list[PostingItem]:
        repost = repost_source_indices or set()
        return [
            PostingItem(
                source_indices=[row["source_item_index"] for row in group],
                container=group[0]["container"],
                bl=group[0]["bl"],
                original_fee=group[0]["fee"],
                selected_fee=group[0]["fee"],
                rule=group[0]["rule"],
                amount=sum(row["amount"] for row in group),
                vessel_voyage_raw=group[0].get("vessel_voyage_raw"),
                vessel_name=group[0].get("vessel_name"),
                voyage_no=group[0].get("voyage_no"),
                source_items=[dict(row) for row in group],
                invoice_candidates=_unique_invoice_values(
                    [row.get("invoice_no") for row in group]
                ),
                carrier_candidates=unique_carriers(
                    [row.get("carrier") for row in group]
                ),
                force_repost=any(
                    row["source_item_index"] in repost for row in group
                ),
            )
            for group in groups
        ]

    @staticmethod
    def _vessel_voyage_key(*values: Any) -> str:
        text = " ".join(str(value) for value in values if value not in (None, ""))
        text = unicodedata.normalize("NFD", text.upper())
        text = "".join(char for char in text if unicodedata.category(char) != "Mn")
        text = re.sub(r"\b(?:VOYAGE|CHUYEN\s*SO|CHUYEN|VOY|V)\b", " ", text)
        return " ".join(re.findall(r"[A-Z0-9]+", text))

    def _ensure_bang_ke_columns(self, workbook: Any) -> set[str]:
        """Xóa VAT và bảo đảm Gia hạn sau cặp Sửa chữa/HĐ trên mọi sheet tháng."""

        from .bang_ke import BangKeColumnError, ensure_bang_ke_fee_columns

        changed: set[str] = set()
        for worksheet in workbook.worksheets:
            if self.months.parse_target_sheet(worksheet.title) is None:
                continue
            base = self._resolve_base_headers(worksheet)
            try:
                if ensure_bang_ke_fee_columns(
                    worksheet, header_row=base.row_end
                ):
                    changed.add(worksheet.title)
            except BangKeColumnError as exc:
                raise ExpensePostingError(str(exc)) from exc
        calculation = getattr(workbook, "calculation", None)
        if calculation is not None and changed:
            calculation.fullCalcOnLoad = True
            calculation.forceFullCalc = True
            calculation.calcMode = "auto"
        return changed

    def _analyze_bang_ke_items(
        self,
        workbook: Any,
        items: list[PostingItem],
        *,
        batch_hash: str,
    ) -> tuple[list[PostingItem], list[PostingConflict]]:
        sheets: dict[str, tuple[Any, HeaderResolution, dict[str, int]]] = {}
        all_candidates: list[RowCandidate] = []
        by_container: dict[str, list[RowCandidate]] = defaultdict(list)
        for name in workbook.sheetnames:
            if self.months.parse_target_sheet(name) is None:
                continue
            worksheet = workbook[name]
            base = self._resolve_base_headers(worksheet)
            fee_columns = self._resolve_fee_columns(worksheet, base)
            sheets[name] = (worksheet, base, fee_columns)
            candidates = [
                candidate
                for values in self._container_index(worksheet, base).values()
                for candidate in values
            ]
            all_candidates.extend(candidates)
            for candidate in candidates:
                if candidate.container:
                    by_container[candidate.container].append(candidate)

        conflicts: list[PostingConflict] = []
        for item_index, item in enumerate(items):
            if item.status is not PostingItemStatus.PLANNED:
                continue
            item.target_row = item.target_column = None
            item.target_cell = None
            item.current_value = None
            item.cell_state = None
            source_key = self._vessel_voyage_key(
                item.vessel_name,
                item.voyage_no,
                item.vessel_voyage_raw if not (item.vessel_name or item.voyage_no) else None,
            )
            candidates = list(by_container.get(item.container or "", ()))
            item.row_candidates = candidates
            selected: RowCandidate | None = None
            if item.selected_source_row is not None:
                selected = next(
                    (
                        candidate
                        for candidate in all_candidates
                        if candidate.row == item.selected_source_row
                        and candidate.source_sheet == item.selected_source_sheet
                    ),
                    None,
                )
                if selected is None:
                    raise ExpensePostingError("Dòng BK đã chọn không còn hợp lệ.")
            elif candidates:
                exact = [
                    candidate
                    for candidate in candidates
                    if source_key
                    and self._vessel_voyage_key(candidate.vessel) == source_key
                ]
                if len(exact) == 1:
                    selected = exact[0]
                elif len(candidates) == 1:
                    item.row_candidates = candidates
                    conflicts.append(
                        self._item_conflict(
                            batch_hash,
                            item_index,
                            item,
                            ConflictType.PARTIAL_KEY_MATCH,
                            "Container chỉ có một dòng nhưng tàu/chuyến trống hoặc khác; hãy xác nhận dòng BK.",
                            (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                            row_candidates=candidates,
                            details={"source_vessel_voyage": source_key},
                        )
                    )
                    continue
                else:
                    choices = exact if exact else candidates
                    item.row_candidates = choices
                    conflicts.append(
                        self._item_conflict(
                            batch_hash,
                            item_index,
                            item,
                            ConflictType.MULTIPLE_CONTAINER_MATCH,
                            f"Container {item.container} có nhiều dòng trong BK; hãy chọn đúng tàu/chuyến.",
                            (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                            row_candidates=choices,
                            details={"source_vessel_voyage": source_key},
                        )
                    )
                    continue
            else:
                item.row_candidates = all_candidates
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.CONTAINER_NOT_FOUND,
                        f"Không tìm thấy container {item.container or 'trống'} trên toàn bộ BK.",
                        (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                        row_candidates=all_candidates,
                    )
                )
                continue

            assert selected is not None
            item.sheet_name = selected.source_sheet
            item.selected_source_sheet = selected.source_sheet
            item.selected_source_row = selected.row
            item.target_row = selected.row
            item.source_sqt = selected.sqt
            worksheet, _base, fee_columns = sheets[selected.source_sheet]
            item.target_column = fee_columns.get(item.selected_fee)
            if item.target_column is None:
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.FEE_COLUMN_MISSING,
                        f"Không nhận diện được cột phí {item.selected_fee} trên {selected.source_sheet}.",
                        (ResolutionAction.SKIP, ResolutionAction.CANCEL_ALL),
                    )
                )
                continue
            self._set_cell_state(worksheet, item)
            if item.amount < 0:
                state = item.cell_state.kind if item.cell_state else None
                numeric = state in {TargetCellKind.NUMBER, TargetCellKind.ZERO, TargetCellKind.SAME_VALUE}
                before = item.current_value
                after = before + item.amount if numeric else None
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.NEGATIVE_ADJUSTMENT,
                        (
                            f"Áp dụng điều chỉnh giảm: {before:,} + ({item.amount:,}) = {after:,}."
                            + (" CẢNH BÁO: kết quả sau điều chỉnh bị âm." if after is not None and after < 0 else "")
                            if numeric
                            else "Ô đích không phải số; hãy chọn lại dòng hoặc bỏ qua khoản điều chỉnh giảm."
                        ),
                        (
                            (ResolutionAction.ADD, ResolutionAction.SKIP, ResolutionAction.CANCEL_ALL)
                            if numeric
                            else (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP, ResolutionAction.CANCEL_ALL)
                        ),
                        row_candidates=item.row_candidates,
                        details={
                            "value_before": before,
                            "adjustment": item.amount,
                            "value_after": after,
                            "negative_after": bool(after is not None and after < 0),
                        },
                    )
                )
            else:
                conflict = self._cell_conflict(batch_hash, item_index, item)
                if conflict is not None:
                    conflicts.append(conflict)
        return items, conflicts

    def _resolve_base_headers(self, worksheet: Any) -> HeaderResolution:
        return self.headers.resolve(
            worksheet,
            BASE_HEADER_ALIASES,
            required=("sqt", "closing_date", "container"),
        )

    def _resolve_fee_columns(
        self, worksheet: Any, base: HeaderResolution
    ) -> dict[str, int]:
        from .payment_sync import find_summary_start

        normalized = base.headers
        summary_start = find_summary_start(worksheet)
        result: dict[str, int] = {}
        notes_candidates = [
            column
            for column, header in normalized.items()
            if header in {
                normalize_header(alias)
                for alias in BASE_HEADER_ALIASES["notes"]
            }
        ]
        notes_column = min(notes_candidates) if notes_candidates else None
        for fee, aliases in FEE_HEADER_ALIASES.items():
            alias_keys = {normalize_header(alias) for alias in aliases}
            matches = [
                column
                for column, header in normalized.items()
                if header in alias_keys
                and (summary_start is None or column < summary_start)
            ]
            if len(matches) != 1:
                continue
            column = matches[0]
            if normalized[column] in INVOICE_HEADER_NAMES:
                continue
            if (
                fee == "CBDH"
                and normalized[column] == normalize_header("ĐƠN GIÁ")
                and notes_column is not None
                and column >= notes_column
            ):
                # "ĐƠN GIÁ" trong vùng sau GHI CHÚ thuộc thông tin hóa đơn,
                # không phải cột Cước bộ đóng hàng.
                continue
            if fee == "LL" and (
                notes_column is None or column >= notes_column
            ):
                # Không đoán cột LÀM LỆNH nếu không xác định được ranh giới
                # vùng cước chính trước GHI CHÚ.
                continue
            result[fee] = column
        return result

    def _resolve_invoice_columns(
        self,
        worksheet: Any,
        base: HeaderResolution,
        fee_columns: Mapping[str, int],
    ) -> dict[str, int]:
        from .payment_sync import find_summary_start

        result: dict[str, int] = {}
        summary_start = find_summary_start(worksheet)
        for fee, aliases in FEE_INVOICE_HEADER_ALIASES.items():
            if fee not in fee_columns:
                continue
            alias_keys = {normalize_header(alias) for alias in aliases}
            matches = [
                column
                for column, header in base.headers.items()
                if header in alias_keys
                and (summary_start is None or column < summary_start)
            ]
            if len(matches) == 1:
                result[fee] = matches[0]

        for fee, fee_column in fee_columns.items():
            if fee == "LL" or fee in FEE_INVOICE_HEADER_ALIASES:
                continue
            invoice_column = fee_column + 1
            if base.headers.get(invoice_column) in INVOICE_NUMBER_HEADER_NAMES:
                result[fee] = invoice_column
        return result

    @staticmethod
    def _resolve_carrier_columns(base: HeaderResolution) -> dict[str, int]:
        result: dict[str, int] = {}
        sea = base.columns.get("carrier_sea")
        road = base.columns.get("carrier_road")
        nam = base.columns.get("carrier_nam")
        if sea is not None:
            result["SEA"] = sea
        if road is not None:
            result["ROAD"] = road
        if nam is not None:
            result["NAM"] = nam
        return result

    @staticmethod
    def _update_column(
        base: HeaderResolution,
    ) -> int | None:
        matches = [
            column
            for column, header in base.headers.items()
            if header in UPDATE_HEADER_NAMES
        ]
        if len(matches) > 1:
            raise ExpensePostingError(
                "Sheet BK có nhiều cột Date cập nhật; không thể chọn an toàn."
            )
        return matches[0] if matches else None

    def _ensure_update_column(self, worksheet: Any) -> int:
        base = self._resolve_base_headers(worksheet)
        existing = self._update_column(base)
        if existing is not None:
            return existing
        used_headers = [
            column for column, header in base.headers.items() if header
        ]
        column = (max(used_headers) if used_headers else worksheet.max_column) + 1
        header_cell = worksheet.cell(base.row_end, column)
        previous = worksheet.cell(base.row_end, column - 1)
        if previous.has_style:
            header_cell._style = copy.copy(previous._style)
        header_cell.font = copy.copy(previous.font)
        header_cell.fill = copy.copy(previous.fill)
        header_cell.border = copy.copy(previous.border)
        header_cell.alignment = copy.copy(previous.alignment)
        header_cell.protection = copy.copy(previous.protection)
        header_cell.value = UPDATE_HEADER_CANONICAL
        return column

    def _container_index(
        self,
        worksheet: Any,
        base: HeaderResolution,
        *,
        plan_header: HeaderResolution | None = None,
    ) -> dict[str, list[RowCandidate]]:
        result: dict[str, list[RowCandidate]] = defaultdict(list)
        columns = base.columns
        for row in range(base.row_end + 1, worksheet.max_row + 1):
            raw = worksheet.cell(row, columns["container"]).value
            try:
                container = normalize_container(
                    str(raw) if raw not in (None, "") else None
                )
            except TypeError:
                continue
            if not container:
                continue
            cargo = (
                worksheet.cell(row, columns["cargo_type"]).value
                if "cargo_type" in columns
                else None
            )
            plan_values = (
                tuple(
                    worksheet.cell(row, plan_header.columns[field]).value
                    for field in SYNC_FIELDS
                )
                if plan_header is not None
                else ()
            )
            carrier = join_carriers(
                *(
                    worksheet.cell(row, columns[field]).value
                    for field in ("carrier_sea", "carrier_road", "carrier_nam")
                    if field in columns
                )
            )
            result[container].append(
                RowCandidate(
                    row=row,
                    sqt=self._positive_int(
                        worksheet.cell(row, columns["sqt"]).value
                    ),
                    container=container,
                    source_sheet=worksheet.title,
                    source_row=row,
                    plan_values=plan_values,
                    source_signature=(
                        self._plan_signature(plan_values) if plan_values else ""
                    ),
                    cargo_type=cargo,
                    closing_date=worksheet.cell(
                        row, columns["closing_date"]
                    ).value,
                    vessel=(
                        worksheet.cell(row, columns["vessel"]).value
                        if "vessel" in columns
                        else None
                    ),
                    recipient=(
                        worksheet.cell(row, columns["recipient"]).value
                        if "recipient" in columns
                        else None
                    ),
                    carrier=carrier,
                    is_ron="ron" in normalize_header(cargo),
                )
            )
        return dict(result)

    @staticmethod
    def _plan_signature(values: Sequence[Any]) -> str:
        payload = json.dumps(
            list(values),
            ensure_ascii=False,
            default=str,
            separators=(",", ":"),
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _plan_header(self, worksheet: Any) -> HeaderResolution:
        return self.headers.resolve(
            worksheet,
            SOURCE_HEADER_ALIASES,
            required=SYNC_FIELDS,
        )

    def _source_window_names(self, workbook: Any, target_sheet: str) -> list[str]:
        parsed_target = self.months.parse_target_sheet(target_sheet)
        if parsed_target is None:
            raise ExpensePostingError(f"Sheet {target_sheet!r} không có dạng TMM YY.")
        month, year = parsed_target
        by_period: dict[tuple[int, int], str] = {}
        for name in workbook.sheetnames:
            parsed = self.months.parse_target_sheet(name)
            if parsed is None:
                continue
            if parsed in by_period:
                raise ExpensePostingError(
                    f"Có nhiều sheet BK cho tháng {parsed[0]:02d}/{parsed[1]}."
                )
            by_period[parsed] = name
        return [
            by_period[period]
            for offset in range(3)
            if (period := _period_offset(month, year, offset)) in by_period
        ]

    def _carry_forward_record(
        self,
        candidate: RowCandidate,
        target_sheet: str,
    ) -> Any | None:
        getter = getattr(self.posting_repository, "get_carry_forward", None)
        if not callable(getter) or candidate.sqt is None or not candidate.container:
            return None
        return getter(
            workbook_path=self.bk_path,
            source_sheet=candidate.source_sheet,
            source_row=int(candidate.source_row or candidate.row),
            source_sqt=candidate.sqt,
            container=candidate.container,
            target_sheet=target_sheet,
        )

    @staticmethod
    def _row_plan_values(
        worksheet: Any,
        header: HeaderResolution,
        row: int,
    ) -> tuple[Any, ...]:
        return tuple(
            worksheet.cell(row, header.columns[field]).value
            for field in SYNC_FIELDS
        )

    def _window_index(
        self,
        workbook: Any,
        target_sheet: str,
    ) -> tuple[dict[str, list[RowCandidate]], list[RowCandidate], HeaderResolution | None]:
        window_names = self._source_window_names(workbook, target_sheet)
        indexes: dict[str, dict[str, list[RowCandidate]]] = {}
        plan_headers: dict[str, HeaderResolution | None] = {}
        for name in window_names:
            worksheet = workbook[name]
            base = self._resolve_base_headers(worksheet)
            try:
                plan_header = self._plan_header(worksheet)
            except HeaderResolutionError:
                plan_header = None
                if name != target_sheet:
                    raise ExpensePostingError(
                        f"Sheet nguồn {name} không đủ 12 cột thông tin kế hoạch."
                    )
            plan_headers[name] = plan_header
            indexes[name] = self._container_index(
                worksheet,
                base,
                plan_header=plan_header,
            )

        target_worksheet = workbook[target_sheet]
        target_plan_header = plan_headers.get(target_sheet)
        target_candidates = [
            candidate
            for candidates in indexes[target_sheet].values()
            for candidate in candidates
        ]
        carried_target_rows: set[int] = set()
        for name in window_names[1:]:
            for candidates in indexes[name].values():
                for candidate in candidates:
                    record = self._carry_forward_record(candidate, target_sheet)
                    if record is None:
                        continue
                    if str(getattr(record, "source_signature", "")) != candidate.source_signature:
                        candidate.mapping_invalid = True
                        continue
                    target_row = int(getattr(record, "target_row"))
                    matches_recorded_row = (
                        target_plan_header is not None
                        and self._row_plan_values(
                            target_worksheet, target_plan_header, target_row
                        )
                        == candidate.plan_values
                    )
                    if matches_recorded_row:
                        candidate.carried_target_row = target_row
                        carried_target_rows.add(target_row)
                        continue
                    signature_matches = [
                        item.row
                        for item in target_candidates
                        if item.source_signature == candidate.source_signature
                    ]
                    if len(signature_matches) == 1:
                        candidate.carried_target_row = signature_matches[0]
                        carried_target_rows.add(signature_matches[0])
                    else:
                        candidate.mapping_invalid = True

        combined: dict[str, list[RowCandidate]] = defaultdict(list)
        all_candidates: list[RowCandidate] = []
        for name in window_names:
            for container, candidates in indexes[name].items():
                for candidate in candidates:
                    if name == target_sheet and candidate.row in carried_target_rows:
                        continue
                    combined[container].append(candidate)
                    all_candidates.append(candidate)
        return dict(combined), all_candidates, target_plan_header

    @staticmethod
    def _origin_key(candidate: RowCandidate) -> tuple[str, int, int | None, str | None]:
        return (
            candidate.source_sheet,
            int(candidate.source_row or candidate.row),
            candidate.sqt,
            candidate.container,
        )

    def _sheet_candidates(
        self,
        workbook: Any,
        sheet_names: Sequence[str],
        groups: Sequence[Sequence[dict[str, Any]]],
        latest_sheet: str | None,
    ) -> tuple[list[MonthCandidate], dict[str, dict[str, list[RowCandidate]]]]:
        containers = {
            row["container"]
            for group in groups
            for row in group
            if row["container"]
        }
        candidates: list[MonthCandidate] = []
        indexes: dict[str, dict[str, list[RowCandidate]]] = {}
        for name in sheet_names:
            base = self._resolve_base_headers(workbook[name])
            index = self._container_index(workbook[name], base)
            indexes[name] = index
            parsed = self.months.parse_target_sheet(name)
            candidates.append(
                MonthCandidate(
                    month=parsed[0] if parsed else 0,
                    source_sheet="",
                    target_sheet=name,
                    match_count=sum(container in index for container in containers),
                    recently_synced=name == latest_sheet,
                )
            )
        return candidates, indexes

    @staticmethod
    def _choose_sheet(
        requested: str | None,
        candidates: Sequence[MonthCandidate],
        groups: Sequence[Sequence[dict[str, Any]]],
        latest_sheet: str | None,
    ) -> str | None:
        names = {candidate.target_sheet for candidate in candidates}
        if requested is not None:
            if requested not in names:
                raise ExpensePostingError(f"Sheet {requested!r} không hợp lệ.")
            return requested
        return None

    def _analyze_items(
        self,
        workbook: Any,
        target_sheet: str,
        items: list[PostingItem],
        *,
        batch_hash: str,
    ) -> tuple[list[PostingItem], list[PostingConflict]]:
        worksheet = workbook[target_sheet]
        base = self._resolve_base_headers(worksheet)
        fee_columns = self._resolve_fee_columns(worksheet, base)
        invoice_columns = self._resolve_invoice_columns(
            worksheet, base, fee_columns
        )
        carrier_columns = self._resolve_carrier_columns(base)
        index, manual_candidates, target_plan_header = self._window_index(
            workbook, target_sheet
        )
        manual_candidates = sorted(
            manual_candidates,
            key=lambda candidate: (
                self._source_window_names(workbook, target_sheet).index(
                    candidate.source_sheet
                ),
                candidate.row,
            ),
        )
        next_target_row = (
            DailySyncService._last_data_row(worksheet, target_plan_header) + 1
            if target_plan_header is not None
            else base.row_end + 1
        )
        allocated_rows: dict[tuple[str, int, int | None, str | None], int] = {}
        analyzed_items: list[PostingItem] = []
        conflicts: list[PostingConflict] = []
        target_groups: dict[tuple[str, int, int], list[int]] = defaultdict(list)

        for item in items:
            if item.status is not PostingItemStatus.PLANNED:
                analyzed_items.append(item)
                continue
            item.sheet_name = target_sheet
            item.target_column = None
            item.target_cell = None
            item.current_value = None
            item.cell_state = None
            item.carry_forward_required = False
            item.carrier_group = carrier_group_for_fee(item.selected_fee)
            item.carrier_column = None
            item.carrier_cell = None
            item.carrier_current_value = None
            item.carrier_value_after = None
            # carrier_action là quyết định của user, không phải trạng thái đọc
            # từ workbook; giữ nó qua các vòng phân tích lại.
            if item.selected_fee not in FEE_HEADER_ALIASES:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.UNKNOWN_FEE_CODE,
                        "Mã phí chưa xác định; hãy chọn một trong 12 mã phí.",
                        (ResolutionAction.SELECT_FEE, ResolutionAction.SKIP),
                        details={"fees": sorted(FEE_HEADER_ALIASES)},
                    )
                )
                continue

            manually_selected: RowCandidate | None = None
            selected_source_row = item.selected_source_row
            selected_source_sheet = item.selected_source_sheet
            if selected_source_row is None and item.target_row is not None:
                selected_source_row = item.target_row
                selected_source_sheet = target_sheet
            if selected_source_row is not None:
                manually_selected = next(
                    (
                        candidate
                        for candidate in manual_candidates
                        if candidate.row == int(selected_source_row)
                        and candidate.source_sheet
                        == str(selected_source_sheet or target_sheet)
                    ),
                    None,
                )
                if manually_selected is None:
                    raise ExpensePostingError(
                        f"Dòng nguồn {selected_source_sheet or target_sheet}!"
                        f"{selected_source_row} không còn hợp lệ."
                    )

            if item.container is None and manually_selected is None:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflict_type = (
                    ConflictType.BL_ONLY_NO_CONTAINER
                    if item.bl and item.selected_fee == "CB"
                    else ConflictType.CONTAINER_NOT_FOUND
                )
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        conflict_type,
                        "Khoản chi chưa có container; hãy chọn đúng dòng kế hoạch.",
                        (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                        row_candidates=manual_candidates,
                    )
                )
                continue

            candidates = (
                [manually_selected]
                if manually_selected is not None
                else list(index.get(item.container or "", ()))
            )
            item.row_candidates = candidates
            selected = manually_selected or self._automatic_row(candidates)
            if not candidates:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.CONTAINER_NOT_FOUND,
                        f"Không tìm thấy container {item.container} trong tháng đích và hai tháng trước.",
                        (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                        row_candidates=manual_candidates,
                    )
                )
                continue
            if selected is None:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.MULTIPLE_CONTAINER_MATCH,
                        f"Container {item.container} có nhiều dòng kế hoạch; hãy chọn đúng tháng và SQT.",
                        (ResolutionAction.SELECT_ROW, ResolutionAction.SKIP),
                        row_candidates=candidates,
                    )
                )
                continue

            item.selected_source_sheet = selected.source_sheet
            item.selected_source_row = int(selected.source_row or selected.row)
            item.source_sqt = selected.sqt
            item.plan_values = tuple(selected.plan_values)
            item.source_signature = selected.source_signature
            if selected.mapping_invalid:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.CARRY_FORWARD_MAPPING_INVALID,
                        "Ánh xạ dòng đã mang sang không còn khớp; cần kiểm tra BK trước khi tiếp tục.",
                        (ResolutionAction.SKIP, ResolutionAction.CANCEL_ALL),
                        default=ResolutionAction.SKIP,
                    )
                )
                continue

            if selected.source_sheet == target_sheet:
                item.target_row = selected.row
            elif selected.carried_target_row is not None:
                item.target_row = selected.carried_target_row
            else:
                if target_plan_header is None or not selected.plan_values:
                    raise ExpensePostingError(
                        f"Không thể mang dòng {selected.source_sheet}!{selected.row}: "
                        "sheet đích không đủ 12 cột kế hoạch."
                    )
                origin_key = self._origin_key(selected)
                if origin_key not in allocated_rows:
                    allocated_rows[origin_key] = next_target_row
                    next_target_row += 1
                item.target_row = allocated_rows[origin_key]
                item.carry_forward_required = True

            item.target_column = fee_columns.get(item.selected_fee)
            if item.target_column is None:
                item_index = len(analyzed_items)
                analyzed_items.append(item)
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        item_index,
                        item,
                        ConflictType.FEE_COLUMN_MISSING,
                        f"Không nhận diện duy nhất cột phí {item.selected_fee}.",
                        (ResolutionAction.SKIP,),
                        default=ResolutionAction.SKIP,
                    )
                )
                continue

            item_index = len(analyzed_items)
            analyzed_items.append(item)
            target_groups[(target_sheet, item.target_row, item.target_column)].append(
                item_index
            )

        for (_sheet, _row, _column), item_indexes in target_groups.items():
            if len(item_indexes) > 1:
                first_index = item_indexes[0]
                first = analyzed_items[first_index]
                first.target_cell = f"{get_column_letter(int(first.target_column))}{first.target_row}"
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        first_index,
                        first,
                        ConflictType.MULTIPLE_EXPENSE_SAME_CELL,
                        "Nhiều dòng JSON cùng trỏ tới một ô phí; hãy chọn đúng một dòng để ghi.",
                        (ResolutionAction.SELECT_SOURCE_ITEM,),
                        default=ResolutionAction.SELECT_SOURCE_ITEM,
                        details={
                            "item_indexes": list(item_indexes),
                            "source_item_options": [
                                {
                                    "source_item_index": analyzed_items[index].source_indices[0],
                                    "amount": analyzed_items[index].amount,
                                    "invoice_no": (
                                        analyzed_items[index].invoice_candidates[0]
                                        if len(analyzed_items[index].invoice_candidates) == 1
                                        else None
                                    ),
                                }
                                for index in item_indexes
                            ],
                        },
                    )
                )
                continue

            item_index = item_indexes[0]
            item = analyzed_items[item_index]
            self._set_cell_state(worksheet, item)
            cell_conflict = self._cell_conflict(batch_hash, item_index, item)
            if cell_conflict is not None:
                conflicts.append(cell_conflict)
            invoice_conflict = self._set_invoice_state(
                worksheet,
                item,
                item_index=item_index,
                batch_hash=batch_hash,
                invoice_columns=invoice_columns,
            )
            if invoice_conflict is not None:
                conflicts.append(invoice_conflict)
        conflicts.extend(
            self._analyze_carrier_groups(
                worksheet,
                analyzed_items,
                batch_hash=batch_hash,
                carrier_columns=carrier_columns,
            )
        )
        return analyzed_items, conflicts

    def _analyze_carrier_groups(
        self,
        worksheet: Any,
        items: Sequence[PostingItem],
        *,
        batch_hash: str,
        carrier_columns: Mapping[str, int],
    ) -> list[PostingConflict]:
        grouped: dict[tuple[int, str], list[int]] = defaultdict(list)
        conflicts: list[PostingConflict] = []
        for index, item in enumerate(items):
            group = carrier_group_for_fee(item.selected_fee)
            item.carrier_group = group
            if (
                group is None
                or item.target_row is None
                or item.status not in {PostingItemStatus.PLANNED, PostingItemStatus.ALREADY_EXISTS}
            ):
                continue
            column = carrier_columns.get(group)
            item.carrier_column = column
            if column is None:
                if not item.carrier_candidates:
                    continue
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        index,
                        item,
                        ConflictType.CARRIER_COLUMN_MISSING,
                        f"Không nhận diện được cột bên vận tải cho nhóm {group}.",
                        (ResolutionAction.KEEP_EXISTING, ResolutionAction.CANCEL_ALL),
                        default=ResolutionAction.KEEP_EXISTING,
                        details={"scope": "carrier", "carrier_group": group},
                    )
                )
                continue
            cell = worksheet.cell(int(item.target_row), column)
            item.carrier_cell = cell.coordinate
            item.carrier_current_value = cell.value
            grouped[(int(item.target_row), group)].append(index)

        for (row, group), indexes in grouped.items():
            first = items[indexes[0]]
            column = int(carrier_columns[group])
            cell = worksheet.cell(row, column)
            incoming = unique_carriers(
                carrier
                for index in indexes
                for carrier in (
                    [items[index].selected_carrier]
                    if items[index].selected_carrier
                    else items[index].carrier_candidates
                )
            )
            invoice_carriers: dict[str, list[str]] = defaultdict(list)
            invoice_labels: dict[str, str] = {}
            for index in indexes:
                item = items[index]
                invoices = item.invoice_candidates or ([item.selected_invoice] if item.selected_invoice else [])
                for invoice in invoices:
                    key = _invoice_key(invoice)
                    if key is None:
                        continue
                    invoice_labels[key] = str(invoice)
                    invoice_carriers[key].extend(
                        [item.selected_carrier]
                        if item.selected_carrier
                        else item.carrier_candidates
                    )
            ambiguous = {
                invoice_labels[key]: unique_carriers(values)
                for key, values in invoice_carriers.items()
                if len(unique_carriers(values)) > 1
            }
            details = {
                "scope": "carrier",
                "carrier_group": group,
                "item_indexes": list(indexes),
                "carrier_candidates": incoming,
                "invoice_carriers": {
                    invoice_labels[key]: unique_carriers(values)
                    for key, values in invoice_carriers.items()
                },
                "carrier_column": column,
                "carrier_cell": cell.coordinate,
                "carrier_current_value": cell.value,
            }
            if group in DAILY_MANAGED_CARRIER_GROUPS:
                # Carrier GPT của CB/CBDH/VTN đã bị bỏ khi tiếp nhận. Một giá
                # trị còn xuất hiện ở đây vì thế là ngoại lệ user chủ động nhập.
                # Có giá trị thì dùng, để trống thì giữ nguồn chuẩn Hàng ngày.
                for index in indexes:
                    items[index].carrier_action = (
                        ResolutionAction.OVERWRITE
                        if incoming
                        else ResolutionAction.KEEP_EXISTING
                    )
                    items[index].carrier_value_after = (
                        join_carriers(incoming) if incoming else cell.value
                    )
                continue
            if ambiguous:
                ambiguous_candidates = unique_carriers(
                    carrier
                    for values in ambiguous.values()
                    for carrier in values
                )
                conflicts.append(
                    self._item_conflict(
                        batch_hash,
                        indexes[0],
                        first,
                        ConflictType.MULTIPLE_SOURCE_CARRIERS,
                        "Cùng một Số HĐ có nhiều bên vận tải trong JSON; hãy chọn một mã.",
                        (ResolutionAction.SELECT_CARRIER,),
                        default=ResolutionAction.SELECT_CARRIER,
                        details={
                            **details,
                            "carrier_candidates": ambiguous_candidates,
                            "incoming_carriers": incoming,
                            "ambiguous_invoice_carriers": ambiguous,
                        },
                    )
                )
                continue
            current = split_carriers(cell.value)
            if not current:
                for index in indexes:
                    items[index].carrier_action = (
                        ResolutionAction.OVERWRITE
                        if incoming
                        else ResolutionAction.KEEP_EXISTING
                    )
                    items[index].carrier_value_after = join_carriers(incoming)
                continue
            if not incoming and len(current) == 1:
                for index in indexes:
                    items[index].carrier_action = ResolutionAction.KEEP_EXISTING
                    items[index].carrier_value_after = cell.value
                continue
            if incoming and all(
                carrier_key(value) in {carrier_key(item) for item in current}
                for value in incoming
            ):
                for index in indexes:
                    items[index].carrier_action = ResolutionAction.KEEP_EXISTING
                    items[index].carrier_value_after = cell.value
                continue
            current_invoice_keys = {
                _invoice_key(value)
                for index in indexes
                for value in items[index].invoice_candidates
                if _invoice_key(value) is not None
                and _invoice_key(value) == _invoice_key(items[index].invoice_current_value)
            }
            incoming_invoice_keys = set(invoice_carriers)
            actions = [ResolutionAction.KEEP_EXISTING]
            if incoming:
                actions.append(ResolutionAction.OVERWRITE)
            if incoming and not current_invoice_keys.intersection(incoming_invoice_keys):
                actions.insert(1, ResolutionAction.APPEND_CARRIER)
            requires_existing_selection = (
                len(current) > 1
                and not any(
                    carrier_key(value) in {carrier_key(existing) for existing in current}
                    for value in incoming
                )
            )
            conflicts.append(
                self._item_conflict(
                    batch_hash,
                    indexes[0],
                    first,
                    ConflictType.CARRIER_VALUE_CONFLICT,
                    f"Ô bên vận tải nhóm {group} đang có mã khác với JSON.",
                    tuple(actions),
                    default=ResolutionAction.KEEP_EXISTING,
                    details={
                        **details,
                        "existing_carrier_candidates": current,
                        "requires_existing_carrier_selection": requires_existing_selection,
                    },
                )
            )
        return conflicts

    def _set_invoice_state(
        self,
        worksheet: Any,
        item: PostingItem,
        *,
        item_index: int,
        batch_hash: str,
        invoice_columns: Mapping[str, int],
    ) -> PostingConflict | None:
        chosen_action = item.invoice_action
        item.invoice_column = None
        item.invoice_cell = None
        item.invoice_current_value = None
        item.invoice_value_after = None

        if item.selected_fee == "LL" or not item.invoice_candidates:
            item.selected_invoice = None
            item.invoice_action = None
            return None

        if item.selected_invoice is not None and _invoice_key(
            item.selected_invoice
        ) not in {_invoice_key(value) for value in item.invoice_candidates}:
            item.selected_invoice = None

        invoice_column = invoice_columns.get(item.selected_fee)
        if invoice_column is None:
            if chosen_action in {
                ResolutionAction.SKIP_INVOICE,
                ResolutionAction.KEEP_EXISTING,
            }:
                item.invoice_action = chosen_action
                return None
            item.invoice_action = None
            return self._item_conflict(
                batch_hash,
                item_index,
                item,
                ConflictType.INVOICE_COLUMN_MISSING,
                "Không nhận diện được cột Số HĐ tương ứng với loại phí.",
                (ResolutionAction.SKIP_INVOICE, ResolutionAction.CANCEL_ALL),
                default=ResolutionAction.SKIP_INVOICE,
                details={
                    "scope": "invoice",
                    "invoice_candidates": list(item.invoice_candidates),
                },
            )

        item.invoice_column = invoice_column
        invoice_cell = worksheet.cell(int(item.target_row), invoice_column)
        item.invoice_cell = invoice_cell.coordinate
        item.invoice_current_value = invoice_cell.value

        if len(item.invoice_candidates) > 1 and item.selected_invoice is None:
            item.invoice_action = None
            return self._item_conflict(
                batch_hash,
                item_index,
                item,
                ConflictType.MULTIPLE_SOURCE_INVOICES,
                "Nhiều Số HĐ khác nhau cùng trỏ tới một ô BK; hãy chọn một Số HĐ.",
                (ResolutionAction.SELECT_INVOICE,),
                default=ResolutionAction.SELECT_INVOICE,
                details={
                    "scope": "invoice",
                    "invoice_candidates": list(item.invoice_candidates),
                    "invoice_column": invoice_column,
                    "invoice_cell": invoice_cell.coordinate,
                    "invoice_current_value": invoice_cell.value,
                },
            )

        selected_invoice = item.selected_invoice or item.invoice_candidates[0]
        item.selected_invoice = selected_invoice
        if _invoice_text(invoice_cell.value) is None:
            item.invoice_action = ResolutionAction.OVERWRITE
            item.invoice_value_after = selected_invoice
            return None
        if _invoice_key(invoice_cell.value) == _invoice_key(selected_invoice):
            item.invoice_action = ResolutionAction.KEEP_EXISTING
            item.invoice_value_after = invoice_cell.value
            return None

        if chosen_action is ResolutionAction.OVERWRITE:
            item.invoice_action = chosen_action
            # Ghi đúng chuỗi nguồn để bảo toàn số 0 đầu.
            item.invoice_value_after = str(selected_invoice).strip()
            return None
        if chosen_action in {
            ResolutionAction.KEEP_EXISTING,
            ResolutionAction.SKIP_INVOICE,
        }:
            item.invoice_action = chosen_action
            item.invoice_value_after = invoice_cell.value
            return None

        item.invoice_action = None
        return self._item_conflict(
            batch_hash,
            item_index,
            item,
            ConflictType.INVOICE_VALUE_CONFLICT,
            "Ô Số HĐ đang có giá trị khác với Số HĐ từ JSON.",
            (ResolutionAction.KEEP_EXISTING, ResolutionAction.OVERWRITE),
            default=ResolutionAction.KEEP_EXISTING,
            details={
                "scope": "invoice",
                "invoice_candidates": [selected_invoice],
                "selected_invoice": selected_invoice,
                "invoice_column": invoice_column,
                "invoice_cell": invoice_cell.coordinate,
                "invoice_current_value": invoice_cell.value,
            },
        )

    @staticmethod
    def _automatic_row(candidates: Sequence[RowCandidate]) -> RowCandidate | None:
        return candidates[0] if len(candidates) == 1 else None

    @staticmethod
    def _set_cell_state(worksheet: Any, item: PostingItem) -> None:
        cell = worksheet.cell(item.target_row, item.target_column)
        item.target_cell = cell.coordinate
        item.current_value = cell.value
        item.cell_state = classify_target_cell(cell, item.amount)
        if item.cell_state.kind in {TargetCellKind.EMPTY, TargetCellKind.ZERO}:
            item.action = ResolutionAction.OVERWRITE
        elif item.cell_state.kind is TargetCellKind.SAME_VALUE:
            if item.force_repost:
                item.action = ResolutionAction.OVERWRITE
            else:
                item.status = PostingItemStatus.ALREADY_EXISTS

    def _cell_conflict(
        self, batch_hash: str, item_index: int, item: PostingItem
    ) -> PostingConflict | None:
        if item.cell_state is None:
            return None
        if item.cell_state.kind is TargetCellKind.NUMBER:
            return self._item_conflict(
                batch_hash,
                item_index,
                item,
                ConflictType.TARGET_CELL_OCCUPIED,
                "Ô đích đang có số khác.",
                (
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.OVERWRITE,
                ),
            )
        if item.cell_state.kind is TargetCellKind.FORMULA:
            return self._item_conflict(
                batch_hash,
                item_index,
                item,
                ConflictType.TARGET_CELL_FORMULA,
                "Ô đích chứa công thức.",
                (
                    ResolutionAction.KEEP_FORMULA,
                    ResolutionAction.OVERWRITE,
                ),
            )
        if item.cell_state.kind is TargetCellKind.TEXT:
            return self._item_conflict(
                batch_hash,
                item_index,
                item,
                ConflictType.TARGET_CELL_TEXT,
                "Ô đích chứa văn bản.",
                (
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.OVERWRITE,
                ),
            )
        return None

    def _item_conflict(
        self,
        batch_hash: str,
        item_index: int,
        item: PostingItem,
        conflict_type: ConflictType,
        message: str,
        actions: tuple[ResolutionAction, ...],
        *,
        default: ResolutionAction | None = None,
        row_candidates: Sequence[RowCandidate] = (),
        details: Mapping[str, Any] | None = None,
    ) -> PostingConflict:
        conflict_details = dict(details or {})
        vessel_voyage = " ".join(
            str(value).strip()
            for value in (item.vessel_name, item.voyage_no)
            if value not in (None, "") and str(value).strip()
        )
        if not vessel_voyage and item.vessel_voyage_raw not in (None, ""):
            vessel_voyage = str(item.vessel_voyage_raw).strip()
        if item.source_items:
            source = item.source_items[0]
            conflict_details.setdefault(
                "source_document_id", source.get("source_document_id")
            )
            conflict_details.setdefault(
                "source_document_name", source.get("source_document_name")
            )
            conflict_details.setdefault("invoice_no", source.get("invoice_no"))
            conflict_details.setdefault("target_sheet", item.sheet_name)
        invoice_scope = conflict_details.get("scope") == "invoice"
        carrier_scope = conflict_details.get("scope") == "carrier"
        target_column = (
            conflict_details.get("invoice_column")
            if invoice_scope
            else conflict_details.get("carrier_column")
            if carrier_scope
            else item.target_column
        )
        target_cell = (
            conflict_details.get("invoice_cell")
            if invoice_scope
            else conflict_details.get("carrier_cell")
            if carrier_scope
            else item.target_cell
        )
        current_value = (
            conflict_details.get("invoice_current_value")
            if invoice_scope
            else conflict_details.get("carrier_current_value")
            if carrier_scope
            else item.current_value
        )
        return PostingConflict(
            conflict_id=_stable_id(
                conflict_type.value,
                batch_hash,
                item.source_indices,
                item.container,
                item.selected_fee,
                item.sheet_name,
                target_cell,
            ),
            conflict_type=conflict_type,
            message=message,
            item_index=item_index,
            container=item.container,
            bl=item.bl,
            fee=item.selected_fee,
            amount=item.amount,
            carrier=join_carriers(item.carrier_candidates),
            vessel_voyage=vessel_voyage or None,
            sheet_name=item.sheet_name,
            target_row=item.target_row,
            target_column=target_column,
            target_cell=target_cell,
            current_value=current_value,
            allowed_actions=actions,
            default_action=default,
            row_candidates=list(row_candidates),
            details=conflict_details,
        )

    def _resolve_apply_actions(
        self,
        worksheet: Any,
        items: list[PostingItem],
        conflicts: Sequence[PostingConflict],
        resolutions: Mapping[str, Any],
        *,
        batch_hash: str,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        if worksheet is None:
            return [], []
        base = self._resolve_base_headers(worksheet)
        fee_columns = self._resolve_fee_columns(worksheet, base)
        invoice_columns = self._resolve_invoice_columns(
            worksheet, base, fee_columns
        )
        index = self._container_index(worksheet, base)
        conflicts_by_item: dict[int, list[PostingConflict]] = defaultdict(list)
        for conflict in conflicts:
            if conflict.item_index is not None:
                conflicts_by_item[conflict.item_index].append(conflict)
        if any(
            conflict.conflict_type is ConflictType.MULTIPLE_EXPENSE_SAME_CELL
            for conflict in conflicts
        ):
            raise ExpensePostingError(
                "Xung đột nhiều dòng JSON cùng ô phí phải được refine trước khi ghi."
            )

        invoice_conflict_types = {
            ConflictType.MULTIPLE_SOURCE_INVOICES,
            ConflictType.INVOICE_VALUE_CONFLICT,
            ConflictType.INVOICE_COLUMN_MISSING,
        }
        amount_conflict_types = {
            ConflictType.TARGET_CELL_OCCUPIED,
            ConflictType.TARGET_CELL_FORMULA,
            ConflictType.TARGET_CELL_TEXT,
        }
        carrier_conflict_types = {
            ConflictType.MULTIPLE_SOURCE_CARRIERS,
            ConflictType.CARRIER_VALUE_CONFLICT,
            ConflictType.CARRIER_COLUMN_MISSING,
        }
        actions: list[dict[str, Any]] = []
        history: list[dict[str, Any]] = []
        for item_index, item in enumerate(items):
            selected_fee = item.selected_fee
            selected_row = item.target_row
            selected_invoice = item.selected_invoice
            amount_action = item.action
            invoice_action = item.invoice_action
            skip_status: PostingItemStatus | None = (
                item.status
                if item.status
                in {
                    PostingItemStatus.USER_SKIPPED,
                    PostingItemStatus.NOT_MATCHED,
                    PostingItemStatus.UNRESOLVED,
                }
                else None
            )

            for conflict in conflicts_by_item.get(item_index, ()):
                value = resolutions.get(conflict.conflict_id)
                action = (
                    self._action(value)
                    if value is not None
                    else conflict.default_action
                )
                if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                if conflict.conflict_type in carrier_conflict_types:
                    continue
                if conflict.conflict_type in invoice_conflict_types:
                    if action is ResolutionAction.SELECT_INVOICE:
                        selected = self._resolution_attr(value, "selected_invoice")
                        if _invoice_key(selected) not in {
                            _invoice_key(candidate)
                            for candidate in item.invoice_candidates
                        }:
                            raise ExpensePostingError("Số HĐ được chọn không hợp lệ.")
                        selected_invoice = str(selected)
                    elif action in {
                        ResolutionAction.KEEP_EXISTING,
                        ResolutionAction.OVERWRITE,
                        ResolutionAction.SKIP_INVOICE,
                    }:
                        invoice_action = action
                    continue
                if action is ResolutionAction.SELECT_FEE:
                    selected_fee = self._resolution_attr(value, "selected_fee")
                elif action is ResolutionAction.SELECT_ROW:
                    row_value = self._resolution_attr(value, "selected_row")
                    selected_row = int(row_value) if row_value is not None else None
                elif conflict.conflict_type in amount_conflict_types and action in {
                    ResolutionAction.OVERWRITE,
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.KEEP_FORMULA,
                    ResolutionAction.SKIP,
                }:
                    amount_action = action
                if conflict.conflict_type not in amount_conflict_types and action in {
                    None,
                    ResolutionAction.SKIP,
                    ResolutionAction.KEEP_EXISTING,
                    ResolutionAction.KEEP_FORMULA,
                }:
                    skip_status = PostingItemStatus.USER_SKIPPED

            if selected_fee not in fee_columns:
                skip_status = PostingItemStatus.UNRESOLVED
            if selected_row is None:
                selected = self._automatic_row(index.get(item.container or "", ()))
                selected_row = selected.row if selected else None
            if selected_row is None:
                skip_status = PostingItemStatus.NOT_MATCHED

            column = fee_columns.get(selected_fee)
            current_value = (
                worksheet.cell(selected_row, column).value
                if selected_row is not None and column is not None
                else None
            )
            amount_after = current_value
            amount_status = skip_status
            recorded_amount_action = amount_action or ResolutionAction.SKIP
            amount_write = False
            if skip_status is None and column is not None and selected_row is not None:
                cell = worksheet.cell(selected_row, column)
                state = classify_target_cell(cell, item.amount)
                if state.kind is TargetCellKind.SAME_VALUE and not item.force_repost:
                    recorded_amount_action = ResolutionAction.KEEP_EXISTING
                    amount_status = PostingItemStatus.ALREADY_EXISTS
                else:
                    chosen = amount_action
                    if state.kind in {TargetCellKind.EMPTY, TargetCellKind.ZERO} or (
                        state.kind is TargetCellKind.SAME_VALUE and item.force_repost
                    ):
                        chosen = ResolutionAction.OVERWRITE
                    if chosen is ResolutionAction.OVERWRITE:
                        amount_after = item.amount
                        amount_write = True
                        amount_status = PostingItemStatus.POSTED
                    elif chosen in {
                        ResolutionAction.SKIP,
                        ResolutionAction.KEEP_EXISTING,
                        ResolutionAction.KEEP_FORMULA,
                    }:
                        recorded_amount_action = chosen
                        amount_status = PostingItemStatus.USER_SKIPPED
                    else:
                        raise ExpensePostingError(
                            f"Ô tiền {cell.coordinate} chưa có quyết định được áp dụng."
                        )
                    recorded_amount_action = chosen or recorded_amount_action

            invoice_column: int | None = None
            invoice_cell: Any = None
            invoice_before: Any = None
            invoice_after: Any = None
            invoice_write = False
            recorded_invoice_action = invoice_action
            if (
                skip_status is None
                and selected_row is not None
                and selected_fee != "LL"
                and item.invoice_candidates
            ):
                if selected_invoice is None:
                    if len(item.invoice_candidates) > 1:
                        raise ExpensePostingError(
                            "Chưa chọn Số HĐ cho khoản có nhiều hóa đơn."
                        )
                    selected_invoice = item.invoice_candidates[0]
                invoice_column = invoice_columns.get(str(selected_fee))
                if invoice_column is not None:
                    invoice_cell = worksheet.cell(selected_row, invoice_column)
                    invoice_before = invoice_cell.value
                    invoice_after = invoice_before
                    if _invoice_key(invoice_before) == _invoice_key(selected_invoice):
                        recorded_invoice_action = ResolutionAction.KEEP_EXISTING
                    elif _invoice_text(invoice_before) is None:
                        recorded_invoice_action = ResolutionAction.OVERWRITE
                        invoice_after = selected_invoice
                        invoice_write = True
                    elif invoice_action is ResolutionAction.OVERWRITE:
                        recorded_invoice_action = ResolutionAction.OVERWRITE
                        invoice_after = selected_invoice
                        invoice_write = True
                    elif invoice_action in {
                        ResolutionAction.SKIP_INVOICE,
                        ResolutionAction.KEEP_EXISTING,
                    }:
                        recorded_invoice_action = (
                            ResolutionAction.SKIP_INVOICE
                            if invoice_action is ResolutionAction.SKIP_INVOICE
                            else ResolutionAction.KEEP_EXISTING
                        )
                    else:
                        raise ExpensePostingError(
                            f"Ô Số HĐ {invoice_cell.coordinate} chưa có quyết định được áp dụng."
                        )
                else:
                    if invoice_action not in {
                        ResolutionAction.SKIP_INVOICE,
                        ResolutionAction.KEEP_EXISTING,
                    }:
                        raise ExpensePostingError(
                            "Cột Số HĐ chưa được nhận diện và chưa có quyết định bỏ qua."
                        )
                    recorded_invoice_action = invoice_action

            final_status = amount_status or PostingItemStatus.USER_SKIPPED
            if amount_write or invoice_write:
                final_status = PostingItemStatus.POSTED
            action_record = self._action_record(
                item,
                str(selected_fee),
                selected_row,
                column,
                current_value,
                amount_after,
                recorded_amount_action,
                final_status,
            )
            action_record.update(
                {
                    "amount_write": amount_write,
                    "invoice_no": item.invoice_candidates[0]
                    if len(item.invoice_candidates) == 1
                    else None,
                    "invoice_selected": selected_invoice,
                    "invoice_target_column": invoice_column,
                    "invoice_target_cell": (
                        invoice_cell.coordinate if invoice_cell is not None else None
                    ),
                    "invoice_value_before": invoice_before,
                    "invoice_value_after": invoice_after,
                    "invoice_action": recorded_invoice_action,
                    "invoice_write": invoice_write,
                }
            )
            actions.append(action_record)
        self._resolve_carrier_actions(
            worksheet,
            items,
            actions,
            conflicts,
            resolutions,
        )
        for action, item in zip(actions, items, strict=True):
            history.extend(self._history_rows(action, item))
        return actions, history

    def _resolve_carrier_actions(
        self,
        worksheet: Any,
        items: Sequence[PostingItem],
        actions: list[dict[str, Any]],
        conflicts: Sequence[PostingConflict],
        resolutions: Mapping[str, Any],
    ) -> None:
        carrier_conflicts: dict[tuple[int, str], PostingConflict] = {}
        for conflict in conflicts:
            if conflict.conflict_type not in {
                ConflictType.MULTIPLE_SOURCE_CARRIERS,
                ConflictType.CARRIER_VALUE_CONFLICT,
                ConflictType.CARRIER_COLUMN_MISSING,
            }:
                continue
            row = conflict.target_row
            group = conflict.details.get("carrier_group")
            if row is not None and group in {"SEA", "ROAD", "NAM"}:
                carrier_conflicts[(int(row), str(group))] = conflict

        grouped: dict[tuple[int, str], list[int]] = defaultdict(list)
        for index, (item, action) in enumerate(zip(items, actions, strict=True)):
            group = carrier_group_for_fee(action.get("fee_selected"))
            row = action.get("target_row")
            if (
                group is None
                or row is None
                or (
                    item.carrier_column is None
                    and not item.carrier_candidates
                )
            ):
                continue
            grouped[(int(row), group)].append(index)

        base = self._resolve_base_headers(worksheet)
        carrier_columns = self._resolve_carrier_columns(base)
        for (row, group), indexes in grouped.items():
            column = carrier_columns.get(group)
            incoming = unique_carriers(
                carrier
                for index in indexes
                for carrier in (
                    [items[index].selected_carrier]
                    if items[index].selected_carrier
                    else items[index].carrier_candidates
                )
            )
            conflict = carrier_conflicts.get((row, group))
            selected_carrier: str | None = None
            chosen: ResolutionAction | None = None
            if conflict is not None:
                value = resolutions.get(conflict.conflict_id)
                chosen = self._action(value) if value is not None else conflict.default_action
                if chosen in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                if chosen is ResolutionAction.SELECT_CARRIER:
                    selected_carrier = self._resolution_attr(value, "selected_carrier")
                    if carrier_key(selected_carrier) not in {
                        carrier_key(candidate)
                        for candidate in conflict.details.get("carrier_candidates", ())
                    }:
                        raise ExpensePostingError("Bên vận tải được chọn không hợp lệ.")
                    ambiguous_invoices = {
                        _invoice_key(invoice)
                        for invoice in conflict.details.get(
                            "ambiguous_invoice_carriers", {}
                        )
                    }
                    incoming = unique_carriers(
                        carrier
                        for index in indexes
                        for carrier in (
                            [selected_carrier]
                            if ambiguous_invoices.intersection(
                                _invoice_key(invoice)
                                for invoice in items[index].invoice_candidates
                            )
                            else items[index].carrier_candidates
                        )
                    )
                    chosen = ResolutionAction.OVERWRITE
            if chosen is None:
                persisted_actions = {
                    items[index].carrier_action
                    for index in indexes
                    if items[index].carrier_action is not None
                }
                if len(persisted_actions) == 1:
                    chosen = next(iter(persisted_actions))
            if column is None:
                for index in indexes:
                    actions[index].update(
                        {
                            "carrier_group": group,
                            "carrier_source": join_carriers(items[index].carrier_candidates),
                            "carrier_effective": None,
                            "carrier_action": ResolutionAction.KEEP_EXISTING,
                            "carrier_write": False,
                        }
                    )
                continue
            cell = worksheet.cell(row, column)
            before = cell.value
            if chosen is None:
                if group in DAILY_MANAGED_CARRIER_GROUPS:
                    chosen = (
                        ResolutionAction.OVERWRITE
                        if incoming
                        else ResolutionAction.KEEP_EXISTING
                    )
                elif not split_carriers(before):
                    chosen = ResolutionAction.OVERWRITE
                elif all(
                    carrier_key(value) in {carrier_key(current) for current in split_carriers(before)}
                    for value in incoming
                ):
                    chosen = ResolutionAction.KEEP_EXISTING
                else:
                    raise ExpensePostingError(
                        f"Ô bên vận tải {cell.coordinate} chưa có quyết định được áp dụng."
                    )
            if chosen is ResolutionAction.OVERWRITE:
                after = join_carriers(incoming)
            elif chosen is ResolutionAction.APPEND_CARRIER:
                after = join_carriers(before, incoming)
            else:
                after = before
            write = carrier_text(after) != carrier_text(before)
            current_parts = split_carriers(before)
            selected_existing_carrier = None
            if (
                conflict is not None
                and chosen is ResolutionAction.KEEP_EXISTING
                and conflict.details.get("requires_existing_carrier_selection")
            ):
                value = resolutions.get(conflict.conflict_id)
                selected_existing_carrier = self._resolution_attr(
                    value, "selected_carrier"
                )
                if carrier_key(selected_existing_carrier) not in {
                    carrier_key(candidate) for candidate in current_parts
                }:
                    raise ExpensePostingError(
                        "Phải chọn một bên vận tải hiện có khi giữ nguyên ô có nhiều mã."
                    )
            for position, index in enumerate(indexes):
                source_carrier = (
                    items[index].selected_carrier
                    or join_carriers(items[index].carrier_candidates)
                )
                if chosen in {ResolutionAction.OVERWRITE, ResolutionAction.APPEND_CARRIER}:
                    effective = source_carrier
                elif len(current_parts) == 1:
                    effective = current_parts[0]
                elif carrier_key(source_carrier) in {carrier_key(value) for value in current_parts}:
                    effective = source_carrier
                elif selected_existing_carrier is not None:
                    effective = str(selected_existing_carrier)
                else:
                    effective = None
                actions[index].update(
                    {
                        "source_items": [dict(value) for value in items[index].source_items],
                        "carrier_group": group,
                        "carrier_source": source_carrier,
                        "carrier_effective": effective,
                        "carrier_target_column": column,
                        "carrier_target_cell": cell.coordinate,
                        "carrier_value_before": before,
                        "carrier_value_after": after,
                        "carrier_action": chosen,
                        "carrier_write": bool(write and position == 0),
                    }
                )

    @staticmethod
    def _action_record(
        item: PostingItem,
        fee: str,
        row: int | None,
        column: int | None,
        before: Any,
        after: Any,
        action: ResolutionAction,
        status: PostingItemStatus,
    ) -> dict[str, Any]:
        return {
            "source_indices": list(item.source_indices),
            "container": item.container,
            "fee_original": item.original_fee,
            "fee_selected": fee,
            "sheet_name": item.sheet_name,
            "selected_source_sheet": item.selected_source_sheet,
            "selected_source_row": item.selected_source_row,
            "source_sqt": item.source_sqt,
            "plan_values": tuple(item.plan_values),
            "source_signature": item.source_signature,
            "carry_forward_required": item.carry_forward_required,
            "target_row": row,
            "target_column": column,
            "target_cell": (
                f"{get_column_letter(column)}{row}" if row and column else None
            ),
            "value_before": before,
            "value_after": after,
            "action": action,
            "status": status,
            "source_items": [dict(value) for value in item.source_items],
        }

    @staticmethod
    def _history_rows(
        action: Mapping[str, Any], item: PostingItem
    ) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        sources = {
            source["source_item_index"]: source for source in item.source_items
        }
        for source_index in item.source_indices:
            source = sources.get(source_index, {})
            result.append(
                {
                    "source_item_index": source_index,
                    "origin_batch_id": source.get("origin_batch_id"),
                    "origin_batch_hash": source.get("origin_batch_hash"),
                    "origin_source_item_index": source.get(
                        "origin_source_item_index", source_index
                    ),
                    "reconciliation_group_id": source.get(
                        "reconciliation_group_id"
                    ),
                    "container": source.get("container", item.container),
                    "bl": source.get("bl", item.bl),
                    "fee_original": source.get("fee", item.original_fee),
                    "fee_selected": action["fee_selected"],
                    "rule": source.get("rule", item.rule),
                    "amount": source.get("amount", item.amount),
                    "sheet_name": action["sheet_name"],
                    "target_row": action["target_row"],
                    "target_column": action["target_column"],
                    "target_cell": action["target_cell"],
                    "value_before": action["value_before"],
                    "value_after": action["value_after"],
                    "action": action["action"],
                    "invoice_no": source.get("invoice_no"),
                    "invoice_selected": action.get("invoice_selected"),
                    "invoice_target_column": action.get("invoice_target_column"),
                    "invoice_target_cell": action.get("invoice_target_cell"),
                    "invoice_value_before": action.get("invoice_value_before"),
                    "invoice_value_after": action.get("invoice_value_after"),
                    "invoice_action": action.get("invoice_action"),
                    "carrier_source": source.get("carrier"),
                    "carrier_effective": action.get("carrier_effective"),
                    "carrier_group": action.get("carrier_group"),
                    "carrier_target_column": action.get("carrier_target_column"),
                    "carrier_target_cell": action.get("carrier_target_cell"),
                    "carrier_value_before": action.get("carrier_value_before"),
                    "carrier_value_after": action.get("carrier_value_after"),
                    "carrier_action": action.get("carrier_action"),
                    "status": action["status"],
                }
            )
        return result

    def _write_carried_plan_rows(
        self,
        worksheet: Any,
        actions: Sequence[Mapping[str, Any]],
    ) -> list[Mapping[str, Any]]:
        unique: dict[tuple[str, int, int | None, str | None], Mapping[str, Any]] = {}
        for action in actions:
            if not action.get("carry_forward_required"):
                continue
            key = (
                str(action.get("selected_source_sheet") or ""),
                int(action.get("selected_source_row") or 0),
                action.get("source_sqt"),
                action.get("container"),
            )
            unique.setdefault(key, action)
        carried = sorted(unique.values(), key=lambda value: int(value["target_row"]))
        if not carried:
            return []

        target_header = self._plan_header(worksheet)
        last_data_row = DailySyncService._last_data_row(worksheet, target_header)
        expected_rows = list(
            range(last_data_row + 1, last_data_row + len(carried) + 1)
        )
        actual_rows = [int(action["target_row"]) for action in carried]
        if actual_rows != expected_rows:
            raise ExpensePostingError(
                "Vị trí dòng kế hoạch cần mang sang đã thay đổi; hãy phân tích lại."
            )
        template_row = max(target_header.row_end + 1, last_data_row)
        max_column = DailySyncService._actual_max_column(worksheet)
        for action in carried:
            target_row = int(action["target_row"])
            values = tuple(action.get("plan_values") or ())
            if len(values) != len(SYNC_FIELDS):
                raise ExpensePostingError("Dòng nguồn không đủ 12 trường kế hoạch.")
            DailySyncService._copy_row_style(
                worksheet,
                template_row,
                target_row,
                max_column=max_column,
            )
            for index, field in enumerate(SYNC_FIELDS):
                worksheet.cell(
                    target_row, TARGET_EXPECTED_COLUMNS[field]
                ).value = values[index]
        return carried

    def _record_carry_forwards(
        self,
        plan: PostingPlan,
        actions: Sequence[Mapping[str, Any]],
    ) -> None:
        saver = getattr(self.posting_repository, "save_carry_forward", None)
        if not callable(saver):
            return
        saved: set[tuple[str, int, int | None, str | None]] = set()
        for action in actions:
            source_sheet = str(action.get("selected_source_sheet") or "")
            source_row = int(action.get("selected_source_row") or 0)
            source_sqt = action.get("source_sqt")
            container = action.get("container")
            if (
                not source_sheet
                or source_sheet == action.get("sheet_name")
                or source_row <= 0
                or not isinstance(source_sqt, int)
                or source_sqt <= 0
                or not container
            ):
                continue
            key = (source_sheet, source_row, source_sqt, str(container))
            if key in saved:
                continue
            saver(
                workbook_path=plan.target_path,
                source_sheet=source_sheet,
                source_row=source_row,
                source_sqt=source_sqt,
                container=str(container),
                source_signature=str(action.get("source_signature") or ""),
                target_sheet=str(action.get("sheet_name") or ""),
                target_row=int(action["target_row"]),
            )
            saved.add(key)

    def _verify_posting(
        self,
        path: Path,
        sheet_name: str,
        actions: Sequence[Mapping[str, Any]],
        *,
        detail_rows: Sequence[Mapping[str, Any]] = (),
        update_column: int | None = None,
        update_timestamp: datetime | None = None,
    ) -> None:
        workbook = self.gateway.load(path, read_only=False)
        try:
            worksheet = workbook[sheet_name]
            base = self._resolve_base_headers(worksheet)
            fee_columns = self._resolve_fee_columns(worksheet, base)
            invoice_columns = self._resolve_invoice_columns(
                worksheet, base, fee_columns
            )
            carrier_columns = self._resolve_carrier_columns(base)
            target_plan_header = (
                self._plan_header(worksheet)
                if any(action.get("carry_forward_required") for action in actions)
                else None
            )
            for action in actions:
                if action.get("carry_forward_required"):
                    actual_plan_values = self._row_plan_values(
                        worksheet,
                        target_plan_header,
                        int(action["target_row"]),
                    )
                    if actual_plan_values != tuple(action.get("plan_values") or ()):
                        raise ExpensePostingError(
                            f"Dòng kế hoạch {action['target_row']} không qua verify."
                        )
                fee = action["fee_selected"]
                column = int(action["target_column"])
                if fee_columns.get(fee) != column:
                    raise ExpensePostingError(f"Cột phí {fee} không qua verify.")
                actual = worksheet.cell(int(action["target_row"]), column).value
                if action.get("amount_write") and actual != action["value_after"]:
                    raise ExpensePostingError(
                        f"Ô {action['target_cell']} không có giá trị dự kiến."
                    )
                if action.get("invoice_write"):
                    invoice_column = int(action["invoice_target_column"])
                    if invoice_columns.get(fee) != invoice_column:
                        raise ExpensePostingError(
                            f"Cột Số HĐ của phí {fee} không qua verify."
                        )
                    invoice_actual = worksheet.cell(
                        int(action["target_row"]), invoice_column
                    ).value
                    if invoice_actual != action["invoice_value_after"]:
                        raise ExpensePostingError(
                            f"Ô {action['invoice_target_cell']} không có Số HĐ dự kiến."
                        )
                if action.get("carrier_write"):
                    carrier_group = str(action["carrier_group"])
                    carrier_column = int(action["carrier_target_column"])
                    if carrier_columns.get(carrier_group) != carrier_column:
                        raise ExpensePostingError(
                            f"Cột bên vận tải {carrier_group} không qua verify."
                        )
                    carrier_actual = worksheet.cell(
                        int(action["target_row"]), carrier_column
                    ).value
                    if carrier_text(carrier_actual) != carrier_text(
                        action.get("carrier_value_after")
                    ):
                        raise ExpensePostingError(
                            f"Ô {action['carrier_target_cell']} không có bên vận tải dự kiến."
                        )
            if actions:
                if update_column is None or update_timestamp is None:
                    raise ExpensePostingError("Bản lưu thiếu thông tin Date cập nhật.")
                actual_update_column = self._update_column(base)
                if actual_update_column != update_column:
                    raise ExpensePostingError("Cột Date cập nhật không qua verify.")
                for target_row in {
                    int(action["target_row"]) for action in actions
                }:
                    actual = worksheet.cell(target_row, update_column).value
                    if actual != update_timestamp:
                        raise ExpensePostingError(
                            f"Date cập nhật tại dòng {target_row} không đúng."
                        )
            if detail_rows:
                actual_details = read_bk_detail_rows(workbook, sheet_name)
                actual_by_key = {
                    (str(row["Batch hash"]), int(row["Dòng JSON"])): row
                    for row in actual_details
                }
                for expected in detail_rows:
                    key = (
                        str(expected["batch_hash"]),
                        int(expected["source_item_index"]),
                    )
                    actual_detail = actual_by_key.get(key)
                    if actual_detail is None:
                        raise ExpensePostingError(
                            "Thiếu dòng chi tiết bên vận tải sau khi lưu BK."
                        )
                    if (
                        carrier_text(actual_detail["Carrier hiệu lực"])
                        != carrier_text(
                            expected.get("carrier_effective") or "CHƯA XÁC ĐỊNH"
                        )
                        or actual_detail["Số tiền"] != expected.get("amount")
                    ):
                        raise ExpensePostingError(
                            "Chi tiết bên vận tải không đúng sau khi lưu BK."
                        )
        finally:
            workbook.close()

    def _record_history(
        self, plan: PostingPlan, history: Sequence[Mapping[str, Any]]
    ) -> None:
        if (
            self.posting_repository is None
            or plan.run_id is None
            or plan.batch_id is None
            or not history
        ):
            return
        grouped: dict[tuple[int, str], list[dict[str, Any]]] = defaultdict(list)
        for raw in history:
            payload = dict(raw)
            origin_batch = payload.pop("origin_batch_id", None)
            origin_hash = payload.pop("origin_batch_hash", None)
            origin_index = payload.pop("origin_source_item_index", None)
            payload.pop("reconciliation_group_id", None)
            effective_batch = int(origin_batch or plan.batch_id)
            effective_hash = str(origin_hash or plan.batch_hash)
            if origin_index is not None:
                payload["source_item_index"] = int(origin_index)
            grouped[(effective_batch, effective_hash)].append(payload)
        for (origin_batch, origin_hash), payloads in grouped.items():
            self.posting_repository.create_items(
                payloads,
                run_id=plan.run_id,
                batch_id=origin_batch,
                batch_hash=origin_hash,
            )

    def _mark_completed_reconciliations(self, plan: PostingPlan) -> None:
        if (
            self.sea_freight_repository is None
            or self.posting_repository is None
            or plan.run_id is None
        ):
            return
        lookup = getattr(
            self.posting_repository, "successful_source_indices_for_batch", None
        )
        if not callable(lookup):
            return
        for member in plan.reconciliation_members:
            group_id = int(member["group_id"])
            child_batch_id = int(member["batch_id"])
            row_count = int(member["row_count"])
            successful = set(lookup(child_batch_id))
            if not set(range(row_count)).issubset(successful):
                continue
            group = self.sea_freight_repository.get_group(group_id)
            if group is not None and str(group.status.value) == "ALLOCATED":
                self.sea_freight_repository.mark_posted(group_id, plan.run_id)

    def _ready_path(self, batch_id: int | None) -> Path | None:
        if batch_id is None:
            return self.provider.get_latest_ready_json_path()
        return self.provider.get_ready_json_path(batch_id)

    def _batch_id_for_path(self, path: Path) -> int | None:
        list_method = getattr(self.provider, "list_ready_batches", None)
        if not callable(list_method):
            return None
        resolved = path.resolve()
        for metadata in list_method():
            current_path = getattr(metadata, "ready_path", None)
            if current_path is not None and Path(current_path).resolve() == resolved:
                return int(getattr(metadata, "id", getattr(metadata, "batch_id")))
        return None

    def _successful_indices(self, batch_hash: str) -> set[int]:
        if self.posting_repository is None:
            return set()
        return set(self.posting_repository.successful_source_indices(batch_hash))

    def _previously_posted_items(
        self,
        batch_hash: str,
        document_rows: Sequence[Mapping[str, Any]],
        successful_indices: set[int],
    ) -> list[dict[str, Any]]:
        records: Sequence[Any] = ()
        if document_rows and any(
            row.get("origin_batch_id") is not None for row in document_rows
        ):
            lookup = getattr(
                self.posting_repository,
                "latest_successful_items_for_batch",
                None,
            )
            cache: dict[int, dict[int, Any]] = {}
            result: list[dict[str, Any]] = []
            for source_index in sorted(successful_indices):
                if source_index >= len(document_rows):
                    continue
                source = document_rows[source_index]
                origin_batch = source.get("origin_batch_id")
                if origin_batch is None or not callable(lookup):
                    continue
                origin_batch = int(origin_batch)
                if origin_batch not in cache:
                    cache[origin_batch] = {
                        int(getattr(record, "source_item_index")): record
                        for record in lookup(origin_batch)
                    }
                record = cache[origin_batch].get(
                    int(source.get("origin_source_item_index", source_index))
                )

                def previous(name: str, default: Any = None) -> Any:
                    return getattr(record, name, default) if record is not None else default

                result.append(
                    {
                        "source_item_index": source_index,
                        "container": previous("container", source.get("container")),
                        "bl": previous("bl", source.get("bl")),
                        "fee": previous("fee_selected", source.get("fee")) or source.get("fee"),
                        "amount": previous("amount", source.get("amount")),
                        "sheet_name": previous("sheet_name"),
                        "target_row": previous("target_row"),
                        "target_cell": previous("target_cell"),
                        "created_at": previous("created_at"),
                    }
                )
            return result
        latest = getattr(
            self.posting_repository,
            "latest_successful_items",
            None,
        )
        if callable(latest):
            records = latest(batch_hash)
        by_index = {
            int(getattr(record, "source_item_index")): record
            for record in records
        }
        result: list[dict[str, Any]] = []
        for source_index in sorted(successful_indices):
            if source_index >= len(document_rows):
                continue
            source = document_rows[source_index]
            record = by_index.get(source_index)

            def previous(name: str, default: Any = None) -> Any:
                return getattr(record, name, default) if record is not None else default

            result.append(
                {
                    "source_item_index": source_index,
                    "container": previous("container", source.get("container")),
                    "bl": previous("bl", source.get("bl")),
                    "fee": previous(
                        "fee_selected",
                        source.get("fee"),
                    )
                    or source.get("fee"),
                    "amount": previous("amount", source.get("amount")),
                    "sheet_name": previous("sheet_name"),
                    "target_row": previous("target_row"),
                    "target_cell": previous("target_cell"),
                    "created_at": previous("created_at"),
                }
            )
        return result

    def _latest_sync_sheet(self) -> str | None:
        if self.run_repository is None:
            return None
        return self.run_repository.get_latest_sync_sheet()

    @staticmethod
    def _positive_int(value: Any) -> int | None:
        if isinstance(value, bool):
            return None
        if isinstance(value, int) and value > 0:
            return value
        if isinstance(value, float) and value.is_integer() and value > 0:
            return int(value)
        if isinstance(value, str) and value.strip().isdigit():
            parsed = int(value.strip())
            return parsed if parsed > 0 else None
        return None

    def _selected_sheet(
        self, plan: PostingPlan, resolutions: Mapping[str, Any]
    ) -> str | None:
        if plan.selected_sheet is not None:
            return plan.selected_sheet
        for conflict in plan.conflicts:
            if conflict.conflict_type is ConflictType.TARGET_SHEET_AMBIGUOUS:
                value = resolutions.get(conflict.conflict_id)
                if value is None:
                    continue
                action = self._action(value)
                if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                    raise ExpensePostingError("Người dùng đã hủy nhập khoản chi.")
                if action is ResolutionAction.SELECT_SHEET:
                    selected = self._resolution_attr(value, "selected_sheet")
                    valid = {c.target_sheet for c in plan.sheet_candidates}
                    if selected not in valid:
                        raise ExpensePostingError("Sheet được chọn không hợp lệ.")
                    return str(selected)
        return None

    def _has_selector_resolution(
        self,
        plan: PostingPlan,
        resolutions: Mapping[str, Any],
    ) -> bool:
        for conflict in plan.conflicts:
            value = resolutions.get(conflict.conflict_id)
            if value is None:
                continue
            if self._action(value) in {
                ResolutionAction.SELECT_SHEET,
                ResolutionAction.SELECT_ROW,
                ResolutionAction.SELECT_FEE,
                ResolutionAction.SELECT_INVOICE,
                ResolutionAction.SELECT_CARRIER,
                ResolutionAction.SELECT_SOURCE_ITEM,
            }:
                return True
        return False

    def _check_batch_resolution(
        self, plan: PostingPlan, resolutions: Mapping[str, Any]
    ) -> None:
        for conflict in plan.conflicts:
            if conflict.conflict_type is not ConflictType.BATCH_ALREADY_POSTED:
                continue
            value = resolutions.get(conflict.conflict_id)
            action = self._action(value) if value is not None else conflict.default_action
            if action in {ResolutionAction.CANCEL, ResolutionAction.CANCEL_ALL}:
                raise ExpensePostingError("Người dùng đã hủy nhập lại batch.")

    @staticmethod
    def _action(value: Any) -> ResolutionAction:
        if isinstance(value, (ResolutionAction, str)):
            return ResolutionAction(value)
        action = getattr(value, "action", None)
        if isinstance(value, Mapping):
            action = value.get("action", action)
        return ResolutionAction(action)

    @staticmethod
    def _resolution_attr(value: Any, name: str) -> Any:
        result = getattr(value, name, None)
        if isinstance(value, Mapping):
            result = value.get(name, result)
        return result

    def _create_run(self, source: Path, target: Path) -> int | None:
        if self.run_repository is None:
            return None
        record = self.run_repository.create_run(
            operation=ExcelOperation.EXPENSE_POSTING,
            source_path=source,
            target_path=target,
            status=ExcelRunStatus.ANALYZING,
        )
        return int(getattr(record, "id", record))

    def _update_run(self, run_id: int | None, **changes: Any) -> None:
        if self.run_repository is not None and run_id is not None:
            self.run_repository.update_run(run_id, **changes)

    def _finish_result(self, result: PostingResult) -> None:
        if self.run_repository is None or result.run_id is None:
            return
        self.run_repository.finish_run(
            result.run_id,
            status=result.status,
            sheet_name=result.sheet_name,
            backup_path=result.backup_path,
            target_fingerprint_after=result.fingerprint_after,
            total_items=(
                result.posted_source_items
                + result.skipped_source_items
                + result.already_existing_items
            ),
            changed_items=result.posted_source_items,
            skipped_items=result.skipped_source_items,
            conflict_count=result.conflict_count,
            item_outcomes=result.item_outcomes,
        )

    def _finish_failed(self, run_id: int | None, exc: Exception) -> None:
        if self.run_repository is not None and run_id is not None:
            cancelled = "người dùng đã hủy" in str(exc).casefold()
            self.run_repository.finish_run(
                run_id,
                status=(
                    ExcelRunStatus.CANCELLED
                    if cancelled
                    else ExcelRunStatus.FAILED
                ),
                error_message=None if cancelled else str(exc),
            )
