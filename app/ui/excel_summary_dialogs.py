"""Dialog tóm tắt thống nhất cho ba luồng xử lý Excel.

Module này chỉ chuyển ``Plan``/``Result`` thành dữ liệu trình bày và dựng UI.
Nó không thay đổi quyết định nghiệp vụ hoặc trực tiếp ghi workbook.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True, slots=True)
class SummaryMetric:
    label: str
    value: int | str
    tone: str = "primary"
    note: str = ""


@dataclass(frozen=True, slots=True)
class SummaryNotice:
    text: str
    tone: str = "info"


@dataclass(frozen=True, slots=True)
class SummaryTable:
    title: str
    headers: tuple[str, ...]
    rows: tuple[tuple[Any, ...], ...]


@dataclass(frozen=True, slots=True)
class SummaryDetails:
    title: str
    lines: tuple[str, ...]
    expanded: bool = False
    tone: str = "neutral"


@dataclass(frozen=True, slots=True)
class ExcelOperationSummary:
    title: str
    subtitle: str
    state: str = "preview"
    metrics: tuple[SummaryMetric, ...] = ()
    notices: tuple[SummaryNotice, ...] = ()
    tables: tuple[SummaryTable, ...] = ()
    details: tuple[SummaryDetails, ...] = ()
    change_count: int = 0
    issue_count: int = 0


def _attribute(source: Any, *names: str, default: Any = None) -> Any:
    if source is None:
        return default
    if isinstance(source, Mapping):
        for name in names:
            if name in source:
                return source[name]
        return default
    for name in names:
        if hasattr(source, name):
            return getattr(source, name)
    return default


def _sequence(value: Any) -> tuple[Any, ...]:
    if value is None:
        return ()
    if isinstance(value, Mapping):
        return tuple(value.values())
    if isinstance(value, Sequence) and not isinstance(
        value, (str, bytes, bytearray)
    ):
        return tuple(value)
    try:
        return tuple(value)
    except TypeError:
        return (value,)


def _code(value: Any) -> str:
    raw = getattr(value, "value", value)
    return str(raw or "").split(".")[-1].upper()


def _integer(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _display(value: Any) -> str:
    return "—" if value in (None, "") else str(value)


def _item_count(item: Any) -> int:
    source_indices = _sequence(_attribute(item, "source_indices", default=()))
    return len(source_indices) or 1


def _outcome_status_counts(outcomes: Sequence[Any]) -> dict[str, int]:
    result = {
        code: 0
        for code in (
            "WRITTEN",
            "PARTIAL",
            "UNCHANGED",
            "USER_KEPT",
            "USER_SKIPPED",
            "INVALID_SOURCE",
            "FAILED",
        )
    }
    for outcome in outcomes:
        code = _code(_attribute(outcome, "status", default=""))
        if code in result:
            result[code] += 1
    return result


def _posting_bucket(item: Any) -> str:
    status = _code(_attribute(item, "status", default=""))
    action = _code(_attribute(item, "action", default=""))
    cell_state = _attribute(item, "cell_state", default=None)
    cell_kind = _code(_attribute(cell_state, "kind", default=""))
    if status == "ALREADY_EXISTS":
        return "already"
    if action in {"KEEP_EXISTING", "KEEP_FORMULA"}:
        return "keep"
    if action == "SKIP" or status in {"USER_SKIPPED", "NOT_MATCHED", "UNRESOLVED"}:
        return "skip"
    if status == "PLANNED":
        return "write_new" if cell_kind in {"EMPTY", "ZERO"} else "overwrite"
    return "skip"


def daily_confirmation_summary(
    plan: Any,
    *,
    candidates: Sequence[Any] = (),
) -> ExcelOperationSummary:
    selected = list(candidates)
    if not selected:
        selected = list(_sequence(_attribute(plan, "month_candidates", default=())))
    rows: list[tuple[Any, ...]] = []
    totals = defaultdict(int)
    for candidate in selected:
        values = {
            "update": _integer(_attribute(candidate, "update_count", default=0)),
            "insert": _integer(_attribute(candidate, "new_row_count", default=0)),
            "unchanged": _integer(_attribute(candidate, "unchanged_count", default=0)),
            "target_only": _integer(_attribute(candidate, "target_only_count", default=0)),
            "invalid": _integer(_attribute(candidate, "invalid_count", default=0)),
        }
        for key, value in values.items():
            totals[key] += value
        rows.append(
            (
                _attribute(candidate, "source_sheet", default="—"),
                _attribute(candidate, "target_sheet", "sheet_name", default="—"),
                values["update"],
                values["insert"],
                values["unchanged"],
                values["target_only"],
                values["invalid"],
            )
        )
    if not rows:
        totals.update(
            update=_integer(_attribute(plan, "update_count", default=0)),
            insert=_integer(_attribute(plan, "insert_count", default=0)),
            unchanged=_integer(_attribute(plan, "unchanged_count", default=0)),
            target_only=_integer(_attribute(plan, "target_only_count", default=0)),
            invalid=_integer(_attribute(plan, "invalid_count", default=0)),
        )
        rows.append(
            (
                "—",
                _attribute(plan, "selected_sheet", default="—"),
                totals["update"],
                totals["insert"],
                totals["unchanged"],
                totals["target_only"],
                totals["invalid"],
            )
        )
    changes = totals["update"] + totals["insert"]
    notices: list[SummaryNotice] = []
    if totals["target_only"]:
        notices.append(
            SummaryNotice(
                f"{totals['target_only']} dòng chỉ có ở BK sẽ được giữ nguyên.",
                "info",
            )
        )
    if totals["invalid"]:
        notices.append(
            SummaryNotice(
                f"{totals['invalid']} dòng thiếu hoặc sai SQT sẽ không được đồng bộ.",
                "warning",
            )
        )
    return ExcelOperationSummary(
        title="Xác nhận đồng bộ Hàng ngày → BK",
        subtitle=(
            f"Kiểm tra {changes} thay đổi trước khi ghi vào "
            f"{len(rows)} sheet BK."
        ),
        state="warning" if totals["invalid"] else "preview",
        metrics=(
            SummaryMetric("Cập nhật", totals["update"], "primary", "dòng"),
            SummaryMetric("Thêm mới", totals["insert"], "success", "dòng"),
            SummaryMetric("Đã giống nhau", totals["unchanged"], "neutral", "dòng"),
            SummaryMetric("Thiếu / sai SQT", totals["invalid"], "warning", "dòng"),
        ),
        notices=tuple(notices),
        tables=(
            SummaryTable(
                "Phân bổ theo sheet",
                ("Sheet nguồn", "Sheet BK", "Cập nhật", "Thêm", "Đã giống", "Giữ ở BK", "Sai SQT"),
                tuple(rows),
            ),
        ),
        change_count=changes,
        issue_count=totals["invalid"],
    )


def posting_confirmation_summary(plan: Any) -> ExcelOperationSummary:
    sheet_counts: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    changed_cells = 0
    for item in _sequence(_attribute(plan, "items", default=())):
        bucket = _posting_bucket(item)
        count = _item_count(item)
        sheet = str(
            _attribute(
                item,
                "sheet_name",
                default=_attribute(plan, "selected_sheet", default="Chưa xác định"),
            )
            or "Chưa xác định"
        )
        sheet_counts[sheet][bucket] += count
        if bucket in {"write_new", "overwrite"}:
            changed_cells += 1
    totals = defaultdict(int)
    rows: list[tuple[Any, ...]] = []
    for sheet in sorted(sheet_counts):
        counts = sheet_counts[sheet]
        for key, value in counts.items():
            totals[key] += value
        rows.append(
            (
                sheet,
                counts["write_new"],
                counts["overwrite"],
                counts["already"],
                counts["keep"],
                counts["skip"],
            )
        )
    write_count = totals["write_new"] + totals["overwrite"]
    no_write = totals["keep"] + totals["skip"]
    declared_sheets = {
        str(value)
        for value in _sequence(_attribute(plan, "target_sheets", default=()))
        if str(value or "").strip()
    }
    selected_sheet = _attribute(plan, "selected_sheet", default=None)
    if selected_sheet not in (None, ""):
        declared_sheets.add(str(selected_sheet))
    sheet_total = len(declared_sheets) or len(rows) or 1
    notices: list[SummaryNotice] = []
    if totals["overwrite"]:
        notices.append(
            SummaryNotice(
                f"{totals['overwrite']} khoản sẽ thay thế số tiền đang có trong file BK.",
                "warning",
            )
        )
    if not write_count:
        notices.append(SummaryNotice("Không có ô tiền nào cần thay đổi.", "info"))
    return ExcelOperationSummary(
        title="Xác nhận nhập khoản chi vào BK",
        subtitle=(
            f"{changed_cells} ô BK sẽ thay đổi trên "
            f"{sheet_total} sheet."
        ),
        state="warning" if totals["overwrite"] else "preview",
        metrics=(
            SummaryMetric("Ô BK thay đổi", changed_cells, "primary", "ô"),
            SummaryMetric("Khoản sẽ ghi", write_count, "success", "khoản"),
            SummaryMetric("Đã đúng sẵn", totals["already"], "neutral", "khoản"),
            SummaryMetric("Không ghi", no_write, "warning", "khoản"),
        ),
        notices=tuple(notices),
        tables=(
            SummaryTable(
                "Kết quả theo sheet",
                ("Sheet BK", "Ghi mới", "Ghi đè", "Đúng sẵn", "Giữ lựa chọn", "Bỏ qua"),
                tuple(rows),
            ),
        ) if rows else (),
        change_count=write_count,
        issue_count=totals["overwrite"],
    )


def _payment_plan_entries(plan: Any) -> list[tuple[str, Any]]:
    month_plans = _attribute(plan, "month_plans", default={}) or {}
    if isinstance(month_plans, Mapping) and month_plans:
        return [(str(source), month_plan) for source, month_plan in month_plans.items()]
    return [(str(_attribute(plan, "source_sheet", default="—")), plan)]


def payment_confirmation_summary(
    plan: Any,
    *,
    selected_new_ids: Sequence[str] | None = None,
) -> ExcelOperationSummary:
    selected_ids = None if selected_new_ids is None else {str(value) for value in selected_new_ids}
    totals = defaultdict(int)
    rows: list[tuple[Any, ...]] = []
    created_rows: list[str] = []
    carrier_changes = 0
    for source_sheet, month_plan in _payment_plan_entries(plan):
        targets = _attribute(month_plan, "targets", default={}) or {}
        for target_type in ("HP", "NAM"):
            target = _attribute(targets, target_type, default=None)
            if target is None:
                continue
            items = list(_sequence(_attribute(target, "items", default=())))
            new_items = list(
                _sequence(_attribute(target, "new_rows", default=()))
            ) or [item for item in items if bool(_attribute(item, "is_new", default=False))]
            declared_new_count = _integer(
                _attribute(target, "new_count", default=len(new_items))
            )
            if new_items:
                selected_new = sum(
                    1
                    for item in new_items
                    if selected_ids is None
                    or str(_attribute(item, "item_id", "id", default="")) in selected_ids
                )
            else:
                selected_new = declared_new_count if selected_ids is None else 0
            skipped_new = max(0, declared_new_count - selected_new)
            update_count = _integer(_attribute(target, "update_count", default=0))
            unchanged = _integer(_attribute(target, "unchanged_count", default=0))
            invoices = _integer(_attribute(target, "invoice_change_count", default=0))
            created = bool(_attribute(target, "sheet_to_create", default=False))
            template = _attribute(target, "template_sheet", default="—")
            target_name = _attribute(target, "sheet_name", default=target_type)
            if created:
                created_rows.append(f"{target_name} sẽ được tạo từ sheet mẫu {template}.")
            carrier_changes += sum(
                _attribute(item, "carrier_difference", default=None) is not None
                or _attribute(item, "carrier_review_value", default=None) is not None
                for item in items
            )
            totals["update"] += update_count
            totals["selected_new"] += selected_new
            totals["skipped_new"] += skipped_new
            totals["unchanged"] += unchanged
            totals["invoice"] += invoices
            rows.append(
                (
                    source_sheet,
                    target_name,
                    f"Tạo từ {template}" if created else "Đã có",
                    update_count,
                    selected_new,
                    unchanged,
                    invoices,
                )
            )
    normalize_sheets = _integer(
        _attribute(plan, "normalization_sheet_count", default=0)
    )
    conflicts = _integer(_attribute(plan, "conflict_count", default=0))
    notices = [SummaryNotice(text, "info") for text in created_rows]
    if normalize_sheets:
        notices.append(
            SummaryNotice(
                f"{normalize_sheets} sheet BK cần được chuẩn hóa trước khi đồng bộ.",
                "warning",
            )
        )
    if carrier_changes:
        notices.append(
            SummaryNotice(
                f"Bên vận tải sẽ được cập nhật tại {carrier_changes} dòng.",
                "info",
            )
        )
    if conflicts:
        notices.append(
            SummaryNotice(
                f"Còn {conflicts} xung đột cần xử lý trước khi ghi file.",
                "error",
            )
        )
    return ExcelOperationSummary(
        title="Xác nhận đồng bộ BK → Thanh toán",
        subtitle="Tác vụ có thể cập nhật cả file BK và file Thanh toán.",
        state="warning" if normalize_sheets or conflicts else "preview",
        metrics=(
            SummaryMetric("Dòng cập nhật", totals["update"], "primary", "dòng"),
            SummaryMetric("Dòng mới đã chọn", totals["selected_new"], "success", "dòng"),
            SummaryMetric("Ô HĐ thay đổi", totals["invoice"], "primary", "ô"),
            SummaryMetric("Dòng mới bỏ chọn", totals["skipped_new"], "warning", "dòng"),
        ),
        notices=tuple(notices),
        tables=(
            SummaryTable(
                "Hướng đồng bộ",
                ("BK nguồn", "Sheet Thanh toán", "Trạng thái", "Cập nhật", "Thêm", "Đã giống", "Ô HĐ"),
                tuple(rows),
            ),
        ),
        change_count=totals["update"] + totals["selected_new"],
        issue_count=conflicts,
    )


def daily_completion_summary(result: Any) -> ExcelOperationSummary:
    outcomes = list(_sequence(_attribute(result, "item_outcomes", default=())))
    per_sheet: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for outcome in outcomes:
        status = _code(_attribute(outcome, "status", default=""))
        fields = _sequence(_attribute(outcome, "fields", default=()))
        action = _code(_attribute(fields[0], "action", default="")) if fields else ""
        sheet = str(_attribute(outcome, "target_sheet", default="") or "")
        if not sheet:
            continue
        if action == "INSERT":
            per_sheet[sheet]["insert"] += 1
        elif action == "UPDATE":
            per_sheet[sheet]["update"] += 1
        elif status == "UNCHANGED":
            per_sheet[sheet]["unchanged"] += 1
        elif status == "USER_KEPT":
            per_sheet[sheet]["target_only"] += 1
    updated = _integer(_attribute(result, "updated_rows", default=0))
    inserted = _integer(_attribute(result, "inserted_rows", "added_rows", default=0))
    unchanged = _integer(_attribute(result, "unchanged_rows", default=0))
    target_only = _integer(_attribute(result, "target_only_rows", default=0))
    invalid = _integer(_attribute(result, "invalid_rows", "skipped_rows", default=0))
    rows = [
        (
            sheet,
            counts["update"],
            counts["insert"],
            counts["unchanged"],
            counts["target_only"],
        )
        for sheet, counts in sorted(per_sheet.items())
    ]
    if not rows:
        rows.append(
            (
                _attribute(result, "sheet_name", default="—"),
                updated,
                inserted,
                unchanged,
                target_only,
            )
        )
    changes = updated + inserted
    status = _code(_attribute(result, "status", default=""))
    no_changes = status == "NO_CHANGES" or (not status and changes == 0)
    notices: list[SummaryNotice] = []
    if target_only:
        notices.append(SummaryNotice(f"{target_only} dòng chỉ có ở BK đã được giữ nguyên.", "success"))
    if invalid:
        notices.append(SummaryNotice(f"{invalid} dòng thiếu hoặc sai SQT chưa được đồng bộ.", "warning"))
    return ExcelOperationSummary(
        title=(
            "Dữ liệu Hàng ngày và BK đã đồng bộ"
            if no_changes
            else "Đồng bộ Hàng ngày → BK hoàn tất"
        ),
        subtitle=(
            "Không có dòng nào cần ghi vào file BK."
            if no_changes
            else f"Đã ghi {changes} thay đổi vào file BK."
        ),
        state="no_changes" if no_changes else "warning" if invalid else "success",
        metrics=(
            SummaryMetric("Đã cập nhật", updated, "primary", "dòng"),
            SummaryMetric("Đã thêm", inserted, "success", "dòng"),
            SummaryMetric("Đã giống nhau", unchanged, "neutral", "dòng"),
            SummaryMetric("Bỏ qua nguồn", invalid, "warning", "dòng"),
        ),
        notices=tuple(notices),
        tables=(SummaryTable("Kết quả theo sheet", ("Sheet BK", "Cập nhật", "Thêm", "Đã giống", "Giữ ở BK"), tuple(rows)),),
        change_count=changes,
        issue_count=invalid,
    )


def posting_completion_summary(result: Any) -> ExcelOperationSummary:
    outcomes = list(_sequence(_attribute(result, "item_outcomes", default=())))
    counts = _outcome_status_counts(outcomes)
    per_sheet: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for outcome in outcomes:
        sheet = str(_attribute(outcome, "target_sheet", default="") or "Chưa xác định")
        per_sheet[sheet][_code(_attribute(outcome, "status", default=""))] += 1
    written = counts["WRITTEN"] + counts["PARTIAL"]
    issues = counts["INVALID_SOURCE"] + counts["FAILED"]
    rows = [
        (
            sheet,
            values["WRITTEN"] + values["PARTIAL"],
            values["UNCHANGED"],
            values["USER_KEPT"],
            values["USER_SKIPPED"],
            values["INVALID_SOURCE"] + values["FAILED"],
        )
        for sheet, values in sorted(per_sheet.items())
    ]
    if not outcomes:
        written = _integer(_attribute(result, "posted_source_items", default=0))
        already = _integer(_attribute(result, "already_existing_items", default=0))
        skipped = _integer(_attribute(result, "skipped_source_items", default=0))
        counts["UNCHANGED"] = already
        counts["USER_SKIPPED"] = skipped
        rows.append((_attribute(result, "sheet_name", default="—"), written, already, 0, skipped, 0))
    status = _code(_attribute(result, "status", default=""))
    no_changes = status == "NO_CHANGES" or (not status and written == 0)
    notices = (
        (SummaryNotice(f"Có {issues} khoản cần kiểm tra trong kết quả.", "error"),)
        if issues
        else (SummaryNotice("Không có lỗi hoặc dữ liệu nguồn không hợp lệ.", "success"),)
    )
    return ExcelOperationSummary(
        title="Khoản chi đã có trong BK" if no_changes else "Nhập khoản chi vào BK hoàn tất",
        subtitle=(
            "Không có ô tiền nào cần ghi thêm."
            if no_changes
            else f"Đã ghi {written} khoản vào {len([row for row in rows if row[0] != 'Chưa xác định']) or 1} sheet BK."
        ),
        state="error" if issues else "no_changes" if no_changes else "success",
        metrics=(
            SummaryMetric("Đã ghi", written, "success", "khoản"),
            SummaryMetric("Đã đúng sẵn", counts["UNCHANGED"], "neutral", "khoản"),
            SummaryMetric("Giữ theo lựa chọn", counts["USER_KEPT"], "warning", "khoản"),
            SummaryMetric("Bỏ qua", counts["USER_SKIPPED"], "warning", "khoản"),
        ),
        notices=notices,
        tables=(SummaryTable("Kết quả theo sheet", ("Sheet BK", "Đã ghi", "Đúng sẵn", "Giữ lại", "Bỏ qua", "Cần kiểm tra"), tuple(rows)),),
        change_count=written,
        issue_count=issues,
    )


def payment_completion_summary(result: Any) -> ExcelOperationSummary:
    target_results = _attribute(result, "target_results", default={}) or {}
    rows: list[tuple[Any, ...]] = []
    for key, target in target_results.items():
        rows.append(
            (
                _attribute(target, "sheet_name", default=key),
                "Tạo mới" if bool(_attribute(target, "sheet_created", default=False)) else "Đã có",
                _integer(_attribute(target, "updated_rows", default=0)),
                _integer(_attribute(target, "inserted_rows", default=0)),
                _integer(_attribute(target, "unchanged_rows", default=0)),
                _integer(_attribute(target, "invoice_written_cells", default=0)),
                _integer(_attribute(target, "skipped_rows", default=0)),
            )
        )
    updated = _integer(_attribute(result, "updated_rows", default=0))
    inserted = _integer(_attribute(result, "inserted_rows", default=0))
    unchanged = _integer(_attribute(result, "unchanged_rows", default=0))
    invoice = _integer(_attribute(result, "invoice_written_cells", default=0))
    skipped = _integer(_attribute(result, "skipped_rows", default=0))
    if not rows:
        rows.append((_attribute(result, "sheet_name", default="—"), "Đã có", updated, inserted, unchanged, invoice, skipped))
    carrier = _attribute(result, "carrier_summary", default=None)
    detail_lines: list[str] = []
    carrier_issues = 0
    if carrier is not None:
        total = _attribute(carrier, "selected_period_total", default=0)
        try:
            total_text = f"{float(total):,.0f}".replace(",", ".")
        except (TypeError, ValueError):
            total_text = str(total)
        detail_lines = [
            f"Kỳ: {_display(_attribute(carrier, 'period', default='—'))}",
            f"Bên vận tải: {_integer(_attribute(carrier, 'carrier_count', default=0))}",
            f"Số hóa đơn: {_integer(_attribute(carrier, 'invoice_count', default=0))}",
            f"Tổng tiền kỳ: {total_text}",
            f"Khoản chưa có hóa đơn: {_integer(_attribute(carrier, 'missing_invoice_items', default=0))}",
            f"Khoản chưa xác định vận tải: {_integer(_attribute(carrier, 'unmapped_carrier_items', default=0))}",
            f"Hóa đơn thuộc nhiều bên: {_integer(_attribute(carrier, 'cross_carrier_invoices', default=0))}",
            f"Khoản nghi trùng: {_integer(_attribute(carrier, 'suspected_duplicate_items', default=0))}",
        ]
        carrier_issues = sum(
            _integer(_attribute(carrier, name, default=0))
            for name in (
                "missing_invoice_items",
                "unmapped_carrier_items",
                "cross_carrier_invoices",
                "suspected_duplicate_items",
            )
        )
    outcomes = list(_sequence(_attribute(result, "item_outcomes", default=())))
    outcome_counts = _outcome_status_counts(outcomes)
    execution_issues = outcome_counts["INVALID_SOURCE"] + outcome_counts["FAILED"]
    issues = carrier_issues + execution_issues
    changes = updated + inserted
    status = _code(_attribute(result, "status", default=""))
    no_changes = status == "NO_CHANGES" or (
        not status and changes == 0 and carrier is None
    )
    notices: list[SummaryNotice] = []
    source_sheets = str(
        _attribute(result, "source_sheet_name", default="") or ""
    ).strip()
    if source_sheets:
        notices.append(SummaryNotice(f"Sheet BK nguồn: {source_sheets}.", "info"))
    if issues:
        notices.append(SummaryNotice("Đồng bộ đã hoàn tất nhưng có mục cần kiểm tra.", "warning"))
    else:
        notices.append(SummaryNotice("Không phát hiện vấn đề trong dữ liệu đã đồng bộ.", "success"))
    return ExcelOperationSummary(
        title=(
            "Dữ liệu BK và Thanh toán đã đồng bộ"
            if no_changes
            else "Đồng bộ BK → Thanh toán hoàn tất"
        ),
        subtitle=(
            "Không có dòng nào cần ghi sang file Thanh toán."
            if no_changes
            else f"Đã ghi {changes} thay đổi vào {len(rows)} sheet Thanh toán."
        ),
        state="warning" if issues else "no_changes" if no_changes else "success",
        metrics=(
            SummaryMetric("Đã cập nhật", updated, "primary", "dòng"),
            SummaryMetric("Đã thêm", inserted, "success", "dòng"),
            SummaryMetric("Ô HĐ đã ghi", invoice, "primary", "ô"),
            SummaryMetric("Bỏ qua", skipped, "warning", "dòng"),
        ),
        notices=tuple(notices),
        tables=(SummaryTable("Kết quả theo sheet Thanh toán", ("Sheet", "Trạng thái", "Cập nhật", "Thêm", "Đã giống", "Ô HĐ", "Bỏ qua"), tuple(rows)),),
        details=(
            SummaryDetails(
                "Tổng hợp hóa đơn và bên vận tải",
                tuple(detail_lines),
                expanded=bool(carrier_issues),
                tone="warning" if carrier_issues else "neutral",
            ),
        ) if detail_lines else (),
        change_count=changes,
        issue_count=issues,
    )


class _ExcelSummaryDialog(QDialog):
    """Khung trình bày chung; footer do lớp con cung cấp."""

    _STATE_ICONS = {
        "preview": QStyle.StandardPixmap.SP_ArrowRight,
        "success": QStyle.StandardPixmap.SP_DialogApplyButton,
        "warning": QStyle.StandardPixmap.SP_MessageBoxWarning,
        "error": QStyle.StandardPixmap.SP_MessageBoxCritical,
        "no_changes": QStyle.StandardPixmap.SP_DialogApplyButton,
    }

    def __init__(self, summary: ExcelOperationSummary, parent: QWidget | None = None) -> None:
        # Một số unit test gọi slot của MainWindow bằng owner tối giản không
        # phải QWidget; ở runtime parent vẫn luôn là cửa sổ chính thực tế.
        super().__init__(parent if isinstance(parent, QWidget) else None)
        self.summary = summary
        self.setObjectName("excelSummaryDialog")
        self.setWindowTitle(summary.title)
        self.setMinimumSize(760, 520)
        self.resize(860, 620)
        self._outer = QVBoxLayout(self)
        self._outer.setContentsMargins(18, 18, 18, 14)
        self._outer.setSpacing(14)
        self._build_header()
        scroll = QScrollArea()
        scroll.setObjectName("excelSummaryScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        body = QWidget()
        body.setObjectName("excelSummaryBody")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(14)
        self._build_metrics(body_layout)
        self._build_notices(body_layout)
        self._build_tables(body_layout)
        self._build_details(body_layout)
        body_layout.addStretch(1)
        scroll.setWidget(body)
        self._outer.addWidget(scroll, 1)

    def _build_header(self) -> None:
        header = QFrame()
        header.setObjectName("excelSummaryHeader")
        header.setProperty("state", self.summary.state)
        row = QHBoxLayout(header)
        row.setContentsMargins(16, 14, 16, 14)
        row.setSpacing(14)
        icon = QLabel()
        icon.setObjectName("excelSummaryIcon")
        icon.setProperty("state", self.summary.state)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(42, 42)
        icon_kind = self._STATE_ICONS.get(
            self.summary.state, QStyle.StandardPixmap.SP_ArrowRight
        )
        icon.setPixmap(self.style().standardIcon(icon_kind).pixmap(22, 22))
        row.addWidget(icon)
        text_layout = QVBoxLayout()
        text_layout.setSpacing(3)
        title = QLabel(self.summary.title)
        title.setObjectName("excelSummaryTitle")
        subtitle = QLabel(self.summary.subtitle)
        subtitle.setObjectName("excelSummarySubtitle")
        subtitle.setWordWrap(True)
        text_layout.addWidget(title)
        text_layout.addWidget(subtitle)
        row.addLayout(text_layout, 1)
        self._outer.addWidget(header)

    def _build_metrics(self, layout: QVBoxLayout) -> None:
        if not self.summary.metrics:
            return
        grid = QGridLayout()
        grid.setSpacing(10)
        for index, metric in enumerate(self.summary.metrics):
            card = QFrame()
            card.setObjectName("excelSummaryMetric")
            card.setProperty("tone", metric.tone)
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(13, 10, 13, 10)
            card_layout.setSpacing(2)
            label = QLabel(metric.label.upper())
            label.setObjectName("excelSummaryMetricLabel")
            label.setProperty("tone", metric.tone)
            value = QLabel(str(metric.value))
            value.setObjectName("excelSummaryMetricValue")
            value.setProperty("tone", metric.tone)
            card_layout.addWidget(label)
            card_layout.addWidget(value)
            if metric.note:
                note = QLabel(metric.note)
                note.setObjectName("excelSummaryMetricNote")
                card_layout.addWidget(note)
            grid.addWidget(card, index // 4, index % 4)
        layout.addLayout(grid)

    def _build_notices(self, layout: QVBoxLayout) -> None:
        icons = {
            "success": QStyle.StandardPixmap.SP_DialogApplyButton,
            "warning": QStyle.StandardPixmap.SP_MessageBoxWarning,
            "error": QStyle.StandardPixmap.SP_MessageBoxCritical,
            "info": QStyle.StandardPixmap.SP_MessageBoxInformation,
        }
        for notice in self.summary.notices:
            frame = QFrame()
            frame.setObjectName("excelSummaryNotice")
            frame.setProperty("tone", notice.tone)
            row = QHBoxLayout(frame)
            row.setContentsMargins(12, 9, 12, 9)
            row.setSpacing(9)
            icon = QLabel()
            icon.setObjectName("excelSummaryNoticeIcon")
            icon.setProperty("tone", notice.tone)
            icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
            icon.setFixedWidth(20)
            icon_kind = icons.get(
                notice.tone, QStyle.StandardPixmap.SP_MessageBoxInformation
            )
            icon.setPixmap(self.style().standardIcon(icon_kind).pixmap(16, 16))
            text = QLabel(notice.text)
            text.setWordWrap(True)
            text.setObjectName("excelSummaryNoticeText")
            text.setProperty("tone", notice.tone)
            row.addWidget(icon)
            row.addWidget(text, 1)
            layout.addWidget(frame)

    def _build_tables(self, layout: QVBoxLayout) -> None:
        for table_data in self.summary.tables:
            if not table_data.rows:
                continue
            title = QLabel(table_data.title.upper())
            title.setObjectName("excelSummarySectionTitle")
            layout.addWidget(title)
            table = QTableWidget(len(table_data.rows), len(table_data.headers))
            table.setObjectName("excelSummaryTable")
            table.setHorizontalHeaderLabels(list(table_data.headers))
            table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
            table.setAlternatingRowColors(True)
            table.verticalHeader().setVisible(False)
            table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
            for row_index, row_values in enumerate(table_data.rows):
                for column, value in enumerate(row_values):
                    item = QTableWidgetItem(_display(value))
                    item.setToolTip(_display(value))
                    if column >= 2:
                        item.setTextAlignment(
                            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
                        )
                    table.setItem(row_index, column, item)
            header = table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            if table_data.headers:
                header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            header.setStretchLastSection(False)
            table.setMinimumHeight(min(260, 72 + len(table_data.rows) * 34))
            table.setMaximumHeight(min(300, 76 + len(table_data.rows) * 34))
            layout.addWidget(table)

    def _build_details(self, layout: QVBoxLayout) -> None:
        for detail in self.summary.details:
            toggle = QToolButton()
            toggle.setObjectName("excelSummaryDetailsToggle")
            toggle.setProperty("tone", detail.tone)
            toggle.setCheckable(True)
            toggle.setChecked(detail.expanded)
            toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
            detail_frame = QFrame()
            detail_frame.setObjectName("excelSummaryDetails")
            detail_frame.setProperty("tone", detail.tone)
            detail_layout = QVBoxLayout(detail_frame)
            detail_layout.setContentsMargins(14, 10, 14, 10)
            for line in detail.lines:
                label = QLabel(line)
                label.setWordWrap(True)
                detail_layout.addWidget(label)
            detail_frame.setVisible(detail.expanded)

            def update(checked: bool, *, button: QToolButton = toggle, frame: QFrame = detail_frame, title: str = detail.title) -> None:
                button.setText(("▾ " if checked else "▸ ") + title)
                frame.setVisible(checked)

            toggle.toggled.connect(update)
            update(detail.expanded)
            layout.addWidget(toggle)
            layout.addWidget(detail_frame)


class ExcelConfirmationDialog(_ExcelSummaryDialog):
    def __init__(
        self,
        summary: ExcelOperationSummary,
        *,
        confirm_label: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(summary, parent)
        footer = QHBoxLayout()
        footer.addStretch(1)
        self.cancel_button = QPushButton("Hủy")
        self.cancel_button.setObjectName("excelSummaryCancelButton")
        self.confirm_button = QPushButton(confirm_label)
        self.confirm_button.setObjectName("excelSummaryConfirmButton")
        self.confirm_button.setProperty("primary", True)
        self.confirm_button.setAutoDefault(False)
        self.confirm_button.setDefault(False)
        self.cancel_button.clicked.connect(self.reject)
        self.confirm_button.clicked.connect(self.accept)
        footer.addWidget(self.cancel_button)
        footer.addWidget(self.confirm_button)
        self._outer.addLayout(footer)


class ExcelCompletionDialog(_ExcelSummaryDialog):
    OPEN_TARGET = "open_target"
    VIEW_DETAILS = "view_details"

    def __init__(
        self,
        summary: ExcelOperationSummary,
        *,
        open_label: str,
        details_available: bool,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(summary, parent)
        self.selected_action: str | None = None
        footer = QHBoxLayout()
        self.close_button = QPushButton("Đóng")
        self.close_button.setObjectName("excelSummaryCloseButton")
        self.close_button.clicked.connect(self.accept)
        footer.addWidget(self.close_button)
        footer.addStretch(1)
        self.details_button: QPushButton | None = None
        if details_available:
            self.details_button = QPushButton("Xem chi tiết")
            self.details_button.setObjectName("excelSummaryDetailsButton")
            self.details_button.clicked.connect(
                lambda: self._select_action(self.VIEW_DETAILS)
            )
            footer.addWidget(self.details_button)
        self.open_button = QPushButton(open_label)
        self.open_button.setObjectName("excelSummaryOpenButton")
        self.open_button.setProperty("primary", True)
        self.open_button.clicked.connect(lambda: self._select_action(self.OPEN_TARGET))
        footer.addWidget(self.open_button)
        self._outer.addLayout(footer)

    def _select_action(self, action: str) -> None:
        self.selected_action = action
        self.accept()


__all__ = [
    "ExcelCompletionDialog",
    "ExcelConfirmationDialog",
    "ExcelOperationSummary",
    "SummaryDetails",
    "SummaryMetric",
    "SummaryNotice",
    "SummaryTable",
    "daily_completion_summary",
    "daily_confirmation_summary",
    "payment_completion_summary",
    "payment_confirmation_summary",
    "posting_completion_summary",
    "posting_confirmation_summary",
]
