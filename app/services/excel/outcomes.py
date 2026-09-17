"""Build user-facing execution manifests from the actions actually applied."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import Any

from .models import (
    FieldWriteOutcome,
    ItemWriteOutcome,
    OutcomeStatus,
    PostingItemStatus,
    ResolutionAction,
    SyncAction,
    SyncActionType,
)


def _code(value: Any) -> str | None:
    if value is None:
        return None
    return str(getattr(value, "value", value))


def _aggregate(fields: Sequence[FieldWriteOutcome]) -> OutcomeStatus:
    statuses = {field.status for field in fields}
    if not statuses:
        return OutcomeStatus.UNCHANGED
    if OutcomeStatus.FAILED in statuses:
        return OutcomeStatus.FAILED
    if OutcomeStatus.WRITTEN in statuses:
        return (
            OutcomeStatus.PARTIAL
            if statuses.intersection(
                {OutcomeStatus.USER_KEPT, OutcomeStatus.USER_SKIPPED}
            )
            else OutcomeStatus.WRITTEN
        )
    if OutcomeStatus.USER_SKIPPED in statuses:
        return OutcomeStatus.USER_SKIPPED
    if OutcomeStatus.USER_KEPT in statuses:
        return OutcomeStatus.USER_KEPT
    if OutcomeStatus.INVALID_SOURCE in statuses:
        return OutcomeStatus.INVALID_SOURCE
    return OutcomeStatus.UNCHANGED


def posting_outcomes(actions: Sequence[Mapping[str, Any]]) -> list[ItemWriteOutcome]:
    results: list[ItemWriteOutcome] = []
    for position, action in enumerate(actions):
        status = action.get("status")
        amount_action = action.get("action")
        if action.get("amount_write"):
            amount_status = OutcomeStatus.WRITTEN
            amount_reason = "Đã ghi giá trị tiền theo quyết định đã xác nhận."
        elif status is PostingItemStatus.ALREADY_EXISTS:
            amount_status = OutcomeStatus.UNCHANGED
            amount_reason = "Giá trị tiền đã giống dữ liệu nguồn."
        elif amount_action in {
            ResolutionAction.KEEP_EXISTING,
            ResolutionAction.KEEP_FORMULA,
        }:
            amount_status = OutcomeStatus.USER_KEPT
            amount_reason = "User chủ động giữ giá trị hoặc công thức hiện tại."
        elif amount_action is ResolutionAction.SKIP:
            amount_status = OutcomeStatus.USER_SKIPPED
            amount_reason = "User chủ động bỏ qua khoản chi."
        else:
            amount_status = OutcomeStatus.FAILED
            amount_reason = "Không xác định được action tiền trước khi ghi."
        fields = [
            FieldWriteOutcome(
                field_name="Số tiền",
                source_value=action.get("value_after")
                if action.get("amount_write")
                else action.get("source_items", [{}])[0].get("amount")
                if action.get("source_items")
                else None,
                target_value_before=action.get("value_before"),
                target_value_after=action.get("value_after"),
                action=_code(amount_action),
                status=amount_status,
                source_sheet=action.get("selected_source_sheet"),
                source_row=action.get("selected_source_row"),
                target_sheet=action.get("sheet_name"),
                target_row=action.get("target_row"),
                target_cell=action.get("target_cell"),
                reason_code=f"AMOUNT_{amount_status.value}",
                reason=amount_reason,
            )
        ]
        if action.get("invoice_selected") is not None or action.get(
            "invoice_target_cell"
        ) is not None:
            invoice_action = action.get("invoice_action")
            invoice_status = (
                OutcomeStatus.WRITTEN
                if action.get("invoice_write")
                else OutcomeStatus.USER_KEPT
                if invoice_action
                in {ResolutionAction.KEEP_EXISTING, ResolutionAction.SKIP_INVOICE}
                else OutcomeStatus.UNCHANGED
            )
            fields.append(
                FieldWriteOutcome(
                    field_name="Số HĐ",
                    source_value=action.get("invoice_selected"),
                    target_value_before=action.get("invoice_value_before"),
                    target_value_after=action.get("invoice_value_after"),
                    action=_code(invoice_action),
                    status=invoice_status,
                    target_sheet=action.get("sheet_name"),
                    target_row=action.get("target_row"),
                    target_cell=action.get("invoice_target_cell"),
                    reason_code=f"INVOICE_{invoice_status.value}",
                    reason=(
                        "Đã ghi chính xác chuỗi Số HĐ từ nguồn."
                        if invoice_status is OutcomeStatus.WRITTEN
                        else "Giữ Số HĐ hiện tại theo quyết định của user."
                        if invoice_status is OutcomeStatus.USER_KEPT
                        else "Số HĐ không thay đổi."
                    ),
                )
            )
        if action.get("carrier_group") is not None:
            carrier_action = action.get("carrier_action")
            carrier_status = (
                OutcomeStatus.WRITTEN
                if action.get("carrier_write")
                else OutcomeStatus.USER_KEPT
                if carrier_action is ResolutionAction.KEEP_EXISTING
                else OutcomeStatus.UNCHANGED
            )
            fields.append(
                FieldWriteOutcome(
                    field_name="Bên vận tải",
                    source_value=action.get("carrier_source"),
                    target_value_before=action.get("carrier_value_before"),
                    target_value_after=action.get("carrier_value_after"),
                    action=_code(carrier_action),
                    status=carrier_status,
                    target_sheet=action.get("sheet_name"),
                    target_row=action.get("target_row"),
                    target_cell=action.get("carrier_target_cell"),
                    reason_code=f"CARRIER_{carrier_status.value}",
                    reason=(
                        "Đã cập nhật bên vận tải."
                        if carrier_status is OutcomeStatus.WRITTEN
                        else "Giữ bên vận tải hiện tại theo quyết định của user."
                        if carrier_status is OutcomeStatus.USER_KEPT
                        else "Bên vận tải không thay đổi."
                    ),
                )
            )
        sources = action.get("source_items") or [{}]
        sources_by_index = {
            source.get("source_item_index"): source
            for source in sources
            if source.get("source_item_index") is not None
        }
        source_indices = action.get("source_indices") or [position]
        for source_index in source_indices:
            source = sources_by_index.get(source_index, sources[0])
            results.append(
                ItemWriteOutcome(
                    item_id=f"posting:{source_index}",
                    source_index=int(source_index),
                    container=source.get("container", action.get("container")),
                    sqt=action.get("source_sqt"),
                    bl=source.get("bl"),
                    fee=action.get("fee_selected"),
                    source_sheet=action.get("selected_source_sheet"),
                    source_row=action.get("selected_source_row"),
                    target_sheet=action.get("sheet_name"),
                    target_row=action.get("target_row"),
                    status=_aggregate(fields),
                    fields=list(fields),
                )
            )
    return results


def sync_outcomes(
    action_groups: Iterable[tuple[str, Sequence[SyncAction]]],
    *,
    invalid_rows: Sequence[int] = (),
    invalid_reasons: Mapping[int, str] | None = None,
) -> list[ItemWriteOutcome]:
    results: list[ItemWriteOutcome] = []
    for sheet_name, actions in action_groups:
        for action in actions:
            if action.action in {SyncActionType.INSERT, SyncActionType.UPDATE}:
                status = OutcomeStatus.WRITTEN
                reason = "Đã đồng bộ dòng Hàng ngày vào BK."
            elif action.action is SyncActionType.UNCHANGED:
                status = OutcomeStatus.UNCHANGED
                reason = "Dữ liệu nguồn và BK đã giống nhau."
            else:
                status = OutcomeStatus.USER_KEPT
                reason = "Dòng chỉ có ở BK được giữ theo chính sách đồng bộ."
            source = action.source
            field = FieldWriteOutcome(
                field_name="Dòng đồng bộ các cột nghiệp vụ BK",
                source_value=source.values if source is not None else None,
                target_value_before=action.target_values,
                target_value_after=(
                    source.values
                    if source is not None
                    and action.action in {SyncActionType.INSERT, SyncActionType.UPDATE}
                    else action.target_values
                ),
                action=action.action.value,
                status=status,
                source_sheet=source.source_sheet if source is not None else None,
                source_row=source.source_row if source is not None else None,
                target_sheet=sheet_name,
                target_row=action.target_row,
                reason_code=f"SYNC_{action.action.value}",
                reason=reason,
            )
            results.append(
                ItemWriteOutcome(
                    item_id=f"sync:{sheet_name}:{action.sqt}:{action.target_row}",
                    source_index=source.source_row if source is not None else None,
                    container=(str(source.container) if source and source.container else None),
                    sqt=action.sqt,
                    source_sheet=source.source_sheet if source is not None else None,
                    source_row=source.source_row if source is not None else None,
                    target_sheet=sheet_name,
                    target_row=action.target_row,
                    status=status,
                    fields=[field],
                )
            )
    reasons = dict(invalid_reasons or {})
    for row in invalid_rows:
        reason = reasons.get(int(row), "Dòng nguồn thiếu hoặc sai SQT.")
        field = FieldWriteOutcome(
            field_name="Dòng nguồn",
            source_row=int(row),
            status=OutcomeStatus.INVALID_SOURCE,
            reason_code="INVALID_SOURCE",
            reason=reason,
        )
        results.append(
            ItemWriteOutcome(
                item_id=f"sync:invalid:{row}",
                source_index=int(row),
                source_row=int(row),
                status=OutcomeStatus.INVALID_SOURCE,
                fields=[field],
            )
        )
    return results


def payment_outcomes(plan: Any) -> list[ItemWriteOutcome]:
    if hasattr(plan, "month_plans"):
        return [
            outcome
            for month_plan in plan.month_plans.values()
            for outcome in payment_outcomes(month_plan)
        ]
    results: list[ItemWriteOutcome] = []
    for target in plan.targets.values():
        for item in target.items:
            fields: list[FieldWriteOutcome] = []
            if item.skip_item:
                fields.append(
                    FieldWriteOutcome(
                        field_name="Dòng thanh toán",
                        source_value=item.values,
                        action=ResolutionAction.SKIP.value,
                        status=OutcomeStatus.USER_SKIPPED,
                        source_sheet=plan.source_sheet,
                        source_row=item.source_row,
                        target_sheet=target.sheet_name,
                        target_row=item.target_row,
                        reason_code="PAYMENT_USER_SKIPPED",
                        reason="User chủ động bỏ qua item.",
                    )
                )
            else:
                amount_fields = set(item.differences) | set(
                    item.amount_review_values
                )
                for field_name in sorted(amount_fields):
                    amount_values = item.differences.get(field_name)
                    if amount_values is None:
                        amount_values = item.amount_review_values[field_name]
                    before, after = amount_values
                    action = item.amount_actions.get(
                        field_name, ResolutionAction.OVERWRITE
                    )
                    kept = action in {
                        ResolutionAction.KEEP_EXISTING,
                        ResolutionAction.SKIP,
                    }
                    fields.append(
                        FieldWriteOutcome(
                            field_name=field_name,
                            source_value=after,
                            target_value_before=before,
                            target_value_after=before if kept else after,
                            action=action.value,
                            status=(
                                OutcomeStatus.USER_KEPT
                                if kept
                                else OutcomeStatus.WRITTEN
                            ),
                            source_sheet=plan.source_sheet,
                            source_row=item.source_row,
                            target_sheet=target.sheet_name,
                            target_row=item.target_row,
                            reason_code="PAYMENT_AMOUNT_KEPT" if kept else "PAYMENT_AMOUNT_WRITTEN",
                            reason="Giữ giá trị hiện tại theo quyết định của user."
                            if kept
                            else "Đã ghi giá trị từ BK.",
                        )
                    )
                invoice_fields = (
                    set(item.invoice_differences)
                    | set(item.invoice_review_values)
                    | set(item.invoice_actions)
                )
                for field_name in sorted(invoice_fields):
                    before, after = item.invoice_differences.get(
                        field_name,
                        item.invoice_review_values.get(
                            field_name,
                            (
                                None,
                                item.selected_invoices.get(field_name)
                                or item.invoice_values.get(field_name),
                            ),
                        ),
                    )
                    action = item.invoice_actions.get(
                        field_name, ResolutionAction.OVERWRITE
                    )
                    kept = action in {
                        ResolutionAction.KEEP_EXISTING,
                        ResolutionAction.SKIP_INVOICE,
                    }
                    fields.append(
                        FieldWriteOutcome(
                            field_name=f"Số HĐ ({field_name})",
                            source_value=after,
                            target_value_before=before,
                            target_value_after=before if kept else after,
                            action=action.value,
                            status=OutcomeStatus.USER_KEPT if kept else OutcomeStatus.WRITTEN,
                            source_sheet=plan.source_sheet,
                            source_row=item.source_row,
                            target_sheet=target.sheet_name,
                            target_row=item.target_row,
                            reason_code="PAYMENT_INVOICE_KEPT" if kept else "PAYMENT_INVOICE_WRITTEN",
                            reason="Giữ Số HĐ hiện tại theo quyết định của user."
                            if kept
                            else "Đã ghi Số HĐ từ BK.",
                        )
                    )
                if (
                    item.carrier_difference is not None
                    or item.carrier_review_value is not None
                    or item.carrier_action is not None
                ):
                    before, after = (
                        item.carrier_difference
                        or item.carrier_review_value
                        or (None, item.carrier_value)
                    )
                    action = item.carrier_action or ResolutionAction.OVERWRITE
                    kept = action is ResolutionAction.KEEP_EXISTING
                    fields.append(
                        FieldWriteOutcome(
                            field_name="Bên vận tải",
                            source_value=after,
                            target_value_before=before,
                            target_value_after=before if kept else after,
                            action=action.value,
                            status=OutcomeStatus.USER_KEPT if kept else OutcomeStatus.WRITTEN,
                            source_sheet=plan.source_sheet,
                            source_row=item.source_row,
                            target_sheet=target.sheet_name,
                            target_row=item.target_row,
                            reason_code="PAYMENT_CARRIER_KEPT" if kept else "PAYMENT_CARRIER_WRITTEN",
                            reason="Giữ bên vận tải hiện tại theo quyết định của user."
                            if kept
                            else "Đã ghi bên vận tải từ BK.",
                        )
                    )
                if not fields:
                    fields.append(
                        FieldWriteOutcome(
                            field_name="Dòng thanh toán",
                            source_value=item.values,
                            target_value_after=item.values,
                            action="UNCHANGED" if item.status == "UNCHANGED" else "WRITE",
                            status=OutcomeStatus.UNCHANGED
                            if item.status == "UNCHANGED"
                            else OutcomeStatus.WRITTEN,
                            source_sheet=plan.source_sheet,
                            source_row=item.source_row,
                            target_sheet=target.sheet_name,
                            target_row=item.target_row,
                            reason_code="PAYMENT_UNCHANGED"
                            if item.status == "UNCHANGED"
                            else "PAYMENT_WRITTEN",
                            reason="Dữ liệu đã giống nhau."
                            if item.status == "UNCHANGED"
                            else "Đã ghi dòng Thanh toán.",
                        )
                    )
            results.append(
                ItemWriteOutcome(
                    item_id=item.item_id,
                    source_index=item.source_row,
                    container=item.container,
                    sqt=item.sqt,
                    source_sheet=plan.source_sheet,
                    source_row=item.source_row,
                    target_sheet=target.sheet_name,
                    target_row=item.target_row,
                    status=_aggregate(fields),
                    fields=fields,
                )
            )
    return results


def outcome_counts(outcomes: Sequence[ItemWriteOutcome]) -> dict[OutcomeStatus, int]:
    return {
        status: sum(item.status is status for item in outcomes)
        for status in OutcomeStatus
    }


__all__ = [
    "outcome_counts",
    "payment_outcomes",
    "posting_outcomes",
    "sync_outcomes",
]
