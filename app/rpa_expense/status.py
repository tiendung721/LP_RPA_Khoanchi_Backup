"""Cập nhật Trạng thái RPA sau khi PAD xác nhận web đã lưu."""

from __future__ import annotations

import copy
import json
import os
import re
import shutil
from pathlib import Path
from typing import Any

from app.services.excel.headers import normalize_header
from app.services.excel.payment_sync import find_summary_start
from app.services.excel.workbook import (
    ExcelBackupService,
    ExcelLockService,
    WorkbookGateway,
    ensure_supported_workbook,
)

from .contracts import (
    RPA_EXPENSE_OPERATION,
    RPA_STATUS_IMPORTED,
    RPA_STATUS_NOT_IMPORTED,
)
from .service import STATUS_HEADER, SUMMARY_HEADERS, normalize_sqt


class RpaExpenseStatusError(RuntimeError):
    pass


class RpaExpenseStatusService:
    def __init__(
        self,
        *,
        backup_dir: str | Path | None = None,
        gateway: WorkbookGateway | None = None,
        lock_service: ExcelLockService | None = None,
        tracking_repository: Any | None = None,
    ) -> None:
        self.backup_dir = Path(backup_dir) if backup_dir else None
        self.gateway = gateway or WorkbookGateway()
        self.lock_service = lock_service or ExcelLockService()
        self.tracking_repository = tracking_repository

    def mark_imported(
        self,
        selection_path: str | Path,
        sqt: str,
    ) -> dict[str, Any]:
        selection = self._read_selection(selection_path)
        target_sqt = normalize_sqt(sqt)
        if not target_sqt:
            raise RpaExpenseStatusError(f"SQT không hợp lệ: {sqt!r}")
        item = next(
            (
                value
                for value in selection["items"]
                if normalize_sqt(value.get("sqt")) == target_sqt
            ),
            None,
        )
        if item is None:
            raise RpaExpenseStatusError(
                f"SQT {target_sqt} không thuộc request RPA."
            )
        rows = self._source_rows(item)
        target = ensure_supported_workbook(selection["bk_file"]).resolve()
        if not target.is_file():
            raise RpaExpenseStatusError(f"Không tìm thấy file BK: {target}")
        sheet_name = str(selection["sheet_name"])
        run_id = str(selection["run_id"])
        raw_revision = item.get("pending_revision")
        try:
            expected_revision = (
                int(raw_revision) if raw_revision is not None else None
            )
        except (TypeError, ValueError):
            expected_revision = None
        if self.tracking_repository is not None:
            current_revision = self.tracking_repository.pending_revision(
                target,
                sheet_name,
                target_sqt,
            )
            if current_revision is not None and current_revision != expected_revision:
                return {
                    "success": True,
                    "operation": RPA_EXPENSE_OPERATION,
                    "run_id": run_id,
                    "bk_file": str(target),
                    "sheet_name": sheet_name,
                    "sqt": target_sqt,
                    "source_rows": rows,
                    "status": RPA_STATUS_NOT_IMPORTED,
                    "pending_cleared": False,
                    "stale_selection": True,
                    "message": (
                        "BK đã thay đổi sau khi gửi PAD; giữ SQT trong nhóm chờ."
                    ),
                }

        # Chỉ dùng khóa Windows như bước preflight. Phải nhả khóa trước khi
        # copy/replace vì Windows không cho sao chép file đang bị khóa byte.
        with self.lock_service.acquire(target):
            pass
        fingerprint = self.gateway.fingerprint(target)
        backup = self._ensure_run_backup(target, run_id)
        self.gateway.assert_unchanged(target, fingerprint, label="File BK")
        backups = ExcelBackupService(
            self._backup_directory(target),
            working_dir=target.parent,
        )
        working = backups.create_working_copy(target, run_id=run_id)
        workbook = None
        try:
            workbook = self.gateway.load(
                working, read_only=False, data_only=False
            )
            if sheet_name not in workbook.sheetnames:
                raise RpaExpenseStatusError(
                    f"Không tìm thấy sheet {sheet_name}."
                )
            worksheet = workbook[sheet_name]
            status_column = self._ensure_status_column(worksheet)
            sqt_column = self._source_sqt_column(worksheet)
            for row in rows:
                actual = normalize_sqt(
                    worksheet.cell(row, sqt_column).value
                )
                if actual != target_sqt:
                    raise RpaExpenseStatusError(
                        f"Dòng {row} có SQT {actual or 'trống'}, "
                        f"không phải {target_sqt}."
                    )
                worksheet.cell(row, status_column).value = RPA_STATUS_IMPORTED
            self.gateway.save(workbook, working)
            workbook.close()
            workbook = None
            self.gateway.verify_openable(working)
            self.gateway.atomic_replace(
                working,
                target,
                expected=fingerprint,
            )
        finally:
            if workbook is not None:
                workbook.close()
            working.unlink(missing_ok=True)
        pending_cleared = False
        if self.tracking_repository is not None:
            pending_cleared = self.tracking_repository.mark_pad_succeeded(
                target,
                sheet_name,
                target_sqt,
                expected_revision=expected_revision,
            )
        return {
            "success": True,
            "operation": RPA_EXPENSE_OPERATION,
            "run_id": run_id,
            "bk_file": str(target),
            "sheet_name": sheet_name,
            "sqt": target_sqt,
            "source_rows": rows,
            "status": RPA_STATUS_IMPORTED,
            "backup_path": str(backup),
            "pending_cleared": pending_cleared,
        }

    @staticmethod
    def _read_selection(path: str | Path) -> dict[str, Any]:
        source = Path(path)
        if not source.is_file():
            raise RpaExpenseStatusError(
                f"Không tìm thấy JSON lựa chọn RPA: {source}"
            )
        try:
            payload = json.loads(source.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            raise RpaExpenseStatusError(
                f"Không đọc được JSON lựa chọn RPA: {exc}"
            ) from exc
        if not isinstance(payload, dict):
            raise RpaExpenseStatusError("JSON lựa chọn RPA phải là object.")
        if payload.get("version") != 1:
            raise RpaExpenseStatusError("Version JSON RPA không được hỗ trợ.")
        if payload.get("operation") != RPA_EXPENSE_OPERATION:
            raise RpaExpenseStatusError("JSON không đúng nghiệp vụ khoản chi BK.")
        for field in ("run_id", "bk_file", "sheet_name", "items"):
            if field not in payload:
                raise RpaExpenseStatusError(
                    f"JSON RPA thiếu trường {field}."
                )
        if not isinstance(payload["items"], list):
            raise RpaExpenseStatusError("Trường items phải là mảng.")
        return payload

    @staticmethod
    def _source_rows(item: dict[str, Any]) -> list[int]:
        raw = item.get("source_rows")
        if not isinstance(raw, list) or not raw:
            raise RpaExpenseStatusError("SQT không có danh sách dòng nguồn.")
        rows: list[int] = []
        for value in raw:
            if isinstance(value, bool):
                raise RpaExpenseStatusError("Số dòng nguồn không hợp lệ.")
            try:
                row = int(value)
            except (TypeError, ValueError) as exc:
                raise RpaExpenseStatusError(
                    f"Số dòng nguồn không hợp lệ: {value!r}"
                ) from exc
            if row <= 1:
                raise RpaExpenseStatusError(
                    f"Số dòng nguồn phải lớn hơn 1: {row}"
                )
            rows.append(row)
        return list(dict.fromkeys(rows))

    def _backup_directory(self, target: Path) -> Path:
        return self.backup_dir or (
            target.parent / "_system" / "Excel" / "Backup"
        )

    def _ensure_run_backup(self, target: Path, run_id: str) -> Path:
        directory = self._backup_directory(target)
        directory.mkdir(parents=True, exist_ok=True)
        safe_run_id = re.sub(r"[^0-9A-Za-z_.-]+", "_", run_id)
        backup = directory / (
            f"{target.stem}_before_rpa_{safe_run_id}{target.suffix}"
        )
        if backup.exists():
            return backup
        # Giữ phần mở rộng Excel ở cuối để WorkbookGateway nhận diện đúng định dạng.
        temporary = backup.with_name(
            f".{backup.stem}.tmp{backup.suffix}"
        )
        try:
            shutil.copy2(target, temporary)
            self.gateway.verify_openable(temporary)
            try:
                os.replace(temporary, backup)
            except FileExistsError:
                temporary.unlink(missing_ok=True)
        finally:
            temporary.unlink(missing_ok=True)
        return backup

    @staticmethod
    def _source_sqt_column(worksheet: Any) -> int:
        summary_start = find_summary_start(worksheet)
        if summary_start is None:
            raise RpaExpenseStatusError(
                "Không nhận diện được khối tổng hợp BK."
            )
        matches = [
            column
            for column in range(1, summary_start)
            if normalize_header(worksheet.cell(1, column).value) == "sqt"
        ]
        if len(matches) != 1:
            raise RpaExpenseStatusError(
                "Không xác định duy nhất cột SQT nguồn."
            )
        return matches[0]

    @staticmethod
    def _ensure_status_column(worksheet: Any) -> int:
        expected = normalize_header(STATUS_HEADER)
        matches = [
            column
            for column in range(1, int(worksheet.max_column or 0) + 1)
            if normalize_header(worksheet.cell(1, column).value) == expected
        ]
        if len(matches) > 1:
            raise RpaExpenseStatusError(
                "Có nhiều cột Trạng thái RPA."
            )
        if matches:
            return matches[0]
        summary_start = find_summary_start(worksheet)
        if summary_start is None:
            raise RpaExpenseStatusError(
                "Không nhận diện được khối tổng hợp BK."
            )
        combined_column = summary_start + len(SUMMARY_HEADERS) - 1
        status_column = combined_column + 1
        header = worksheet.cell(1, status_column)
        if header.value not in (None, ""):
            raise RpaExpenseStatusError(
                "Không có cột trống để tạo Trạng thái RPA."
            )
        source = worksheet.cell(1, combined_column)
        if source.has_style:
            header._style = copy.copy(source._style)
        header.font = copy.copy(source.font)
        header.fill = copy.copy(source.fill)
        header.border = copy.copy(source.border)
        header.alignment = copy.copy(source.alignment)
        header.protection = copy.copy(source.protection)
        header.number_format = source.number_format
        header.value = STATUS_HEADER
        return status_column


__all__ = [
    "RpaExpenseStatusError",
    "RpaExpenseStatusService",
]
