"""Các value object dùng chung giữa phần mềm, BAT và PAD."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from app.services.excel.models import WorkbookFingerprint


RPA_EXPENSE_OPERATION = "NHAP_KHOAN_CHI_BK"
RPA_STATUS_NOT_IMPORTED = "Chưa nhập"
RPA_STATUS_IMPORTED = "Đã nhập"


@dataclass(frozen=True, slots=True)
class RpaSheetCandidate:
    sheet_name: str
    month: int
    year: int

    @property
    def target_sheet(self) -> str:
        return self.sheet_name

    @property
    def source_sheet(self) -> str:
        return self.sheet_name


@dataclass(frozen=True, slots=True)
class RpaExpenseAmounts:
    cuoc_bo_dong_hang: int = 0
    nang_ha_dong_hang: int = 0
    cuoc_bien: int = 0
    nang_do_vs_lam_lenh: int = 0
    cuoc_bo_tra_hang: int = 0
    tien_hang: int = 0
    cong_nhan_boc_xep: int = 0
    luu_cont_qua_tai: int = 0
    sua_chua_cont: int = 0

    @property
    def total(self) -> int:
        return sum(self.to_dict().values())

    def to_dict(self) -> dict[str, int]:
        return {
            "cuoc_bo_dong_hang": self.cuoc_bo_dong_hang,
            "nang_ha_dong_hang": self.nang_ha_dong_hang,
            "cuoc_bien": self.cuoc_bien,
            "nang_do_vs_lam_lenh": self.nang_do_vs_lam_lenh,
            "cuoc_bo_tra_hang": self.cuoc_bo_tra_hang,
            "tien_hang": self.tien_hang,
            "cong_nhan_boc_xep": self.cong_nhan_boc_xep,
            "luu_cont_qua_tai": self.luu_cont_qua_tai,
            "sua_chua_cont": self.sua_chua_cont,
        }


@dataclass(frozen=True, slots=True)
class RpaSqtItem:
    sqt: str
    source_rows: tuple[int, ...]
    status: str
    amounts: RpaExpenseAmounts
    errors: tuple[str, ...] = ()

    @property
    def row_count(self) -> int:
        return len(self.source_rows)

    @property
    def can_run(self) -> bool:
        # Một SQT chưa phát sinh khoản chi vẫn phải chạy được để PAD có thể
        # mở đúng quyết toán và nhập bộ giá trị 0. Tổng bằng 0 chỉ là cảnh báo;
        # chỉ lỗi đọc/kiểm tra dữ liệu mới được phép chặn lựa chọn.
        return not self.errors

    @property
    def validation_message(self) -> str:
        if self.errors:
            return "; ".join(self.errors)
        if self.amounts.total <= 0:
            return "Tất cả khoản tiền đều bằng 0."
        return ""

    def to_payload(self) -> dict[str, Any]:
        return {
            "sqt": self.sqt,
            "source_rows": list(self.source_rows),
            "status_before": self.status,
            "amounts": self.amounts.to_dict(),
        }


@dataclass(frozen=True, slots=True)
class RpaExpensePlan:
    bk_path: Path
    sheet_name: str
    fingerprint: WorkbookFingerprint
    items: tuple[RpaSqtItem, ...]
    pending_revisions: tuple[tuple[str, int], ...] = ()
    latest_pad_sqt: tuple[str, ...] = ()

    @property
    def runnable_count(self) -> int:
        return sum(item.can_run for item in self.items)

    def item_map(self) -> dict[str, RpaSqtItem]:
        return {item.sqt: item for item in self.items}

    @property
    def pending_sqt(self) -> tuple[str, ...]:
        return tuple(sqt for sqt, _revision in self.pending_revisions)

    def pending_revision(self, sqt: object) -> int | None:
        target = str(sqt).strip()
        return next(
            (
                int(revision)
                for value, revision in self.pending_revisions
                if value == target
            ),
            None,
        )


@dataclass(frozen=True, slots=True)
class PreparedRpaSelection:
    selection_path: Path
    run_id: str
    item_count: int
    payload: dict[str, Any]


@dataclass(frozen=True, slots=True)
class RpaExpenseLaunchResult:
    success: bool
    message: str
    bat_path: Path
    selection_path: Path
    run_id: str
    item_count: int
    process_id: int | None = None
