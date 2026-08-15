"""Quy tắc nguồn dữ liệu cho trường bên vận tải của kết quả GPT."""

from __future__ import annotations

from app.constants import DAILY_SYNC_CARRIER_FEE_CODES
from app.models import BatchDocument


def carrier_is_managed_by_daily_sync(fee: object) -> bool:
    """Trả về True khi bên vận tải của mã cước do file Hàng ngày quản lý."""

    return (
        isinstance(fee, str)
        and fee.strip().upper() in DAILY_SYNC_CARRIER_FEE_CODES
    )


def strip_received_gpt_carriers(
    document: BatchDocument,
) -> tuple[BatchDocument, int]:
    """Bỏ carrier GPT trùng nguồn, giữ nguyên thứ tự và các trường còn lại.

    Hàm này chỉ dùng tại ranh giới tiếp nhận file GPT. Carrier người dùng nhập
    sau đó không đi qua hàm này nên vẫn được lưu như một ngoại lệ thủ công.
    """

    stripped = 0
    rows = []
    for row in document.rows:
        if carrier_is_managed_by_daily_sync(row.fee) and row.carrier is not None:
            rows.append(row.copy_with(carrier=None))
            stripped += 1
        else:
            rows.append(row)
    if not stripped:
        return document, 0
    return BatchDocument(v=document.v, rows=rows), stripped


__all__ = [
    "carrier_is_managed_by_daily_sync",
    "strip_received_gpt_carriers",
]
