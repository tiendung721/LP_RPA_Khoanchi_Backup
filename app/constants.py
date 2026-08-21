"""Các hằng số nghiệp vụ và mặc định dùng chung trong ứng dụng."""

from __future__ import annotations

from types import MappingProxyType
from typing import Final

APP_NAME: Final = "Trợ lý Dữ liệu Quyết toán"
APP_VENDOR: Final = "Kikai"
APP_SLUG: Final = "TroLyDuLieuQuyetToan"
LEGACY_SCHEMA_VERSION: Final = 1
OBJECT_SCHEMA_VERSION: Final = 2
SCHEMA_VERSION: Final = 3

FEE_CATALOG = MappingProxyType(
    {
        "CB": "Cước biển",
        "CBDH": "Cước bộ đóng hàng; gồm DtD, Dr-to-Dr và Door-to-Door",
        "VTN": "Cước bộ trả hàng",
        "NV": "Nâng vỏ, nâng container rỗng",
        "HH": "Hạ hàng, hạ container có hàng từ xe xuống bãi",
        "NH": "Nâng hàng, nâng container có hàng từ bãi lên xe",
        "HV": "Hạ vỏ, hạ container rỗng từ xe xuống bãi",
        "VSDL": (
            "Vệ sinh, D/O, phí chứng từ, lệnh giao hàng, điện thả hàng, "
            "seal, THC hoặc terminal độc lập"
        ),
        "LC": (
            "Lưu container, lưu vỏ, lưu hàng xuất, demurrage, "
            "detention, storage"
        ),
        "QT": "Quá tải, quá trọng lượng hoặc phụ thu trọng lượng",
        "GH": "Gia hạn",
        "LL": "Phí/công làm lệnh riêng",
        "SC": "Sửa chữa hoặc hư hỏng container",
        "CXD": "Chưa đủ căn cứ hoặc loại phí chưa có mã chính thức",
    }
)
FEE_CODES: Final = frozenset(FEE_CATALOG)

# Ba loại cước này lấy bên vận tải chuẩn từ workbook Hàng ngày. ``carrier`` do
# GPT bóc cho chúng chỉ là dữ liệu trùng nguồn và phải được bỏ trước màn hình
# review; người dùng vẫn có thể nhập lại thủ công khi cần xử lý ngoại lệ.
DAILY_SYNC_CARRIER_FEE_CODES: Final = frozenset({"CB", "CBDH", "VTN"})

RULE_CATALOG = MappingProxyType(
    {
        "HD": "Lấy tổng cộng tiền thanh toán của toàn hóa đơn; chỉ dùng cho CB",
        "ST": "Lấy trực tiếp số tiền đã sau VAT",
        "CV": "Lấy tiền trước thuế rồi cộng VAT thực tế",
        "GV": (
            "Gộp nhiều dòng cùng chứng từ, container và loại cước rồi lấy "
            "số tiền cuối cùng"
        ),
    }
)
RULE_CODES: Final = frozenset(RULE_CATALOG)
UNDETERMINED_RULE_LABEL: Final = "Không xác định"

BATCH_STATUS_LABELS = MappingProxyType(
    {
        "RECEIVED": "Đã tiếp nhận",
        "REVIEWING": "Đang kiểm tra",
        "READY": "Đã xác nhận",
        "INVALID": "Không hợp lệ",
        "ARCHIVED": "Đã lưu trữ",
    }
)

DEFAULT_FILE_PATTERN: Final = "ket_qua_boc_tach*.json"
DEFAULT_STABLE_SECONDS: Final = 3.0
DEFAULT_STABILITY_TIMEOUT_SECONDS: Final = 60.0
DEFAULT_POLL_INTERVAL_SECONDS: Final = 0.25
DEFAULT_MAX_FILE_SIZE_MB: Final = 50
DEFAULT_MAX_FILE_SIZE_BYTES: Final = DEFAULT_MAX_FILE_SIZE_MB * 1024 * 1024

CANONICAL_JSON_FILENAME: Final = "ket_qua_boc_tach.json"
TEMPORARY_FILE_SUFFIXES: Final = (
    ".crdownload",
    ".part",
    ".partial",
    ".tmp",
    ".temp",
    ".download",
)
CONTAINER_PATTERN: Final = r"^[A-Z]{4}[0-9]{7}$"

APP_STATE_ACTIVE_BATCH_ID: Final = "active_batch_id"
APP_STATE_LAST_OUTPUT_SCAN: Final = "last_output_scan_at"
APP_STATE_MAIN_WINDOW_GEOMETRY: Final = "main_window_geometry"
APP_STATE_MAIN_WINDOW_STATE: Final = "main_window_state"
APP_STATE_LAST_PAGE: Final = "last_page"

SQLITE_SCHEMA_VERSION: Final = 19
