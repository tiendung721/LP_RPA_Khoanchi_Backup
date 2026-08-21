"""Model bảng review và proxy tìm kiếm/lọc/sắp xếp.

Model nguồn luôn giữ thứ tự nghiệp vụ. Mọi thao tác sort/filter chỉ xảy ra ở
``ReviewFilterProxyModel`` nên dữ liệu serialize không bị đổi thứ tự ngoài ý muốn.
"""

from __future__ import annotations

from collections import Counter, OrderedDict
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import asdict, dataclass, field, is_dataclass
from enum import Enum
from datetime import date
from typing import Any, Final
from uuid import uuid4

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QColor

from .theme import ERROR_BG, MUTED_TEXT, SUCCESS_BG, WARNING_BG

try:
    from app.constants import (
        FEE_CATALOG as _CORE_FEE_CATALOG,
        RULE_CATALOG as _CORE_RULE_CATALOG,
        SCHEMA_VERSION as _SCHEMA_VERSION,
        UNDETERMINED_RULE_LABEL as _UNDETERMINED_RULE_LABEL,
    )

    FEE_CATALOG: Final[OrderedDict[str, str]] = OrderedDict(
        _CORE_FEE_CATALOG.items()
    )
    RULE_CATALOG: Final[OrderedDict[str | None, str]] = OrderedDict(
        [(None, _UNDETERMINED_RULE_LABEL), *_CORE_RULE_CATALOG.items()]
    )
except ImportError:
    # Fallback giúp module UI vẫn độc lập khi được preview riêng trong Designer.
    FEE_CATALOG = OrderedDict(
        [
            ("CB", "Cước biển"),
            ("CBDH", "Cước bộ đóng hàng; gồm DtD, Dr-to-Dr và Door-to-Door"),
            ("VTN", "Cước bộ trả hàng"),
            ("NV", "Nâng vỏ, nâng container rỗng"),
            ("HH", "Hạ hàng, hạ container có hàng từ xe xuống bãi"),
            ("NH", "Nâng hàng, nâng container có hàng từ bãi lên xe"),
            ("HV", "Hạ vỏ, hạ container rỗng từ xe xuống bãi"),
            ("VSDL", "Vệ sinh, D/O, chứng từ, lệnh, điện, seal, THC/terminal"),
            ("LC", "Lưu container/vỏ/hàng, gia hạn, demurrage/detention/storage"),
            ("QT", "Quá tải, quá trọng lượng hoặc phụ thu trọng lượng"),
            ("LL", "Phí/công làm lệnh riêng"),
            ("SC", "Sửa chữa hoặc hư hỏng container"),
            ("CXD", "Chưa đủ căn cứ hoặc chưa có mã chính thức"),
        ]
    )
    RULE_CATALOG = OrderedDict(
        [
            (None, "Không xác định"),
            ("HD", "Tổng cộng tiền thanh toán của toàn hóa đơn"),
            ("ST", "Số tiền trực tiếp đã sau VAT"),
            ("CV", "Tiền trước thuế cộng VAT thực tế"),
            ("GV", "Gộp các dòng phù hợp rồi lấy tiền cuối cùng"),
        ]
    )
    _SCHEMA_VERSION = 3


class RowStatus(str, Enum):
    """Trạng thái kiểm tra của một dòng."""

    VALID = "valid"
    WARNING = "warning"
    ERROR = "error"


STATUS_LABELS: Final[dict[RowStatus, str]] = {
    RowStatus.VALID: "Hợp lệ",
    RowStatus.WARNING: "Cảnh báo",
    RowStatus.ERROR: "Lỗi",
}


@dataclass(slots=True)
class ReviewRow:
    """Biểu diễn nội bộ của một dòng hóa đơn schema v3."""

    cont: Any = None
    bl: Any = None
    fee: Any = "CXD"
    rule: Any = None
    amount: Any = None
    invoice_no: Any = None
    carrier: Any = None
    vessel_voyage_raw: Any = None
    vessel_name: Any = None
    voyage_no: Any = None
    invoice_container_count: Any = None
    container_count_basis: Any = "UNKNOWN"
    invoice_date: Any = None
    source_document_id: Any = "MANUAL"
    source_document_name: Any = "Dòng thêm thủ công"
    runtime_id: str = field(default_factory=lambda: uuid4().hex)

    def as_array(self) -> list[Any]:
        return [
            self.cont,
            self.bl,
            self.fee,
            self.rule,
            self.invoice_no,
            self.carrier,
            self.amount,
        ]

    def to_object(self) -> dict[str, Any]:
        return {
            "source_document_id": self.source_document_id,
            "source_document_name": self.source_document_name,
            "container": self.cont,
            "bl": self.bl,
            "vessel_voyage_raw": self.vessel_voyage_raw,
            "vessel_name": self.vessel_name,
            "voyage_no": self.voyage_no,
            "invoice_container_count": self.invoice_container_count,
            "container_count_basis": self.container_count_basis,
            "fee": self.fee,
            "rule": self.rule,
            "invoice_no": self.invoice_no,
            "invoice_date": self.invoice_date,
            "carrier": self.carrier,
            "amount": self.amount,
        }

    @classmethod
    def from_mapping(
        cls, value: Mapping[str, Any], *, runtime_id: str | None = None
    ) -> "ReviewRow":
        return cls(
            cont=_mapping_value(value, "cont", "container"),
            bl=_mapping_value(value, "bl", "bill_of_lading", "bill"),
            fee=_mapping_value(value, "fee", "fee_code", default="CXD"),
            rule=_mapping_value(value, "rule", "rule_code"),
            amount=_mapping_value(value, "amount"),
            invoice_no=_mapping_value(value, "invoice_no", "invoice_number", "invoice"),
            carrier=_mapping_value(value, "carrier", "transport_provider", "transport"),
            vessel_voyage_raw=_mapping_value(value, "vessel_voyage_raw"),
            vessel_name=_mapping_value(value, "vessel_name"),
            voyage_no=_mapping_value(value, "voyage_no"),
            invoice_container_count=_mapping_value(value, "invoice_container_count"),
            container_count_basis=_mapping_value(
                value, "container_count_basis", default="UNKNOWN"
            ),
            invoice_date=_mapping_value(value, "invoice_date"),
            source_document_id=_mapping_value(
                value, "source_document_id", default="MANUAL"
            ),
            source_document_name=_mapping_value(
                value, "source_document_name", default="Dòng thêm thủ công"
            ),
            runtime_id=runtime_id or uuid4().hex,
        )

    @classmethod
    def from_sequence(
        cls,
        value: Sequence[Any],
        *,
        runtime_id: str | None = None,
    ) -> "ReviewRow":
        values = list(value)
        if len(values) != 7:
            raise ValueError("Mỗi dòng dữ liệu phải có đúng 7 giá trị.")
        return cls(
            cont=values[0],
            bl=values[1],
            fee=values[2],
            rule=values[3],
            invoice_no=values[4],
            carrier=values[5],
            amount=values[6],
            runtime_id=runtime_id or uuid4().hex,
        )


@dataclass(frozen=True, slots=True)
class RowValidation:
    """Kết quả validation hiển thị cho một dòng."""

    errors: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()

    @property
    def status(self) -> RowStatus:
        if self.errors:
            return RowStatus.ERROR
        if self.warnings:
            return RowStatus.WARNING
        return RowStatus.VALID

    @property
    def messages(self) -> tuple[str, ...]:
        return self.errors + self.warnings


@dataclass(frozen=True, slots=True)
class ReviewStats:
    """Thống kê toàn bộ model sau validation."""

    total: int
    valid: int
    warning: int
    error: int
    with_container: int
    with_bl: int
    with_amount: int
    total_amount: int
    fee_counts: dict[str, int]


@dataclass(frozen=True, slots=True)
class ContainerLoadPresentation:
    status: str = ""
    message: str = ""
    session_id: str | None = None


def _mapping_value(value: Mapping[str, Any], *names: str, default: Any = None) -> Any:
    for name in names:
        if name in value:
            return value[name]
    return default


def coerce_review_row(value: Any) -> ReviewRow:
    """Chuyển dataclass/dict/list/object từ core thành ``ReviewRow``."""

    if isinstance(value, ReviewRow):
        return ReviewRow.from_mapping(value.to_object(), runtime_id=value.runtime_id)
    if isinstance(value, Mapping):
        return ReviewRow.from_mapping(value)
    if is_dataclass(value):
        return coerce_review_row(asdict(value))
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        values = list(value)
        return ReviewRow.from_sequence(values)
    if any(hasattr(value, name) for name in ("cont", "container", "fee", "fee_code")):
        return ReviewRow.from_mapping(
            {
                "container": getattr(value, "cont", getattr(value, "container", None)),
                "bl": getattr(value, "bl", getattr(value, "bill_of_lading", None)),
                "fee": getattr(value, "fee", getattr(value, "fee_code", "CXD")),
                "rule": getattr(value, "rule", getattr(value, "rule_code", None)),
                "amount": getattr(value, "amount", None),
                "invoice_no": getattr(value, "invoice_no", getattr(value, "invoice_number", None)),
                "carrier": getattr(value, "carrier", getattr(value, "transport_provider", None)),
                "vessel_voyage_raw": getattr(value, "vessel_voyage_raw", None),
                "vessel_name": getattr(value, "vessel_name", None),
                "voyage_no": getattr(value, "voyage_no", None),
                "invoice_container_count": getattr(value, "invoice_container_count", None),
                "container_count_basis": getattr(value, "container_count_basis", "UNKNOWN"),
                "invoice_date": getattr(value, "invoice_date", None),
                "source_document_id": getattr(value, "source_document_id", "MANUAL"),
                "source_document_name": getattr(
                    value, "source_document_name", "Dòng thêm thủ công"
                ),
            }
        )
    raise TypeError(f"Không thể chuyển kiểu {type(value).__name__} thành dòng dữ liệu.")


def validate_row(row: ReviewRow, *, allow_negative: bool = False) -> RowValidation:
    """Kiểm tra các lỗi chặn/cảnh báo độc lập của một dòng."""

    errors: list[str] = []
    warnings: list[str] = []

    if not isinstance(row.source_document_id, str) or not row.source_document_id.strip():
        errors.append("Mã chứng từ nguồn phải là chuỗi không rỗng.")
    if not isinstance(row.source_document_name, str) or not row.source_document_name.strip():
        errors.append("Tên chứng từ nguồn phải là chuỗi không rỗng.")

    if row.cont is not None and not isinstance(row.cont, str):
        errors.append("Container phải là chuỗi hoặc null.")
    if row.bl is not None and not isinstance(row.bl, str):
        errors.append("B/L phải là chuỗi hoặc null.")
    if not isinstance(row.fee, str) or row.fee not in FEE_CATALOG:
        errors.append("Mã loại cước không thuộc danh mục chính thức.")
    if row.rule is not None and (not isinstance(row.rule, str) or row.rule not in RULE_CATALOG):
        errors.append("Mã xử lý tiền không hợp lệ.")
    if row.invoice_no is not None and not isinstance(row.invoice_no, str):
        errors.append("Số HĐ phải là chuỗi hoặc null.")
    if row.carrier is not None and not isinstance(row.carrier, str):
        errors.append("Bên vận tải phải là chuỗi hoặc null.")
    for label, value in (
        ("Tàu/chuyến nguyên văn", row.vessel_voyage_raw),
        ("Tên tàu", row.vessel_name),
        ("Số chuyến", row.voyage_no),
        ("Ngày HĐ", row.invoice_date),
    ):
        if value is not None and not isinstance(value, str):
            errors.append(f"{label} phải là chuỗi hoặc null.")
    if row.invoice_container_count is not None and (
        type(row.invoice_container_count) is not int
        or row.invoice_container_count <= 0
    ):
        errors.append("SL cont HĐ phải là số nguyên dương hoặc null.")
    if row.container_count_basis not in {"EXPLICIT", "CALCULATED", "UNKNOWN"}:
        errors.append("Căn cứ SL cont không hợp lệ.")
    elif row.invoice_container_count is None and row.container_count_basis != "UNKNOWN":
        errors.append("Không có SL cont thì căn cứ phải là UNKNOWN.")
    elif row.invoice_container_count is not None and row.container_count_basis == "UNKNOWN":
        errors.append("Có SL cont thì căn cứ phải là EXPLICIT hoặc CALCULATED.")
    if isinstance(row.invoice_date, str):
        try:
            valid_date = date.fromisoformat(row.invoice_date).isoformat() == row.invoice_date
        except ValueError:
            valid_date = False
        if not valid_date:
            errors.append("Ngày HĐ phải có định dạng YYYY-MM-DD hợp lệ.")
    if row.amount is not None:
        if isinstance(row.amount, bool) or not isinstance(row.amount, int):
            errors.append("Số tiền phải là số nguyên hoặc null.")
        elif row.amount < 0 and not allow_negative:
            errors.append("Số tiền không được âm.")
        elif row.amount < 0:
            warnings.append("Khoản điều chỉnh giảm.")

    if row.rule == "HD" and row.fee != "CB":
        errors.append("Quy tắc HD chỉ được dùng cho cước biển (CB).")
    if row.fee == "CB" and row.rule != "HD":
        errors.append("Cước biển (CB) bắt buộc dùng quy tắc HD.")

    if isinstance(row.cont, str) and row.cont:
        import re

        if re.fullmatch(r"[A-Z]{4}[0-9]{7}", row.cont) is None:
            warnings.append("Container không đúng mẫu 4 chữ cái và 7 chữ số.")
    if row.cont is None and row.bl is None:
        warnings.append("Cả container và B/L đều chưa xác định.")
    if row.fee == "CXD":
        warnings.append("Loại cước chưa đủ căn cứ (CXD).")
    if row.amount is None:
        warnings.append("Số tiền chưa xác định.")
    elif isinstance(row.amount, int) and not isinstance(row.amount, bool) and row.amount == 0:
        warnings.append("Số tiền bằng 0.")
    if row.rule is None:
        warnings.append("Quy tắc xử lý tiền chưa xác định.")
    if row.fee == "CB" and row.bl is None:
        warnings.append("Cước biển chưa có số B/L.")
    if row.fee == "CB" and row.cont is None and row.bl is not None:
        if not row.vessel_name or not row.voyage_no:
            warnings.append("Thiếu tên tàu hoặc số chuyến để đối soát BK.")
        if row.invoice_container_count is None:
            warnings.append("Thiếu SL cont hóa đơn để đối soát BK.")
    return RowValidation(tuple(errors), tuple(warnings))


class ReviewTableModel(QAbstractTableModel):
    """Model nguồn cho dữ liệu bóc tách, validation và dirty state."""

    dirtyChanged = Signal(bool)
    validationChanged = Signal(object)
    rowsChanged = Signal()

    COLUMN_NO = 0
    COLUMN_CONT = 1
    COLUMN_BL = 2
    COLUMN_VESSEL_VOYAGE = 3
    COLUMN_INVOICE_CONTAINER_COUNT = 4
    COLUMN_CONTAINER_COUNT_BASIS = 5
    COLUMN_FEE = 6
    COLUMN_FEE_NAME = 7
    COLUMN_RULE = 8
    COLUMN_RULE_NAME = 9
    COLUMN_INVOICE_NO = 10
    COLUMN_INVOICE_DATE = 11
    COLUMN_CARRIER = 12
    COLUMN_AMOUNT = 13
    COLUMN_STATUS = 14
    COLUMN_MESSAGES = 15
    COLUMN_LOOKUP_ACTION = 16

    ACTION_VISIBLE_ROLE = int(Qt.ItemDataRole.UserRole) + 10
    ACTION_ENABLED_ROLE = int(Qt.ItemDataRole.UserRole) + 11
    RUNTIME_ID_ROLE = int(Qt.ItemDataRole.UserRole) + 12
    LOAD_SESSION_ID_ROLE = int(Qt.ItemDataRole.UserRole) + 13

    HEADERS: Final[tuple[str, ...]] = (
        "STT",
        "Container",
        "B/L",
        "Tàu/chuyến",
        "SL cont HĐ",
        "Căn cứ SL",
        "Mã cước",
        "Tên loại cước",
        "Mã xử lý",
        "Diễn giải xử lý tiền",
        "Số HĐ",
        "Ngày HĐ",
        "Bên vận tải",
        "Số tiền cuối cùng (VND)",
        "Trạng thái",
        "Cảnh báo / lỗi",
        "",
    )

    def __init__(
        self,
        rows: Iterable[Any] | None = None,
        parent: Any = None,
        *,
        validator: Callable[[list[list[Any]]], Any] | Any | None = None,
        allow_negative: bool = False,
    ) -> None:
        super().__init__(parent)
        self._rows: list[ReviewRow] = []
        self._validation: list[RowValidation] = []
        self._stats = ReviewStats(0, 0, 0, 0, 0, 0, 0, 0, {})
        self._dirty = False
        self._validator = validator
        self._allow_negative = bool(allow_negative)
        self._lookup_presentations: dict[str, ContainerLoadPresentation] = {}
        self._contextual_warnings: dict[str, tuple[str, ...]] = {}
        if rows is not None:
            self.set_rows(rows, mark_dirty=False)
        else:
            self._revalidate(emit_signal=False)

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(  # noqa: N802
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        if orientation == Qt.Orientation.Vertical:
            return section + 1
        return None

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        row = self._rows[index.row()]
        result = self._validation[index.row()]
        lookup = self._lookup_presentations.get(
            row.runtime_id, ContainerLoadPresentation()
        )
        column = index.column()

        if role == Qt.ItemDataRole.UserRole:
            return self._raw_value(row, result, column, index.row())
        if role == Qt.ItemDataRole.UserRole + 1:
            return result.status.value
        if role == Qt.ItemDataRole.UserRole + 2:
            return row.as_array()
        if role == self.RUNTIME_ID_ROLE:
            return row.runtime_id
        if role == self.LOAD_SESSION_ID_ROLE:
            return lookup.session_id
        if role == self.ACTION_VISIBLE_ROLE:
            return self._lookup_eligible(row) and column == self.COLUMN_LOOKUP_ACTION
        if role == self.ACTION_ENABLED_ROLE:
            return self._lookup_eligible(row) and column == self.COLUMN_LOOKUP_ACTION
        if role == Qt.ItemDataRole.DisplayRole:
            return self._display_value(row, result, lookup, column, index.row())
        if role == Qt.ItemDataRole.ToolTipRole:
            if column == self.COLUMN_LOOKUP_ACTION:
                return (
                    "Mở hồ sơ đối soát số cont của tàu/chuyến này."
                    if lookup.session_id
                    else "Đối soát số cont cho dòng cước biển này."
                )
            messages = "\n".join(result.messages)
            return messages or "Dòng hợp lệ."
        if role == Qt.ItemDataRole.BackgroundRole:
            if result.status is RowStatus.ERROR:
                return QColor(ERROR_BG)
            if result.status is RowStatus.WARNING:
                return QColor(WARNING_BG)
            return QColor(SUCCESS_BG)
        if role == Qt.ItemDataRole.ForegroundRole:
            if column in (
                self.COLUMN_CONT,
                self.COLUMN_BL,
                self.COLUMN_INVOICE_NO,
                self.COLUMN_CARRIER,
            ) and self._raw_value(
                row, result, column, index.row()
            ) is None:
                return QColor(MUTED_TEXT)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if column in (
                self.COLUMN_NO,
                self.COLUMN_FEE,
                self.COLUMN_RULE,
                self.COLUMN_STATUS,
                self.COLUMN_LOOKUP_ACTION,
            ):
                return int(Qt.AlignmentFlag.AlignCenter)
            if column == self.COLUMN_AMOUNT:
                return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return None

    def _raw_value(
        self, row: ReviewRow, result: RowValidation, column: int, source_row: int
    ) -> Any:
        values: tuple[Any, ...] = (
            source_row + 1,
            row.cont,
            row.bl,
            " ".join(
                part
                for part in (
                    str(row.vessel_name).strip() if isinstance(row.vessel_name, str) else "",
                    str(row.voyage_no).strip() if isinstance(row.voyage_no, str) else "",
                )
                if part
            ) or None,
            row.invoice_container_count,
            row.container_count_basis,
            row.fee,
            FEE_CATALOG.get(row.fee, "") if isinstance(row.fee, str) else "",
            row.rule,
            RULE_CATALOG.get(row.rule, "")
            if row.rule is None or isinstance(row.rule, str)
            else "",
            row.invoice_no,
            row.invoice_date,
            row.carrier,
            row.amount,
            result.status.value,
            "\n".join(result.messages),
            "",
        )
        return values[column] if 0 <= column < len(values) else None

    def _display_value(
        self,
        row: ReviewRow,
        result: RowValidation,
        lookup: ContainerLoadPresentation,
        column: int,
        source_row: int,
    ) -> str | int:
        if column == self.COLUMN_LOOKUP_ACTION:
            return "Xem hồ sơ" if lookup.session_id else "Đối soát"
        value = self._raw_value(row, result, column, source_row)
        if column in (
            self.COLUMN_CONT,
            self.COLUMN_BL,
            self.COLUMN_VESSEL_VOYAGE,
            self.COLUMN_RULE,
            self.COLUMN_INVOICE_NO,
            self.COLUMN_INVOICE_DATE,
            self.COLUMN_CARRIER,
        ) and value is None:
            return "—"
        if column == self.COLUMN_AMOUNT:
            if value is None:
                return "—"
            if type(value) is int:
                return f"{value:,}".replace(",", ".")
            return str(value)
        if column == self.COLUMN_STATUS:
            return STATUS_LABELS[result.status]
        if column == self.COLUMN_MESSAGES:
            return "; ".join(result.messages) if result.messages else "Không có"
        return "" if value is None else value

    @property
    def dirty(self) -> bool:
        return self._dirty

    def is_dirty(self) -> bool:
        return self._dirty

    @property
    def stats(self) -> ReviewStats:
        return self._stats

    def validation_at(self, row: int) -> RowValidation:
        return self._validation[row]

    def row_at(self, row: int) -> ReviewRow:
        source = self._rows[row]
        return ReviewRow.from_mapping(source.to_object(), runtime_id=source.runtime_id)

    def rows(self) -> list[ReviewRow]:
        return [
            ReviewRow.from_mapping(row.to_object(), runtime_id=row.runtime_id)
            for row in self._rows
        ]

    def runtime_id_at(self, row: int) -> str:
        return self._rows[row].runtime_id

    def find_runtime_id(self, runtime_id: str) -> int | None:
        return next(
            (
                index
                for index, row in enumerate(self._rows)
                if row.runtime_id == runtime_id
            ),
            None,
        )

    def lookup_presentation(self, runtime_id: str) -> ContainerLoadPresentation:
        return self._lookup_presentations.get(
            runtime_id, ContainerLoadPresentation()
        )

    def rows_as_arrays(self) -> list[list[Any]]:
        return [row.as_array() for row in self._rows]

    def to_document(self) -> dict[str, Any]:
        return {"v": _SCHEMA_VERSION, "d": [row.to_object() for row in self._rows]}

    def set_rows(self, rows: Iterable[Any], *, mark_dirty: bool = False) -> None:
        converted = [coerce_review_row(row) for row in rows]
        self.beginResetModel()
        self._rows = converted
        self._lookup_presentations.clear()
        self._contextual_warnings.clear()
        self._revalidate(emit_signal=False)
        self.endResetModel()
        self._set_dirty(mark_dirty)
        self.validationChanged.emit(self._stats)
        self.rowsChanged.emit()

    def set_document(self, document: Any, *, mark_dirty: bool = False) -> None:
        if isinstance(document, Mapping):
            rows = document.get("d", document.get("rows", []))
        else:
            rows = getattr(document, "d", getattr(document, "rows", document))
        self.set_rows(rows, mark_dirty=mark_dirty)

    def add_row(self, row: Any, position: int | None = None) -> int:
        converted = coerce_review_row(row)
        target = len(self._rows) if position is None else max(0, min(position, len(self._rows)))
        self.beginInsertRows(QModelIndex(), target, target)
        self._rows.insert(target, converted)
        self._validation.insert(target, RowValidation())
        self.endInsertRows()
        self._after_mutation()
        return target

    def update_row(self, position: int, row: Any) -> None:
        if not (0 <= position < len(self._rows)):
            raise IndexError("Dòng cần sửa không tồn tại.")
        self._rows[position] = coerce_review_row(row)
        self._after_mutation()

    def remove_row(self, position: int) -> ReviewRow:
        return self.remove_rows([position])[0]

    def remove_rows(self, positions: Iterable[int]) -> list[ReviewRow]:
        """Xóa nhiều dòng nguồn trong một lần cập nhật validation/dirty state."""

        targets = sorted(set(positions))
        if not targets:
            return []
        if targets[0] < 0 or targets[-1] >= len(self._rows):
            raise IndexError("Dòng cần xóa không tồn tại.")

        removed = [self._rows[position] for position in targets]
        ranges: list[tuple[int, int]] = []
        start = end = targets[0]
        for position in targets[1:]:
            if position == end + 1:
                end = position
                continue
            ranges.append((start, end))
            start = end = position
        ranges.append((start, end))

        # Xóa từ cuối lên đầu để chỉ mục của các vùng phía trước không bị dịch.
        for start, end in reversed(ranges):
            self.beginRemoveRows(QModelIndex(), start, end)
            removed_range = self._rows[start : end + 1]
            del self._rows[start : end + 1]
            del self._validation[start : end + 1]
            for row in removed_range:
                self._lookup_presentations.pop(row.runtime_id, None)
                self._contextual_warnings.pop(row.runtime_id, None)
            self.endRemoveRows()

        self._after_mutation()
        return removed

    def replace_row(self, position: int, rows: Iterable[Any]) -> tuple[int, int]:
        if not (0 <= position < len(self._rows)):
            raise IndexError("Dòng cần thay thế không tồn tại.")
        replacements = [coerce_review_row(row) for row in rows]
        if not replacements:
            raise ValueError("Cần ít nhất một dòng thay thế.")
        removed_runtime_id = self._rows[position].runtime_id
        self.beginResetModel()
        self._rows[position : position + 1] = replacements
        self._lookup_presentations.pop(removed_runtime_id, None)
        self._contextual_warnings.pop(removed_runtime_id, None)
        self._revalidate(emit_signal=False)
        self.endResetModel()
        self._set_dirty(True)
        self.validationChanged.emit(self._stats)
        self.rowsChanged.emit()
        return position, position + len(replacements) - 1

    def set_lookup_presentation(
        self,
        runtime_id: str,
        *,
        status: str,
        message: str = "",
        session_id: str | None = None,
    ) -> None:
        source_row = self.find_runtime_id(runtime_id)
        if source_row is None:
            return
        self._lookup_presentations[runtime_id] = ContainerLoadPresentation(
            status=str(status).upper(),
            message=str(message),
            session_id=session_id,
        )
        self.dataChanged.emit(
            self.index(source_row, self.COLUMN_LOOKUP_ACTION),
            self.index(source_row, self.COLUMN_LOOKUP_ACTION),
        )

    def replace_lookup_presentations(
        self,
        presentations: Mapping[str, ContainerLoadPresentation],
    ) -> None:
        self._lookup_presentations = {
            runtime_id: value
            for runtime_id, value in presentations.items()
            if self.find_runtime_id(runtime_id) is not None
        }
        if self.rowCount():
            self.dataChanged.emit(
                self.index(0, self.COLUMN_LOOKUP_ACTION),
                self.index(self.rowCount() - 1, self.COLUMN_LOOKUP_ACTION),
            )

    def set_contextual_warnings(
        self,
        warnings: Mapping[str, Sequence[str]],
    ) -> None:
        normalized = {
            runtime_id: tuple(dict.fromkeys(str(item) for item in values if str(item)))
            for runtime_id, values in warnings.items()
            if self.find_runtime_id(runtime_id) is not None
        }
        if normalized == self._contextual_warnings:
            return
        self._contextual_warnings = normalized
        self._revalidate(emit_signal=False)
        if self.rowCount():
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
            )
        self.validationChanged.emit(self._stats)

    def set_lookup_busy(self, busy: bool) -> None:
        del busy

    @staticmethod
    def _load_is_active(presentation: ContainerLoadPresentation) -> bool:
        return False

    @staticmethod
    def _lookup_eligible(row: ReviewRow) -> bool:
        container_missing = row.cont is None or (
            isinstance(row.cont, str) and not row.cont.strip()
        )
        return (
            container_missing
            and row.fee == "CB"
        )

    def mark_clean(self) -> None:
        self._set_dirty(False)

    def mark_dirty(self) -> None:
        self._set_dirty(True)

    def revalidate(self) -> ReviewStats:
        self._revalidate(emit_signal=True)
        if self.rowCount():
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
            )
        return self._stats

    def set_allow_negative(self, allow: bool) -> None:
        value = bool(allow)
        if self._allow_negative == value:
            return
        self._allow_negative = value
        self.revalidate()

    def first_error_row(self) -> int | None:
        return next(
            (index for index, result in enumerate(self._validation) if result.errors),
            None,
        )

    def _set_dirty(self, dirty: bool) -> None:
        if self._dirty == dirty:
            return
        self._dirty = dirty
        self.dirtyChanged.emit(dirty)

    def _after_mutation(self) -> None:
        self._revalidate(emit_signal=False)
        if self.rowCount():
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
            )
        self._set_dirty(True)
        self.validationChanged.emit(self._stats)
        self.rowsChanged.emit()

    def _revalidate(self, *, emit_signal: bool) -> None:
        results = [
            validate_row(row, allow_negative=self._allow_negative)
            for row in self._rows
        ]
        duplicate_keys = [
            tuple((key, repr(value)) for key, value in row.to_object().items())
            for row in self._rows
        ]
        duplicate_counts = Counter(duplicate_keys)
        for index, row in enumerate(self._rows):
            if duplicate_counts[duplicate_keys[index]] > 1:
                result = results[index]
                results[index] = RowValidation(
                    result.errors,
                    result.warnings + ("Dòng có đủ 7 giá trị trùng hoàn toàn với dòng khác.",),
                )

        external = self._run_external_validator()
        if external:
            results = self._merge_external_results(results, external)
        for index, row in enumerate(self._rows):
            contextual = self._contextual_warnings.get(row.runtime_id, ())
            if not contextual:
                continue
            old = results[index]
            results[index] = RowValidation(
                old.errors,
                tuple(dict.fromkeys(old.warnings + contextual)),
            )
        self._validation = results
        status_counts = Counter(result.status for result in results)
        valid_amounts = [
            row.amount
            for row in self._rows
            if isinstance(row.amount, int) and not isinstance(row.amount, bool)
        ]
        self._stats = ReviewStats(
            total=len(self._rows),
            valid=status_counts[RowStatus.VALID],
            warning=status_counts[RowStatus.WARNING],
            error=status_counts[RowStatus.ERROR],
            with_container=sum(row.cont is not None for row in self._rows),
            with_bl=sum(row.bl is not None for row in self._rows),
            with_amount=len(valid_amounts),
            total_amount=sum(valid_amounts),
            fee_counts=dict(Counter(str(row.fee) for row in self._rows)),
        )
        if emit_signal:
            self.validationChanged.emit(self._stats)

    def _run_external_validator(self) -> Any:
        if self._validator is None:
            return None
        try:
            if callable(self._validator):
                return self._validator(self.rows_as_arrays())
            for name in ("validate_rows", "validate_document", "validate"):
                method = getattr(self._validator, name, None)
                if callable(method):
                    payload: Any = self.to_document() if name == "validate_document" else self.rows_as_arrays()
                    if name == "validate_document" and self._allow_negative:
                        try:
                            return method(payload, allow_negative=True)
                        except TypeError:
                            pass
                    return method(payload)
        except Exception:
            # Validation nội bộ vẫn đảm bảo UI hoạt động; lỗi service sẽ được lớp điều
            # phối ghi log/hiển thị thay vì làm model Qt sập.
            return None
        return None

    @staticmethod
    def _merge_external_results(
        current: list[RowValidation], external: Any
    ) -> list[RowValidation]:
        row_results = getattr(external, "row_results", getattr(external, "rows", external))
        if not isinstance(row_results, Sequence) or isinstance(row_results, (str, bytes)):
            return current
        merged = list(current)
        for index, value in enumerate(row_results[: len(merged)]):
            errors = getattr(value, "errors", ())
            warnings = getattr(value, "warnings", ())
            if isinstance(value, Mapping):
                errors = value.get("errors", errors)
                warnings = value.get("warnings", warnings)
            extra_errors = [str(item) for item in (errors or ())]
            extra_warnings = [str(item) for item in (warnings or ())]
            for issue in getattr(value, "issues", ()) or ():
                message = str(getattr(issue, "message", issue))
                severity = str(getattr(issue, "severity", "")).casefold()
                if "error" in severity or "lỗi" in severity:
                    extra_errors.append(message)
                elif "warning" in severity or "cảnh báo" in severity:
                    extra_warnings.append(message)
            if extra_errors or extra_warnings:
                old = merged[index]
                merged[index] = RowValidation(
                    tuple(dict.fromkeys(old.errors + tuple(extra_errors))),
                    tuple(dict.fromkeys(old.warnings + tuple(extra_warnings))),
                )
        return merged


class ReviewFilterProxyModel(QSortFilterProxyModel):
    """Proxy tìm thông tin chứng từ, lọc fee/trạng thái và sort an toàn."""

    def __init__(self, parent: Any = None) -> None:
        super().__init__(parent)
        self._search_text = ""
        self._fee = ""
        self._status = ""
        self._source_document_id = ""
        self.setDynamicSortFilter(True)
        self.setSortCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

    def set_search_text(self, text: str) -> None:
        modern = self._begin_filter_update()
        self._search_text = text.strip().casefold()
        self._end_filter_update(modern)

    def set_fee_filter(self, fee: str | None) -> None:
        modern = self._begin_filter_update()
        self._fee = (fee or "").strip().upper()
        self._end_filter_update(modern)

    def set_status_filter(self, status: str | RowStatus | None) -> None:
        modern = self._begin_filter_update()
        if isinstance(status, RowStatus):
            self._status = status.value
        else:
            normalized = (status or "").strip().casefold()
            labels = {
                "hợp lệ": RowStatus.VALID.value,
                "cảnh báo": RowStatus.WARNING.value,
                "lỗi": RowStatus.ERROR.value,
            }
            self._status = labels.get(normalized, normalized)
        self._end_filter_update(modern)

    def set_document_filter(self, source_document_id: str | None) -> None:
        modern = self._begin_filter_update()
        self._source_document_id = (source_document_id or "").strip()
        self._end_filter_update(modern)

    def clear_filters(self) -> None:
        modern = self._begin_filter_update()
        self._search_text = ""
        self._fee = ""
        self._status = ""
        self._source_document_id = ""
        self._end_filter_update(modern)

    def _begin_filter_update(self) -> bool:
        begin = getattr(self, "beginFilterChange", None)
        if callable(begin):
            begin()
            return True
        return False

    def _end_filter_update(self, modern: bool) -> None:
        if modern:
            self.endFilterChange(QSortFilterProxyModel.Direction.Rows)
        else:
            self.invalidateFilter()

    def filterAcceptsRow(  # noqa: N802
        self, source_row: int, source_parent: QModelIndex
    ) -> bool:
        model = self.sourceModel()
        if model is None:
            return False
        if self._search_text:
            cont = model.index(source_row, ReviewTableModel.COLUMN_CONT, source_parent).data(
                Qt.ItemDataRole.UserRole
            )
            bl = model.index(source_row, ReviewTableModel.COLUMN_BL, source_parent).data(
                Qt.ItemDataRole.UserRole
            )
            vessel_voyage = model.index(
                source_row, ReviewTableModel.COLUMN_VESSEL_VOYAGE, source_parent
            ).data(Qt.ItemDataRole.UserRole)
            invoice_no = model.index(
                source_row, ReviewTableModel.COLUMN_INVOICE_NO, source_parent
            ).data(Qt.ItemDataRole.UserRole)
            carrier = model.index(
                source_row, ReviewTableModel.COLUMN_CARRIER, source_parent
            ).data(Qt.ItemDataRole.UserRole)
            haystack = (
                f"{cont or ''} {bl or ''} {vessel_voyage or ''} "
                f"{invoice_no or ''} {carrier or ''}"
            ).casefold()
            if self._search_text not in haystack:
                return False
        if self._fee:
            fee = model.index(source_row, ReviewTableModel.COLUMN_FEE, source_parent).data(
                Qt.ItemDataRole.UserRole
            )
            if str(fee).upper() != self._fee:
                return False
        if self._status:
            status = model.index(source_row, ReviewTableModel.COLUMN_STATUS, source_parent).data(
                Qt.ItemDataRole.UserRole + 1
            )
            if str(status).casefold() != self._status:
                return False
        if self._source_document_id:
            row = getattr(model, "row_at", lambda _index: None)(source_row)
            if row is None or row.source_document_id != self._source_document_id:
                return False
        return True

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:  # noqa: N802
        left_value = left.data(Qt.ItemDataRole.UserRole)
        right_value = right.data(Qt.ItemDataRole.UserRole)
        if left_value is None and right_value is not None:
            return False
        if right_value is None and left_value is not None:
            return True
        if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
            return left_value < right_value
        return str(left_value or "").casefold() < str(right_value or "").casefold()


# Alias dễ hiểu cho test/tích hợp bên ngoài.
ReviewSortFilterProxyModel = ReviewFilterProxyModel
