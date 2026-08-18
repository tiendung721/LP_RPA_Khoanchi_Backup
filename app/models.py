"""Các model thuần Python dùng xuyên suốt tầng dữ liệu và giao diện."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, ClassVar, Iterable, Mapping, Sequence

from app.constants import SCHEMA_VERSION


class BatchStatus(str, Enum):
    """Vòng đời của một lô JSON."""

    RECEIVED = "RECEIVED"
    REVIEWING = "REVIEWING"
    READY = "READY"
    INVALID = "INVALID"
    ARCHIVED = "ARCHIVED"


class Severity(str, Enum):
    """Mức nghiêm trọng, đồng thời là trạng thái hiển thị của một dòng."""

    VALID = "VALID"
    WARNING = "WARNING"
    ERROR = "ERROR"


@dataclass(slots=True)
class DataRow:
    """Một dòng khoản chi dùng chung cho schema v1, v2 và v3.

    JSON ghi mới là object schema v3. Mảng
    ``[cont, bl, fee, rule, invoice_no, carrier, amount]`` và object 13 khóa
    chỉ còn là adapter tương thích v1/v2. Thuộc tính ``amount`` vẫn giữ vị trí cũ trong dataclass để
    mã Python dùng năm đối số vị trí không bị hiểu sai.
    """

    FIELD_NAMES: ClassVar[tuple[str, ...]] = (
        "cont",
        "bl",
        "fee",
        "rule",
        "amount",
        "invoice_no",
        "carrier",
        "vessel_voyage_raw",
        "vessel_name",
        "voyage_no",
        "invoice_container_count",
        "container_count_basis",
        "invoice_date",
        "source_document_id",
        "source_document_name",
    )

    cont: str | None
    bl: str | None
    fee: str
    rule: str | None
    amount: int | None
    invoice_no: str | None = None
    carrier: str | None = None
    vessel_voyage_raw: str | None = None
    vessel_name: str | None = None
    voyage_no: str | None = None
    invoice_container_count: int | None = None
    container_count_basis: str = "UNKNOWN"
    invoice_date: str | None = None
    source_document_id: str = "MANUAL"
    source_document_name: str = "Dòng thêm thủ công"

    @classmethod
    def from_sequence(cls, value: Sequence[Any]) -> "DataRow":
        if isinstance(value, (str, bytes, bytearray)) or len(value) != 7:
            raise ValueError("Mỗi dòng dữ liệu phải là một mảng có đúng 7 phần tử.")
        return cls(
            cont=value[0],
            bl=value[1],
            fee=value[2],
            rule=value[3],
            invoice_no=value[4],
            carrier=value[5],
            amount=value[6],
        )

    def to_list(self) -> list[Any]:
        """Trả mảng v1 cho các adapter nội bộ cũ."""

        return [
            self.cont,
            self.bl,
            self.fee,
            self.rule,
            self.invoice_no,
            self.carrier,
            self.amount,
        ]

    @classmethod
    def from_mapping(cls, value: Mapping[str, Any]) -> "DataRow":
        return cls(
            cont=value.get("container", value.get("cont")),
            bl=value.get("bl"),
            fee=value.get("fee"),
            rule=value.get("rule"),
            amount=value.get("amount"),
            invoice_no=value.get("invoice_no"),
            carrier=value.get("carrier"),
            vessel_voyage_raw=value.get("vessel_voyage_raw"),
            vessel_name=value.get("vessel_name"),
            voyage_no=value.get("voyage_no"),
            invoice_container_count=value.get("invoice_container_count"),
            container_count_basis=value.get("container_count_basis", "UNKNOWN"),
            invoice_date=value.get("invoice_date"),
            source_document_id=value.get("source_document_id", "MANUAL"),
            source_document_name=value.get("source_document_name", "Dòng thêm thủ công"),
        )

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

    def copy_with(self, **changes: Any) -> "DataRow":
        unknown = set(changes).difference(self.FIELD_NAMES)
        if unknown:
            raise TypeError(f"Trường DataRow không hợp lệ: {', '.join(sorted(unknown))}")
        values = {name: getattr(self, name) for name in self.FIELD_NAMES}
        values.update(changes)
        return DataRow(**values)


@dataclass(slots=True, init=False)
class BatchDocument:
    """Tài liệu JSON đã ánh xạ sang model nội bộ."""

    v: int
    rows: list[DataRow]

    def __init__(
        self,
        v: int = SCHEMA_VERSION,
        rows: Iterable[DataRow] | None = None,
        *,
        version: int | None = None,
        data: Iterable[DataRow] | None = None,
        d: Iterable[DataRow] | None = None,
    ) -> None:
        if version is not None:
            v = version
        supplied = [candidate is not None for candidate in (rows, data, d)]
        if sum(supplied) > 1:
            raise TypeError("Chỉ truyền một trong rows, data hoặc d.")
        selected = rows if rows is not None else data if data is not None else d
        self.v = v
        self.rows = list(selected or ())

    @property
    def version(self) -> int:
        return self.v

    @version.setter
    def version(self, value: int) -> None:
        self.v = value

    @property
    def data(self) -> list[DataRow]:
        return self.rows

    @property
    def d(self) -> list[DataRow]:
        return self.rows

    def to_dict(self) -> dict[str, Any]:
        # Thứ tự chèn khóa là một phần của hợp đồng serialize: v trước d.
        return {"v": SCHEMA_VERSION, "d": [row.to_object() for row in self.rows]}


# Tên ngắn thân thiện cho code tích hợp.
JsonDocument = BatchDocument
Document = BatchDocument


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    severity: Severity
    code: str
    message: str
    row_index: int | None = None
    field: str | None = None

    @property
    def is_blocking(self) -> bool:
        return self.severity is Severity.ERROR


@dataclass(slots=True)
class RowValidation:
    row_index: int
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def status(self) -> Severity:
        if any(issue.severity is Severity.ERROR for issue in self.issues):
            return Severity.ERROR
        if any(issue.severity is Severity.WARNING for issue in self.issues):
            return Severity.WARNING
        return Severity.VALID

    @property
    def messages(self) -> list[str]:
        return [issue.message for issue in self.issues]

    @property
    def has_errors(self) -> bool:
        return self.status is Severity.ERROR

    @property
    def has_warnings(self) -> bool:
        return any(issue.severity is Severity.WARNING for issue in self.issues)


@dataclass(slots=True)
class ValidationSummary:
    total_rows: int = 0
    valid_count: int = 0
    warning_count: int = 0
    error_count: int = 0
    container_count: int = 0
    bl_count: int = 0
    amount_count: int = 0
    total_amount: int = 0
    fee_counts: dict[str, int] = field(default_factory=dict)

    @property
    def row_count(self) -> int:
        return self.total_rows


@dataclass(slots=True)
class ValidationResult:
    issues: list[ValidationIssue] = field(default_factory=list)
    row_results: list[RowValidation] = field(default_factory=list)
    summary: ValidationSummary = field(default_factory=ValidationSummary)

    @property
    def is_valid(self) -> bool:
        return not self.has_errors

    @property
    def has_errors(self) -> bool:
        return any(issue.severity is Severity.ERROR for issue in self.issues)

    @property
    def has_warnings(self) -> bool:
        return any(issue.severity is Severity.WARNING for issue in self.issues)

    @property
    def error_count(self) -> int:
        return max(
            self.summary.error_count,
            sum(
                1
                for issue in self.issues
                if issue.severity is Severity.ERROR and issue.row_index is None
            ),
        )

    @property
    def warning_count(self) -> int:
        return self.summary.warning_count

    @property
    def valid_count(self) -> int:
        return self.summary.valid_count

    @property
    def first_error_row(self) -> int | None:
        for issue in self.issues:
            if issue.severity is Severity.ERROR and issue.row_index is not None:
                return issue.row_index
        return None

    def for_row(self, row_index: int) -> RowValidation:
        if row_index < 0 or row_index >= len(self.row_results):
            raise IndexError("Chỉ số dòng nằm ngoài phạm vi.")
        return self.row_results[row_index]


@dataclass(slots=True)
class BatchMetadata:
    id: int
    source_filename: str
    source_output_path: Path | None
    original_archive_path: Path
    working_path: Path
    ready_path: Path | None
    sha256: str
    status: BatchStatus
    received_at: str
    last_opened_at: str | None = None
    last_saved_at: str | None = None
    confirmed_at: str | None = None
    row_count: int = 0
    valid_count: int = 0
    warning_count: int = 0
    error_count: int = 0
    total_amount: int = 0
    last_error: str | None = None
    source_kind: str = "ASSISTANT"
    reconciliation_group_id: int | None = None

    @property
    def batch_id(self) -> int:
        return self.id

    @property
    def filename(self) -> str:
        return self.source_filename

    @property
    def is_ready(self) -> bool:
        return self.status is BatchStatus.READY and self.ready_path is not None


@dataclass(slots=True)
class BatchReview:
    metadata: BatchMetadata
    document: BatchDocument
    validation: ValidationResult

    @property
    def batch(self) -> BatchMetadata:
        return self.metadata

    @property
    def rows(self) -> list[DataRow]:
        return self.document.rows


@dataclass(slots=True)
class ReceiveResult:
    batch: BatchMetadata
    duplicate: bool
    message: str
    review: BatchReview | None = None

    @property
    def metadata(self) -> BatchMetadata:
        return self.batch

    @property
    def batch_id(self) -> int:
        return self.batch.id

    @property
    def created(self) -> bool:
        return not self.duplicate


@dataclass(frozen=True, slots=True)
class FileStabilityResult:
    path: Path
    size: int
    mtime_ns: int
    elapsed_seconds: float


def rows_from_iterable(values: Iterable[DataRow | Sequence[Any]]) -> list[DataRow]:
    """Chuyển một iterable linh hoạt thành danh sách ``DataRow`` mới."""

    return [
        value if isinstance(value, DataRow) else DataRow.from_sequence(value)
        for value in values
    ]


def metadata_as_dict(metadata: BatchMetadata) -> Mapping[str, Any]:
    """Biểu diễn chỉ-đọc hữu ích cho các adapter không dùng dataclass."""

    return {
        "id": metadata.id,
        "source_filename": metadata.source_filename,
        "source_output_path": metadata.source_output_path,
        "original_archive_path": metadata.original_archive_path,
        "working_path": metadata.working_path,
        "ready_path": metadata.ready_path,
        "sha256": metadata.sha256,
        "status": metadata.status,
        "received_at": metadata.received_at,
        "last_opened_at": metadata.last_opened_at,
        "last_saved_at": metadata.last_saved_at,
        "confirmed_at": metadata.confirmed_at,
        "row_count": metadata.row_count,
        "valid_count": metadata.valid_count,
        "warning_count": metadata.warning_count,
        "error_count": metadata.error_count,
        "total_amount": metadata.total_amount,
        "last_error": metadata.last_error,
        "source_kind": metadata.source_kind,
        "reconciliation_group_id": metadata.reconciliation_group_id,
    }
