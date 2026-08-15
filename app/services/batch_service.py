"""Điều phối một file JSON hiện hành duy nhất trong thư mục Output."""

from __future__ import annotations

import logging
import hashlib
import os
import shutil
import sqlite3
import stat
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

from app.config import AppPaths, AppSettings
from app.constants import DEFAULT_MAX_FILE_SIZE_BYTES
from app.database import Database
from app.models import (
    BatchDocument,
    BatchMetadata,
    BatchReview,
    BatchStatus,
    DataRow,
    ReceiveResult,
    ValidationResult,
)
from app.repositories.batch_repository import BatchRepository, local_now_iso
from app.schema import coerce_document, document_to_dict
from app.services.carrier_policy import strip_received_gpt_carriers
from app.services.file_stability import file_sha256, is_temporary_file
from app.services.json_codec import JsonCodec, JsonCodecError
from app.services.validation_service import ValidationService

LOGGER = logging.getLogger(__name__)


class BatchServiceError(RuntimeError):
    """Lỗi nghiệp vụ có thông điệp phù hợp để controller chuyển cho UI."""


class BatchNotFoundError(BatchServiceError):
    """Không có batch tương ứng trong lịch sử."""


class BatchDataError(BatchServiceError):
    """Working file không thể đọc hoặc ghi an toàn."""


class BatchValidationError(BatchServiceError):
    def __init__(self, message: str, validation: ValidationResult) -> None:
        super().__init__(message)
        self.validation = validation
        self.first_error_row = validation.first_error_row


def calculate_sha256(path: str | Path) -> str:
    return file_sha256(path)


class BatchService:
    def __init__(
        self,
        paths: AppPaths | AppSettings | str | Path | BatchRepository,
        repository: BatchRepository | AppPaths | AppSettings | str | Path | None = None,
        *,
        codec: JsonCodec | None = None,
        validation_service: ValidationService | None = None,
        validator: ValidationService | None = None,
        max_file_size_bytes: int | None = None,
        output_write_callback: Callable[[Path], None] | None = None,
    ) -> None:
        # Hỗ trợ cả BatchService(paths, repo) và BatchService(repo, paths).
        if isinstance(paths, BatchRepository):
            actual_repository = paths
            if repository is None:
                raise TypeError("Cần truyền AppPaths khi tham số đầu là repository.")
            actual_paths = self._coerce_paths(repository)
        else:
            actual_paths = self._coerce_paths(paths)
            actual_repository = (
                repository if isinstance(repository, BatchRepository) else None
            )
            if repository is not None and actual_repository is None:
                raise TypeError("repository phải là BatchRepository.")

        actual_paths.ensure_directories()
        self.paths = actual_paths
        self._cleanup_legacy_json_storage()
        self._owns_repository = actual_repository is None
        self.repository = actual_repository or BatchRepository(
            Database(actual_paths.database_path)
        )
        self.codec = codec or JsonCodec()
        self.validation_service = (
            validation_service or validator or ValidationService()
        )
        self.max_file_size_bytes = (
            max_file_size_bytes
            if max_file_size_bytes is not None
            else DEFAULT_MAX_FILE_SIZE_BYTES
        )
        if self.max_file_size_bytes <= 0:
            raise ValueError("Giới hạn kích thước file phải lớn hơn 0.")
        self._output_write_callback = output_write_callback
        self._repair_reconciliation_batches()
        self._current_output_batch_id = self._find_current_output_batch_id()

    def set_output_write_callback(
        self, callback: Callable[[Path], None] | None
    ) -> None:
        self._output_write_callback = callback

    def _repair_reconciliation_batches(self) -> None:
        """Chuẩn hóa batch con beta mà không biến nó thành file Output hiện hành."""

        rows = self.repository.database.query_all(
            """
            SELECT b.id, b.ready_path, b.status, g.is_current, g.status AS group_status
            FROM batches AS b
            JOIN sea_freight_reconciliation_groups AS g
              ON g.id = b.reconciliation_group_id
            WHERE b.source_kind = 'SEA_FREIGHT_RECONCILIATION'
            """
        )
        for row in rows:
            batch_id = int(row["id"])
            if not bool(row["is_current"]) or str(row["group_status"]) == "CANCELLED":
                if str(row["status"]) != BatchStatus.ARCHIVED.value:
                    self.repository.update_batch(batch_id, status=BatchStatus.ARCHIVED)
                continue
            path = Path(str(row["ready_path"] or ""))
            if not path.is_file():
                continue
            try:
                document = self.codec.load(path)
                validation = self.validation_service.validate_document(document)
            except Exception:
                continue
            if not validation.has_errors and str(row["status"]) != BatchStatus.READY.value:
                self.repository.update_batch(
                    batch_id,
                    status=BatchStatus.READY,
                    ready_path=path,
                    working_path=path,
                    row_count=validation.summary.total_rows,
                    valid_count=validation.summary.valid_count,
                    warning_count=validation.summary.warning_count,
                    error_count=validation.error_count,
                    total_amount=validation.summary.total_amount,
                    last_error=None,
                )

    def update_paths(self, paths: AppPaths) -> None:
        paths.ensure_directories()
        output_changed = self._path_key(paths.output_dir) != self._path_key(
            self.paths.output_dir
        )
        self.paths = paths
        self._cleanup_legacy_json_storage()
        if output_changed:
            self._current_output_batch_id = self._find_current_output_batch_id()
            self.repository.set_active_batch_id(self._current_output_batch_id)

    @property
    def current_output_batch_id(self) -> int | None:
        return self._current_output_batch_id

    def get_current_output_batch(self) -> BatchMetadata | None:
        batch_id = self._current_output_batch_id
        if batch_id is None:
            return None
        metadata = self.repository.get_by_id(batch_id)
        if metadata is None or metadata.status is not BatchStatus.INVALID:
            return metadata
        try:
            return self.load_batch(batch_id).metadata
        except BatchDataError:
            # File vẫn sai cấu trúc theo schema hiện hành; giữ nguyên INVALID.
            return self.repository.get_by_id(batch_id)

    @staticmethod
    def _coerce_paths(
        value: AppPaths | AppSettings | str | Path,
    ) -> AppPaths:
        if isinstance(value, AppPaths):
            return value
        if isinstance(value, AppSettings):
            return value.paths
        return AppPaths.from_data_root(value)

    def receive_file(self, path: str | Path) -> ReceiveResult:
        """Tiếp nhận file ổn định và thay thế JSON hiện hành trong Output."""

        source = Path(path)
        LOGGER.info("Bắt đầu tiếp nhận file: %s", source.name)
        if is_temporary_file(source):
            raise BatchServiceError("File vẫn mang hậu tố tải xuống tạm thời.")
        if source.suffix.lower() != ".json":
            raise BatchServiceError("Chỉ có thể tiếp nhận file có đuôi .json.")
        try:
            stat_result = source.stat()
        except OSError as exc:
            raise BatchServiceError(f"Không thể truy cập file: {source}.") from exc
        if not source.is_file():
            raise BatchServiceError("Đường dẫn đã chọn không phải là file.")
        if stat_result.st_size > self.max_file_size_bytes:
            raise BatchServiceError(
                "File vượt giới hạn kích thước đã cấu hình."
            )
        if self._is_output_candidate(source) and not self._is_latest_output_candidate(
            source
        ):
            raise BatchServiceError(
                "Đã bỏ qua file cũ vì Output có một file kết quả mới hơn."
            )

        try:
            sha256 = calculate_sha256(source)
        except OSError as exc:
            raise BatchServiceError("Không thể đọc file để tính SHA-256.") from exc
        LOGGER.info("SHA-256 file %s: %s", source.name, sha256)
        current = self.get_current_output_batch()
        if (
            current is not None
            and current.source_output_path is not None
            and self._path_key(current.source_output_path) == self._path_key(source)
        ):
            LOGGER.info("Bỏ qua sự kiện lặp của JSON hiện hành: %s", source.name)
            return ReceiveResult(
                batch=current,
                duplicate=True,
                message="File hiện hành không thay đổi.",
            )

        source_kind = (
            "BANG_KE"
            if source.name.casefold().startswith("ket_qua_boc_tach_bang_ke_")
            else "ASSISTANT"
        )
        received_at = local_now_iso()
        output_path = self._promote_output_file(source, received_at)
        duplicate = self.repository.get_by_sha256(sha256)
        try:
            if duplicate is None:
                batch = self.repository.create_batch(
                    source_filename=output_path.name,
                    source_output_path=output_path,
                    original_archive_path=output_path,
                    working_path=output_path,
                    sha256=sha256,
                    status=BatchStatus.RECEIVED,
                    received_at=received_at,
                    source_kind=source_kind,
                )
            else:
                batch = self.repository.update_batch(
                    duplicate.id,
                    source_filename=output_path.name,
                    source_output_path=output_path,
                    original_archive_path=output_path,
                    working_path=output_path,
                    ready_path=None,
                    status=BatchStatus.RECEIVED,
                    received_at=received_at,
                    last_opened_at=None,
                    last_saved_at=None,
                    confirmed_at=None,
                    row_count=0,
                    valid_count=0,
                    warning_count=0,
                    error_count=0,
                    total_amount=0,
                    last_error=None,
                    source_kind=source_kind,
                )
        except sqlite3.IntegrityError:
            # Hai event có thể đồng thời vượt qua bước tra SHA-256.
            existing = self.repository.get_by_sha256(sha256)
            if existing is None:
                raise BatchServiceError("Không thể tạo metadata cho file JSON.")
            batch = self.repository.update_batch(
                existing.id,
                source_filename=output_path.name,
                source_output_path=output_path,
                original_archive_path=output_path,
                working_path=output_path,
                ready_path=None,
                status=BatchStatus.RECEIVED,
                received_at=received_at,
                last_opened_at=None,
                last_saved_at=None,
                confirmed_at=None,
                last_error=None,
                source_kind=source_kind,
            )

        self._current_output_batch_id = batch.id
        try:
            document = self.codec.load(output_path)
        except JsonCodecError as exc:
            LOGGER.warning("Batch %s không parse/schema được: %s", batch.id, exc)
            metadata = self.repository.update_batch(
                batch.id,
                status=BatchStatus.INVALID,
                error_count=1,
                last_error=str(exc),
            )
            self.repository.set_active_batch_id(batch.id)
            return ReceiveResult(
                batch=metadata,
                duplicate=False,
                message="File JSON không đọc được hoặc sai cấu trúc gốc.",
            )

        document = self._prepare_received_gpt_document(output_path, document)

        validation = self._validate_document(document, source_kind=source_kind)
        batch = self.repository.update_batch(
            batch.id,
            row_count=validation.summary.total_rows,
            valid_count=validation.summary.valid_count,
            warning_count=validation.summary.warning_count,
            error_count=validation.error_count,
            total_amount=validation.summary.total_amount,
            last_error=None,
        )
        self._activate_if_appropriate(batch)
        review = BatchReview(batch, document, validation)
        LOGGER.info(
            "Đã tiếp nhận batch %s: %s dòng, %s cảnh báo, %s lỗi",
            batch.id,
            batch.row_count,
            batch.warning_count,
            batch.error_count,
        )
        return ReceiveResult(
            batch=batch,
            duplicate=False,
            message="Đã tiếp nhận file JSON thành công.",
            review=review,
        )

    ingest_file = receive_file
    process_file = receive_file
    receive = receive_file

    def load_batch(self, batch_id: int) -> BatchReview:
        metadata = self._require_batch(batch_id)
        try:
            document = self.codec.load(metadata.working_path)
        except JsonCodecError as exc:
            self.repository.update_batch(
                batch_id,
                status=BatchStatus.INVALID,
                error_count=max(1, metadata.error_count),
                last_error=str(exc),
            )
            raise BatchDataError(
                "Bản làm việc không còn là JSON có cấu trúc hợp lệ."
            ) from exc
        if metadata.last_saved_at is None:
            document = self._prepare_received_gpt_document(
                metadata.working_path, document
            )
        validation = self._validate_document(
            document, source_kind=metadata.source_kind
        )
        new_status = (
            BatchStatus.REVIEWING
            if metadata.status in {BatchStatus.RECEIVED, BatchStatus.INVALID}
            else metadata.status
        )
        metadata = self.repository.update_batch(
            batch_id,
            status=new_status,
            last_opened_at=local_now_iso(),
            row_count=validation.summary.total_rows,
            valid_count=validation.summary.valid_count,
            warning_count=validation.summary.warning_count,
            error_count=validation.error_count,
            total_amount=validation.summary.total_amount,
            last_error=None,
        )
        if batch_id == self._current_output_batch_id:
            self.repository.set_active_batch_id(batch_id)
        return BatchReview(metadata, document, validation)

    open_batch = load_batch

    def save_working(
        self,
        batch_id: int,
        document: BatchDocument
        | dict[str, Any]
        | Iterable[DataRow | Sequence[Any]],
    ) -> BatchReview:
        metadata = self._require_current_batch(batch_id)
        normalized = coerce_document(document)
        output_path, reloaded = self._write_current_document(metadata, normalized)
        revalidation = self._validate_document(
            reloaded, source_kind=metadata.source_kind
        )
        metadata = self.repository.mark_saved(
            batch_id,
            revalidation.summary,
            status=BatchStatus.REVIEWING,
            error_count=revalidation.error_count,
        )
        metadata = self.repository.update_batch(
            batch_id,
            source_filename=output_path.name,
            source_output_path=output_path,
            original_archive_path=output_path,
            working_path=output_path,
            ready_path=None,
        )
        self.repository.set_active_batch_id(batch_id)
        LOGGER.info("Đã lưu JSON hiện hành của batch %s", batch_id)
        return BatchReview(metadata, reloaded, revalidation)

    save_working_copy = save_working
    save_batch = save_working

    def confirm_batch(
        self,
        batch_id: int,
        document: BatchDocument
        | dict[str, Any]
        | Iterable[DataRow | Sequence[Any]]
        | None = None,
    ) -> BatchReview:
        metadata = self._require_current_batch(batch_id)
        document_to_save = (
            self.codec.load(metadata.working_path)
            if document is None
            else coerce_document(document)
        )
        validation = self._validate_document(
            document_to_save, source_kind=metadata.source_kind
        )
        if validation.has_errors:
            raise BatchValidationError(
                "Không thể xác nhận vì lô vẫn còn lỗi chặn.",
                validation,
            )
        output_path, reloaded = self._write_current_document(
            metadata, document_to_save
        )
        revalidation = self._validate_document(
            reloaded, source_kind=metadata.source_kind
        )
        if revalidation.has_errors:
            raise BatchValidationError(
                "Dữ liệu đọc lại vẫn còn lỗi chặn.",
                revalidation,
            )
        timestamp = local_now_iso()
        metadata = self.repository.update_batch(
            batch_id,
            source_filename=output_path.name,
            source_output_path=output_path,
            original_archive_path=output_path,
            working_path=output_path,
            ready_path=output_path,
            status=BatchStatus.READY,
            last_saved_at=timestamp,
            confirmed_at=timestamp,
            row_count=revalidation.summary.total_rows,
            valid_count=revalidation.summary.valid_count,
            warning_count=revalidation.summary.warning_count,
            error_count=revalidation.error_count,
            total_amount=revalidation.summary.total_amount,
            last_error=None,
        )
        self.repository.set_active_batch_id(batch_id)
        LOGGER.info("Đã lưu và xác nhận JSON hiện hành của batch %s", batch_id)
        return BatchReview(metadata, reloaded, revalidation)

    confirm = confirm_batch
    finalize_batch = confirm_batch

    def create_reconciliation_batch(
        self, group_id: int, rows: Iterable[DataRow]
    ) -> BatchReview:
        """Tạo một batch READY bất biến từ kết quả phân bổ cước biển."""

        document = BatchDocument(v=2, rows=list(rows))
        validation = self.validation_service.validate_document(document)
        if validation.has_errors:
            raise BatchValidationError(
                "Batch phân bổ cước biển không hợp lệ.", validation
            )
        existing = self.repository.database.query_one(
            "SELECT id FROM batches WHERE reconciliation_group_id = ? LIMIT 1",
            (group_id,),
        )
        if existing is not None:
            batch_id = int(existing["id"])
            current = self.load_batch(batch_id)
            if document_to_dict(current.document) == document_to_dict(document):
                return current
            if current.metadata.status is not BatchStatus.READY:
                raise BatchServiceError(
                    "Batch phân bổ đã được sử dụng; không thể cập nhật lại tự động."
                )
            target = current.metadata.ready_path or current.metadata.working_path
            self.codec.dump_atomic(target, document, create_backup=True)
            timestamp = local_now_iso()
            metadata = self.repository.update_batch(
                batch_id,
                sha256=self._reconciliation_sha(target, group_id),
                last_saved_at=timestamp,
                confirmed_at=timestamp,
                row_count=validation.summary.total_rows,
                valid_count=validation.summary.valid_count,
                warning_count=validation.summary.warning_count,
                error_count=validation.error_count,
                total_amount=validation.summary.total_amount,
            )
            return BatchReview(metadata, document, validation)
        target = self.paths.ready_dir / f"sea_freight_group_{group_id}.json"
        self.codec.dump_atomic(target, document, create_backup=False)
        digest = self._reconciliation_sha(target, group_id)
        timestamp = local_now_iso()
        metadata = self.repository.create_batch(
            source_filename=target.name,
            source_output_path=None,
            original_archive_path=target,
            working_path=target,
            ready_path=target,
            sha256=digest,
            status=BatchStatus.READY,
            received_at=timestamp,
            source_kind="SEA_FREIGHT_RECONCILIATION",
            reconciliation_group_id=group_id,
        )
        metadata = self.repository.update_batch(
            metadata.id,
            last_saved_at=timestamp,
            confirmed_at=timestamp,
            row_count=validation.summary.total_rows,
            valid_count=validation.summary.valid_count,
            warning_count=validation.summary.warning_count,
            error_count=validation.error_count,
            total_amount=validation.summary.total_amount,
        )
        return BatchReview(metadata, document, validation)

    @staticmethod
    def _reconciliation_sha(path: Path, group_id: int) -> str:
        return hashlib.sha256(
            f"SEA_FREIGHT_RECONCILIATION:{group_id}:{calculate_sha256(path)}".encode("ascii")
        ).hexdigest()

    def reopen_batch(self, batch_id: int) -> BatchReview:
        review = self.load_batch(batch_id)
        if review.metadata.status is BatchStatus.READY:
            metadata = self.repository.update_batch(
                batch_id,
                status=BatchStatus.REVIEWING,
                last_opened_at=local_now_iso(),
            )
            review = BatchReview(metadata, review.document, review.validation)
        if batch_id == self._current_output_batch_id:
            self.repository.set_active_batch_id(batch_id)
        return review

    def restore_active_batch(self) -> BatchReview | None:
        candidate = self.repository.restore_active_batch()
        while candidate is not None:
            try:
                return self.load_batch(candidate.id)
            except BatchDataError:
                LOGGER.exception(
                    "Không thể khôi phục working file batch %s", candidate.id
                )
                self.repository.set_active_batch_id(None)
                candidate = next(
                    (
                        item
                        for item in self.repository.list_recoverable()
                        if item.id != candidate.id and item.working_path.is_file()
                    ),
                    None,
                )
                if candidate is not None:
                    self.repository.set_active_batch_id(candidate.id)
        return None

    restore = restore_active_batch

    def set_active_batch(self, batch_id: int | None) -> None:
        if batch_id is not None:
            self._require_batch(batch_id)
        self.repository.set_active_batch_id(batch_id)

    def get_active_batch(self) -> BatchMetadata | None:
        batch_id = self.repository.get_active_batch_id()
        return self.repository.get_by_id(batch_id) if batch_id is not None else None

    def get_batch(self, batch_id: int) -> BatchMetadata | None:
        return self.repository.get_by_id(batch_id)

    def list_batches(
        self,
        *,
        search: str | None = None,
        status: BatchStatus | str | None = None,
        limit: int | None = None,
    ) -> list[BatchMetadata]:
        return self.repository.list_batches(
            search=search,
            status=status,
            limit=limit,
        )

    def validate_batch(self, batch_id: int) -> ValidationResult:
        return self.load_batch(batch_id).validation

    def close(self) -> None:
        if self._owns_repository:
            self.repository.database.close()

    def _activate_if_appropriate(self, new_batch: BatchMetadata) -> None:
        # File tải mới luôn là phiên bản duy nhất được phép tiếp tục chỉnh sửa.
        self.repository.set_active_batch_id(new_batch.id)

    def _require_batch(self, batch_id: int) -> BatchMetadata:
        metadata = self.repository.get_by_id(batch_id)
        if metadata is None:
            raise BatchNotFoundError(f"Không tìm thấy batch {batch_id}.")
        return metadata

    @staticmethod
    def _filename_timestamp(iso_value: str) -> str:
        try:
            parsed = datetime.fromisoformat(iso_value)
        except ValueError:
            parsed = datetime.now().astimezone()
        return parsed.strftime("%Y%m%d_%H%M%S")

    def _promote_output_file(self, source: Path, received_at: str) -> Path:
        """Chỉ giữ candidate mới và đặt tên theo timestamp tiếp nhận."""

        try:
            in_output = source.resolve().parent == self.paths.output_dir.resolve()
        except OSError:
            in_output = False
        if not in_output:
            return source

        target = self._timestamped_output_path(received_at)
        for candidate in self.paths.output_dir.glob("ket_qua_boc_tach*.json"):
            try:
                same_as_source = candidate.resolve() == source.resolve()
            except OSError:
                same_as_source = candidate == source
            if same_as_source:
                continue
            try:
                candidate.unlink(missing_ok=True)
                LOGGER.info("Đã dọn file Output cũ: %s", candidate)
            except OSError as exc:
                raise BatchServiceError(
                    f"Không thể xóa file Output cũ đang bị khóa: {candidate.name}"
                ) from exc

        if source != target:
            try:
                os.replace(source, target)
            except OSError as exc:
                raise BatchServiceError(
                    f"Không thể đổi tên file Output thành {target.name}."
                ) from exc
        self._notify_output_written(target)
        return target

    def _find_current_output_batch_id(self) -> int | None:
        candidates = [
            path
            for path in self.paths.output_dir.glob("ket_qua_boc_tach*.json")
            if path.is_file()
        ]
        if not candidates:
            return None
        current_path = max(
            candidates,
            key=lambda path: (path.stat().st_mtime_ns, path.name.casefold()),
        )
        for metadata in self.repository.list_batches():
            source_path = metadata.source_output_path
            if (
                source_path is not None
                and self._path_key(source_path) == self._path_key(current_path)
            ):
                return metadata.id
        try:
            sha256 = calculate_sha256(current_path)
        except OSError:
            LOGGER.warning(
                "Không thể nhận diện batch tương ứng với Output hiện hành: %s",
                current_path,
            )
            return None
        metadata = self.repository.get_by_sha256(sha256)
        return metadata.id if metadata is not None else None

    def _is_output_candidate(self, source: Path) -> bool:
        try:
            return source.resolve().parent == self.paths.output_dir.resolve()
        except OSError:
            return False

    def _is_latest_output_candidate(self, source: Path) -> bool:
        try:
            candidates = [
                candidate
                for candidate in self.paths.output_dir.glob(
                    "ket_qua_boc_tach*.json"
                )
                if candidate.is_file()
            ]
            latest = max(
                candidates,
                key=lambda candidate: (
                    candidate.stat().st_mtime_ns,
                    candidate.name.casefold(),
                ),
            )
        except (OSError, ValueError):
            return True
        return self._path_key(latest) == self._path_key(source)

    def _require_current_batch(self, batch_id: int) -> BatchMetadata:
        metadata = self._require_batch(batch_id)
        if self._current_output_batch_id != batch_id:
            raise BatchDataError(
                "Batch này đã bị file tải mới thay thế và không còn dữ liệu để lưu."
            )
        if not metadata.working_path.is_file():
            raise BatchDataError("File JSON hiện hành không còn tồn tại.")
        return metadata

    def _write_current_document(
        self,
        metadata: BatchMetadata,
        document: BatchDocument,
    ) -> tuple[Path, BatchDocument]:
        old_path = metadata.working_path
        timestamp = local_now_iso()
        output_path = self._timestamped_output_path(timestamp)
        try:
            self.codec.dump_atomic(
                output_path,
                document,
                create_backup=False,
                validate=True,
            )
            reloaded = self.codec.load(output_path)
            if self._path_key(old_path) != self._path_key(output_path):
                self._make_writable(old_path)
                old_path.unlink(missing_ok=True)
            self._remove_other_output_json(output_path)
        except (JsonCodecError, OSError) as exc:
            LOGGER.exception("Không thể ghi JSON hiện hành: %s", output_path)
            raise BatchDataError("Không thể lưu dữ liệu JSON.") from exc
        self._notify_output_written(output_path)
        return output_path, reloaded

    def _prepare_received_gpt_document(
        self,
        path: Path,
        document: BatchDocument,
    ) -> BatchDocument:
        """Làm sạch carrier trùng nguồn trước lần review đầu tiên.

        Chỉ dữ liệu vừa nhận từ GPT (chưa từng được user lưu) đi qua đây. Sau
        khi user tự nhập carrier cho CB/CBDH/VTN, bản đã lưu sẽ được giữ nguyên.
        """

        prepared, stripped = strip_received_gpt_carriers(document)
        if not stripped:
            return document
        try:
            self.codec.dump_atomic(
                path,
                prepared,
                create_backup=False,
                validate=True,
            )
            reloaded = self.codec.load(path)
        except (JsonCodecError, OSError) as exc:
            LOGGER.exception("Không thể làm sạch carrier GPT trong %s", path)
            raise BatchDataError(
                "Không thể chuẩn hóa bên vận tải trước khi hiển thị."
            ) from exc
        LOGGER.info(
            "Đã bỏ carrier GPT của %s dòng lấy bên vận tải từ file Hàng ngày.",
            stripped,
        )
        self._notify_output_written(path)
        return reloaded

    def _validate_document(
        self,
        document: BatchDocument,
        *,
        source_kind: str,
    ) -> ValidationResult:
        return self.validation_service.validate_document(
            document,
            allow_negative=source_kind == "BANG_KE",
        )

    def _timestamped_output_path(self, iso_value: str) -> Path:
        timestamp = self._filename_timestamp(iso_value)
        return self.paths.output_dir / f"ket_qua_boc_tach_{timestamp}.json"

    def _remove_other_output_json(self, keep: Path) -> None:
        for candidate in self.paths.output_dir.glob("ket_qua_boc_tach*.json"):
            if self._path_key(candidate) == self._path_key(keep):
                continue
            self._make_writable(candidate)
            candidate.unlink(missing_ok=True)

    def _cleanup_legacy_json_storage(self) -> None:
        for legacy_dir in (
            self.paths.system_dir / "Archive",
            self.paths.workspace_dir,
            self.paths.ready_dir,
            self.paths.rejected_dir,
        ):
            if not legacy_dir.exists():
                continue
            try:
                shutil.rmtree(legacy_dir, onerror=self._remove_readonly)
                LOGGER.info("Đã xóa vùng lưu JSON cũ: %s", legacy_dir)
            except OSError:
                LOGGER.exception("Không thể xóa vùng lưu JSON cũ: %s", legacy_dir)
        for backup in self.paths.output_dir.glob("ket_qua_boc_tach*.json.bak"):
            try:
                self._make_writable(backup)
                backup.unlink(missing_ok=True)
            except OSError:
                LOGGER.exception("Không thể xóa backup JSON cũ: %s", backup)

    @staticmethod
    def _remove_readonly(
        function: Callable[[str], None], path: str, _exc: object
    ) -> None:
        os.chmod(path, stat.S_IWRITE)
        function(path)

    def _notify_output_written(self, path: Path) -> None:
        callback = self._output_write_callback
        if callback is not None:
            try:
                callback(path)
            except Exception:
                LOGGER.exception("Không đánh dấu được file Output do ứng dụng ghi.")

    @staticmethod
    def _path_key(path: str | Path) -> str:
        return os.path.normcase(os.path.abspath(os.fspath(path)))

    @staticmethod
    def _make_writable(path: Path) -> None:
        if path.exists():
            path.chmod(path.stat().st_mode | stat.S_IWUSR)


hash_file = calculate_sha256
