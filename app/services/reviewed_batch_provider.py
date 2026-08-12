"""Cung cấp file JSON hiện hành đã được người dùng xác nhận."""

from __future__ import annotations

import logging
from pathlib import Path

from app.database import Database
from app.models import BatchMetadata, BatchStatus
from app.repositories.batch_repository import BatchRepository
from app.services.batch_service import BatchService

LOGGER = logging.getLogger(__name__)


class ReviewedBatchProvider:
    def __init__(
        self,
        source: BatchRepository | BatchService | Database | str | Path,
        *,
        ready_root: str | Path | None = None,
    ) -> None:
        if isinstance(source, BatchService):
            self.repository = source.repository
            inferred_root = source.paths.output_dir
            self._owns_repository = False
        elif isinstance(source, BatchRepository):
            self.repository = source
            inferred_root = None
            self._owns_repository = False
        elif isinstance(source, Database):
            self.repository = BatchRepository(source)
            inferred_root = None
            self._owns_repository = False
        else:
            self.repository = BatchRepository(Database(source))
            inferred_root = None
            self._owns_repository = True
        self.ready_root = (
            Path(ready_root) if ready_root is not None else inferred_root
        )

    def get_latest_ready_json_path(self) -> Path | None:
        for metadata in self.repository.list_ready_batches():
            if metadata.source_kind == "SEA_FREIGHT_RECONCILIATION":
                continue
            path = self._usable_path(metadata)
            if path is not None:
                return path
        return None

    def get_ready_json_path(self, batch_id: int) -> Path | None:
        metadata = self.repository.get_by_id(batch_id)
        if metadata is None or metadata.status is not BatchStatus.READY:
            return None
        return self._usable_path(metadata)

    def list_ready_batches(self) -> list[BatchMetadata]:
        return [
            metadata
            for metadata in self.repository.list_ready_batches()
            if metadata.source_kind != "SEA_FREIGHT_RECONCILIATION"
            and self._usable_path(metadata) is not None
        ]

    def close(self) -> None:
        if self._owns_repository:
            self.repository.database.close()

    def _usable_path(self, metadata: BatchMetadata) -> Path | None:
        path = metadata.ready_path
        if path is None or not path.is_file():
            if path is not None:
                LOGGER.warning(
                    "Bỏ qua JSON hiện hành bị thiếu của batch %s: %s",
                    metadata.id,
                    path,
                )
            return None
        if self.ready_root is not None:
            try:
                path.resolve().relative_to(self.ready_root.resolve())
            except (OSError, ValueError):
                LOGGER.error(
                    "Từ chối JSON nằm ngoài thư mục Output: %s",
                    path,
                )
                return None
        return path
