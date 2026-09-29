"""Lưu nhóm SQT chờ nhập quyết toán và lượt PAD gần nhất."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from app.database import Database


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _path_parts(path: str | Path) -> tuple[str, str]:
    resolved = Path(path).expanduser().resolve(strict=False)
    return str(resolved).casefold(), str(resolved)


def _normalized_sqt(values: Iterable[object]) -> tuple[str, ...]:
    result: list[str] = []
    for value in values:
        text = str(value).strip()
        if text and text not in result:
            result.append(text)
    return tuple(result)


@dataclass(frozen=True, slots=True)
class RpaTrackingSnapshot:
    pending_revisions: dict[str, int]
    latest_pad_sqt: tuple[str, ...]

    @property
    def pending_sqt(self) -> tuple[str, ...]:
        return tuple(self.pending_revisions)


class RpaTrackingRepository:
    """Kho trạng thái nhỏ, độc lập với lịch sử chi tiết của các lần chạy Excel."""

    def __init__(self, database: Database | str | Path) -> None:
        self.database = database if isinstance(database, Database) else Database(database)

    def mark_bk_changed(
        self,
        workbook_path: str | Path,
        sheet_name: str,
        sqt_values: Iterable[object],
        *,
        source_run_id: int | None = None,
        updated_at: str | None = None,
    ) -> dict[str, int]:
        workbook_key, display_path = _path_parts(workbook_path)
        values = _normalized_sqt(sqt_values)
        if not values:
            return {}
        timestamp = updated_at or _now_iso()
        with self.database.transaction(immediate=True) as connection:
            for sqt in values:
                connection.execute(
                    """
                    INSERT INTO rpa_pending_sqt (
                        workbook_key, workbook_path, sheet_name, sqt,
                        revision, source_run_id, updated_at
                    ) VALUES (?, ?, ?, ?, 1, ?, ?)
                    ON CONFLICT(workbook_key, sheet_name, sqt) DO UPDATE SET
                        workbook_path = excluded.workbook_path,
                        revision = rpa_pending_sqt.revision + 1,
                        source_run_id = excluded.source_run_id,
                        updated_at = excluded.updated_at
                    """,
                    (
                        workbook_key,
                        display_path,
                        str(sheet_name),
                        sqt,
                        source_run_id,
                        timestamp,
                    ),
                )
            rows = connection.execute(
                """
                SELECT sqt, revision
                FROM rpa_pending_sqt
                WHERE workbook_key = ? AND sheet_name = ?
                  AND sqt IN ({})
                """.format(",".join("?" for _ in values)),
                (workbook_key, str(sheet_name), *values),
            ).fetchall()
        return {str(row["sqt"]): int(row["revision"]) for row in rows}

    def record_latest_pad(
        self,
        workbook_path: str | Path,
        sheet_name: str,
        sqt_values: Iterable[object],
        *,
        run_id: str,
        launched_at: str | None = None,
    ) -> None:
        workbook_key, display_path = _path_parts(workbook_path)
        values = _normalized_sqt(sqt_values)
        timestamp = launched_at or _now_iso()
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                """
                DELETE FROM rpa_latest_pad_items
                WHERE workbook_key = ? AND sheet_name = ?
                """,
                (workbook_key, str(sheet_name)),
            )
            connection.executemany(
                """
                INSERT INTO rpa_latest_pad_items (
                    workbook_key, workbook_path, sheet_name, sqt,
                    run_id, launched_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        workbook_key,
                        display_path,
                        str(sheet_name),
                        sqt,
                        str(run_id),
                        timestamp,
                    )
                    for sqt in values
                ],
            )

    def snapshot(
        self,
        workbook_path: str | Path,
        sheet_name: str,
    ) -> RpaTrackingSnapshot:
        workbook_key, _display_path = _path_parts(workbook_path)
        pending_rows = self.database.query_all(
            """
            SELECT sqt, revision
            FROM rpa_pending_sqt
            WHERE workbook_key = ? AND sheet_name = ?
            ORDER BY updated_at, sqt
            """,
            (workbook_key, str(sheet_name)),
        )
        latest_rows = self.database.query_all(
            """
            SELECT sqt
            FROM rpa_latest_pad_items
            WHERE workbook_key = ? AND sheet_name = ?
            ORDER BY launched_at, sqt
            """,
            (workbook_key, str(sheet_name)),
        )
        return RpaTrackingSnapshot(
            pending_revisions={
                str(row["sqt"]): int(row["revision"]) for row in pending_rows
            },
            latest_pad_sqt=tuple(str(row["sqt"]) for row in latest_rows),
        )

    def mark_pad_succeeded(
        self,
        workbook_path: str | Path,
        sheet_name: str,
        sqt: object,
        *,
        expected_revision: int | None,
    ) -> bool:
        if expected_revision is None:
            return False
        workbook_key, _display_path = _path_parts(workbook_path)
        removed = self.database.execute(
            """
            DELETE FROM rpa_pending_sqt
            WHERE workbook_key = ? AND sheet_name = ?
              AND sqt = ? AND revision = ?
            """,
            (
                workbook_key,
                str(sheet_name),
                str(sqt).strip(),
                int(expected_revision),
            ),
        )
        return removed == 1

    def pending_revision(
        self,
        workbook_path: str | Path,
        sheet_name: str,
        sqt: object,
    ) -> int | None:
        workbook_key, _display_path = _path_parts(workbook_path)
        row = self.database.query_one(
            """
            SELECT revision
            FROM rpa_pending_sqt
            WHERE workbook_key = ? AND sheet_name = ? AND sqt = ?
            """,
            (workbook_key, str(sheet_name), str(sqt).strip()),
        )
        return int(row["revision"]) if row is not None else None


__all__ = ["RpaTrackingRepository", "RpaTrackingSnapshot"]
