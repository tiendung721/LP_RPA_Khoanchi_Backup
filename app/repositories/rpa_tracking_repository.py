"""Lưu lượt ghi BK gần nhất, phiên bản SQT và lượt PAD gần nhất."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Mapping

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
    latest_bk_revisions: dict[str, int]
    latest_pad_sqt: tuple[str, ...]

    @property
    def latest_bk_sqt(self) -> tuple[str, ...]:
        return tuple(self.latest_bk_revisions)


@dataclass(frozen=True, slots=True)
class RpaRowTrackingSnapshot:
    latest_bk_rows: tuple[tuple[str, int], ...]
    latest_pad_rows: tuple[tuple[str, int], ...]


class RpaTrackingRepository:
    """Kho trạng thái nhỏ, độc lập với lịch sử chi tiết của các lần chạy Excel."""

    def __init__(self, database: Database | str | Path) -> None:
        self.database = database if isinstance(database, Database) else Database(database)

    def record_bk_run(
        self,
        workbook_path: str | Path,
        changed_by_sheet: Mapping[str, Iterable[object]],
        *,
        changed_rows: Iterable[tuple[str, int, object]] = (),
        source_run_id: int | None = None,
        updated_at: str | None = None,
    ) -> dict[str, dict[str, int]]:
        workbook_key, display_path = _path_parts(workbook_path)
        timestamp = updated_at or _now_iso()
        normalized = {
            str(sheet): _normalized_sqt(values)
            for sheet, values in changed_by_sheet.items()
            if str(sheet).strip()
        }
        normalized_rows = tuple(
            (str(sheet), int(row), str(sqt).strip())
            for sheet, row, sqt in changed_rows
            if str(sheet).strip() and int(row) > 1 and str(sqt).strip()
        )
        revisions: dict[str, dict[str, int]] = {}
        with self.database.transaction(immediate=True) as connection:
            previous = connection.execute(
                "SELECT generation FROM rpa_latest_bk_run WHERE workbook_key = ?",
                (workbook_key,),
            ).fetchone()
            generation = int(previous["generation"]) + 1 if previous else 1
            connection.execute(
                """
                INSERT INTO rpa_latest_bk_run (
                    workbook_key, workbook_path, generation, source_run_id, updated_at
                ) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(workbook_key) DO UPDATE SET
                    workbook_path = excluded.workbook_path,
                    generation = excluded.generation,
                    source_run_id = excluded.source_run_id,
                    updated_at = excluded.updated_at
                """,
                (workbook_key, display_path, generation, source_run_id, timestamp),
            )
            connection.execute(
                "DELETE FROM rpa_latest_bk_rows WHERE workbook_key = ?",
                (workbook_key,),
            )
            connection.executemany(
                """
                INSERT INTO rpa_latest_bk_rows (
                    workbook_key, sheet_name, row_number, sqt
                ) VALUES (?, ?, ?, ?)
                """,
                [
                    (workbook_key, sheet, row, sqt)
                    for sheet, row, sqt in dict.fromkeys(normalized_rows)
                ],
            )
            for sheet_name, values in normalized.items():
                if not values:
                    continue
                for sqt in values:
                    connection.execute(
                        """
                        INSERT INTO rpa_pending_sqt (
                            workbook_key, workbook_path, sheet_name, sqt,
                            revision, source_run_id, updated_at, latest_generation
                        ) VALUES (?, ?, ?, ?, 1, ?, ?, ?)
                        ON CONFLICT(workbook_key, sheet_name, sqt) DO UPDATE SET
                            workbook_path = excluded.workbook_path,
                            revision = rpa_pending_sqt.revision + 1,
                            source_run_id = excluded.source_run_id,
                            updated_at = excluded.updated_at,
                            latest_generation = excluded.latest_generation
                        """,
                        (
                            workbook_key, display_path, sheet_name, sqt,
                            source_run_id, timestamp, generation,
                        ),
                    )
                rows = connection.execute(
                    """
                    SELECT sqt, revision FROM rpa_pending_sqt
                    WHERE workbook_key = ? AND sheet_name = ?
                      AND latest_generation = ?
                    """,
                    (workbook_key, sheet_name, generation),
                ).fetchall()
                revisions[sheet_name] = {
                    str(row["sqt"]): int(row["revision"]) for row in rows
                }
        return revisions

    def record_latest_pad_rows(
        self,
        workbook_path: str | Path,
        selected_rows: Iterable[tuple[str, int, object]],
        *,
        run_id: str,
        launched_at: str | None = None,
    ) -> None:
        workbook_key, _display_path = _path_parts(workbook_path)
        timestamp = launched_at or _now_iso()
        rows = tuple(
            (str(sheet), int(row), str(sqt).strip())
            for sheet, row, sqt in selected_rows
        )
        with self.database.transaction(immediate=True) as connection:
            connection.execute(
                "DELETE FROM rpa_latest_pad_rows WHERE workbook_key = ?",
                (workbook_key,),
            )
            connection.executemany(
                """
                INSERT INTO rpa_latest_pad_rows (
                    workbook_key, sheet_name, row_number, sqt, run_id, launched_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                [
                    (workbook_key, sheet, row, sqt, str(run_id), timestamp)
                    for sheet, row, sqt in rows
                ],
            )

    def snapshot_rows(self, workbook_path: str | Path) -> RpaRowTrackingSnapshot:
        workbook_key, _display_path = _path_parts(workbook_path)
        bk_rows = self.database.query_all(
            """
            SELECT sheet_name, row_number FROM rpa_latest_bk_rows
            WHERE workbook_key = ? ORDER BY sheet_name, row_number
            """,
            (workbook_key,),
        )
        pad_rows = self.database.query_all(
            """
            SELECT sheet_name, row_number FROM rpa_latest_pad_rows
            WHERE workbook_key = ? ORDER BY sheet_name, row_number
            """,
            (workbook_key,),
        )
        return RpaRowTrackingSnapshot(
            latest_bk_rows=tuple(
                (str(row["sheet_name"]), int(row["row_number"])) for row in bk_rows
            ),
            latest_pad_rows=tuple(
                (str(row["sheet_name"]), int(row["row_number"])) for row in pad_rows
            ),
        )

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
        latest_bk_rows = self.database.query_all(
            """
            SELECT item.sqt, item.revision
            FROM rpa_pending_sqt AS item
            JOIN rpa_latest_bk_run AS run
              ON run.workbook_key = item.workbook_key
             AND run.generation = item.latest_generation
            WHERE item.workbook_key = ? AND item.sheet_name = ?
            ORDER BY item.sqt
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
            latest_bk_revisions={
                str(row["sqt"]): int(row["revision"]) for row in latest_bk_rows
            },
            latest_pad_sqt=tuple(str(row["sqt"]) for row in latest_rows),
        )

    def bk_revision(
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
