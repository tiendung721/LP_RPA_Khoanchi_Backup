"""CLI để PAD đánh dấu một SQT đã nhập thành công lên web."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Cho phép PAD gọi trực tiếp script bằng đường dẫn tuyệt đối, không phụ thuộc
# thư mục làm việc hiện tại của tiến trình.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.rpa_expense.status import (
    RpaExpenseStatusError,
    RpaExpenseStatusService,
)
from app.database import Database
from app.repositories.rpa_tracking_repository import RpaTrackingRepository


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="rpa_excel_helper.py")
    subparsers = parser.add_subparsers(dest="command", required=True)
    mark = subparsers.add_parser("mark-imported")
    mark.add_argument("--selection", required=True)
    mark.add_argument("--sqt", required=True)
    mark.add_argument("--backup-dir")
    return parser


def _write(stream: object, payload: dict[str, object]) -> None:
    data = json.dumps(
        payload,
        ensure_ascii=False,
        separators=(",", ":"),
    )
    target = getattr(stream, "buffer", stream)
    target.write((data + "\n").encode("utf-8"))
    target.flush()


def main(argv: list[str] | None = None) -> int:
    database: Database | None = None
    try:
        args = build_parser().parse_args(argv)
        if args.command != "mark-imported":
            raise RpaExpenseStatusError(
                f"Command không được hỗ trợ: {args.command}"
            )
        selection_path = Path(args.selection)
        selection = json.loads(selection_path.read_text(encoding="utf-8-sig"))
        project_root = Path(str(selection.get("project_root") or PROJECT_ROOT))
        database = Database(project_root / "Database" / "app_state.db")
        service = RpaExpenseStatusService(
            backup_dir=Path(args.backup_dir) if args.backup_dir else None,
            tracking_repository=RpaTrackingRepository(database),
        )
        result = service.mark_imported(args.selection, args.sqt)
        _write(sys.stdout, result)
        return 0
    except RpaExpenseStatusError as exc:
        _write(
            sys.stderr,
            {
                "success": False,
                "error_code": "RPA_STATUS_ERROR",
                "message": str(exc),
            },
        )
        return 2
    except Exception as exc:  # PAD cần mã lỗi khác 0 cho lỗi không xác định.
        _write(
            sys.stderr,
            {
                "success": False,
                "error_code": "UNKNOWN_ERROR",
                "message": str(exc),
            },
        )
        return 1
    finally:
        if database is not None:
            database.close()


if __name__ == "__main__":
    raise SystemExit(main())
