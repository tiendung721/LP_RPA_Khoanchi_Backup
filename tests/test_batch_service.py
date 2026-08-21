from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pytest

from app.config import AppSettings
from app.models import BatchDocument, BatchStatus, DataRow
from app.services.batch_service import (
    BatchService,
    BatchServiceError,
    BatchValidationError,
)
from app.services.reviewed_batch_provider import ReviewedBatchProvider

FIXTURE = Path(__file__).parent / "fixtures" / "ket_qua_boc_tach.json"


def _settings(tmp_path: Path) -> AppSettings:
    return AppSettings(
        data_root=tmp_path / "runtime",
        output_dir=tmp_path / "Output",
    )


def _write_json(path: Path, rows: list[list[object]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    normalized_rows = [
        [*row[:4], None, None, row[4]] if len(row) == 5 else row
        for row in rows
    ]
    path.write_text(
        json.dumps(
            {"v": 1, "d": normalized_rows},
            ensure_ascii=False,
            separators=(",", ":"),
        ),
        encoding="utf-8",
    )
    return path


def _valid_rows() -> list[list[object]]:
    return [
        ["DRYU3026167", None, "VTN", "CV", None, None, 13_554_000],
        [None, "BL123456789", "CB", "HD", None, None, 27_500_000],
    ]


def _current_json(settings: AppSettings) -> Path:
    files = list(settings.output_dir.glob("ket_qua_boc_tach*.json"))
    assert len(files) == 1
    return files[0]


def _assert_timestamped(path: Path) -> None:
    assert re.fullmatch(
        r"ket_qua_boc_tach_\d{8}_\d{6}\.json",
        path.name,
    )


def test_receive_keeps_one_timestamped_json_without_legacy_copies(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(settings.output_dir / "ket_qua_boc_tach.json", _valid_rows())
    original_payload = source.read_bytes()
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.created
    assert result.review is not None
    current = _current_json(settings)
    _assert_timestamped(current)
    assert current.read_bytes() == original_payload
    assert not source.exists()
    assert result.batch.original_archive_path == current
    assert result.batch.working_path == current
    assert result.batch.row_count == 2
    assert result.batch.last_saved_at is None
    assert result.batch.source_output_path == current
    assert not settings.paths.archive_original_dir.exists()
    assert not settings.paths.workspace_dir.exists()
    assert not settings.paths.ready_dir.exists()
    assert not settings.paths.rejected_dir.exists()
    service.close()


def test_bang_ke_filename_sets_source_kind_before_timestamp_rename(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach_bang_ke_ha_vo.json",
        [["DRYU3045911", "VS26020793", "VSDL", "GV", None, None, 282_000]],
    )
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.batch.source_kind == "BANG_KE"
    _assert_timestamped(result.batch.working_path)
    service.close()


def test_receive_real_fixture_records_all_47_rows(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = settings.output_dir / "ket_qua_boc_tach.json"
    source.parent.mkdir(parents=True)
    source.write_bytes(FIXTURE.read_bytes())
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.review is not None
    assert result.batch.row_count == 47
    assert len(result.review.document.rows) == 47
    service.close()


def test_receive_strips_only_gpt_carriers_managed_by_daily_sync(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach.json",
        [
            ["DRYU3026167", None, "VTN", "CV", "HD-VTN", "GPT VTN DÀI", 100],
            ["GAOU2112422", None, "CBDH", "CV", "HD-CBDH", "GPT BỘ DÀI", 200],
            [None, "BL123456789", "CB", "HD", "HD-CB", "GPT BIỂN DÀI", 300],
            ["VSGU2250713", None, "NH", "CV", "HD-PHI", "NHS", 400],
        ],
    )
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.review is not None
    assert [row.carrier for row in result.review.document.rows] == [
        None,
        None,
        None,
        "NHS",
    ]
    persisted = json.loads(_current_json(settings).read_text(encoding="utf-8"))
    assert [row["carrier"] for row in persisted["d"]] == [
        None,
        None,
        None,
        "NHS",
    ]
    service.close()


def test_receive_removes_legacy_vat_rows_and_persists_remaining_rows(
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach_bang_ke_old.json",
        [
            ["DRYU3026167", None, "VSDL", "GV", None, None, 200_000],
            ["DRYU3026167", None, " vat ", "GV", None, None, 80_000],
            ["GAOU2112422", None, "GH", "GV", None, None, 150_000],
        ],
    )
    service = BatchService(settings)

    with caplog.at_level("WARNING"):
        result = service.receive_file(source)

    assert result.review is not None
    assert [row.fee for row in result.review.document.rows] == ["VSDL", "GH"]
    assert result.batch.row_count == 2
    assert result.batch.total_amount == 350_000
    persisted = json.loads(_current_json(settings).read_text(encoding="utf-8"))
    assert [row["fee"] for row in persisted["d"]] == ["VSDL", "GH"]
    assert "Đã loại 1 dòng fee=VAT" in caplog.text
    service.close()


def test_receive_only_legacy_vat_rows_creates_empty_valid_batch(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach_bang_ke_old.json",
        [["DRYU3026167", None, "VAT", "GV", None, None, 80_000]],
    )
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.review is not None
    assert result.review.document.rows == []
    assert result.batch.row_count == 0
    assert result.batch.error_count == 0
    persisted = json.loads(_current_json(settings).read_text(encoding="utf-8"))
    assert persisted == {"v": 3, "d": []}
    service.close()


def test_bang_ke_prompt_explicitly_ignores_vat_fee() -> None:
    prompt = (
        Path(__file__).parents[1] / "gpt_custom_instructions_bang_ke.txt"
    ).read_text(encoding="utf-8")

    assert 'Không xử lý cột VAT hoặc THUẾ GTGT' in prompt
    assert '- VAT: amount =' not in prompt
    assert 'VSDL, QT, VAT, GH' not in prompt


def test_manual_carrier_for_daily_sync_fee_survives_save_and_reload(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach.json",
        [["DRYU3026167", None, "VTN", "CV", "HD-1", "GPT DÀI", 100]],
    )
    service = BatchService(settings)
    received = service.receive_file(source)
    received.review.document.rows[0].carrier = "USER CHỌN"

    saved = service.save_working(received.batch.id, received.review.document)
    reloaded = service.load_batch(received.batch.id)

    assert saved.document.rows[0].carrier == "USER CHỌN"
    assert reloaded.document.rows[0].carrier == "USER CHỌN"
    service.close()


def test_startup_removes_legacy_json_storage_but_preserves_ready_results(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    legacy_files = (
        settings.paths.archive_original_dir / "original.json",
        settings.paths.workspace_dir / "1" / "working.json",
        settings.paths.workspace_dir / "1" / "working.json.bak",
        settings.paths.rejected_dir / "bad.json",
    )
    for path in legacy_files:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("legacy", encoding="utf-8")
    excel_backup = settings.paths.excel_backup_dir / "BK_backup.xlsx"
    excel_backup.parent.mkdir(parents=True, exist_ok=True)
    excel_backup.write_bytes(b"excel")
    ready_result = settings.paths.ready_dir / "sea_freight_group_7.json"
    ready_result.parent.mkdir(parents=True, exist_ok=True)
    ready_result.write_text('{"v":3,"d":[]}', encoding="utf-8")

    service = BatchService(settings)

    assert not (settings.paths.system_dir / "Archive").exists()
    assert not settings.paths.workspace_dir.exists()
    assert ready_result.read_text(encoding="utf-8") == '{"v":3,"d":[]}'
    assert not settings.paths.rejected_dir.exists()
    assert excel_backup.read_bytes() == b"excel"
    service.close()


def test_same_hash_new_download_replaces_current_and_is_processed_again(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    first = _write_json(settings.output_dir / "ket_qua_boc_tach.json", _valid_rows())
    service = BatchService(settings)

    created = service.receive_file(first)
    second = _write_json(
        settings.output_dir / "ket_qua_boc_tach (1).json", _valid_rows()
    )
    replaced = service.receive_file(second)

    assert replaced.created
    assert not replaced.duplicate
    assert replaced.review is not None
    assert replaced.batch.id == created.batch.id
    assert len(service.list_batches()) == 1
    _assert_timestamped(_current_json(settings))
    service.close()


def test_invalid_root_is_recorded_without_rejected_copy(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = settings.output_dir / "ket_qua_boc_tach.json"
    source.parent.mkdir(parents=True)
    source.write_text('{"v":1,"rows":[]}', encoding="utf-8")
    service = BatchService(settings)

    result = service.receive_file(source)

    assert result.batch.status is BatchStatus.INVALID
    assert result.batch.last_error
    assert result.batch.source_output_path == _current_json(settings)
    assert not settings.paths.rejected_dir.exists()
    service.close()


def test_invalid_new_download_still_replaces_old_output(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    service = BatchService(settings)
    old = _write_json(
        settings.output_dir / "ket_qua_boc_tach.json", _valid_rows()
    )
    service.receive_file(old)
    incoming = settings.output_dir / "ket_qua_boc_tach (1).json"
    incoming.write_text('{"v":1,"rows":[]}', encoding="utf-8")

    result = service.receive_file(incoming)
    current = _current_json(settings)

    assert result.batch.status is BatchStatus.INVALID
    assert result.batch.last_saved_at is None
    assert current.read_text(encoding="utf-8") == '{"v":1,"rows":[]}'
    assert not incoming.exists()
    service.close()


def test_current_invalid_batch_recovers_after_file_matches_current_schema(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = settings.output_dir / "ket_qua_boc_tach.json"
    source.parent.mkdir(parents=True)
    source.write_text(
        '{"v":1,"d":[["DRYU3026167",null,"VTN","CV",100]]}',
        encoding="utf-8",
    )
    service = BatchService(settings)
    invalid = service.receive_file(source)
    assert invalid.batch.status is BatchStatus.INVALID

    current = _current_json(settings)
    current.write_text(
        '{"v":1,"d":[["DRYU3026167",null,"VTN","CV","HD-130",'
        '"Vận tải ABC",100]]}',
        encoding="utf-8",
    )

    recovered = service.get_current_output_batch()

    assert recovered is not None
    assert recovered.status is BatchStatus.REVIEWING
    assert recovered.last_error is None
    assert recovered.row_count == 1
    service.close()


def test_saving_superseded_batch_does_not_overwrite_new_output(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    service = BatchService(settings)
    first = service.receive_file(
        _write_json(
            settings.output_dir / "ket_qua_boc_tach.json",
            [["DRYU3026167", None, "VTN", "CV", 1]],
        )
    )
    second = service.receive_file(
        _write_json(
            settings.output_dir / "ket_qua_boc_tach (1).json",
            [["GAOU2112422", None, "VTN", "CV", 2]],
        )
    )
    with pytest.raises(BatchServiceError, match="đã bị file tải mới thay thế"):
        service.save_working(
            first.batch.id,
            BatchDocument(rows=[DataRow("DRYU3026167", None, "VTN", "CV", 99)]),
        )

    output = json.loads(
        _current_json(settings).read_text(encoding="utf-8")
    )
    assert second.batch.id != first.batch.id
    assert output["d"][0][-1] == 2
    _assert_timestamped(_current_json(settings))
    service.close()


def test_reapplying_same_output_keeps_review_sync_enabled(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    service = BatchService(settings)
    received = service.receive_file(
        _write_json(
            settings.output_dir / "ket_qua_boc_tach.json",
            [["DRYU3026167", None, "VTN", "CV", 1]],
        )
    )

    service.update_paths(settings.paths)
    service.save_working(
        received.batch.id,
        BatchDocument(
            rows=[DataRow("DRYU3026167", None, "VTN", "CV", 88)]
        ),
    )

    output = json.loads(
        _current_json(settings).read_text(encoding="utf-8")
    )
    assert output["v"] == 3
    assert output["d"][0]["amount"] == 88
    service.close()


def test_changing_to_empty_output_clears_current_batch(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    service = BatchService(settings)
    received = service.receive_file(
        _write_json(
            settings.output_dir / "ket_qua_boc_tach.json",
            [["DRYU3026167", None, "VTN", "CV", 1]],
        )
    )
    assert service.current_output_batch_id == received.batch.id

    empty_output = tmp_path / "Output mới"
    service.update_paths(
        AppSettings(
            data_root=settings.data_root,
            output_dir=empty_output,
        ).paths
    )

    assert service.current_output_batch_id is None
    assert service.get_current_output_batch() is None
    assert service.get_active_batch() is None
    service.close()


def test_restart_identifies_current_output_before_watcher_scan(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    first_service = BatchService(settings)
    received = first_service.receive_file(
        _write_json(
            settings.output_dir / "ket_qua_boc_tach.json",
            [["DRYU3026167", None, "VTN", "CV", 1]],
        )
    )
    first_service.close()

    restarted = BatchService(settings)

    assert restarted.current_output_batch_id == received.batch.id
    current = restarted.get_current_output_batch()
    assert current is not None
    assert current.id == received.batch.id
    restarted.save_working(
        received.batch.id,
        BatchDocument(
            rows=[DataRow("DRYU3026167", None, "VTN", "CV", 77)]
        ),
    )
    output = json.loads(
        _current_json(settings).read_text(encoding="utf-8")
    )
    assert output["v"] == 3
    assert output["d"][0]["amount"] == 77
    restarted.close()


def test_older_candidate_is_rejected_when_newer_download_exists(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    older = _write_json(
        settings.output_dir / "ket_qua_boc_tach.json",
        [["DRYU3026167", None, "VTN", "CV", 1]],
    )
    newer = _write_json(
        settings.output_dir / "ket_qua_boc_tach (1).json",
        [["GAOU2112422", None, "VTN", "CV", 2]],
    )
    older_mtime = newer.stat().st_mtime_ns - 1_000_000
    os.utime(older, ns=(older_mtime, older_mtime))
    service = BatchService(settings)

    with pytest.raises(BatchServiceError, match="file cũ"):
        service.receive_file(older)

    assert newer.is_file()
    assert service.list_batches() == []
    service.close()


def test_save_and_confirmation_reuse_single_timestamped_json(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(settings.output_dir / "ket_qua_boc_tach.json", _valid_rows())
    service = BatchService(settings)
    received = service.receive_file(source)
    edited = BatchDocument(
        rows=[
            DataRow(None, "BL123456789", "CB", "HD", 27_500_000),
            DataRow("DRYU3026167", None, "VTN", "CV", 13_554_001),
        ]
    )

    saved = service.save_working(received.batch.id, edited)
    confirmed = service.confirm_batch(received.batch.id, saved.document)
    provider = ReviewedBatchProvider(service)

    assert [row.to_list() for row in saved.document.rows] == [
        row.to_list() for row in edited.rows
    ]
    current = _current_json(settings)
    _assert_timestamped(current)
    assert not list(settings.output_dir.glob("*.bak"))
    assert json.loads(current.read_text(encoding="utf-8")) == saved.document.to_dict()
    assert confirmed.metadata.status is BatchStatus.READY
    assert confirmed.metadata.source_output_path == current
    assert confirmed.metadata.ready_path == current
    assert provider.get_latest_ready_json_path() == current
    assert provider.get_ready_json_path(received.batch.id) == current
    assert [item.id for item in provider.list_ready_batches()] == [received.batch.id]
    assert not settings.paths.ready_dir.exists()
    service.close()


def test_blocking_validation_prevents_confirmation(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = _write_json(
        settings.output_dir / "ket_qua_boc_tach.json",
        [["DRYU3026167", None, "CB", "CV", 1]],
    )
    service = BatchService(settings)
    received = service.receive_file(source)

    with pytest.raises(BatchValidationError):
        service.confirm_batch(received.batch.id)

    assert not settings.paths.ready_dir.exists()
    service.close()


def test_ready_batch_can_be_reopened_and_saved_without_snapshot(
    tmp_path: Path,
) -> None:
    settings = _settings(tmp_path)
    source = _write_json(settings.output_dir / "ket_qua_boc_tach.json", _valid_rows())
    service = BatchService(settings)
    received = service.receive_file(source)
    first = service.confirm_batch(received.batch.id)

    reopened = service.reopen_batch(received.batch.id)
    reopened.document.rows[0].amount = 13_554_001
    second = service.confirm_batch(received.batch.id, reopened.document)

    assert reopened.metadata.status is BatchStatus.REVIEWING
    assert second.metadata.status is BatchStatus.READY
    assert second.metadata.source_output_path == _current_json(settings)
    assert len(list(settings.output_dir.glob("ket_qua_boc_tach*.json"))) == 1
    assert not settings.paths.ready_dir.exists()
    service.close()


def test_restart_restores_active_reviewing_batch(tmp_path: Path) -> None:
    settings = _settings(tmp_path)
    source = _write_json(settings.output_dir / "ket_qua_boc_tach.json", _valid_rows())
    first_service = BatchService(settings)
    received = first_service.receive_file(source)
    first_service.load_batch(received.batch.id)
    first_service.close()

    restarted = BatchService(settings)
    restored = restarted.restore_active_batch()

    assert restored is not None
    assert restored.metadata.id == received.batch.id
    assert restored.metadata.status is BatchStatus.REVIEWING
    restarted.close()
