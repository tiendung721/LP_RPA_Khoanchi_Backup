from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

from .container_numbers import validate_iso6346
from app.services.excel.daily_sync import SOURCE_HEADER_ALIASES, SYNC_FIELDS, parse_sqt
from app.services.excel.headers import HeaderResolver
from app.services.excel.workbook import WorkbookGateway

from .contracts import BkContainerSnapshot, ContainerRecord, VesselVoyageSuggestion

_NON_ALNUM = re.compile(r"[^A-Z0-9]+")
_INVALID_BK_VALUES = frozenset({"TP", "GND", "CUOCBO", "CUOC BO", "NM", "KBB"})


class SeaFreightMatchError(ValueError):
    pass


def normalize_match_key(value: object) -> str:
    if not isinstance(value, str):
        return ""
    decomposed = unicodedata.normalize("NFKD", value.strip().upper())
    ascii_like = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return _NON_ALNUM.sub("", ascii_like)


def vessel_voyage_text(vessel_name: object, voyage_no: object) -> str:
    vessel = " ".join(str(vessel_name or "").strip().split())
    voyage = " ".join(str(voyage_no or "").strip().split())
    return " ".join(part for part in (vessel, voyage) if part)


def vessel_voyage_keys(
    vessel_name: object,
    voyage_no: object,
) -> tuple[str, str, str]:
    vessel_key = normalize_match_key(vessel_name)
    voyage_key = normalize_match_key(voyage_no)
    if not vessel_key or not voyage_key:
        raise SeaFreightMatchError("Thiếu tên tàu hoặc số chuyến để đối soát BK.")
    if not any(character.isdigit() for character in voyage_key):
        raise SeaFreightMatchError("Số chuyến phải chứa ít nhất một chữ số.")
    combined_key = vessel_key + voyage_key
    if combined_key in _INVALID_BK_VALUES or vessel_key in _INVALID_BK_VALUES:
        raise SeaFreightMatchError("Giá trị tàu/chuyến không hợp lệ.")
    return vessel_key, voyage_key, combined_key


def validate_vessel_voyage(
    raw: object,
    vessel_name: object,
    voyage_no: object,
) -> tuple[str, str, str]:
    """API tương thích; chuỗi AI nguyên văn không còn chi phối khóa đối soát."""

    del raw
    return vessel_voyage_keys(vessel_name, voyage_no)


def _ocr_similarity_key(value: str) -> str:
    return value.translate(str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5"}))


def _suggestion_score(target_key: str, voyage_key: str, candidate_key: str) -> float:
    direct = SequenceMatcher(None, target_key, candidate_key).ratio()
    ocr = SequenceMatcher(
        None,
        _ocr_similarity_key(target_key),
        _ocr_similarity_key(candidate_key),
    ).ratio() * 0.95
    voyage_bonus = 0.15 if candidate_key.endswith(voyage_key) else 0.0
    return min(1.0, max(direct, ocr) + voyage_bonus)


def _date_text(value: object) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


class BkVesselMatcher:
    def __init__(
        self,
        *,
        gateway: WorkbookGateway | None = None,
        headers: HeaderResolver | None = None,
    ) -> None:
        self.gateway = gateway or WorkbookGateway()
        self.headers = headers or HeaderResolver()

    def sheet_names(self, bk_path: str | Path) -> tuple[str, ...]:
        workbook = self.gateway.load(bk_path, read_only=False, data_only=False)
        try:
            return tuple(
                name
                for name in workbook.sheetnames
                if re.fullmatch(r"T\d{2}\s+\d{2}\s*", name, re.IGNORECASE)
            )
        finally:
            workbook.close()

    def snapshot(
        self,
        bk_path: str | Path,
        bk_sheet: str,
        *,
        vessel_voyage_raw: str,
        vessel_name: str,
        voyage_no: str,
    ) -> BkContainerSnapshot:
        path = Path(bk_path).expanduser().resolve()
        vessel_key, voyage_key, combined_key = vessel_voyage_keys(
            vessel_name, voyage_no
        )
        effective_text = vessel_voyage_text(vessel_name, voyage_no)
        fingerprint = self.gateway.fingerprint(path)
        workbook = self.gateway.load(path, read_only=False, data_only=False)
        try:
            if bk_sheet not in workbook.sheetnames:
                raise SeaFreightMatchError(f"Không tìm thấy sheet BK {bk_sheet}.")
            worksheet = workbook[bk_sheet]
            header = self.headers.resolve(
                worksheet,
                SOURCE_HEADER_ALIASES,
                required=SYNC_FIELDS,
            )
            seen: set[str] = set()
            records: list[ContainerRecord] = []
            invalid_count = 0
            duplicate_count = 0
            for row in range(header.row_end + 1, worksheet.max_row + 1):
                raw_vessel = worksheet.cell(row, header.columns["vessel"]).value
                if normalize_match_key(raw_vessel) != combined_key:
                    continue
                raw_container = worksheet.cell(row, header.columns["container"]).value
                try:
                    container = validate_iso6346(
                        str(raw_container) if raw_container not in (None, "") else ""
                    )
                except ValueError:
                    if raw_container not in (None, ""):
                        invalid_count += 1
                    continue
                if container in seen:
                    duplicate_count += 1
                    continue
                seen.add(container)
                records.append(
                    ContainerRecord(
                        container=container,
                        source_sheet=bk_sheet,
                        source_row=row,
                        source_sqt=parse_sqt(
                            worksheet.cell(row, header.columns["sqt"]).value
                        ),
                        departure_date=_date_text(
                            worksheet.cell(row, header.columns["departure_date"]).value
                        ),
                    )
                )
        finally:
            workbook.close()
        fingerprint_text = json.dumps(
            fingerprint.to_dict(), sort_keys=True, separators=(",", ":")
        )
        snapshot_payload = [
            {
                "container": record.container,
                "sheet": record.source_sheet,
                "row": record.source_row,
                "sqt": record.source_sqt,
                "departure_date": record.departure_date,
            }
            for record in records
        ]
        snapshot_hash = hashlib.sha256(
            json.dumps(
                snapshot_payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        ).hexdigest()
        return BkContainerSnapshot(
            bk_path=str(path),
            bk_sheet=bk_sheet,
            vessel_voyage_raw=effective_text,
            vessel_key=vessel_key,
            voyage_key=voyage_key,
            combined_key=combined_key,
            workbook_fingerprint=fingerprint_text,
            snapshot_hash=snapshot_hash,
            containers=tuple(records),
            invalid_container_count=invalid_count,
            duplicate_container_count=duplicate_count,
        )

    def suggestions(
        self,
        bk_path: str | Path,
        bk_sheet: str,
        *,
        vessel_name: str,
        voyage_no: str,
        limit: int = 5,
    ) -> tuple[VesselVoyageSuggestion, ...]:
        """Xếp hạng gợi ý từ đúng sheet BK; không dùng để tự động match."""

        path = Path(bk_path).expanduser().resolve()
        _, voyage_key, target_key = vessel_voyage_keys(vessel_name, voyage_no)
        workbook = self.gateway.load(path, read_only=False, data_only=False)
        try:
            if bk_sheet not in workbook.sheetnames:
                return ()
            worksheet = workbook[bk_sheet]
            header = self.headers.resolve(
                worksheet,
                SOURCE_HEADER_ALIASES,
                required=SYNC_FIELDS,
            )
            displays: dict[str, str] = {}
            containers: dict[str, set[str]] = {}
            for row in range(header.row_end + 1, worksheet.max_row + 1):
                raw_vessel = worksheet.cell(row, header.columns["vessel"]).value
                candidate_key = normalize_match_key(raw_vessel)
                if not candidate_key:
                    continue
                displays.setdefault(candidate_key, " ".join(str(raw_vessel).strip().split()))
                containers.setdefault(candidate_key, set())
                raw_container = worksheet.cell(row, header.columns["container"]).value
                try:
                    container = validate_iso6346(
                        str(raw_container) if raw_container not in (None, "") else ""
                    )
                except ValueError:
                    continue
                containers[candidate_key].add(container)
        finally:
            workbook.close()

        ranked = [
            VesselVoyageSuggestion(
                vessel_voyage=display,
                container_count=len(containers[candidate_key]),
                score=_suggestion_score(target_key, voyage_key, candidate_key),
            )
            for candidate_key, display in displays.items()
        ]
        ranked.sort(key=lambda item: (-item.score, item.vessel_voyage.casefold()))
        return tuple(item for item in ranked if item.score >= 0.45)[: max(0, limit)]
