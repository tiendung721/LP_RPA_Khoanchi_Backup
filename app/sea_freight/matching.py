from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

from .container_numbers import validate_iso6346
from app.services.excel.daily_sync import SOURCE_HEADER_ALIASES, SYNC_FIELDS, parse_sqt
from app.services.excel.headers import HeaderResolver
from app.services.excel.workbook import WorkbookGateway

from .contracts import (
    BkContainerSnapshot,
    ContainerRecord,
    VesselVoyageResolution,
    VesselVoyageResolutionKind,
    VesselVoyageSuggestion,
)

_NON_ALNUM = re.compile(r"[^A-Z0-9]+")
_VOYAGE_KEY = re.compile(r"^(?P<prefix>[A-Z]*)(?P<number>\d+)(?P<suffix>[A-Z]*)$")
_INVALID_BK_VALUES = frozenset({"TP", "GND", "CUOCBO", "CUOC BO", "NM", "KBB"})


@dataclass(frozen=True, slots=True)
class _VoyageIdentity:
    prefix: str
    number: str
    suffix: str


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


def _voyage_identity(value: object) -> _VoyageIdentity | None:
    key = normalize_match_key(value)
    match = _VOYAGE_KEY.fullmatch(key)
    if match is None:
        return None
    return _VoyageIdentity(
        prefix=match.group("prefix"),
        number=match.group("number"),
        suffix=match.group("suffix"),
    )


def vessel_voyage_alias_equivalent(
    left_vessel_name: object,
    left_voyage_no: object,
    right_vessel_name: object,
    right_voyage_no: object,
) -> bool:
    """Khớp exact hoặc chỉ khác việc một bên thiếu tiền tố chuyến."""

    if normalize_match_key(left_vessel_name) != normalize_match_key(right_vessel_name):
        return False
    left = _voyage_identity(left_voyage_no)
    right = _voyage_identity(right_voyage_no)
    if left is None or right is None:
        return False
    if (left.number, left.suffix) != (right.number, right.suffix):
        return False
    return left.prefix == right.prefix or not left.prefix or not right.prefix


def _split_bk_vessel_voyage(
    raw_value: object,
    expected_vessel_key: str,
) -> tuple[str, str, str] | None:
    if not isinstance(raw_value, str):
        return None
    display = " ".join(raw_value.strip().split())
    combined_key = normalize_match_key(display)
    if not expected_vessel_key or not combined_key.startswith(expected_vessel_key):
        return None
    voyage_key = combined_key[len(expected_vessel_key) :]
    if _voyage_identity(voyage_key) is None:
        return None

    separators = " \t\r\n/-–—|:"
    for index in range(1, len(display)):
        vessel = display[:index].rstrip(separators)
        voyage = display[index:].lstrip(separators)
        if (
            normalize_match_key(vessel) == expected_vessel_key
            and normalize_match_key(voyage) == voyage_key
        ):
            return vessel, voyage, combined_key
    return None


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

    def resolve(
        self,
        bk_path: str | Path,
        bk_sheet: str,
        *,
        vessel_name: str,
        voyage_no: str,
    ) -> VesselVoyageResolution:
        """Tìm identity canonical trong đúng sheet, ưu tiên exact tuyệt đối."""

        path = Path(bk_path).expanduser().resolve()
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
            return self._resolve_worksheet(
                worksheet,
                header,
                vessel_name=vessel_name,
                voyage_no=voyage_no,
            )
        finally:
            workbook.close()

    @staticmethod
    def _resolve_worksheet(
        worksheet: object,
        header: object,
        *,
        vessel_name: str,
        voyage_no: str,
    ) -> VesselVoyageResolution:
        vessel_key, voyage_key, combined_key = vessel_voyage_keys(
            vessel_name, voyage_no
        )
        displays: dict[str, str] = {}
        parsed: dict[str, tuple[str, str]] = {}
        containers: dict[str, set[str]] = {}
        for row in range(header.row_end + 1, worksheet.max_row + 1):
            raw_vessel = worksheet.cell(row, header.columns["vessel"]).value
            candidate_key = normalize_match_key(raw_vessel)
            if not candidate_key:
                continue
            displays.setdefault(
                candidate_key, " ".join(str(raw_vessel).strip().split())
            )
            containers.setdefault(candidate_key, set())
            split = _split_bk_vessel_voyage(raw_vessel, vessel_key)
            if split is not None:
                parsed.setdefault(candidate_key, (split[0], split[1]))
            raw_container = worksheet.cell(row, header.columns["container"]).value
            try:
                container = validate_iso6346(
                    str(raw_container) if raw_container not in (None, "") else ""
                )
            except ValueError:
                continue
            containers[candidate_key].add(container)

        if combined_key in displays:
            canonical_vessel, canonical_voyage = parsed.get(
                combined_key, (str(vessel_name).strip(), str(voyage_no).strip())
            )
            return VesselVoyageResolution(
                kind=VesselVoyageResolutionKind.EXACT,
                input_vessel_name=str(vessel_name),
                input_voyage_no=str(voyage_no),
                canonical_vessel_name=canonical_vessel,
                canonical_voyage_no=canonical_voyage,
                canonical_vessel_voyage=displays[combined_key],
                vessel_key=normalize_match_key(canonical_vessel),
                voyage_key=normalize_match_key(canonical_voyage),
                combined_key=combined_key,
            )

        alias_candidates: list[VesselVoyageSuggestion] = []
        for candidate_key, (candidate_vessel, candidate_voyage) in parsed.items():
            if not containers.get(candidate_key):
                continue
            if not vessel_voyage_alias_equivalent(
                vessel_name,
                voyage_no,
                candidate_vessel,
                candidate_voyage,
            ):
                continue
            candidate_vessel_key = normalize_match_key(candidate_vessel)
            candidate_voyage_key = normalize_match_key(candidate_voyage)
            alias_candidates.append(
                VesselVoyageSuggestion(
                    vessel_voyage=displays[candidate_key],
                    container_count=len(containers[candidate_key]),
                    score=_suggestion_score(
                        combined_key, voyage_key, candidate_key
                    ),
                    vessel_name=candidate_vessel,
                    voyage_no=candidate_voyage,
                    vessel_key=candidate_vessel_key,
                    voyage_key=candidate_voyage_key,
                    combined_key=candidate_key,
                    selectable=True,
                )
            )
        alias_candidates.sort(key=lambda item: item.vessel_voyage.casefold())

        if len(alias_candidates) == 1:
            selected = alias_candidates[0]
            return VesselVoyageResolution(
                kind=VesselVoyageResolutionKind.AUTO_ALIAS,
                input_vessel_name=str(vessel_name),
                input_voyage_no=str(voyage_no),
                canonical_vessel_name=selected.vessel_name,
                canonical_voyage_no=selected.voyage_no,
                canonical_vessel_voyage=selected.vessel_voyage,
                vessel_key=selected.vessel_key,
                voyage_key=selected.voyage_key,
                combined_key=selected.combined_key,
                candidates=(selected,),
            )
        if len(alias_candidates) > 1:
            return VesselVoyageResolution(
                kind=VesselVoyageResolutionKind.AMBIGUOUS,
                input_vessel_name=str(vessel_name),
                input_voyage_no=str(voyage_no),
                vessel_key=vessel_key,
                voyage_key=voyage_key,
                combined_key=combined_key,
                candidates=tuple(alias_candidates),
            )
        return VesselVoyageResolution(
            kind=VesselVoyageResolutionKind.NOT_FOUND,
            input_vessel_name=str(vessel_name),
            input_voyage_no=str(voyage_no),
            vessel_key=vessel_key,
            voyage_key=voyage_key,
            combined_key=combined_key,
        )

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
        vessel_key, voyage_key, combined_key = vessel_voyage_keys(vessel_name, voyage_no)
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
            resolution = self._resolve_worksheet(
                worksheet,
                header,
                vessel_name=vessel_name,
                voyage_no=voyage_no,
            )
            if resolution.kind in {
                VesselVoyageResolutionKind.EXACT,
                VesselVoyageResolutionKind.AUTO_ALIAS,
            }:
                vessel_key = resolution.vessel_key
                voyage_key = resolution.voyage_key
                combined_key = resolution.combined_key
                effective_text = resolution.canonical_vessel_voyage
            seen: set[str] = set()
            records: list[ContainerRecord] = []
            invalid_count = 0
            duplicate_count = 0
            for row in range(header.row_end + 1, worksheet.max_row + 1):
                raw_vessel = worksheet.cell(row, header.columns["vessel"]).value
                if (
                    resolution.kind
                    not in {
                        VesselVoyageResolutionKind.EXACT,
                        VesselVoyageResolutionKind.AUTO_ALIAS,
                    }
                    or normalize_match_key(raw_vessel) != combined_key
                ):
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
            resolution_kind=resolution.kind,
            canonical_vessel_name=resolution.canonical_vessel_name,
            canonical_voyage_no=resolution.canonical_voyage_no,
            alias_candidates=resolution.candidates,
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
