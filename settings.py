"""Load project settings without supplying inferred locality or school-year data."""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


DEFAULT_SOURCE_COLUMNS = {
    "sequence": "A",
    "family_name": "B",
    "given_name": "C",
    "birth_day": "D",
    "birth_month": "E",
    "birth_year": "F",
    "female": "G",
    "ethnicity": "H",
    "religion": "I",
    "priority_category": "J",
    "head_family_name": "K",
    "head_given_name": "L",
    "address": "M",
    "household_code": "N",
    "residence": "O",
    "residence_status": "P",
    "grade_group": "Q",
    "class_name": "R",
    "school_district": "S",
    "school_name": "T",
    "school_code": "U",
    "graduation_level": "V",
    "supplementary_education": "W",
    "graduation_year": "X",
    "vocational_graduation_level": "Y",
    "diploma_number": "Z",
    "vocational_graduation_year": "AA",
    "completed_education": "AB",
    "completed_education_year": "AC",
    "dropped_out_class": "AD",
    "dropped_out_year": "AE",
    "currently_studying_class": "AF",
    "completed_class": "AG",
    "reliteracy_level": "AH",
    "disability_mobility": "AI",
    "disability_hearing_speech": "AJ",
    "disability_vision": "AK",
    "disability_neurological": "AL",
    "disability_intellectual": "AM",
    "disability_learning": "AN",
    "disability_autism": "AO",
    "disability_other": "AP",
    "disability_certified": "AQ",
    "learning_capacity": "AR",
    "special_circumstances": "AS",
    "special_circumstances_details": "AT",
    "relationship_to_head": "AU",
    "parent_guardian": "AV",
    "phone": "AW",
    "notes": "AX",
}


@dataclass(frozen=True)
class ReaderSettings:
    preferred_sheet: str = "MauNhapLieu"
    header_row: int = 2
    data_start_row: int = 5
    max_header_scan_rows: int = 10
    max_workers: int = 4


def reader_worker_count(value) -> int:
    count = int(value)
    if not 1 <= count <= 8:
        raise ValueError("Số luồng đọc phải từ 1 đến 8.")
    return count


@dataclass(frozen=True)
class Settings:
    commune_name: str = ""
    school_year: str | None = None
    source_path: Path | None = None
    template_path: Path | None = None
    reader: ReaderSettings = field(default_factory=ReaderSettings)
    source_columns: dict[str, str] = field(
        default_factory=lambda: DEFAULT_SOURCE_COLUMNS.copy()
    )
    output_directory: Path = Path("output")

    @property
    def report_path(self) -> Path:
        return self.output_directory / "CANH_BAO_DU_LIEU.xlsx"


def validate_export_settings(settings: Settings) -> None:
    missing = []
    if not isinstance(settings.commune_name, str) or not settings.commune_name.strip():
        missing.append("commune_name")
    if not isinstance(settings.school_year, str) or not settings.school_year.strip():
        missing.append("school_year")
    if missing:
        raise ValueError(
            "Missing required export settings: " + ", ".join(missing)
        )

    match = re.fullmatch(r"(\d{4})-(\d{4})", settings.school_year.strip())
    if match is None or int(match.group(2)) != int(match.group(1)) + 1:
        raise ValueError(
            "Invalid school_year; expected YYYY-YYYY with consecutive years"
        )


def load_settings(path: str | Path = "config/settings.json") -> Settings:
    settings_path = Path(path)
    data: dict[str, Any] = json.loads(settings_path.read_text(encoding="utf-8"))
    reader_data = data.get("reader", {})
    reader = ReaderSettings(
        preferred_sheet=reader_data.get("preferred_sheet", "MauNhapLieu"),
        header_row=int(reader_data.get("header_row", 2)),
        data_start_row=int(reader_data.get("data_start_row", 5)),
        max_header_scan_rows=int(reader_data.get("max_header_scan_rows", 10)),
        max_workers=reader_worker_count(reader_data.get("max_workers", 4)),
    )
    output_directory = Path(data.get("output_directory") or "output")
    if not output_directory.is_absolute():
        output_directory = settings_path.resolve().parent.parent / output_directory
    source_path = _configured_path(data.get("source_path"), settings_path)
    template_path = _configured_path(data.get("template_path"), settings_path)
    return Settings(
        commune_name=(data.get("commune_name") or "").strip(),
        school_year=data.get("school_year"),
        source_path=source_path,
        template_path=template_path,
        reader=reader,
        source_columns={
            **DEFAULT_SOURCE_COLUMNS,
            **data.get("source_columns", {}),
        },
        output_directory=output_directory,
    )


def save_settings(
    path: str | Path = "config/settings.json",
    *,
    commune_name: str | None,
    school_year: str | None,
    source_path: str | Path | None = None,
    template_path: str | Path | None = None,
    output_directory: str | Path | None = None,
    reader_workers: int | str | None = None,
) -> Settings:
    settings_path = Path(path)
    data = json.loads(settings_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Settings JSON must contain an object at its root")
    if commune_name is not None and not isinstance(commune_name, str):
        raise TypeError("commune_name must be a string or None")
    if school_year is not None and not isinstance(school_year, str):
        raise TypeError("school_year must be a string or None")

    normalized_commune = commune_name.strip() if commune_name is not None else ""
    normalized_school_year = school_year.strip() if school_year is not None else ""
    data["commune_name"] = normalized_commune or None
    data["school_year"] = normalized_school_year or None
    if reader_workers is not None:
        data.setdefault("reader", {})["max_workers"] = reader_worker_count(reader_workers)
    for key, value in (
        ("source_path", source_path),
        ("template_path", template_path),
        ("output_directory", output_directory),
    ):
        if value is not None:
            normalized_path = str(value).strip()
            data[key] = normalized_path or None

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=settings_path.parent,
            prefix=f".{settings_path.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            json.dump(data, temporary_file, ensure_ascii=False, indent=2)
            temporary_file.write("\n")
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
        os.replace(temporary_path, settings_path)
    except Exception:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
        raise

    return load_settings(settings_path)


def _configured_path(value: Any, settings_path: Path) -> Path | None:
    if not isinstance(value, str) or not value.strip():
        return None
    configured_path = Path(value.strip()).expanduser()
    if not configured_path.is_absolute():
        configured_path = settings_path.resolve().parent.parent / configured_path
    return configured_path
