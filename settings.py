"""Load project settings without supplying inferred locality or school-year data."""

from __future__ import annotations

import json
import re
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


@dataclass(frozen=True)
class Settings:
    commune_name: str = ""
    school_year: str | None = None
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
    )
    output_directory = Path(data.get("output_directory", "output"))
    if not output_directory.is_absolute():
        output_directory = settings_path.resolve().parent.parent / output_directory
    return Settings(
        commune_name=(data.get("commune_name") or "").strip(),
        school_year=data.get("school_year"),
        reader=reader,
        source_columns={
            **DEFAULT_SOURCE_COLUMNS,
            **data.get("source_columns", {}),
        },
        output_directory=output_directory,
    )
