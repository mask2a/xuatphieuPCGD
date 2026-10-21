"""Load and validate the verified source-to-template mapping."""

from __future__ import annotations

import json
import re
import string
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class TemplateMapping:
    template_sheet: str
    member_start_rows: tuple[int, ...]
    member_rows_per_slot: int
    confirmation_start_row: int
    confirmation_end_row: int
    confirmation_commune_cell: str
    confirmation_merge_area: str
    print_area: str
    last_column_number: int
    full_form_zoom: int
    expected_merge_count: int
    short_form_max_members: int
    fixed_fields: dict[str, dict[str, str]]
    member_fields: dict[str, dict[str, Any]]
    approved_member_fields: dict[str, dict[str, Any]]
    unmapped_household_fields: dict[str, str]
    unmapped_member_fields: dict[str, str]
    school_year_policy: str


def load_template_mapping(path: str | Path) -> TemplateMapping:
    mapping_path = Path(path)
    data = json.loads(mapping_path.read_text(encoding="utf-8"))
    rows = tuple(int(row) for row in data["member_start_rows"])
    if len(rows) != 14 or len(set(rows)) != 14 or tuple(sorted(rows)) != rows:
        raise ValueError("The template mapping must define 14 ordered member slots")
    if int(data["short_form_max_members"]) not in range(1, len(rows) + 1):
        raise ValueError("short_form_max_members is outside the template capacity")

    fixed_fields = data["fixed_fields"]
    expected_fixed_fields = {
        "commune_name",
        "household_code",
        "household_head_name",
        "phone",
    }
    if set(fixed_fields) != expected_fixed_fields:
        raise ValueError("The fixed-field mapping is incomplete or contains unknown fields")

    member_fields = data["member_fields"]
    expected_member_fields = {
        "member_name",
        "relationship_to_head",
        "birth_date",
        "ethnicity",
        "religion",
        "parent_guardian",
    }
    if set(member_fields) != expected_member_fields:
        raise ValueError("The confirmed member-field mapping is incomplete or contains unknown fields")
    for field in member_fields.values():
        if field["mode"] not in {"value", "labelled", "birth_components"}:
            raise ValueError(f"Unsupported mapping mode: {field['mode']}")
        if int(field["row_offset"]) not in range(int(data["member_rows_per_slot"])):
            raise ValueError("A member-field row offset is outside its slot")
        if field["column"] not in string.ascii_uppercase:
            raise ValueError(f"Invalid Excel column in mapping: {field['column']}")

    approved_member_fields = data["approved_member_fields"]
    expected_approved_fields = {
        "class_name": ("class_name", "H"),
        "school_name": ("school_name", "K"),
        "graduation_level": ("graduation_level", "P"),
        "graduation_year": ("graduation_year", "R"),
    }
    if set(approved_member_fields) != set(expected_approved_fields):
        raise ValueError(
            "The approved education-field mapping is incomplete or contains "
            "unapproved fields"
        )
    approved_addresses: set[str] = set()
    for field_name, (expected_source, expected_column) in expected_approved_fields.items():
        field = approved_member_fields[field_name]
        if field.get("source") != expected_source:
            raise ValueError(f"Unexpected source for approved field {field_name}")
        if (
            len(field["cells_by_slot"]) != len(rows)
            or len(field["merge_areas_by_slot"]) != len(rows)
        ):
            raise ValueError(
                f"Approved field {field_name} must define every member slot"
            )
        for slot_index, (address, merge_area, start_row) in enumerate(
            zip(
                field["cells_by_slot"],
                field["merge_areas_by_slot"],
                rows,
                strict=True,
            )
        ):
            match = re.fullmatch(r"([A-Z]{1,3})(\d+)", address)
            if (
                match is None
                or match.group(1) != expected_column
                or int(match.group(2)) != start_row
            ):
                raise ValueError(
                    f"Approved field {field_name} has an unverified target "
                    f"at slot {slot_index + 1}: {address!r}"
                )
            if address in approved_addresses:
                raise ValueError(f"Duplicate approved template cell: {address}")
            approved_addresses.add(address)
            if merge_area is not None:
                merge_match = re.fullmatch(
                    rf"{re.escape(address)}:([A-Z]{{1,3}}){start_row}",
                    merge_area,
                )
                if merge_match is None:
                    raise ValueError(
                        f"Invalid verified merge area for {field_name}: "
                        f"{merge_area!r}"
                    )
        if field_name == "class_name":
            if field.get("comparison_source") != "grade_group":
                raise ValueError("B5 must retain grade_group as its comparison source")
            if not isinstance(field.get("review_values"), list):
                raise ValueError("B5 review_values must be a list")

    return TemplateMapping(
        template_sheet=data["template_sheet"],
        member_start_rows=rows,
        member_rows_per_slot=int(data["member_rows_per_slot"]),
        confirmation_start_row=int(data["confirmation_start_row"]),
        confirmation_end_row=int(data["confirmation_end_row"]),
        confirmation_commune_cell=data["confirmation_commune_cell"],
        confirmation_merge_area=data["confirmation_merge_area"],
        print_area=data["print_area"],
        last_column_number=int(data["last_column_number"]),
        full_form_zoom=int(data["full_form_zoom"]),
        expected_merge_count=int(data["expected_merge_count"]),
        short_form_max_members=int(data["short_form_max_members"]),
        fixed_fields={key: dict(value) for key, value in fixed_fields.items()},
        member_fields={key: dict(value) for key, value in member_fields.items()},
        approved_member_fields={
            key: dict(value) for key, value in approved_member_fields.items()
        },
        unmapped_household_fields=dict(data["unmapped_household_fields"]),
        unmapped_member_fields=dict(data["unmapped_member_fields"]),
        school_year_policy=data["school_year_policy"],
    )
