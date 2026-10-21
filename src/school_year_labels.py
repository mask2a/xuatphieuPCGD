"""Update the verified school-year label cells on a copied template workbook."""

from __future__ import annotations

import re
from typing import Any


YEAR_LABEL_MERGES = (
    ("F9", "F9:G9"),
    ("F10", "F10:G10"),
    ("F11", "F11:G11"),
    ("F12", "F12:G12"),
    ("F13", "F13:G13"),
    ("F14", "F14:G14"),
    ("F15", "F15:G15"),
    ("F16", "F16:G16"),
    ("F17", "F17:G17"),
    ("F18", "F18:G18"),
    ("F19", "F19:G19"),
    ("F20", "F20:G20"),
    ("F21", "F21:G21"),
    ("F22", "F22:G22"),
    ("F23", "F23:G23"),
    ("F24", "F24:G24"),
    ("F25", "F25:G25"),
    ("F26", "F26:G26"),
    ("F27", "F27:G27"),
    ("F28", "F28:G28"),
    ("F29", "F29:G29"),
    ("F30", "F30:G30"),
    ("F31", "F31:G31"),
    ("F32", "F32:G32"),
    ("F33", "F33:G33"),
    ("F34", "F34:G34"),
    ("F35", "F35:G35"),
    ("F36", "F36:G36"),
    ("F37", "F37:G37"),
    ("F38", "F38:G38"),
    ("F39", "F39:G39"),
    ("F40", "F40:G40"),
    ("F41", "F41:G41"),
    ("F42", "F42:G42"),
    ("F43", "F43:G43"),
    ("F47", "F47:G47"),
    ("F48", "F48:G48"),
    ("F49", "F49:G49"),
    ("F50", "F50:G50"),
    ("F51", "F51:G51"),
    ("F52", "F52:G52"),
    ("F53", "F53:G53"),
    ("F54", "F54:G54"),
    ("F55", "F55:G55"),
    ("F56", "F56:G56"),
    ("F57", "F57:G57"),
    ("F58", "F58:G58"),
    ("F59", "F59:G59"),
    ("F60", "F60:G60"),
    ("F61", "F61:G61"),
    ("F62", "F62:G62"),
    ("F63", "F63:G63"),
    ("F64", "F64:G64"),
    ("F65", "F65:G65"),
    ("F66", "F66:G66"),
    ("F67", "F67:G67"),
    ("F68", "F68:G68"),
    ("F69", "F69:G69"),
    ("F70", "F70:G70"),
    ("F71", "F71:G71"),
    ("F72", "F72:G72"),
    ("F73", "F73:G73"),
    ("F74", "F74:G74"),
    ("F75", "F75:G75"),
    ("F76", "F76:G76"),
    ("F77", "F77:G77"),
    ("F78", "F78:G78"),
    ("F79", "F79:G79"),
    ("F80", "F80:G80"),
    ("F81", "F81:G81"),
)

YEAR_LABEL_CELLS = tuple(address for address, _ in YEAR_LABEL_MERGES)
SIGNATURE_YEAR_RANGE = "F82:W82"
_SCHOOL_YEAR_PATTERN = re.compile(r"(\d{4})-(\d{4})\Z")
_TEMPLATE_LABEL_PATTERN = re.compile(r"\d{4}\s*-\s*\d{4}\Z")
_EXPECTED_TEMPLATE_SHEET = "Phiếu điều tra 1 mặt"


def build_school_year_labels(school_year: str) -> tuple[str, ...]:
    if not isinstance(school_year, str):
        raise ValueError("school_year must be YYYY-YYYY with consecutive years")
    match = _SCHOOL_YEAR_PATTERN.fullmatch(school_year.strip())
    if match is None:
        raise ValueError("school_year must be YYYY-YYYY with consecutive years")
    first_year, second_year = (int(part) for part in match.groups())
    if second_year != first_year + 1:
        raise ValueError("school_year must be YYYY-YYYY with consecutive years")
    if first_year < 1000 or first_year + 5 > 9999:
        raise ValueError("school_year years must be four-digit calendar years")
    return tuple(
        f"{first_year + offset}-{first_year + offset + 1}"
        for offset in range(5)
    )


def update_school_year_labels(workbook: Any, school_year: str) -> tuple[str, ...]:
    labels = build_school_year_labels(school_year)
    if len(YEAR_LABEL_MERGES) != 70 or len(set(YEAR_LABEL_CELLS)) != 70:
        raise RuntimeError("The verified school-year cell allowlist is invalid")

    worksheet = workbook.Worksheets.Item(1)
    if worksheet.Name != _EXPECTED_TEMPLATE_SHEET:
        raise ValueError(
            f"Unexpected template sheet: {worksheet.Name!r}; "
            f"expected {_EXPECTED_TEMPLATE_SHEET!r}"
        )

    target_addresses = set(YEAR_LABEL_CELLS)
    original_values = _read_used_range_values(worksheet)
    actual_labels = _find_year_label_cells(worksheet, original_values)
    if actual_labels != target_addresses:
        missing = sorted(target_addresses - actual_labels)
        unexpected = sorted(actual_labels - target_addresses)
        raise ValueError(
            "Template year-label cells do not match the verified allowlist; "
            f"missing={missing}, unexpected={unexpected}"
        )

    targets = []
    for address, expected_merge in YEAR_LABEL_MERGES:
        cell = worksheet.Range(address)
        merge_area = cell.MergeArea.Address.replace("$", "").upper()
        if not cell.MergeCells or merge_area != expected_merge:
            raise ValueError(
                f"Unexpected merge area for {address}: {merge_area!r}; "
                f"expected {expected_merge!r}"
            )
        if cell.HasFormula:
            raise ValueError(f"Refusing to overwrite formula at {address}")
        targets.append(cell)

    signature_before = worksheet.Range(SIGNATURE_YEAR_RANGE).Value2
    for cell, label in zip(targets, labels * 14, strict=True):
        cell.Value2 = label

    updated_values = _read_used_range_values(worksheet)
    changed_cells = _changed_cell_addresses(
        worksheet,
        original_values,
        updated_values,
    )
    if changed_cells != target_addresses:
        raise RuntimeError(
            "Updating school-year labels changed an unexpected set of cells; "
            f"changed={sorted(changed_cells)}, expected={sorted(target_addresses)}"
        )
    if worksheet.Range(SIGNATURE_YEAR_RANGE).Value2 != signature_before:
        raise RuntimeError("Calendar years in the signature area were changed")

    return labels


def _read_used_range_values(worksheet: Any) -> tuple[tuple[Any, ...], ...]:
    values = worksheet.UsedRange.Value2
    if not isinstance(values, tuple):
        return ((values,),)
    if values and not isinstance(values[0], tuple):
        return (values,)
    return values


def _find_year_label_cells(
    worksheet: Any,
    values: tuple[tuple[Any, ...], ...],
) -> set[str]:
    used_range = worksheet.UsedRange
    first_row = int(used_range.Row)
    first_column = int(used_range.Column)
    found = set()
    for row_offset, row_values in enumerate(values):
        for column_offset, value in enumerate(row_values):
            if isinstance(value, str) and _TEMPLATE_LABEL_PATTERN.fullmatch(value):
                row = first_row + row_offset
                column = first_column + column_offset
                found.add(f"{_column_name(column)}{row}")
    return found


def _changed_cell_addresses(
    worksheet: Any,
    before: tuple[tuple[Any, ...], ...],
    after: tuple[tuple[Any, ...], ...],
) -> set[str]:
    if len(before) != len(after) or any(
        len(before_row) != len(after_row)
        for before_row, after_row in zip(before, after, strict=True)
    ):
        raise RuntimeError("The worksheet used range changed during label update")
    used_range = worksheet.UsedRange
    first_row = int(used_range.Row)
    first_column = int(used_range.Column)
    changed = set()
    for row_offset, (before_row, after_row) in enumerate(
        zip(before, after, strict=True)
    ):
        for column_offset, (before_value, after_value) in enumerate(
            zip(before_row, after_row, strict=True)
        ):
            if before_value != after_value:
                row = first_row + row_offset
                column = first_column + column_offset
                changed.add(f"{_column_name(column)}{row}")
    return changed


def _column_name(column_number: int) -> str:
    name = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        name = chr(65 + remainder) + name
    return name
