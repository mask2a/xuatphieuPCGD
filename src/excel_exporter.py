"""Export each household from an untouched copy of the verified Excel template."""

from __future__ import annotations

import logging
import hashlib
import math
import os
import re
import gc
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Callable

from household_grouper import Household
from settings import load_settings, validate_export_settings
from src.school_year_labels import update_school_year_labels
from src.approved_mapping import (
    values as additional_values, verify_targets, household_address, address_label, template_label, key,
)
from src.template_mapper import TemplateMapping, load_template_mapping
from src.print_layout import (
    CHECKED, UNCHECKED, CHECKBOX_FONT, ETHNICITY_FONT_SIZE, PAGE_CAPACITY, PRINT_ZOOM, checkbox,
    normalize_geometry, verify_page_setup, invalid_birth_date, fit_text, class_font_size,
)


LOGGER = logging.getLogger(__name__)
XLSX_FILE_FORMAT = 51
INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class ExportWarning:
    household_code: str | None
    code: str
    message: str


@dataclass(frozen=True)
class ExportedHousehold:
    household_code: str | None
    path: Path
    member_count: int
    form_type: str
    supplementary_sheet_count: int
    warnings: tuple[ExportWarning, ...]


@dataclass(frozen=True)
class ExportFailure:
    household_code: str | None
    error: str


@dataclass(frozen=True)
class ExportReport:
    successes: tuple[ExportedHousehold, ...]
    failures: tuple[ExportFailure, ...]
    warnings: tuple[ExportWarning, ...]

    @property
    def success_count(self) -> int:
        return len(self.successes)

    @property
    def failure_count(self) -> int:
        return len(self.failures)

    @property
    def summary(self) -> dict[str, int]:
        return {
            "success_count": self.success_count,
            "failure_count": self.failure_count,
            "warning_count": len(self.warnings),
        }


@dataclass(frozen=True)
class TemplateSignature:
    merge_count: int
    cell_formats: dict[str, tuple]
    row_heights: tuple[float, ...]
    column_widths: tuple[float, ...]
    page_setup: tuple


class ExcelAutomationUnavailable(RuntimeError):
    pass


class ExportConfigurationError(ValueError):
    pass


def export_households(
    households: Iterable[Household],
    template_path: str | Path,
    output_dir: str | Path,
    mapping_path: str | Path,
    *,
    settings_path: str | Path | None = None,
    apply_school_year_labels: bool = False,
    progress: Callable[[int, int, str], None] | None = None,
    cancel_event=None,
) -> ExportReport:
    """Export one workbook per household, sharing one Excel COM session."""
    template = Path(template_path).resolve()
    destination = Path(output_dir).resolve()
    mapping = load_template_mapping(mapping_path)
    configured_settings_path = (
        Path(settings_path)
        if settings_path is not None
        else Path(__file__).resolve().parents[1] / "config" / "settings.json"
    )
    try:
        settings = load_settings(configured_settings_path)
        validate_export_settings(settings)
    except (OSError, ValueError) as error:
        raise ExportConfigurationError(
            f"Export configuration is invalid at {configured_settings_path}: {error}"
        ) from error
    household_list = tuple(households)
    if not template.is_file():
        raise FileNotFoundError(f"Excel template does not exist: {template}")
    if not household_list:
        return ExportReport((), (), ())
    for household in household_list:
        if _normalized_commune(household.commune_name) != _normalized_commune(
            settings.commune_name
        ):
            raise ExportConfigurationError(
                f"Household {household.household_code!r} commune_name "
                "does not match configured commune_name"
            )

    destination.mkdir(parents=True, exist_ok=True)
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise ExcelAutomationUnavailable(
            "pywin32 is required to export Excel workbooks through Microsoft Excel COM"
        ) from error

    successes: list[ExportedHousehold] = []
    failures: list[ExportFailure] = []
    all_warnings: list[ExportWarning] = []
    excel = None
    template_workbook = None
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        template_workbook = excel.Workbooks.Open(
            str(template),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        template_sheet = _verify_template_sheet(template_workbook, mapping)
        template_signature = _template_signature(template, template_sheet, mapping)

        for index, household in enumerate(household_list):
            if cancel_event is not None and cancel_event.is_set():
                break
            if progress:
                progress(index, len(household_list), f"Excel {index + 1}/{len(household_list)}; thành công {len(successes)}, lỗi {len(failures)}")
            code = _household_code(household)
            final_path = destination / f"ho_{_safe_filename_component(code)}.xlsx"
            if final_path.exists():
                failures.append(
                    ExportFailure(code, f"Output already exists and was not overwritten: {final_path}")
                )
                continue

            warnings = _household_export_warnings(
                household,
                mapping,
                settings.school_year,
                apply_school_year_labels,
            )
            all_warnings.extend(warnings)
            temporary_copy: Path | None = None
            temporary_xlsx: Path | None = None
            workbook = None
            try:
                temporary_copy = _temporary_path(destination, ".xls")
                template_workbook.SaveCopyAs(str(temporary_copy))
                workbook = excel.Workbooks.Open(
                    str(temporary_copy),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                    IgnoreReadOnlyRecommended=True,
                )
                update_school_year_labels(workbook, "2025-2026")
                previous_warning_count = len(warnings)
                sheets = _populate_workbook(workbook, household, mapping, warnings)
                all_warnings.extend(warnings[previous_warning_count:])
                temporary_xlsx = _temporary_path(destination, ".xlsx")
                workbook.SaveAs(
                    str(temporary_xlsx),
                    FileFormat=XLSX_FILE_FORMAT,
                    CreateBackup=False,
                    AddToMru=False,
                )
                workbook.Close(SaveChanges=False)
                workbook = None
                _validate_saved_workbook(
                    excel,
                    temporary_xlsx,
                    household,
                    mapping,
                    len(sheets),
                    template_signature,
                )
                os.rename(temporary_xlsx, final_path)
                temporary_xlsx = None
                successes.append(
                    ExportedHousehold(
                        household_code=code,
                        path=final_path,
                        member_count=len(household.members),
                        form_type=(
                            "short"
                            if len(household.members) <= PAGE_CAPACITY
                            else "full"
                        ),
                        supplementary_sheet_count=max(
                            0,
                            math.ceil(
                                len(household.members)
                                / PAGE_CAPACITY
                            ) - 1,
                        ),
                        warnings=tuple(warnings),
                    )
                )
            except Exception as error:
                LOGGER.error("Failed exporting household %s: %s", code, error)
                failures.append(ExportFailure(code, str(error)))
                if progress:
                    progress(index + 1, len(household_list), f"LỖI Excel hộ {code}: {error}")
            finally:
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception as error:
                        LOGGER.error(
                            "Failed closing workbook for household %s: %s",
                            code,
                            error,
                        )
                        if failures and failures[-1].household_code == code:
                            previous = failures[-1]
                            failures[-1] = ExportFailure(
                                code,
                                f"{previous.error}; workbook close failed: {error}",
                            )
                        else:
                            failures.append(ExportFailure(code, f"Workbook close failed: {error}"))
                    finally:
                        workbook = None
                sheets = None
                for temporary_file in (temporary_copy, temporary_xlsx):
                    try:
                        _remove_if_exists(temporary_file)
                    except OSError as error:
                        LOGGER.error(
                            "Failed removing temporary file for household %s: %s",
                            code,
                            error,
                        )
                        all_warnings.append(
                            ExportWarning(
                                code,
                                "TEMPORARY_FILE_CLEANUP_FAILED",
                                str(error),
                            )
                        )
    finally:
        try:
            template_sheet = None
            if template_workbook is not None:
                template_workbook.Close(SaveChanges=False)
                template_workbook = None
        finally:
            try:
                if excel is not None:
                    excel.Quit()
                    excel = None
            finally:
                gc.collect()
                if com_initialized:
                    pythoncom.CoUninitialize()

    report = ExportReport(
        successes=tuple(successes),
        failures=tuple(failures),
        warnings=tuple(all_warnings),
    )
    if progress:
        progress(len(household_list), len(household_list), f"Excel: {report.success_count} thành công, {report.failure_count} lỗi")
    LOGGER.info(
        "Excel export completed: %d success, %d failed, %d warnings",
        report.success_count,
        report.failure_count,
        len(report.warnings),
    )
    return report


def _populate_workbook(
    workbook,
    household: Household,
    mapping: TemplateMapping,
    layout_warnings=None,
) -> tuple[object, ...]:
    sheet_count = max(
        1,
        math.ceil(len(household.members) / PAGE_CAPACITY),
    )
    sheets = [workbook.Worksheets.Item(mapping.template_sheet)]
    for page_number in range(2, sheet_count + 1):
        sheets[0].Copy(None, workbook.Worksheets.Item(workbook.Worksheets.Count))
        page = workbook.ActiveSheet
        page.Name = f"BoSung_{page_number}"
        sheets.append(page)

    for page_index, sheet in enumerate(sheets):
        first_member_index = page_index * PAGE_CAPACITY
        page_members = household.members[
            first_member_index:first_member_index + PAGE_CAPACITY
        ]
        _populate_header(sheet, household, mapping)
        _populate_member_page(sheet, page_members, first_member_index, mapping)
        _configure_page(sheet, len(page_members), mapping)
        for slot_index, member in enumerate(page_members):
            row = mapping.member_start_rows[slot_index]
            class_address = mapping.approved_member_fields["class_name"]["cells_by_slot"][slot_index]
            class_cell = sheet.Range(class_address)
            original_size = class_cell.Font.Size
            reduced_size = class_font_size(class_cell, original_size)
            if reduced_size < original_size:
                class_cell.Font.Size = reduced_size
                if layout_warnings is not None:
                    layout_warnings.append(ExportWarning(_household_code(household), "LONG_CLASS_FONT_REDUCED", f"Thành viên {first_member_index + slot_index + 1}, ô {class_address}: giảm riêng font lớp từ {original_size:g} xuống {reduced_size:g} pt để giữ đủ chữ."))
            targets = [f'{field["column"]}{row + int(field["row_offset"])}' for field in mapping.member_fields.values()]
            targets += [field["cells_by_slot"][slot_index] for field in mapping.approved_member_fields.values()]
            targets += [f"{column}{row}" for column in additional_values(member)[0]]
            for address in targets:
                if address == class_address and reduced_size < original_size:
                    continue  # Already reported the approved local font reduction.
                if fit_text(sheet.Range(address)) and layout_warnings is not None:
                    layout_warnings.append(ExportWarning(_household_code(household), "LONG_TEXT_REQUIRES_REVIEW", f"Thành viên {first_member_index + slot_index + 1}, ô {address}: nội dung dài có nguy cơ vượt ô; giữ nguyên định dạng mẫu, cần rà soát trước khi in."))
        for address in ("A1", "A3", "H3", "Z1", "Z3"):
            if fit_text(sheet.Range(address)) and layout_warnings is not None:
                layout_warnings.append(ExportWarning(_household_code(household), "LONG_TEXT_REQUIRES_REVIEW", f"Thông tin hộ, ô {address}: nội dung dài có nguy cơ vượt ô; giữ nguyên định dạng mẫu, cần rà soát trước khi in."))
    return tuple(sheets)


def _populate_header(sheet, household: Household, mapping: TemplateMapping) -> None:
    address = household_address(household)
    _write_labelled_cell(sheet.Range("A3"), address_label(sheet), address or "", "address")
    residence_cell = sheet.Range("Z2")
    _ensure_not_formula(residence_cell, "Z2")
    residence_number_format = residence_cell.NumberFormat
    residence = key(household.residence)
    text = str(residence_cell.Value2)
    positions = [i for i, char in enumerate(text) if char == "q"]
    for index, position in enumerate(positions):
        selected = residence == ("thường trú", "tạm trú")[index]
        _replace_text_range(residence_cell, position + 1, 1, CHECKED if selected else UNCHECKED)
        _character_range(residence_cell, position + 1, 1).Font.Name = CHECKBOX_FONT
        _character_range(residence_cell, position + 1, 1).Font.Size = 11
    residence_cell.NumberFormat = residence_number_format
    commune_field = mapping.fixed_fields["commune_name"]
    commune_cell = sheet.Range(commune_field["cell"])
    commune_label = template_label(sheet, commune_field["cell"])
    _ensure_not_formula(commune_cell, commune_field["cell"])
    original_commune = _content_after_label(
        str(commune_cell.Value2 or ""),
        commune_label,
    )
    if not household.commune_name.strip():
        raise ExportConfigurationError("commune_name must be configured before export")
    commune_cell.Value2 = f'{commune_label} {household.commune_name.strip()}'

    confirmation = sheet.Range(mapping.confirmation_commune_cell)
    _ensure_not_formula(confirmation, mapping.confirmation_commune_cell)
    confirmation_text = str(confirmation.Value2 or "")
    if not original_commune:
        confirmation_heading = re.search(
            r"XÁC NHẬN CỦA UBND[^\S\r\n]+(?:XÃ|PHƯỜNG|THỊ TRẤN)[^\S\r\n]+([^\r\n]+)",
            confirmation_text,
            flags=re.IGNORECASE,
        )
        if confirmation_heading is None:
            raise ExportConfigurationError(
                "Cannot identify the template commune in the confirmation heading"
            )
        original_commune = confirmation_heading.group(1).strip()
    if original_commune:
        original_name = original_commune.strip()
        start_index = confirmation_text.find(original_name)
        if start_index < 0:
            raise ExportConfigurationError(
                "The confirmation text does not contain the template commune name"
            )
        number_format = confirmation.NumberFormat
        _replace_text_range(
            confirmation,
            start_index + 1,
            len(original_name),
            household.commune_name.strip(),
        )
        if confirmation.NumberFormat != number_format:
            confirmation.NumberFormat = number_format

    _write_labelled(
        sheet,
        mapping.fixed_fields["household_code"],
        _household_code(household),
    )
    if household.household_head_name:
        _write_labelled(
            sheet,
            mapping.fixed_fields["household_head_name"],
            household.household_head_name,
        )
    if household.phone:
        _write_labelled(sheet, mapping.fixed_fields["phone"], household.phone)


def _populate_member_page(
    sheet,
    members,
    first_member_index: int,
    mapping: TemplateMapping,
) -> None:
    ethnicity_field = mapping.member_fields["ethnicity"]
    for row in mapping.member_start_rows:
        checkbox(sheet.Range(f"C{row + 3}"), False)
        ethnicity_address = f'{ethnicity_field["column"]}{row + int(ethnicity_field["row_offset"])}'
        sheet.Range(ethnicity_address).Font.Size = ETHNICITY_FONT_SIZE
    for slot_index, member in enumerate(members):
        start_row = mapping.member_start_rows[slot_index]
        female = sheet.Range(f"C{start_row + 3}")
        _ensure_not_formula(female, f"C{start_row + 3}")
        checkbox(female, key(member.fields.get("female")) in {"x", "nữ"})
        extra_values, _ = additional_values(member)
        for column, value in extra_values.items():
            if _has_value(value):
                cell = sheet.Range(f"{column}{start_row}")
                _ensure_not_formula(cell, f"{column}{start_row}")
                _write_preserving_source_value(cell, value)
        for field_name, field_mapping in mapping.member_fields.items():
            value = _member_field_value(member, field_mapping["source"])
            if value is None or value == "":
                continue
            target = f'{field_mapping["column"]}{start_row + int(field_mapping["row_offset"])}'
            cell = sheet.Range(target)
            _ensure_not_formula(cell, target)
            if field_mapping["mode"] == "value":
                cell.Value2 = value
            elif field_mapping["mode"] == "labelled":
                _write_labelled_cell(
                    cell,
                    field_mapping["label"],
                    str(value),
                    field_name,
                )
            elif field_mapping["mode"] == "birth_components":
                components = tuple(value)
                if any(component is not None for component in components):
                    cell.Value2 = "/".join(component or "" for component in components)
        for field_mapping in mapping.approved_member_fields.values():
            value = member.fields.get(field_mapping["source"])
            if not _has_value(value):
                continue
            target = field_mapping["cells_by_slot"][slot_index]
            cell = sheet.Range(target)
            _ensure_not_formula(cell, target)
            _write_preserving_source_value(cell, value)
        sequence_cell = sheet.Range(f"A{start_row}")
        _ensure_not_formula(sequence_cell, f"A{start_row}")
        sequence_cell.Value2 = first_member_index + slot_index + 1


def _member_field_value(member, source: str):
    if source == "member_name":
        return member.member_name or None
    if source == "relationship_to_head":
        return member.relationship_to_head
    if source == "birth_components":
        return member.birth_components
    return member.normalized_fields.get(source)


def _configure_page(sheet, member_count: int, mapping: TemplateMapping) -> None:
    normalize_geometry(sheet, member_count, mapping)


def _verify_template_sheet(workbook, mapping: TemplateMapping):
    try:
        sheet = workbook.Worksheets.Item(mapping.template_sheet)
    except Exception as error:
        raise ExportConfigurationError(
            f"Template sheet not found: {mapping.template_sheet}"
        ) from error
    if (
        sheet.Range(mapping.confirmation_commune_cell)
        .MergeArea.Address.replace("$", "")
        != mapping.confirmation_merge_area
    ):
        raise ExportConfigurationError("Template confirmation cell differs from the verified layout")
    for field in mapping.fixed_fields.values():
        cell = sheet.Range(field["cell"])
        if cell.MergeCells is False and field["cell"] in {"A1", "Z1", "H3", "Z3"}:
            raise ExportConfigurationError(
                f"Expected merged template cell at {field['cell']}"
            )
        expected_label = template_label(sheet, field["cell"]) if field["cell"] == mapping.fixed_fields["commune_name"]["cell"] else field["label"]
        if not str(cell.Value2 or "").startswith(expected_label):
            raise ExportConfigurationError(
                f"Template label differs from mapping at {field['cell']}"
            )
    for slot_number, start_row in enumerate(mapping.member_start_rows, start=1):
        if sheet.Range(f"A{start_row}").Value2 != slot_number:
            raise ExportConfigurationError(
                f"Template member numbering differs at row {start_row}"
            )
    verify_targets(sheet, mapping.member_start_rows)
    for start_row in mapping.member_start_rows:
        for field_name, field in mapping.member_fields.items():
            address = f'{field["column"]}{start_row + int(field["row_offset"])}'
            expected_merge = re.sub(r"\d+", lambda match: str(int(match.group()) + start_row - mapping.member_start_rows[0]), field["merge_area"])
            merge_area = sheet.Range(address).MergeArea.Address.replace("$", "")
            if merge_area != expected_merge:
                raise ExportConfigurationError(f"Template merge area differs for {field_name} at {address}")
        for offset in range(5):
            cell = sheet.Range(f"F{start_row + offset}")
            if cell.MergeArea.Address.replace("$", "") != f"F{start_row + offset}:G{start_row + offset}":
                raise ExportConfigurationError("Template year-label merge differs")
            if str(cell.Value2).replace(" ", "") != f"{2025 + offset}-{2026 + offset}":
                raise ExportConfigurationError("Template year labels differ from the approved fixed years")
    for column, year in zip(("F", "J", "N", "R", "W"), range(2025, 2030)):
        if not str(sheet.Range(f"{column}82").Value2 or "").strip().endswith(f"Năm {year}"):
            raise ExportConfigurationError("Template signature years differ")
    if sheet.PageSetup.PaperSize != 9 or sheet.PageSetup.Orientation != 2:
        raise ExportConfigurationError("Template must use A4 landscape")
    if workbook.LinkSources(1) is not None:
        raise ExportConfigurationError("Template contains external workbook links")
    for field_name, field in mapping.approved_member_fields.items():
        for address, expected_merge in zip(
            field["cells_by_slot"],
            field["merge_areas_by_slot"],
            strict=True,
        ):
            cell = sheet.Range(address)
            actual_merge = (
                cell.MergeArea.Address.replace("$", "")
                if cell.MergeCells
                else None
            )
            if actual_merge != expected_merge:
                raise ExportConfigurationError(
                    f"Template merge area differs for approved field "
                    f"{field_name} at {address}: {actual_merge!r}"
                )
    return sheet


def _template_signature(
    template_path: Path,
    sheet,
    mapping: TemplateMapping,
) -> TemplateSignature:
    import xlrd

    try:
        LOGGER.info("Template SHA-256 (informational only): %s", hashlib.sha256(template_path.read_bytes()).hexdigest())
    except OSError:
        LOGGER.warning("Template SHA-256 unavailable; continuing structural validation")

    source = xlrd.open_workbook(str(template_path), formatting_info=True)
    try:
        merge_count = len(source.sheet_by_name(mapping.template_sheet).merged_cells)
    finally:
        source.release_resources()
    if merge_count != mapping.expected_merge_count:
        raise ExportConfigurationError(
            f"Template has {merge_count} merged ranges; expected "
            f"{mapping.expected_merge_count}"
        )

    addresses = {
        "A1", "Z1", "A3", "H3", "Z3", "B9", "B10", "D11", "C12",
        "D12", "E12", "B13", mapping.confirmation_commune_cell,
    }
    addresses.update(
        f'{field["column"]}{mapping.member_start_rows[0] + int(field["row_offset"])}'
        for field in mapping.member_fields.values()
    )
    addresses.update(
        address
        for field in mapping.approved_member_fields.values()
        for address in field["cells_by_slot"]
    )
    cell_formats = {
        address: _com_cell_format(sheet.Range(address))
        for address in addresses
    }
    row_heights = tuple(
        float(sheet.Rows.Item(row).RowHeight)
        for row in range(1, mapping.confirmation_end_row + 1)
    )
    column_widths = tuple(
        float(sheet.Columns.Item(column).ColumnWidth)
        for column in range(1, mapping.last_column_number + 1)
    )
    page_setup = _com_page_setup(sheet)
    return TemplateSignature(
        merge_count=merge_count,
        cell_formats=cell_formats,
        row_heights=row_heights,
        column_widths=column_widths,
        page_setup=page_setup,
    )


def _validate_saved_workbook(
    excel,
    path: Path,
    household: Household,
    mapping: TemplateMapping,
    expected_sheet_count: int,
    signature: TemplateSignature,
) -> None:
    from openpyxl import load_workbook

    workbook = None
    try:
        output_workbook = load_workbook(path, data_only=False)
        try:
            if len(output_workbook._external_links) != 0:
                raise ValueError("Saved workbook contains external workbook links")
            if any(
                len(sheet.merged_cells.ranges) != signature.merge_count
                for sheet in output_workbook.worksheets
            ):
                raise ValueError("Saved workbook has a different merged-cell structure")
            for sheet in output_workbook:
                for row, expected_height in enumerate(signature.row_heights, start=1):
                    dimension = sheet.row_dimensions[row]
                    height = dimension.height if dimension.height is not None else sheet.sheet_format.defaultRowHeight
                    if abs(float(height) - expected_height) > 0.05:
                        raise ValueError(f"Saved workbook stored row height differs at row {row}")
                    if bool(dimension.hidden) != (44 <= row <= 81):
                        raise ValueError(f"Saved workbook pagination visibility differs at row {row}")
        finally:
            output_workbook.close()

        workbook = excel.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        if workbook.Worksheets.Count != expected_sheet_count:
            raise ValueError(
                f"Saved workbook has {workbook.Worksheets.Count} worksheets; "
                f"expected {expected_sheet_count}"
            )
        main_sheet = workbook.Worksheets.Item(mapping.template_sheet)
        if main_sheet.Range(mapping.fixed_fields["commune_name"]["cell"]).Value2 != (
            f'{template_label(main_sheet, mapping.fixed_fields["commune_name"]["cell"])} {household.commune_name.strip()}'
        ):
            raise ValueError("Saved workbook commune value does not match the household")
        if main_sheet.Range(mapping.confirmation_commune_cell).Value2.find(
            household.commune_name.strip()
        ) < 0:
            raise ValueError("Saved workbook confirmation does not contain the configured commune")
        if main_sheet.PageSetup.PrintArea != mapping.print_area:
            raise ValueError("Saved workbook print area differs from the configured area")
        for page_index in range(expected_sheet_count):
            sheet = workbook.Worksheets.Item(
                mapping.template_sheet if page_index == 0 else f"BoSung_{page_index + 1}"
            )
            page_members = household.members[
                page_index * PAGE_CAPACITY:
                (page_index + 1) * PAGE_CAPACITY
            ]
            _validate_sheet_format(sheet, signature, len(page_members), mapping.approved_member_fields["class_name"]["cells_by_slot"][:len(page_members)])
            verify_page_setup(sheet, mapping.print_area)
            residence_cell = sheet.Range("Z2")
            residence_text = str(residence_cell.Value2)
            expected_markers = [CHECKED if key(household.residence) == value else UNCHECKED for value in ("thường trú", "tạm trú")]
            actual_markers = [char for char in residence_text if char in {CHECKED, UNCHECKED}]
            if actual_markers != expected_markers:
                raise ValueError("Saved workbook residence checkboxes differ")
            for position, char in enumerate(residence_text, start=1):
                if char in {CHECKED, UNCHECKED} and _character_range(residence_cell, position, 1).Font.Name != CHECKBOX_FONT:
                    raise ValueError("Saved workbook residence checkbox font differs")
            for slot, row in enumerate(mapping.member_start_rows):
                marker = sheet.Range(f"C{row + 3}")
                expected = CHECKED if slot < len(page_members) and key(page_members[slot].fields.get("female")) in {"x", "nữ"} else UNCHECKED
                if marker.Value2 != expected or marker.Font.Name != CHECKBOX_FONT:
                    raise ValueError("Saved workbook checkbox value/font differs")
                ethnicity_field = mapping.member_fields["ethnicity"]
                ethnicity_address = f'{ethnicity_field["column"]}{row + int(ethnicity_field["row_offset"])}'
                if sheet.Range(ethnicity_address).Font.Size != ETHNICITY_FONT_SIZE:
                    raise ValueError("Saved workbook ethnicity font size differs")
            if (
                len(page_members) <= PAGE_CAPACITY
                and (
                    sheet.PageSetup.Zoom != PRINT_ZOOM
                    or sheet.PageSetup.FitToPagesWide != 1
                    or sheet.PageSetup.FitToPagesTall != 1
                )
            ):
                raise ValueError("Short-form worksheet does not fit to one printed page")
            if sheet.PageSetup.Zoom != PRINT_ZOOM or sheet.PageSetup.FitToPagesWide != 1 or sheet.PageSetup.FitToPagesTall != 1:
                raise ValueError(
                    "Worksheet does not fit to one printed page"
                )
            for slot_index, member in enumerate(page_members):
                start_row = mapping.member_start_rows[slot_index]
                extra_values, _ = additional_values(member)
                for column, expected in extra_values.items():
                    actual = sheet.Range(f"{column}{start_row}").Value2
                    if (actual if _has_value(actual) else None) != (expected if _has_value(expected) else None):
                        raise ValueError(f"Saved workbook additional value differs at {column}{start_row}")
                expected_female = CHECKED if key(member.fields.get("female")) in {"x", "nữ"} else UNCHECKED
                if (sheet.Range(f"C{start_row + 3}").Value2 or "") != expected_female:
                    raise ValueError("Saved workbook female marker differs")
                if sheet.Range(f"B{start_row}").Value2 != member.member_name:
                    raise ValueError(
                        "Saved workbook does not contain member "
                        f"{page_index * PAGE_CAPACITY + slot_index + 1}"
                    )
                for field_name, field_mapping in (
                    mapping.approved_member_fields.items()
                ):
                    target = field_mapping["cells_by_slot"][slot_index]
                    expected_value = member.fields.get(field_mapping["source"])
                    expected_value = (
                        expected_value if _has_value(expected_value) else None
                    )
                    actual_value = sheet.Range(target).Value2
                    if actual_value != expected_value:
                        raise ValueError(
                            f"Saved workbook value differs for {field_name} "
                            f"at {target}"
                        )
        if workbook.LinkSources(1) is not None:
            raise ValueError("Saved workbook contains external workbook links")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
            workbook = None


def _com_cell_format(cell) -> tuple:
    font = cell.Font
    # Reuse COM proxies within this cell only. Each repeated Borders/Item
    # access crosses the Excel process boundary; all properties remain checked.
    collection = cell.Borders
    borders = []
    for index in (7, 8, 9, 10):
        border = collection.Item(index)
        borders.append((border.LineStyle, border.Weight, border.ColorIndex))
    return (
        font.Name,
        font.Size,
        font.Bold,
        font.Italic,
        cell.NumberFormat,
        cell.HorizontalAlignment,
        cell.VerticalAlignment,
        cell.WrapText,
        tuple(borders),
    )


def _com_page_setup(sheet) -> tuple:
    setup = sheet.PageSetup
    return (
        setup.Orientation,
        setup.PaperSize,
        float(setup.LeftMargin),
        float(setup.RightMargin),
        float(setup.TopMargin),
        float(setup.BottomMargin),
        float(setup.HeaderMargin),
        float(setup.FooterMargin),
        setup.CenterHorizontally,
        setup.CenterVertically,
    )


def _validate_sheet_format(sheet, signature: TemplateSignature, member_count=None, class_cells=()) -> None:
    for address, expected in signature.cell_formats.items():
        actual = _com_cell_format(sheet.Range(address))
        if member_count is not None and address == "C12":
            expected = (CHECKBOX_FONT, 11) + expected[2:]
        if member_count is not None and address == "D12":
            expected = (expected[0], ETHNICITY_FONT_SIZE) + expected[2:]
        if address in class_cells:
            expected = (expected[0], class_font_size(sheet.Range(address), expected[1])) + expected[2:]
        if actual != expected:
            raise ValueError(
                f"Saved workbook formatting differs at {address}: "
                f"expected {expected!r}, got {actual!r}"
            )
    for row, expected_height in enumerate(signature.row_heights, start=1):
        # Excel COM reports zero for hidden rows; stored heights are checked above.
        if sheet.Rows.Item(row).Hidden:
            continue
        if abs(float(sheet.Rows.Item(row).RowHeight) - expected_height) > 0.05:
            raise ValueError(f"Saved workbook row height differs at row {row}")
    for column, expected_width in enumerate(signature.column_widths, start=1):
        if abs(float(sheet.Columns.Item(column).ColumnWidth) - expected_width) > 0.05:
            raise ValueError(f"Saved workbook column width differs at column {column}")
    if member_count is not None:
        verify_page_setup(sheet)
    elif _com_page_setup(sheet) != signature.page_setup:
        raise ValueError("Saved workbook page setup differs from the verified template")


def _household_export_warnings(
    household: Household,
    mapping: TemplateMapping,
    school_year: str,
    apply_school_year_labels: bool,
) -> list[ExportWarning]:
    code = _household_code(household)
    warnings = [
        ExportWarning(code, warning.code, warning.message)
        for warning in household.warnings
    ]
    warnings.append(
        ExportWarning(
            code,
            "SCHOOL_YEAR_ROWS_NOT_MAPPED",
            (
                "Nhãn cố định 2025-2026 đến 2029-2030; chỉ ghi dữ liệu "
                "trên dòng năm học thứ nhất."
                if apply_school_year_labels
                else f"Năm học cấu hình {school_year}: {mapping.school_year_policy}"
            ),
        )
    )
    if household.household_head_name is None and any(
        warning.code == "HOUSEHOLD_HEAD_NAME_CONFLICT"
        for warning in household.warnings
    ):
        warnings.append(
            ExportWarning(
                code,
                "HOUSEHOLD_HEAD_NAME_NOT_EXPORTED",
                "Tên chủ hộ mâu thuẫn; giữ nguyên ô mẫu, không tự chọn giá trị.",
            )
        )
    if key(household.residence) not in {"thường trú", "tạm trú"}:
        warnings.append(
            ExportWarning(
                code,
                "HOUSEHOLD_RESIDENCE_NOT_EXPORTED",
                "Diện cư trú trống, không hợp lệ hoặc mâu thuẫn; không đánh dấu.",
            )
        )
    for field_name, reason in mapping.unmapped_household_fields.items():
        if field_name == "residence":
            continue
        if any(_has_value(member.fields.get(field_name)) for member in household.members):
            warnings.append(
                ExportWarning(code, f"UNMAPPED_{field_name.upper()}", reason)
            )
    for field_name, reason in mapping.unmapped_member_fields.items():
        if any(_has_value(member.fields.get(field_name)) for member in household.members):
            warnings.append(
                ExportWarning(code, f"UNMAPPED_{field_name.upper()}", reason)
            )
    for member_index, member in enumerate(household.members, start=1):
        if invalid_birth_date(member):
            warnings.append(ExportWarning(code, "INVALID_BIRTH_DATE", f"Thành viên {member_index}: ngày sinh không hợp lệ; giữ nguyên dữ liệu nguồn."))
        for warning_code, message in additional_values(member)[1]:
            warnings.append(ExportWarning(code, warning_code, f"Thành viên {member_index}: {message}"))
        for warning_code, message in _approved_education_warnings(member, mapping):
            warnings.append(
                ExportWarning(
                    code,
                    warning_code,
                    f"Thành viên {member_index}: {message}",
                )
            )
    if household_address(household) is None:
        warnings.append(ExportWarning(code, "HOUSEHOLD_ADDRESS_CONFLICT", "Địa chỉ hộ mâu thuẫn; không tự chọn địa chỉ."))
    return warnings


def _approved_education_warnings(
    member,
    mapping: TemplateMapping,
) -> list[tuple[str, str]]:
    warnings: list[tuple[str, str]] = []
    class_mapping = mapping.approved_member_fields["class_name"]
    grade = member.fields.get(class_mapping["comparison_source"])
    class_name = member.fields.get(class_mapping["source"])

    if _has_value(grade) and _has_value(class_name):
        grade_text = str(grade).strip()
        class_text = str(class_name).strip()
        grade_match = re.fullmatch(r"\d+", grade_text)
        class_match = re.match(r"(\d+)", class_text)
        if grade_match is None or class_match is None:
            warnings.append(
                (
                    "B5_GRADE_CLASS_NOT_COMPARABLE",
                    "Q/R không có dạng số khối và tiền tố số lớp để đối chiếu; "
                    "giữ nguyên lớp ở R.",
                )
            )
        elif grade_match.group(0) != class_match.group(1):
            warnings.append(
                (
                    "B5_GRADE_CLASS_CONFLICT",
                    f"Khối Q={grade_text!r} không khớp tiền tố lớp R={class_text!r}; "
                    "giữ nguyên lớp ở R.",
                )
            )
    elif _has_value(grade):
        warnings.append(
            (
                "B5_CLASS_VALUE_MISSING",
                "Q có dữ liệu nhưng R trống; không dùng Q để tự điền lớp.",
            )
        )
    elif _has_value(class_name):
        warnings.append(
            (
                "B5_GRADE_GROUP_MISSING",
                "R có dữ liệu nhưng Q trống nên không thể đối chiếu khối.",
            )
        )

    review_values = {
        str(value).strip().casefold()
        for value in class_mapping["review_values"]
    }
    if (
        _has_value(class_name)
        and str(class_name).strip().casefold() in review_values
    ):
        warnings.append(
            (
                "B5_CLASS_VALUE_REQUIRES_REVIEW",
                f"Giá trị lớp R={str(class_name).strip()!r} là giá trị đặc biệt; "
                "được giữ nguyên, không tự chuẩn hóa.",
            )
        )

    school_name = member.fields.get(
        mapping.approved_member_fields["school_name"]["source"]
    )
    if _has_value(school_name) and re.fullmatch(
        r"\d{4}-\d{4}", str(school_name).strip()
    ):
        warnings.append(
            (
                "B6_SCHOOL_NAME_LOOKS_LIKE_YEAR",
                "T có dạng chuỗi năm thay vì tên trường; được giữ nguyên, "
                "không tự sửa.",
            )
        )

    graduation_level = member.fields.get(
        mapping.approved_member_fields["graduation_level"]["source"]
    )
    graduation_year = member.fields.get(
        mapping.approved_member_fields["graduation_year"]["source"]
    )
    if _has_value(graduation_level) != _has_value(graduation_year):
        warnings.append(
            (
                "B7_GRADUATION_PAIR_INCOMPLETE",
                "Chỉ có một trong V (bậc tốt nghiệp) và X (năm tốt nghiệp); "
                "giá trị có sẵn được giữ nguyên.",
            )
        )
    return warnings


def _write_labelled(sheet, field_mapping: dict[str, str], value: str) -> None:
    cell = sheet.Range(field_mapping["cell"])
    _write_labelled_cell(cell, field_mapping["label"], value, field_mapping["cell"])


def _write_labelled_cell(cell, label: str, value: str, field_name: str) -> None:
    _ensure_not_formula(cell, field_name)
    original = str(cell.Value2 or "")
    if not original.startswith(label):
        raise ExportConfigurationError(f"Template label differs at {field_name}")
    cell.Value2 = f"{label} {value}"


def _ensure_not_formula(cell, address: str) -> None:
    if cell.HasFormula:
        raise ExportConfigurationError(
            f"Refusing to overwrite a formula at template cell {address}"
        )


def _write_preserving_source_value(cell, value) -> None:
    if not isinstance(value, str):
        cell.Value2 = value
        return
    number_format = cell.NumberFormat
    try:
        cell.NumberFormat = "@"
        cell.Value2 = value
    finally:
        cell.NumberFormat = number_format


def _content_after_label(value: str, label: str) -> str:
    if not value.startswith(label):
        raise ExportConfigurationError(f"Template label differs at {label}")
    return value[len(label):].strip()


def _replace_text_range(cell, start: int, length: int, value: str) -> None:
    _character_range(cell, start, length).Text = value


def _character_range(cell, start: int, length: int):
    import pythoncom
    import win32com.client

    characters_id = cell._oleobj_.GetIDsOfNames("Characters")
    characters_dispatch = cell._oleobj_.Invoke(
        characters_id,
        0,
        pythoncom.DISPATCH_PROPERTYGET,
        True,
        start,
        length,
    )
    return win32com.client.Dispatch(characters_dispatch)


def _household_code(household: Household) -> str:
    if household.household_code is not None and str(household.household_code).strip():
        return str(household.household_code).strip()
    first_source = household.members[0].source_record.source
    return f"THIEU_{first_source.sheet}_{first_source.row}"


def _normalized_commune(value: str) -> str:
    return " ".join(value.split()).casefold()


def _safe_filename_component(value: str) -> str:
    safe = INVALID_FILENAME_CHARACTERS.sub("_", value).strip(" .")
    safe = safe.rstrip(". ")
    if not safe:
        safe = "khong_ma"
    if safe.split(".", 1)[0].upper() in {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{index}" for index in range(1, 10)),
        *(f"LPT{index}" for index in range(1, 10)),
    }:
        safe = f"_{safe}"
    return safe


def _temporary_path(directory: Path, suffix: str) -> Path:
    return directory / f".pcgd-{uuid.uuid4().hex}{suffix}"


def _has_value(value) -> bool:
    return value is not None and bool(str(value).strip())


def _remove_if_exists(path: Path | None) -> None:
    if path is not None and path.exists():
        path.unlink()
