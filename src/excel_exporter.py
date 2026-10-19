"""Export each household from an untouched copy of the verified Excel template."""

from __future__ import annotations

import logging
import math
import os
import re
import gc
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from household_grouper import Household
from settings import load_settings, validate_export_settings
from src.template_mapper import TemplateMapping, load_template_mapping


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

        for household in household_list:
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
                sheets = _populate_workbook(workbook, household, mapping)
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
                            if len(household.members) <= mapping.short_form_max_members
                            else "full"
                        ),
                        supplementary_sheet_count=max(
                            0,
                            math.ceil(
                                len(household.members)
                                / len(mapping.member_start_rows)
                            ) - 1,
                        ),
                        warnings=tuple(warnings),
                    )
                )
            except Exception as error:
                LOGGER.error("Failed exporting household %s: %s", code, error)
                failures.append(ExportFailure(code, str(error)))
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
) -> tuple[object, ...]:
    sheet_count = max(
        1,
        math.ceil(len(household.members) / len(mapping.member_start_rows)),
    )
    sheets = [workbook.Worksheets.Item(mapping.template_sheet)]
    for page_number in range(2, sheet_count + 1):
        sheets[0].Copy(None, workbook.Worksheets.Item(workbook.Worksheets.Count))
        page = workbook.ActiveSheet
        page.Name = f"BoSung_{page_number}"
        sheets.append(page)

    for page_index, sheet in enumerate(sheets):
        first_member_index = page_index * len(mapping.member_start_rows)
        page_members = household.members[
            first_member_index:first_member_index + len(mapping.member_start_rows)
        ]
        _populate_header(sheet, household, mapping)
        _populate_member_page(sheet, page_members, first_member_index, mapping)
        _configure_page(sheet, len(page_members), mapping)
    return tuple(sheets)


def _populate_header(sheet, household: Household, mapping: TemplateMapping) -> None:
    commune_field = mapping.fixed_fields["commune_name"]
    commune_cell = sheet.Range(commune_field["cell"])
    _ensure_not_formula(commune_cell, commune_field["cell"])
    original_commune = _content_after_label(
        str(commune_cell.Value2 or ""),
        commune_field["label"],
    )
    if not household.commune_name.strip():
        raise ExportConfigurationError("commune_name must be configured before export")
    commune_cell.Value2 = f'{commune_field["label"]} {household.commune_name.strip()}'

    confirmation = sheet.Range(mapping.confirmation_commune_cell)
    _ensure_not_formula(confirmation, mapping.confirmation_commune_cell)
    confirmation_text = str(confirmation.Value2 or "")
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
    for slot_index, member in enumerate(members):
        start_row = mapping.member_start_rows[slot_index]
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
    sheet.PageSetup.PrintArea = mapping.print_area
    if member_count <= mapping.short_form_max_members:
        sheet.Rows(
            f"{mapping.member_start_rows[member_count]}:{mapping.confirmation_start_row - 1}"
        ).Hidden = True
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1


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
        if not str(cell.Value2 or "").startswith(field["label"]):
            raise ExportConfigurationError(
                f"Template label differs from mapping at {field['cell']}"
            )
    for slot_number, start_row in enumerate(mapping.member_start_rows, start=1):
        if sheet.Range(f"A{start_row}").Value2 != slot_number:
            raise ExportConfigurationError(
                f"Template member numbering differs at row {start_row}"
            )
    for field_name, field in mapping.member_fields.items():
        address = f'{field["column"]}{mapping.member_start_rows[0] + int(field["row_offset"])}'
        merge_area = sheet.Range(address).MergeArea.Address.replace("$", "")
        if merge_area != field["merge_area"]:
            raise ExportConfigurationError(
                f"Template merge area differs for {field_name}: {merge_area}"
            )
    return sheet


def _template_signature(
    template_path: Path,
    sheet,
    mapping: TemplateMapping,
) -> TemplateSignature:
    import xlrd

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
            f'{mapping.fixed_fields["commune_name"]["label"]} {household.commune_name.strip()}'
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
            _validate_sheet_format(sheet, signature)
            page_members = household.members[
                page_index * len(mapping.member_start_rows):
                (page_index + 1) * len(mapping.member_start_rows)
            ]
            if (
                len(page_members) <= mapping.short_form_max_members
                and (
                    sheet.PageSetup.Zoom is not False
                    or sheet.PageSetup.FitToPagesWide != 1
                    or sheet.PageSetup.FitToPagesTall != 1
                )
            ):
                raise ValueError("Short-form worksheet does not fit to one printed page")
            if (
                len(page_members) > mapping.short_form_max_members
                and sheet.PageSetup.Zoom != mapping.full_form_zoom
            ):
                raise ValueError(
                    "Full-form worksheet scale differs from the verified template"
                )
            for slot_index, member in enumerate(page_members):
                start_row = mapping.member_start_rows[slot_index]
                if sheet.Range(f"B{start_row}").Value2 != member.member_name:
                    raise ValueError(
                        "Saved workbook does not contain member "
                        f"{page_index * len(mapping.member_start_rows) + slot_index + 1}"
                    )
        if workbook.LinkSources(1) is not None:
            raise ValueError("Saved workbook contains external workbook links")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
            workbook = None


def _com_cell_format(cell) -> tuple:
    font = cell.Font
    borders = tuple(
        (
            cell.Borders.Item(index).LineStyle,
            cell.Borders.Item(index).Weight,
            cell.Borders.Item(index).ColorIndex,
        )
        for index in (7, 8, 9, 10)
    )
    return (
        font.Name,
        font.Size,
        font.Bold,
        font.Italic,
        cell.NumberFormat,
        cell.HorizontalAlignment,
        cell.VerticalAlignment,
        cell.WrapText,
        borders,
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


def _validate_sheet_format(sheet, signature: TemplateSignature) -> None:
    for address, expected in signature.cell_formats.items():
        actual = _com_cell_format(sheet.Range(address))
        if actual != expected:
            raise ValueError(
                f"Saved workbook formatting differs at {address}: "
                f"expected {expected!r}, got {actual!r}"
            )
    for row, expected_height in enumerate(signature.row_heights, start=1):
        if sheet.Rows.Item(row).Hidden:
            continue
        if abs(float(sheet.Rows.Item(row).RowHeight) - expected_height) > 0.05:
            raise ValueError(f"Saved workbook row height differs at row {row}")
    for column, expected_width in enumerate(signature.column_widths, start=1):
        if abs(float(sheet.Columns.Item(column).ColumnWidth) - expected_width) > 0.05:
            raise ValueError(f"Saved workbook column width differs at column {column}")
    if _com_page_setup(sheet) != signature.page_setup:
        raise ValueError("Saved workbook page setup differs from the verified template")


def _household_export_warnings(
    household: Household,
    mapping: TemplateMapping,
    school_year: str,
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
            f"Năm học cấu hình {school_year}: {mapping.school_year_policy}",
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
    if any(member.residence for member in household.members):
        warnings.append(
            ExportWarning(
                code,
                "HOUSEHOLD_RESIDENCE_NOT_EXPORTED",
                mapping.unmapped_household_fields["residence"],
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


def _content_after_label(value: str, label: str) -> str:
    if not value.startswith(label):
        raise ExportConfigurationError(f"Template label differs at {label}")
    return value[len(label):].strip()


def _replace_text_range(cell, start: int, length: int, value: str) -> None:
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
    win32com.client.Dispatch(characters_dispatch).Text = value


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
