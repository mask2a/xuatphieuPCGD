import shutil
import tempfile
import unittest
from pathlib import Path

from analysis import analyze_sources
from reader import REQUIRED_HEADERS
from settings import Settings
from src.school_year_labels import (
    YEAR_LABEL_CELLS,
    build_school_year_labels,
    update_school_year_labels,
)


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "maubaocao.xls"
SOURCE = ROOT / "dulieumau.xls"
MEMBER_START_ROWS = (9, 14, 19, 24, 29, 34, 39, 47, 52, 57, 62, 67, 72, 77)
SIGNATURE_AREAS = ("F82:I82", "J82:M82", "N82:Q82", "R82:V82", "W82:Z82")


class SchoolYearLabelTests(unittest.TestCase):
    def test_builds_five_labels_from_the_configured_first_year(self):
        self.assertEqual(
            build_school_year_labels("2026-2027"),
            (
                "2026-2027",
                "2027-2028",
                "2028-2029",
                "2029-2030",
                "2030-2031",
            ),
        )
        self.assertEqual(
            build_school_year_labels("2027-2028"),
            (
                "2027-2028",
                "2028-2029",
                "2029-2030",
                "2030-2031",
                "2031-2032",
            ),
        )

    def test_rejects_malformed_nonconsecutive_and_out_of_range_school_years(self):
        for school_year in (
            "2026/2027",
            "2026-2028",
            "9995-9996",
            "9999-10000",
            "n/a",
        ):
            with self.subTest(school_year=school_year):
                with self.assertRaises(ValueError):
                    build_school_year_labels(school_year)

    def test_verified_allowlist_contains_seventy_unique_explicit_cells(self):
        self.assertEqual(len(YEAR_LABEL_CELLS), 70)
        self.assertEqual(len(set(YEAR_LABEL_CELLS)), 70)
        self.assertEqual(YEAR_LABEL_CELLS[0], "F9")
        self.assertEqual(YEAR_LABEL_CELLS[-1], "F81")
        self.assertNotIn("F44", YEAR_LABEL_CELLS)
        self.assertNotIn("F82", YEAR_LABEL_CELLS)

    def test_template_has_required_structure_independent_of_historical_hash(self):
        import pythoncom
        import win32com.client
        from src.excel_exporter import _verify_template_sheet, _template_signature
        from src.template_mapper import load_template_mapping

        mapping = load_template_mapping(ROOT / "config/template_mapping.json")
        pythoncom.CoInitialize()
        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            workbook = excel.Workbooks.Open(str(TEMPLATE), UpdateLinks=0, ReadOnly=True, AddToMru=False)
            sheet = _verify_template_sheet(workbook, mapping)
            with self.assertLogs("src.excel_exporter", level="INFO") as logs:
                signature = _template_signature(TEMPLATE, sheet, mapping)
            self.assertIn("SHA-256 (informational only)", logs.output[0])
            self.assertEqual(signature.merge_count, 442)
            self.assertEqual(len(signature.row_heights), 87)
            self.assertEqual(len(signature.column_widths), 33)
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()

    def test_source_has_required_structure_and_can_be_grouped(self):
        # Source content is user-managed. Validate it through the real reader
        # (including required headers), not historical hash/mtime metadata.
        settings = Settings()
        result = analyze_sources([SOURCE], settings)
        self.assertEqual(result.read_result.errors, ())
        self.assertEqual(result.read_result.files_read, (SOURCE.resolve(),))
        records = result.read_result.records
        self.assertTrue(records, "Source must contain readable member records")
        for record in records:
            self.assertEqual(record.source.sheet, settings.reader.preferred_sheet)
            self.assertTrue(set(REQUIRED_HEADERS).issubset(record.values))

        households = result.grouping_result.households
        self.assertTrue(households, "Source must be groupable into households")
        grouped_rows = [
            member.source_record.source
            for household in households
            for member in household.members
        ]
        self.assertCountEqual(grouped_rows, [record.source for record in records])

    @unittest.skipUnless(TEMPLATE.is_file(), "The verified template is not present")
    def test_updates_only_verified_labels_and_preserves_workbook_structure(self):
        import pythoncom
        import win32com.client

        with tempfile.TemporaryDirectory() as directory:
            for school_year in ("2026-2027", "2027-2028"):
                with self.subTest(school_year=school_year):
                    self._exercise_com_copy(
                        Path(directory),
                        school_year,
                        pythoncom,
                        win32com.client,
                    )

    @unittest.skipUnless(TEMPLATE.is_file(), "The verified template is not present")
    def test_refuses_to_update_when_a_verified_merge_area_is_wrong(self):
        import pythoncom
        import win32com.client

        with tempfile.TemporaryDirectory() as directory:
            copied_template = Path(directory) / "wrong-merge.xls"
            shutil.copy2(TEMPLATE, copied_template)
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = None
            try:
                workbook = excel.Workbooks.Open(
                    str(copied_template),
                    UpdateLinks=0,
                    ReadOnly=False,
                    AddToMru=False,
                )
                worksheet = workbook.Worksheets.Item(1)
                original_values = {
                    address: worksheet.Range(address).Value2
                    for address in YEAR_LABEL_CELLS
                }
                worksheet.Range("F9:G9").UnMerge()
                from src.excel_exporter import _verify_template_sheet, ExportConfigurationError
                from src.template_mapper import load_template_mapping
                with self.assertRaises(ExportConfigurationError):
                    _verify_template_sheet(workbook, load_template_mapping(ROOT / "config/template_mapping.json"))
                with self.assertRaisesRegex(ValueError, "Unexpected merge area"):
                    update_school_year_labels(workbook, "2026-2027")
                self.assertEqual(
                    {
                        address: worksheet.Range(address).Value2
                        for address in YEAR_LABEL_CELLS
                    },
                    original_values,
                )
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
                excel.Quit()
                pythoncom.CoUninitialize()

    def _exercise_com_copy(self, directory, school_year, pythoncom, client):
        copied_template = directory / f"template_{school_year}.xls"
        output_path = directory / f"school_year_{school_year}.xlsx"
        shutil.copy2(TEMPLATE, copied_template)

        pythoncom.CoInitialize()
        excel = client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(
                str(copied_template),
                UpdateLinks=0,
                ReadOnly=False,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            worksheet = workbook.Worksheets.Item(1)
            # Give Excel a concrete print scope on this temporary copy. With an
            # empty PrintArea its automatic-break Count can include an invalid
            # item. Keep comparing all breaks and PageSetup before/after.
            worksheet.PageSetup.PrintArea = worksheet.UsedRange.Address
            before = _workbook_snapshot(worksheet)
            labels = update_school_year_labels(workbook, school_year)
            self.assertEqual(labels, build_school_year_labels(school_year))
            after_update = _workbook_snapshot(worksheet)
            self.assertEqual(
                {
                    address
                    for address in before["values"]
                    if before["values"][address] != after_update["values"][address]
                },
                set(YEAR_LABEL_CELLS),
            )
            for start_row in MEMBER_START_ROWS:
                self.assertEqual(
                    tuple(
                        worksheet.Cells(start_row + offset, 6).Value2
                        for offset in range(5)
                    ),
                    labels,
                )
                self.assertEqual(
                    tuple(
                        worksheet.Cells(start_row + offset, 6).MergeArea.Address.replace(
                            "$", ""
                        )
                        for offset in range(5)
                    ),
                    tuple(
                        f"F{start_row + offset}:G{start_row + offset}"
                        for offset in range(5)
                    ),
                )
            workbook.SaveAs(
                str(output_path),
                FileFormat=51,
                CreateBackup=False,
                AddToMru=False,
            )
            workbook.Close(SaveChanges=False)
            workbook = None

            reopened = excel.Workbooks.Open(
                str(output_path),
                UpdateLinks=0,
                ReadOnly=True,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            try:
                result_sheet = reopened.Worksheets.Item(1)
                after_reopen = _workbook_snapshot(result_sheet)
                self.assertEqual(after_update["values"], after_reopen["values"])
                self.assertEqual(before["merges"], after_reopen["merges"])
                self.assertEqual(before["formats"], after_reopen["formats"])
                self.assertEqual(before["row_heights"], after_reopen["row_heights"])
                self.assertEqual(before["column_widths"], after_reopen["column_widths"])
                self.assertEqual(before["page_setup"], after_reopen["page_setup"])
                self.assertEqual(before["horizontal_breaks"], after_reopen["horizontal_breaks"])
                self.assertEqual(before["vertical_breaks"], after_reopen["vertical_breaks"])
                self.assertEqual(before["signature"], after_reopen["signature"])
                self.assertEqual(before["print_area"], after_reopen["print_area"])
                self.assertEqual(before["external_links"], after_reopen["external_links"])
                self.assertEqual(len(after_reopen["merges"]), 442)
                self.assertEqual(after_reopen["page_setup"]["Orientation"], 2.0)
                self.assertEqual(after_reopen["page_setup"]["Zoom"], 68)
                self.assertEqual(after_reopen["page_setup"]["FitToPagesWide"], 1)
                self.assertEqual(after_reopen["page_setup"]["FitToPagesTall"], 1)
                self.assertEqual(after_reopen["print_area"], "")
                self.assertEqual(
                    tuple(
                        result_sheet.Cells(start_row, 6).Value2
                        for start_row in MEMBER_START_ROWS
                    ),
                    (labels[0],) * len(MEMBER_START_ROWS),
                )
                self.assertEqual(
                    _external_link_sources(reopened),
                    None,
                )
            finally:
                reopened.Close(SaveChanges=False)
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()


def _workbook_snapshot(worksheet):
    used_range = worksheet.UsedRange
    first_row = int(used_range.Row)
    first_column = int(used_range.Column)
    value_matrix = used_range.Value2
    values = {}
    for row_offset, row_values in enumerate(value_matrix):
        for column_offset, value in enumerate(row_values):
            address = (
                f"{_column_name(first_column + column_offset)}"
                f"{first_row + row_offset}"
            )
            values[address] = value

    merges = set()
    for row in range(first_row, first_row + used_range.Rows.Count):
        for column in range(first_column, first_column + used_range.Columns.Count):
            cell = worksheet.Cells(row, column)
            if cell.MergeCells and cell.Address == cell.MergeArea.Cells(1, 1).Address:
                merges.add(cell.MergeArea.Address.replace("$", ""))

    formats = {}
    for address in YEAR_LABEL_CELLS:
        cell = worksheet.Range(address)
        formats[address] = (
                cell.Font.Name,
                cell.Font.Size,
                cell.Font.Bold,
                cell.Font.Italic,
                cell.NumberFormat,
                cell.HorizontalAlignment,
                cell.VerticalAlignment,
                cell.WrapText,
                cell.Interior.Color,
                cell.Interior.Pattern,
                tuple(
                    (
                        cell.Borders(index).LineStyle,
                        cell.Borders(index).Weight,
                        cell.Borders(index).Color,
                    )
                    for index in range(7, 13)
                ),
            )
    page_setup = {
        name: getattr(worksheet.PageSetup, name)
        for name in (
            "PrintArea",
            "Zoom",
            "FitToPagesWide",
            "FitToPagesTall",
            "Orientation",
            "PaperSize",
            "Order",
            "CenterHorizontally",
            "CenterVertically",
        )
    }
    return {
        "values": values,
        "merges": frozenset(merges),
        "formats": formats,
        "row_heights": tuple(
            worksheet.Rows(row).RowHeight
            for row in range(1, used_range.Row + used_range.Rows.Count)
        ),
        "column_widths": tuple(
            worksheet.Columns(column).ColumnWidth
            for column in range(1, used_range.Column + used_range.Columns.Count)
        ),
        "page_setup": page_setup,
        "horizontal_breaks": tuple(
            worksheet.HPageBreaks(index).Location.Address.replace("$", "")
            for index in range(1, worksheet.HPageBreaks.Count + 1)
        ),
        "vertical_breaks": tuple(
            worksheet.VPageBreaks(index).Location.Address.replace("$", "")
            for index in range(1, worksheet.VPageBreaks.Count + 1)
        ),
        "signature": tuple(
            (area, worksheet.Range(area).Value2) for area in SIGNATURE_AREAS
        )
        + (("AA82:AG83", worksheet.Range("AA82:AG83").Value2),),
        "print_area": worksheet.PageSetup.PrintArea,
        "external_links": _external_link_sources(worksheet.Parent),
    }


def _column_name(column_number):
    name = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _external_link_sources(workbook):
    return workbook.LinkSources(1)


if __name__ == "__main__":
    unittest.main()
