import tempfile
import unittest
import hashlib
import json
import shutil
from pathlib import Path

from openpyxl import load_workbook

from household_grouper import group_households
from normalizer import normalize_records
from reader import RawRecord, SourceLocation
from settings import Settings
from src.excel_exporter import ExportConfigurationError, export_households


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "maubaocao.xls"
MAPPING = ROOT / "config" / "template_mapping.json"
COMMUNE = "Xã Kiểm thử"


def make_household(
    code,
    count,
    *,
    missing_birth_year_for=(),
    head_phone="0912345678",
    non_head_phones=(),
    head_names=None,
    residences=None,
    unmapped_fields=None,
    commune_name=COMMUNE,
):
    records = []
    for index in range(count):
        family_name = f"Họ{index + 1}"
        given_name = f"Tên{index + 1}"
        birth_year = None if index in missing_birth_year_for else 2010 + index
        head_name = (head_names or ["Chủ Hộ"])[min(index, len(head_names or ["Chủ Hộ"]) - 1)]
        residence = (residences or ["Thường trú"])[min(index, len(residences or ["Thường trú"]) - 1)]
        relationship = "Chủ hộ" if index == 0 else "Thành viên"
        phone = head_phone if index == 0 else (
            non_head_phones[index - 1]
            if index - 1 < len(non_head_phones)
            else None
        )
        values = {
            "B": family_name,
            "C": given_name,
            "D": index + 1,
            "E": index + 2,
            "F": birth_year,
            "G": "X" if index % 2 == 0 else None,
            "H": "Kinh",
            "I": None,
            "K": head_name,
            "L": None,
            "M": None,
            "N": code,
            "O": residence,
            "P": None,
            "Q": None,
            "R": None,
            "S": None,
            "T": None,
            "AU": relationship,
            "AV": "Cha mẹ",
            "AW": phone,
        }
        values.update(unmapped_fields or {})
        displays = {
            column: "" if value is None else str(value)
            for column, value in values.items()
        }
        displays["D"] = f"{index + 1:02}"
        displays["E"] = f"{index + 2:02}"
        records.append(
            RawRecord(
                values=values,
                display_values=displays,
                source=SourceLocation(
                    "synthetic-source.xls", "MauNhapLieu", index + 5
                ),
            )
        )
    members = normalize_records(records, Settings(commune_name=commune_name))
    return group_households(members).households[0]


def write_export_settings(directory, *, commune_name=COMMUNE, school_year="2025-2026"):
    path = Path(directory) / "settings.json"
    path.write_text(
        json.dumps(
            {"commune_name": commune_name, "school_year": school_year},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return path


class ExcelExporterTests(unittest.TestCase):
    def test_exports_short_full_and_supplementary_households_from_real_template(self):
        if not TEMPLATE.exists():
            self.skipTest("The verified source template is not present")
        households = tuple(
            make_household(f"{count:02}", count)
            for count in (1, 6, 7, 14, 15)
        )
        template_hash_before = hashlib.sha256(TEMPLATE.read_bytes()).digest()
        source_path = ROOT / "dulieumau.xls"
        source_hash_before = hashlib.sha256(source_path.read_bytes()).digest()

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "excel"
            report = export_households(
                households,
                TEMPLATE,
                output,
                MAPPING,
                settings_path=write_export_settings(directory),
            )

            self.assertEqual(report.success_count, 5)
            self.assertEqual(report.failure_count, 0)
            self.assertEqual(report.summary["success_count"], 5)
            self.assertEqual(report.summary["failure_count"], 0)
            self.assertGreater(report.summary["warning_count"], 0)
            self.assertTrue(
                all(
                    "SCHOOL_YEAR_ROWS_NOT_MAPPED"
                    in {warning.code for warning in item.warnings}
                    for item in report.successes
                )
            )
            self.assertTrue(
                all(
                    any(
                        "2025-2026" in warning.message
                        for warning in item.warnings
                        if warning.code == "SCHOOL_YEAR_ROWS_NOT_MAPPED"
                    )
                    for item in report.successes
                )
            )
            self.assertEqual(
                {item.member_count for item in report.successes},
                {1, 6, 7, 14, 15},
            )
            self.assertEqual(
                hashlib.sha256(TEMPLATE.read_bytes()).digest(),
                template_hash_before,
            )
            self.assertEqual(
                hashlib.sha256(source_path.read_bytes()).digest(),
                source_hash_before,
            )

            by_code = {str(item.household_code): item for item in report.successes}
            for code, count in (("01", 1), ("06", 6), ("07", 7), ("14", 14), ("15", 15)):
                exported = by_code[code]
                self.assertTrue(exported.path.is_file())
                workbook = load_workbook(exported.path, data_only=False)
                try:
                    sheet = workbook["Phiếu điều tra 1 mặt"]
                    self.assertEqual(sheet["A1"].value, f"Xã (1): {COMMUNE}")
                    self.assertEqual(sheet["Z1"].value, f"Số phiếu (4): {code}")
                    self.assertEqual(
                        sheet["Z3"].value, "Điện thoại (6): 0912345678"
                    )
                    self.assertEqual(
                        sheet["AA82"].value.splitlines()[0],
                        "XÁC NHẬN CỦA UBND XÃ " + COMMUNE,
                    )
                    self.assertEqual(len(sheet.merged_cells.ranges), 442)
                    if count <= 6:
                        first_hidden = (9, 14, 19, 24, 29, 34, 39)[count]
                        self.assertTrue(sheet.row_dimensions[first_hidden].hidden)
                        self.assertEqual(sheet.print_area, "'Phiếu điều tra 1 mặt'!$A$1:$AG$87")
                        self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
                    else:
                        self.assertFalse(sheet.row_dimensions[39].hidden)
                    self.assertIn("Họ1 Tên1", sheet["B9"].value)
                    self.assertEqual(sheet["D11"].value, "01/02/2010")
                    self.assertEqual(sheet["B10"].value, "QH với chủ hộ: Chủ hộ")
                    self.assertEqual(sheet["D12"].value, "DT: Kinh")
                    self.assertEqual(sheet["B13"].value, "Cha, mẹ, Ng đỡ đầu: Cha mẹ")
                    self.assertEqual(sheet["A1"].font.name, "Arial")
                    self.assertTrue(sheet["A1"].font.bold)
                    self.assertEqual(sheet["C12"].font.name, "Wingdings")
                    self.assertTrue(sheet["B9"].border.bottom.style)
                    self.assertTrue(sheet["AA82"].value.startswith("XÁC NHẬN CỦA UBND"))
                    self.assertFalse(workbook._external_links)
                    for member_index in range(count):
                        page_index, slot_index = divmod(member_index, 14)
                        member_sheet = workbook.worksheets[page_index]
                        start_row = (9, 14, 19, 24, 29, 34, 39, 47, 52, 57, 62, 67, 72, 77)[slot_index]
                        self.assertEqual(
                            member_sheet[f"B{start_row}"].value,
                            f"Họ{member_index + 1} Tên{member_index + 1}",
                        )
                    if count == 15:
                        self.assertEqual(len(workbook.worksheets), 2)
                        supplement = workbook.worksheets[1]
                        self.assertEqual(supplement["A9"].value, 15)
                        self.assertEqual(supplement["B9"].value, "Họ15 Tên15")
                        self.assertTrue(supplement.row_dimensions[14].hidden)
                finally:
                    workbook.close()

    def test_missing_and_conflicting_values_are_not_inferred_or_silently_selected(self):
        households = (
            make_household(
                "NO-PHONE",
                2,
                missing_birth_year_for=(0,),
                head_phone=None,
                non_head_phones=("0900111222",),
            ),
            make_household(
                "CONFLICT",
                2,
                head_names=("Chủ Một", "Chủ Hai"),
                residences=("Thường trú", "Tạm trú"),
                unmapped_fields={
                    "J": "Diện ưu tiên",
                    "M": "Tổ 7",
                    "Q": "1",
                    "AX": "Ghi chú",
                },
            ),
        )

        with tempfile.TemporaryDirectory() as directory:
            report = export_households(
                households,
                TEMPLATE,
                Path(directory),
                MAPPING,
                settings_path=write_export_settings(directory),
            )

            self.assertEqual(report.success_count, 2)
            self.assertEqual(report.failure_count, 0)
            by_code = {str(item.household_code): item for item in report.successes}
            missing = load_workbook(by_code["NO-PHONE"].path, data_only=False)
            conflict = load_workbook(by_code["CONFLICT"].path, data_only=False)
            try:
                missing_sheet = missing["Phiếu điều tra 1 mặt"]
                self.assertEqual(
                    missing_sheet["Z3"].value.replace("Điện thoại (6):", "").strip("_ "),
                    "",
                )
                self.assertEqual(missing_sheet["D11"].value, "01/02/")
                self.assertNotIn(
                    "HOUSEHOLD_PHONE_CONFLICT",
                    {warning.code for warning in by_code["NO-PHONE"].warnings},
                )
                conflict_sheet = conflict["Phiếu điều tra 1 mặt"]
                self.assertTrue(conflict_sheet["H3"].value.endswith("_" * 10))
                self.assertIn("q Thường trú", conflict_sheet["Z2"].value)
                conflict_codes = {
                    warning.code for warning in by_code["CONFLICT"].warnings
                }
                self.assertIn("HOUSEHOLD_HEAD_NAME_CONFLICT", conflict_codes)
                self.assertIn("HOUSEHOLD_RESIDENCE_CONFLICT", conflict_codes)
                self.assertIn("UNMAPPED_ADDRESS", conflict_codes)
                self.assertIn("UNMAPPED_GRADE_GROUP", conflict_codes)
                self.assertIn("UNMAPPED_NOTES", conflict_codes)
                self.assertIn("UNMAPPED_PRIORITY_CATEGORY", conflict_codes)
            finally:
                missing.close()
                conflict.close()

    def test_invalid_filename_characters_are_sanitized_and_existing_files_are_not_overwritten(self):
        first = make_household("A/B", 1)
        second = make_household("A:B", 1)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            target = output / "ho_A_B.xlsx"
            report = export_households(
                (first, second),
                TEMPLATE,
                output,
                MAPPING,
                settings_path=write_export_settings(directory),
            )

            self.assertEqual(report.success_count, 1)
            self.assertEqual(report.failure_count, 1)
            self.assertEqual(report.summary["success_count"], 1)
            self.assertEqual(report.summary["failure_count"], 1)
            self.assertTrue(target.is_file())
            original_result = target.read_bytes()
            self.assertIn("already exists", report.failures[0].error)

            repeated = export_households(
                (first,),
                TEMPLATE,
                output,
                MAPPING,
                settings_path=write_export_settings(directory),
            )

            self.assertEqual(repeated.success_count, 0)
            self.assertEqual(repeated.failure_count, 1)
            self.assertEqual(target.read_bytes(), original_result)

    def test_refuses_to_overwrite_a_formula_in_a_mapped_template_cell(self):
        import pythoncom
        import win32com.client

        household = make_household("FORMULA", 1)
        with tempfile.TemporaryDirectory() as directory:
            modified_template = Path(directory) / "formula-template.xls"
            shutil.copyfile(TEMPLATE, modified_template)
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = None
            try:
                workbook = excel.Workbooks.Open(
                    str(modified_template), UpdateLinks=0, ReadOnly=False
                )
                cell = workbook.Worksheets.Item(1).Range("B9")
                cell.NumberFormat = "General"
                cell.Formula = "=1+1"
                workbook.Save()
                workbook.Close(SaveChanges=False)
                workbook = None
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
                excel.Quit()
                excel = None
                pythoncom.CoUninitialize()

            report = export_households(
                (household,),
                modified_template,
                Path(directory) / "output",
                MAPPING,
                settings_path=write_export_settings(directory),
            )

            self.assertEqual(report.success_count, 0)
            self.assertEqual(report.failure_count, 1)
            self.assertIn("formula", report.failures[0].error.lower())
            self.assertFalse((Path(directory) / "output" / "ho_FORMULA.xlsx").exists())

    def test_export_requires_a_configured_commune_instead_of_using_the_template_default(self):
        household = make_household("NO-COMMUNE", 1, commune_name="")

        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ExportConfigurationError):
                export_households(
                    (household,),
                    TEMPLATE,
                    Path(directory) / "output",
                    MAPPING,
                    settings_path=write_export_settings(directory),
                )
            self.assertFalse((Path(directory) / "output").exists())

    def test_export_rejects_missing_commune_and_school_year_before_creating_output(self):
        household = make_household("NO-CONFIG", 1)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "not-created"
            with self.assertRaisesRegex(
                ExportConfigurationError,
                "commune_name.*school_year",
            ):
                export_households(
                    (household,),
                    TEMPLATE,
                    output,
                    MAPPING,
                )
            self.assertFalse(output.exists())

    def test_export_rejects_invalid_school_year_and_commune_mismatch(self):
        household = make_household("BAD-CONFIG", 1)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "not-created"
            with self.assertRaisesRegex(
                ExportConfigurationError,
                "school_year.*YYYY-YYYY",
            ):
                export_households(
                    (household,),
                    TEMPLATE,
                    output,
                    MAPPING,
                    settings_path=write_export_settings(
                        directory,
                        school_year="đồng ý",
                    ),
                )
            self.assertFalse(output.exists())

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "not-created"
            with self.assertRaisesRegex(
                ExportConfigurationError,
                "does not match configured commune_name",
            ):
                export_households(
                    (make_household("OTHER-COMMUNE", 1, commune_name="Xã Khác"),),
                    TEMPLATE,
                    output,
                    MAPPING,
                    settings_path=write_export_settings(directory),
                )
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
