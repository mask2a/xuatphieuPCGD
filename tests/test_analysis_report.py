import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from analysis import analyze_sources
from settings import load_settings
from warning_report import write_warning_report


class AnalysisReportTests(unittest.TestCase):
    def test_settings_do_not_supply_a_hardcoded_commune_or_school_year(self):
        settings = load_settings(Path(__file__).resolve().parents[1] / "config" / "settings.json")

        self.assertEqual(settings.commune_name, "")
        self.assertIsNone(settings.school_year)

    def test_verified_fixture_analysis_and_warning_report(self):
        fixture = Path(__file__).resolve().parents[1] / "dulieumau.xls"
        if not fixture.exists():
            self.skipTest("Verified source workbook is not present")
        with tempfile.TemporaryDirectory() as directory:
            report_path = Path(directory) / "CANH_BAO_DU_LIEU.xlsx"
            result = analyze_sources([fixture])
            write_warning_report(result, report_path)

            self.assertEqual(len(result.read_result.files_read), 1)
            self.assertEqual(len(result.members), 427)
            self.assertEqual(len(result.grouping_result.households), 120)
            self.assertTrue(report_path.exists())
            workbook = load_workbook(report_path, read_only=True, data_only=True)
            try:
                summary = dict(workbook["Tong quan"].iter_rows(min_row=2, values_only=True))
                self.assertEqual(summary["Tổng số tệp đọc"], 1)
                self.assertEqual(summary["Tổng số bản ghi"], 427)
                self.assertEqual(summary["Tổng số hộ"], 120)
                household_rows = list(workbook["So ho"].iter_rows(min_row=2, values_only=True))
                self.assertEqual(len(household_rows), 120)
                self.assertEqual(sum(row[3] for row in household_rows), 427)
                warnings = list(workbook["Canh bao"].iter_rows(min_row=2, values_only=True))
                codes = [row[1] for row in warnings]
                self.assertEqual(codes.count("HOUSEHOLD_HEAD_NAME_CONFLICT"), 3)
                self.assertEqual(codes.count("HOUSEHOLD_RESIDENCE_CONFLICT"), 5)
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
