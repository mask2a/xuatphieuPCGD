import unittest
from dataclasses import replace

from src.print_layout import invalid_birth_date, normalize_geometry, PAGE_CAPACITY, PRINT_ZOOM
from unittest.mock import MagicMock
from src.excel_exporter import _household_export_warnings
from src.template_mapper import load_template_mapping
from tests.test_approved_mapping import full_household
from tests.test_excel_exporter import MAPPING


class PrintLayoutTests(unittest.TestCase):
    def test_seven_slots_fit_without_modifying_row_or_column_dimensions(self):
        sheet = MagicMock()
        mapping = load_template_mapping(MAPPING)
        normalize_geometry(sheet, 1, mapping)
        sheet.Rows.assert_called_once_with("44:81")
        sheet.Rows.Item.assert_not_called()
        sheet.Columns.Item.assert_not_called()
        self.assertEqual(PAGE_CAPACITY, 7)
        self.assertEqual(sheet.PageSetup.Zoom, PRINT_ZOOM)
        self.assertIs(PRINT_ZOOM, False)
        with self.assertRaises(ValueError):
            normalize_geometry(sheet, 8, mapping)

    def test_invalid_calendar_dates_warn_without_changing_source(self):
        household = full_household("DATE-TEST", 1)
        member = household.members[0]
        self.assertFalse(invalid_birth_date(replace(member, birth_components=("29", "02", "2024"))))
        self.assertFalse(invalid_birth_date(replace(member, birth_components=("29", "02", "2024.0"))))
        self.assertTrue(invalid_birth_date(replace(member, birth_components=("29", "02", "2024.5"))))
        bad = replace(member, birth_components=("31", "04", "2024"))
        self.assertTrue(invalid_birth_date(bad))
        source_before = dict(bad.source_record.values)
        warnings = _household_export_warnings(replace(household, members=(bad,)), load_template_mapping(MAPPING), "2025-2026", True)
        self.assertIn("INVALID_BIRTH_DATE", {w.code for w in warnings})
        self.assertEqual(dict(bad.source_record.values), source_before)
        self.assertEqual(bad.birth_components, ("31", "04", "2024"))
