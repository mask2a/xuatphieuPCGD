import unittest
from types import SimpleNamespace

from src.approved_mapping import template_label
from src.excel_exporter import _write_labelled_cell


class TemplateLabelTests(unittest.TestCase):
    def test_editable_labels_are_preserved_when_filling(self):
        for address, label in (("A1", "Xã (1):"), ("A1", "Phường(1):"),
                               ("A3", "Xóm (2):"), ("A3", "Địa chỉ (2):")):
            with self.subTest(label=label):
                cell = SimpleNamespace(Value2=label + "_____", HasFormula=False)
                sheet = SimpleNamespace(Range=lambda _: cell)
                _write_labelled_cell(cell, template_label(sheet, address), "Nội dung mới", address)
                self.assertEqual(cell.Value2, label + " Nội dung mới")

    def test_missing_label_separator_is_rejected(self):
        for text in ("", "Không có nhãn", ":___"):
            sheet = SimpleNamespace(Range=lambda _: SimpleNamespace(Value2=text))
            with self.assertRaises(ValueError):
                template_label(sheet, "A1")
