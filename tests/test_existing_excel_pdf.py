import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from pypdf import PdfReader, PdfWriter
from src.export_workflow import export_existing_excel_pdf


class ExistingExcelPdfTests(unittest.TestCase):
    def test_exports_existing_files_without_template_or_source_data(self):
        excel = MagicMock()
        book = excel.Workbooks.Open.return_value

        def render(kind, path, *args, **kwargs):
            writer = PdfWriter()
            writer.add_blank_page(width=842, height=595)
            writer.write(path)
            writer.close()

        book.ExportAsFixedFormat.side_effect = render
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "saved"
            source.mkdir()
            paths = [source / "ho_2.xlsx", source / "ho_1.xls"]
            for path in paths:
                path.write_bytes(b"existing workbook")
            with patch("win32com.client.DispatchEx", return_value=excel):
                report = export_existing_excel_pdf(paths, root / "out")
            self.assertEqual(report.failures, ())
            self.assertEqual([item.household_code for item in report.successes], ["ho_1", "ho_2"])
            self.assertEqual(len(PdfReader(report.merged_path).pages), 2)
            self.assertEqual(excel.Workbooks.Open.call_count, 2)
            self.assertTrue(all(call.kwargs["ReadOnly"] for call in excel.Workbooks.Open.call_args_list))
            self.assertTrue(all(path.read_bytes() == b"existing workbook" for path in paths))
            excel.Quit.assert_called_once()

    def test_cancel_before_start_does_not_open_excel(self):
        cancel = threading.Event()
        cancel.set()
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / "saved.xlsx"
            path.touch()
            with patch("win32com.client.DispatchEx") as dispatch:
                report = export_existing_excel_pdf([path], root, cancel_event=cancel)
            dispatch.assert_not_called()
            self.assertEqual(report.successes, ())

    def test_invalid_file_is_rejected_before_start(self):
        with self.assertRaisesRegex(ValueError, "Excel"):
            export_existing_excel_pdf(["missing.xlsx"], "output")
