import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from pypdf import PdfReader
from scripts.verify_gui_batch_pdf import legible_fixture
from src.export_workflow import export_excel_batch, export_pdf_batch
from tests.test_excel_exporter import ROOT, MAPPING, COMMUNE

TEMPLATE = ROOT / "tempplates/maubaocao.xls"


def source_household(path, code="H1"):
    return replace(legible_fixture(code, 1), source_file=str(path))


class SourceOutputTests(unittest.TestCase):
    def test_each_input_has_its_own_excel_pdf_and_root_summary_even_with_same_household_code(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            households = [source_household(root / "to2.xlsx"), source_household(root / "to1.xlsx")]
            batch = export_excel_batch(households, TEMPLATE, root, MAPPING, commune_name=COMMUNE)
            self.assertEqual(batch.report.failure_count, 0)
            self.assertEqual({a.path.parent.relative_to(root).as_posix() for a in batch.accepted}, {"excel/to1", "excel/to2"})
            pdf = export_pdf_batch(batch, root)
            self.assertEqual(pdf.failures, ())
            self.assertEqual({p.path.parent.relative_to(root).as_posix() for p in pdf.successes}, {"pdf/to1", "pdf/to2"})
            self.assertEqual({p.name for p in pdf.merged_paths}, {"to1.pdf", "to2.pdf"})
            self.assertTrue(all(p.parent == root for p in pdf.merged_paths))
            self.assertEqual([len(PdfReader(root / name).pages) for name in ("to1.pdf", "to2.pdf")], [1, 1])
            original = {p: p.read_bytes() for p in pdf.merged_paths}
            again = export_pdf_batch(batch, root)
            self.assertEqual(again.failures, ())
            self.assertEqual({p.name for p in again.merged_paths}, {"to1_2.pdf", "to2_2.pdf"})
            self.assertEqual({p.path.parent.name for p in again.successes}, {"to1_2", "to2_2"})
            self.assertTrue(all(p.read_bytes() == content for p, content in original.items()))

    def test_folder_name_override_and_repeat_export_do_not_overwrite(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            households = [source_household(root / "to1.xlsx")]
            first = export_excel_batch(households, TEMPLATE, root, MAPPING, commune_name=COMMUNE, source_name="DuLieu")
            self.assertEqual(first.accepted[0].path.parent, root / "excel/DuLieu")
            original = first.accepted[0].path.read_bytes()
            second = export_excel_batch(households, TEMPLATE, root, MAPPING, commune_name=COMMUNE, source_name="DuLieu")
            self.assertEqual(second.accepted[0].path.parent, root / "excel/DuLieu_2")
            self.assertEqual(first.accepted[0].path.read_bytes(), original)

    def test_equal_filenames_in_different_source_folders_do_not_mix(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            households = [source_household(root / "a/to1.xlsx"), source_household(root / "b/to1.xlsx")]
            batch = export_excel_batch(households, TEMPLATE, root, MAPPING, commune_name=COMMUNE)
            self.assertEqual(batch.report.failure_count, 0)
            self.assertEqual({a.path.parent.name for a in batch.accepted}, {"to1", "to1_2"})
            self.assertEqual(len({a.path for a in batch.accepted}), 2)
