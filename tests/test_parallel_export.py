import json
import os
from pathlib import Path
import tempfile
import unittest
from dataclasses import replace

from pypdf import PdfReader
from openpyxl import load_workbook
from scripts.verify_gui_batch_pdf import legible_fixture
from src.export_workflow import export_excel_batch, export_pdf_batch
from tests.test_excel_exporter import TEMPLATE, MAPPING, COMMUNE


class ParallelExportTests(unittest.TestCase):
    def test_checked_parallel_pdf_rejects_changed_workbooks(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            households = [replace(legible_fixture(f"CHECK-{i}", 1), source_file=str(root / f"source{i}.xlsx")) for i in (1, 2)]
            batch = export_excel_batch(households, TEMPLATE, root, MAPPING, commune_name=COMMUNE, workers=2)
            self.assertEqual(batch.report.failure_count, 0)
            for item in batch.accepted:
                workbook = load_workbook(item.path)
                workbook.active["B9"] = "Changed"
                workbook.save(item.path)
                workbook.close()
            checked = export_pdf_batch(batch, root, workers=2, validate=True)
            self.assertEqual(len(checked.failures), 2)
            self.assertEqual(checked.successes, ())
            self.assertTrue(all("Excel changed" in f.error for f in checked.failures))
            unchecked = export_pdf_batch(batch, root, workers=2, validate=False)
            self.assertEqual(unchecked.failures, ())
            self.assertEqual(unchecked.total_pages, 2)

    def test_two_processes_overlap_preserve_names_and_continue_after_failed_source(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            first = replace(legible_fixture("SAME", 8), source_file=str(root / "a/to.xlsx"))
            bad = replace(legible_fixture("DUP", 1), source_file=str(root / "b/bad.xlsx"))
            last = replace(legible_fixture("SAME", 1), source_file=str(root / "c/to.xlsx"))
            progress = []
            batch = export_excel_batch([last, bad, first, bad], TEMPLATE, root, MAPPING, commune_name=COMMUNE, workers=2, progress=lambda d, t, m: progress.append(d / max(1, t)))
            self.assertEqual(batch.report.success_count, 2, batch.report.failures)
            self.assertEqual(batch.report.failure_count, 2)
            self.assertEqual([g.directory.name for g in batch.groups], ["to", "bad", "to_2"])
            self.assertTrue(progress)
            self.assertEqual(progress[-1], 1)
            self.assertEqual(progress, sorted(progress))
            executions = [json.loads((a.path.parent / "EXPORT_REPORT.json").read_text(encoding="utf-8"))["execution"] for a in batch.accepted]
            self.assertEqual(len({r["worker_pid"] for r in executions}), 2)
            self.assertTrue(all(r["worker_pid"] != os.getpid() for r in executions))
            self.assertLess(max(r["started_ns"] for r in executions), min(r["finished_ns"] for r in executions))
            pdf = export_pdf_batch(batch, root, validate=False, workers=2)
            self.assertEqual(pdf.failures, ())
            self.assertEqual(pdf.total_pages, 3)
            self.assertEqual([p.name for p in pdf.merged_paths], ["to.pdf", "to_2.pdf"])
            self.assertEqual([len(PdfReader(p).pages) for p in pdf.merged_paths], [2, 1])
