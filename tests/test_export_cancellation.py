import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from src.excel_exporter import export_households
from src.export_workflow import export_excel_batch
from src.export_parallel import SourceJob, run_source_jobs
from tests.test_excel_exporter import make_household, MAPPING, COMMUNE, write_export_settings


class ExportCancellationTests(unittest.TestCase):
    def test_parallel_stop_skips_pending_sources(self):
        cancel = threading.Event()
        cancel.set()
        with tempfile.TemporaryDirectory() as root:
            jobs = [SourceJob("excel", (), Path(root), Path(root) / str(i)) for i in range(3)]
            self.assertEqual(run_source_jobs(jobs, 2, cancel_event=cancel), ())

    def test_gui_stop_releases_busy_state_after_worker_cleanup(self):
        from gui import PCGDApp
        import queue
        app = PCGDApp.__new__(PCGDApp)
        app._busy = True
        app._closing = False
        app._cancel_event = threading.Event()
        app.stop_button = MagicMock()
        app.progress_text = MagicMock()
        app.progress = MagicMock()
        app._log = MagicMock()
        app._refresh_export_buttons = MagicMock()
        app._events = queue.Queue()
        app.stop_export()
        self.assertTrue(app._busy)
        self.assertTrue(app._cancel_event.is_set())
        app._events.put(("finished", None))
        app._drain_events()
        self.assertFalse(app._busy)
        app._refresh_export_buttons.assert_called_once()

    def test_stop_after_failed_household_closes_excel_and_skips_remaining(self):
        cancel = threading.Event()
        excel = MagicMock()
        template_book = excel.Workbooks.Open.return_value

        def fail_and_stop(*args):
            cancel.set()
            raise ValueError("Saved workbook confirmation does not contain the configured commune")

        template_book.SaveCopyAs.side_effect = fail_and_stop
        messages = []
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "template.xls"
            template.touch()
            settings = write_export_settings(root)
            with patch("win32com.client.DispatchEx", return_value=excel), patch(
                "src.excel_exporter._verify_template_sheet"
            ), patch("src.excel_exporter._template_signature"):
                report = export_households(
                    [make_household("A", 1), make_household("B", 1)],
                    template, root / "out", MAPPING,
                    settings_path=settings, cancel_event=cancel,
                    progress=lambda done, total, message: messages.append(message),
                )
            self.assertEqual(report.failure_count, 1)
            self.assertEqual(report.failures[0].household_code, "A")
            self.assertTrue(any("LỖI Excel hộ A:" in message and "configured commune" in message for message in messages))
            template_book.SaveCopyAs.assert_called_once()
            template_book.Close.assert_called_once()
            excel.Quit.assert_called_once()
            self.assertEqual(list((root / "out").iterdir()), [])

    def test_cancelled_multi_source_batch_does_not_start_next_source(self):
        cancel = threading.Event()
        cancel.set()
        households = [SimpleNamespace(household_code=str(i), source_file=f"source{i}.xls") for i in range(2)]
        with tempfile.TemporaryDirectory() as root, patch("src.export_workflow.export_households") as exporter:
            batch = export_excel_batch(households, "template.xls", root, MAPPING,
                                       commune_name=COMMUNE, cancel_event=cancel)
            exporter.assert_not_called()
            self.assertEqual(batch.accepted, ())


if __name__ == "__main__":
    unittest.main()
