import tempfile
import tkinter as tk
import time
import unittest
import threading
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from gui import (
    PCGDApp,
    export_one_household,
    filter_households,
)
from settings import load_settings
from tests.test_excel_exporter import TEMPLATE


class GuiTests(unittest.TestCase):
    def test_parallel_worker_count_is_forwarded_and_saved(self):
        self.assertEqual(self.app.export_workers.get(), "2")
        self.app.output_path.set(self.temporary_directory.name)
        self.app.template_path.set(str(TEMPLATE))
        self.app.commune_name.set("Test")
        self.app._analysis_commune = "Test"
        self.app.households = (object(),)
        self.app.analysis_result = object()
        report = SimpleNamespace(success_count=1, failure_count=0, failures=(), warnings=())
        batch = SimpleNamespace(report=report, accepted=(object(),), directory=Path(self.temporary_directory.name))
        pdf = SimpleNamespace(successes=(), failures=(), total_pages=0, merged_path=None, merge_error=None)
        with patch("gui.BATCH_FEATURES_VERIFIED", True), patch("gui.export_excel_batch", return_value=batch) as excel, patch("gui.export_pdf_batch", return_value=pdf) as exporter:
            for count in (1, 2, 4):
                self.app.export_workers.set(str(count))
                self.app.export_all_excel()
                self.assertEqual(str(self.app.export_workers_input["state"]), "disabled")
                self.pump_until_idle()
                self.assertEqual(excel.call_args.kwargs["workers"], count)
                self.app.export_all_pdf()
                self.pump_until_idle()
                self.assertEqual(exporter.call_args.kwargs["workers"], count)
                self.app.save_configuration()
                self.assertEqual(load_settings(self.settings_path).export_workers, count)
            calls = excel.call_count
            self.app.export_workers.set("5")
            self.app.export_all_excel()
            self.assertEqual(excel.call_count, calls)
            self.assertFalse(self.app._busy)

    def test_pdf_validation_checkbox_is_forwarded_and_saved(self):
        self.assertFalse(self.app.pdf_validation.get())
        root = Path(self.temporary_directory.name)
        self.app.output_path.set(str(root))
        self.app.commune_name.set("Xã kiểm thử")
        self.app._analysis_commune = "Xã kiểm thử"
        self.app.households = (object(),)
        self.app.analysis_result = object()
        report = SimpleNamespace(success_count=1, failure_count=0)
        self.app._last_excel_batch = SimpleNamespace(report=report, accepted=(object(),))
        result = SimpleNamespace(successes=(), failures=(), merged_path=None, total_pages=0, merge_error=None)
        with patch("gui.BATCH_FEATURES_VERIFIED", True), patch("gui.export_pdf_batch", return_value=result) as exporter:
            for selected in (True, False):
                self.app.pdf_validation.set(selected)
                self.app.export_all_pdf()
                self.assertEqual(str(self.app.pdf_validation_check["state"]), "disabled")
                self.pump_until_idle()
                self.assertEqual(str(self.app.pdf_validation_check["state"]), "normal")
                self.assertIs(exporter.call_args.kwargs["validate"], selected)
                self.app.save_configuration()
                self.assertIs(load_settings(self.settings_path).pdf_validation, selected)

    def test_folder_label_and_multiple_pdf_summaries_are_connected_to_gui(self):
        root = Path(self.temporary_directory.name)
        self.app.source_path.set(str(root))
        self.app._analysis_source = root
        self.app.output_path.set(str(root))
        self.app.template_path.set(str(TEMPLATE))
        self.app.commune_name.set("Xã kiểm thử")
        self.app._analysis_commune = "Xã kiểm thử"
        self.app.analysis_result = object()
        first = SimpleNamespace(source_file=str(root / "to1.xlsx"), household_code="H1")
        second = SimpleNamespace(source_file=str(root / "to2.xlsx"), household_code="H1")
        self.app.households = (first,)
        report = SimpleNamespace(success_count=1, failure_count=0, failures=(), warnings=())
        batch = SimpleNamespace(report=report, accepted=(object(),), directory=root)
        with patch("gui.BATCH_FEATURES_VERIFIED", True), patch("gui.export_excel_batch", return_value=batch) as exporter:
            self.app.export_all_excel()
            self.pump_until_idle()
            self.assertEqual(exporter.call_args.kwargs["source_name"], root.name)
            self.app.households = (first, second)
            self.app.export_all_excel()
            self.pump_until_idle()
            self.assertIsNone(exporter.call_args.kwargs["source_name"])
            pdf = SimpleNamespace(successes=(), failures=(), total_pages=2, merged_path=None, merged_paths=(root / "to1.pdf", root / "to2.pdf"), merge_error=None)
            with patch("gui.export_pdf_batch", return_value=pdf):
                self.app.export_all_pdf()
                self.pump_until_idle()
            log = self.app.log_text.get("1.0", "end")
            self.assertIn("to1.pdf", log)
            self.assertIn("to2.pdf", log)

    def pump_until_idle(self):
        deadline = time.monotonic() + 10
        while self.app._busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.root.update()
        self.assertFalse(self.app._busy)

    def test_batch_buttons_run_on_worker_disable_concurrency_and_show_results(self):
        household = SimpleNamespace(household_code="TEST", members=(object(),))
        self.app.households = (household,)
        self.app.analysis_result = object()
        self.app.commune_name.set("Xã kiểm thử")
        self.app._analysis_commune = "Xã kiểm thử"
        self.app.output_path.set(self.temporary_directory.name)
        self.app.template_path.set(str(TEMPLATE))
        excel_report = SimpleNamespace(success_count=1, failure_count=0, failures=(), warnings=())
        excel_batch = SimpleNamespace(report=excel_report, accepted=(object(),), directory=Path(self.temporary_directory.name))
        pdf_report = SimpleNamespace(successes=(SimpleNamespace(household_code="TEST", pages=1, path=Path("test.pdf")),), failures=(), total_pages=1, merged_path=Path("TongHop.pdf"), merge_error=None)
        gate = threading.Event()
        threads = []

        def excel_work(*args, **kwargs):
            threads.append(threading.get_ident())
            gate.wait(5)
            kwargs["progress"](1, 1, "Excel hoàn tất")
            return excel_batch

        with patch("gui.BATCH_FEATURES_VERIFIED", True), patch("gui.export_excel_batch", side_effect=excel_work) as exporter, patch("gui.export_pdf_batch", return_value=pdf_report) as pdf:
            self.app._refresh_export_buttons()
            self.assertEqual(str(self.app.export_all_button["state"]), "normal")
            self.assertEqual(str(self.app.export_pdf_button["state"]), "disabled")
            self.app.export_all_button.invoke()
            self.assertTrue(self.app._busy)
            self.assertEqual(str(self.app.export_all_button["state"]), "disabled")
            self.app.export_all_excel()
            self.app.export_all_pdf()
            pdf.assert_not_called()
            gate.set()
            self.pump_until_idle()
            exporter.assert_called_once()
            self.assertNotEqual(threads[0], threading.get_ident())
            self.assertEqual(str(self.app.export_pdf_button["state"]), "normal")
            self.app.export_pdf_button.invoke()
            self.pump_until_idle()
            pdf.assert_called_once()
            self.assertIn("Tổng hợp: 1 trang", self.app.export_stats.get())
            self.assertIn("TongHop.pdf", self.app.log_text.get("1.0", "end"))

    def test_batch_rejects_changed_commune_and_worker_failure_releases_busy(self):
        self.app.households = (object(),)
        self.app.analysis_result = object()
        self.app._analysis_commune = "Cũ"
        self.app.commune_name.set("Mới")
        with patch("gui.BATCH_FEATURES_VERIFIED", True), patch("gui.export_excel_batch") as exporter:
            self.app.export_all_excel()
            exporter.assert_not_called()
        self.app._start_background("Failure", lambda: (_ for _ in ()).throw(RuntimeError("TEST ERROR")), lambda _: None)
        self.pump_until_idle()
        self.assertIn("TEST ERROR", self.app.log_text.get("1.0", "end"))

    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary_directory.cleanup)
        self.addCleanup(self.root.destroy)
        self.settings_path = Path(self.temporary_directory.name) / "settings.json"
        self.settings_path.write_text(
            '{"commune_name":null,"school_year":"2026-2027"}',
            encoding="utf-8",
        )
        self.mapping_path = (
            Path(__file__).resolve().parents[1]
            / "config"
            / "template_mapping.json"
        )
        self.app = PCGDApp(
            self.root,
            settings_path=self.settings_path,
            mapping_path=self.mapping_path,
            show_dialogs=False,
        )

    def test_window_opens_without_starting_excel_and_bulk_actions_are_disabled(self):
        self.root.update()

        self.assertEqual(self.root.title(), "PCGD HỘ GIA ĐÌNH")
        self.assertFalse(self.app._busy)
        self.assertEqual(str(self.app.export_all_button["state"]), "disabled")
        self.assertEqual(str(self.app.export_pdf_button["state"]), "disabled")
        self.assertEqual(str(self.app.export_selected_button["state"]), "disabled")

    def test_search_matches_household_code_and_head_name(self):
        households = (
            SimpleNamespace(
                household_code="MA-001",
                household_head_name="Nguyen Van A",
            ),
            SimpleNamespace(
                household_code="MA-002",
                household_head_name="Tran Thi B",
            ),
        )

        self.assertEqual(
            filter_households(households, "ma-001"),
            (households[0],),
        )
        self.assertEqual(
            filter_households(households, "TRAN THI"),
            (households[1],),
        )
        self.assertEqual(filter_households(households, "khong co"), ())
        self.assertEqual(filter_households(households, ""), households)

    def test_save_configuration_persists_paths_locality_and_school_year(self):
        self.app.commune_name.set("Xã kiểm thử")
        self.app.school_year.set("2027-2028")
        self.app.source_path.set("source.xlsx")
        self.app.template_path.set("template.xls")
        self.app.output_path.set("result")

        self.app.save_configuration()
        saved = load_settings(self.settings_path)

        self.assertEqual(saved.commune_name, "Xã kiểm thử")
        self.assertEqual(saved.school_year, "2027-2028")
        self.assertEqual(
            saved.source_path,
            self.settings_path.parent.parent / "source.xlsx",
        )
        self.assertEqual(
            saved.template_path,
            self.settings_path.parent.parent / "template.xls",
        )
        self.assertEqual(
            saved.output_directory,
            self.settings_path.parent.parent / "result",
        )

    def test_export_is_blocked_when_commune_is_missing(self):
        self.app.template_path.set(
            str(Path(__file__).resolve().parents[1] / "maubaocao.xls")
        )
        self.app.output_path.set(self.temporary_directory.name)
        self.app.analysis_result = object()
        self.app._analysis_commune = ""
        household = SimpleNamespace(household_code="TEST-1", members=(object(),))
        self.app._household_by_id["test-household"] = household
        self.app.household_tree.insert("", "end", iid="test-household")
        self.app.household_tree.selection_set("test-household")

        self.app.export_selected()

        self.assertFalse(self.app._busy)
        self.assertIn("Thiếu xã/phường", self.app.log_text.get("1.0", "end"))

    def test_gui_reads_verified_source_and_updates_household_statistics(self):
        source = Path(__file__).resolve().parents[1] / "dulieumau.xls"
        if not source.is_file():
            self.skipTest("Verified source workbook is not present")
        self.app.source_path.set(str(source))

        self.app.read_data()
        deadline = time.monotonic() + 30
        while self.app._busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.02)
        self.root.update()

        self.assertFalse(self.app._busy)
        self.assertEqual(self.app.member_total.get(), "427")
        self.assertEqual(self.app.household_total.get(), "120")

    def test_single_household_adapter_uses_official_batch_output(self):
        household = object()
        report = object()
        with tempfile.TemporaryDirectory() as directory:
            with patch("gui.load_settings", return_value=SimpleNamespace(commune_name="Xã kiểm thử")), patch("gui.export_excel_batch", return_value=report) as exporter:
                result = export_one_household(
                    household,
                    "template.xls",
                    directory,
                    "mapping.json",
                    settings_path="settings.json",
                )

        self.assertIs(result, report)
        args, kwargs = exporter.call_args
        self.assertEqual(args[0], (household,))
        self.assertEqual(args[1], "template.xls")
        self.assertEqual(args[2], directory)
        self.assertEqual(args[3], "mapping.json")
        self.assertEqual(kwargs["commune_name"], "Xã kiểm thử")
