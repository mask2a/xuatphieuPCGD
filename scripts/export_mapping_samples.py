"""Export only six synthetic verification households; never read real members."""
import hashlib
import json
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pythoncom
import win32com.client

from src.excel_exporter import export_households
from tests.test_approved_mapping import full_household
from tests.test_excel_exporter import ROOT, TEMPLATE, MAPPING, write_export_settings


def main():
    output = ROOT / "output/testing" / ("approved_mapping_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    protected = [ROOT / "dulieumau.xls", TEMPLATE, ROOT / "config/approved_workbook_snapshot.json"]
    before = {str(p.name): hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    households = [full_household(f"SYNTHETIC-{count:02}", count) for count in (1, 5, 7, 8, 14, 15)]
    with tempfile.TemporaryDirectory() as directory:
        report = export_households(households, TEMPLATE, output, MAPPING, settings_path=write_export_settings(directory))
    if report.failure_count or report.success_count != 6:
        raise RuntimeError(str(report.failures))
    pythoncom.CoInitialize()
    excel = None
    results = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        for item in report.successes:
            workbook = excel.Workbooks.Open(str(item.path), UpdateLinks=0, ReadOnly=True, AddToMru=False)
            try:
                pages = []
                for sheet in workbook.Worksheets:
                    sheet.Activate()
                    count = int(excel.ExecuteExcel4Macro("GET.DOCUMENT(50)"))
                    if count != 1:
                        raise ValueError(f"Expected one printed page per sheet; got {count}")
                    pages.append(count)
                results.append({"file": item.path.name, "members": item.member_count, "printed_pages_by_sheet": pages})
            finally:
                workbook.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    after = {str(p.name): hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    assert before == after, "Protected files changed"
    summary = {"synthetic_only": True, "results": results, "protected_files_unchanged": True}
    (output / "VERIFICATION.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(str(output))
    print(json.dumps(summary))


if __name__ == "__main__":
    main()
