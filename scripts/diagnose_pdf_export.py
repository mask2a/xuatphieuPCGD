"""Direct COM probe: no GUI, no merge, no PageSetup setters, no workbook save."""
import argparse
import hashlib
import json
from pathlib import Path
import traceback
import sys
from datetime import datetime
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def diagnose(workbook_path, output, map_paper_size_off=False):
    import pythoncom
    import win32com.client
    from pypdf import PdfReader
    source = Path(workbook_path).resolve()
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    pdf = output / (source.stem + ".pdf")
    if pdf.exists():
        raise FileExistsError(pdf)
    log = {"workbook": str(source), "pdf": str(pdf), "sheets": [], "errors": [], "stage": "start", "sha256_before": hashlib.sha256(source.read_bytes()).hexdigest()}
    excel = book = None
    pythoncom.CoInitialize()
    try:
        log["stage"] = "create_excel"
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        log["map_paper_size"] = excel.MapPaperSize
        if map_paper_size_off:
            excel.MapPaperSize = False
        log["map_paper_size_during_export"] = excel.MapPaperSize
        log["stage"] = "open_workbook_readonly"
        book = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        log["sheet_count"] = book.Worksheets.Count
        for sheet in book.Worksheets:
            log["stage"] = "inspect_sheet"
            setup = sheet.PageSetup
            record = {"name": sheet.Name, "visible": sheet.Visible, "print_area": setup.PrintArea, "paper_size": setup.PaperSize, "orientation": setup.Orientation, "zoom": setup.Zoom, "fit_wide": setup.FitToPagesWide, "fit_tall": setup.FitToPagesTall, "margins": [setup.LeftMargin, setup.RightMargin, setup.TopMargin, setup.BottomMargin], "hidden_rows": [i for i in range(1, 88) if sheet.Rows.Item(i).Hidden]}
            log["sheets"].append(record)
            record["print_width_pt"] = sheet.Range(setup.PrintArea).Width
            record["print_height_pt"] = sheet.Range(setup.PrintArea).Height
            record["printable_a4_width_pt"] = 297 * 72 / 25.4 - setup.LeftMargin - setup.RightMargin
            record["printable_a4_height_pt"] = 210 * 72 / 25.4 - setup.TopMargin - setup.BottomMargin
            if sheet.Visible == -1:
                sheet.Activate()
                try:
                    record["excel_expected_pages"] = int(excel.ExecuteExcel4Macro("GET.DOCUMENT(50)"))
                except Exception:
                    record["page_count_traceback"] = traceback.format_exc()
        book.Worksheets.Item(1).Activate()
        log["stage"] = "ExportAsFixedFormat"
        book.ExportAsFixedFormat(0, str(pdf), 0, True, False, OpenAfterPublish=False)
        log["com_export_success"] = True
        for sheet, record in zip(book.Worksheets, log["sheets"]):
            if sheet.Visible == -1:
                sheet.Activate()
                record["excel_pages_after_export"] = int(excel.ExecuteExcel4Macro("GET.DOCUMENT(50)"))
    except Exception:
        log["errors"].append({"stage": log["stage"], "traceback": traceback.format_exc()})
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                log["errors"].append({"stage": "close_workbook", "traceback": traceback.format_exc()})
        if excel is not None:
            try:
                excel.MapPaperSize = log["map_paper_size"]
                excel.Quit()
            except Exception:
                log["errors"].append({"stage": "quit_excel", "traceback": traceback.format_exc()})
        book = excel = None
        pythoncom.CoUninitialize()
    log["pdf_exists"] = pdf.is_file()
    log["pdf_bytes"] = pdf.stat().st_size if pdf.is_file() else 0
    if pdf.is_file():
        try:
            reader = PdfReader(pdf, strict=True)
            log["pdf_encrypted"] = reader.is_encrypted
            log["pdf_readable"] = not reader.is_encrypted
            if not reader.is_encrypted:
                log["pdf_actual_pages"] = len(reader.pages)
                log["pdf_page_sizes"] = [[float(p.mediabox.width), float(p.mediabox.height)] for p in reader.pages]
                log["pdf_text_nonempty"] = [bool((p.extract_text() or "").strip()) for p in reader.pages]
        except Exception:
            log["pdf_readable"] = False
            log["errors"].append({"stage": "read_pdf", "traceback": traceback.format_exc()})
    log["current_program_expected_pages"] = sum(s["visible"] == -1 for s in log["sheets"])
    log["sha256_after"] = hashlib.sha256(source.read_bytes()).hexdigest()
    log["matches_program_page_count"] = log.get("pdf_actual_pages") == log["current_program_expected_pages"]
    if pdf.is_file():
        try:
            from src.export_workflow import validate_pdf
            validate_pdf(pdf, log["current_program_expected_pages"])
            log["program_validation"] = "PASS"
        except Exception:
            log["program_validation"] = "FAIL"
            log["errors"].append({"stage": "validate_pdf_actual_page_count", "traceback": traceback.format_exc()})
    path = output / "DIAGNOSIS.json"
    path.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=2))
    print("LOG:", path)
    return log


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook")
    parser.add_argument("--output", default="output/diagnostics/pdf_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    parser.add_argument("--map-paper-size-off", action="store_true")
    args = parser.parse_args()
    result = diagnose(args.workbook, args.output, args.map_paper_size_off)
    raise SystemExit(0 if result.get("matches_program_page_count") and not result["errors"] else 1)
