"""Staged PDF recovery: real single household -> boundaries -> authorized 38."""
from pathlib import Path
import argparse
from datetime import datetime
import json
import hashlib
import sys
import traceback

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from analysis import analyze_sources
from settings import load_settings
from src.export_workflow import export_excel_batch, export_pdf_batch, household_order, ExcelBatch, AcceptedExcel, fingerprint
from types import SimpleNamespace
from scripts.verify_gui_batch_pdf import legible_fixture


def main(stage, gate, existing_excel=None):
    settings = load_settings(ROOT / "config/settings.json")
    mapping = ROOT / "config/template_mapping.json"
    protected = (settings.source_path, settings.template_path, mapping)
    hashes = lambda: {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    before = hashes()
    if stage == "batch":
        if not gate:
            raise ValueError("A successful boundary report is required before the 38-household run")
        prior = json.loads(Path(gate).read_text(encoding="utf-8"))
        if prior.get("stage") != "boundaries" or prior.get("failures") != 0 or prior.get("counts") != {"CASE-01": 1, "CASE-07": 1, "CASE-08": 2, "CASE-14": 2, "CASE-15": 3} or prior.get("protected_after") != before:
            raise ValueError("Boundary gate has not passed or protected inputs changed")
    if stage == "boundaries":
        households = [legible_fixture(f"CASE-{n:02}", n) for n in (1, 7, 8, 14, 15)]
        commune = households[0].commune_name
        expected = {"CASE-01": 1, "CASE-07": 1, "CASE-08": 2, "CASE-14": 2, "CASE-15": 3}
    else:
        analysis = analyze_sources([settings.source_path], settings)
        if analysis.read_result.errors:
            raise ValueError(str(analysis.read_result.errors))
        households = analysis.grouping_result.households
        if stage == "one":
            households = [h for h in households if str(h.household_code).strip() == "TNg_QTRI_T18_12"]
            if len(households) != 1:
                raise ValueError("The requested real household was not uniquely found")
        elif len(households) != 38:
            raise ValueError(f"Expected authorized 38 households, found {len(households)}")
        commune = settings.commune_name
        expected = {str(h.household_code): (len(h.members) + 6) // 7 for h in households}
    output = ROOT / "output/pdf_recovery" / (datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + stage)
    output.mkdir(parents=True, exist_ok=False)
    report = {"stage": stage, "protected_before": before}
    try:
        if existing_excel:
            directory = Path(existing_excel).resolve()
            manifest = json.loads((directory / "EXPORT_REPORT.json").read_text(encoding="utf-8"))
            by_code = {str(h.household_code): h for h in households}
            if manifest["failures"] or {f["household_code"] for f in manifest["files"]} != set(by_code):
                raise ValueError("Existing Excel manifest does not match all loaded households")
            accepted = []
            for f in manifest["files"]:
                path = directory / f["file"]
                household = by_code[f["household_code"]]
                if path.parent != directory or fingerprint(path) != f["sha256"] or len(household.members) != f["members"]:
                    raise ValueError("Existing Excel changed or members differ")
                accepted.append(AcceptedExcel(household, path, f["sha256"]))
            excel = ExcelBatch(SimpleNamespace(success_count=len(accepted), failures=()), tuple(accepted), settings.template_path, mapping, directory)
        else:
            excel = export_excel_batch(households, settings.template_path, output, mapping, commune_name=commune)
        pdf = export_pdf_batch(excel, output, progress=lambda done, total, message: print(message, flush=True))
        report.update(excel_success=excel.report.success_count, excel_failures=[vars(f) for f in excel.report.failures], pdf_success=len(pdf.successes), pdf_failures=[vars(f) for f in pdf.failures], merge_error=pdf.merge_error, pages=pdf.total_pages, counts={p.household_code: p.pages for p in pdf.successes}, expected_counts=expected, merged=str(pdf.merged_path) if pdf.merged_path else None, excel_directory=str(excel.directory), pdf_order=[p.household_code for p in pdf.successes], expected_order=[str(h.household_code) for h in sorted(households, key=household_order)])
        report["failures"] = len(excel.report.failures) + len(pdf.failures) + int(bool(pdf.merge_error))
        if report["counts"] != expected or report["pdf_order"] != report["expected_order"]:
            report["failures"] += 1
    except Exception:
        report.update(failures=1, traceback=traceback.format_exc())
    report["protected_after"] = hashes()
    if before != report["protected_after"]:
        report["failures"] += 1
    path = output / "RECOVERY.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("REPORT:", path)
    return report["failures"]


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("stage", choices=("one", "boundaries", "batch"))
    parser.add_argument("--gate")
    parser.add_argument("--existing-excel")
    args = parser.parse_args()
    raise SystemExit(bool(main(args.stage, args.gate, args.existing_excel)))
