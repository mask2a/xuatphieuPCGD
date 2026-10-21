"""Repair the rejected household; retain 37 accepted PDFs without overwriting."""
from pathlib import Path
import hashlib
import json
import math
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from analysis import analyze_sources
from settings import load_settings
from src.export_workflow import (
    export_excel_batch, export_pdf_batch, household_order, merge_pdfs,
    new_directory, PdfResult, validate_pdf, validate_pdf_content,
)
from src.template_mapper import load_template_mapping


def main():
    settings = load_settings(ROOT / "config/settings.json")
    previous = json.loads((ROOT / "output/pdf_recovery/20261021_161747_batch/RECOVERY.json").read_text(encoding="utf-8"))
    protected = (settings.source_path, settings.template_path, ROOT / "config/template_mapping.json")
    hashes = lambda: {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    before = hashes()
    if previous["protected_after"] != before:
        raise ValueError("Protected inputs changed; previous PDFs cannot be reused")
    analysis = analyze_sources([settings.source_path], settings)
    if analysis.read_result.errors:
        raise ValueError(str(analysis.read_result.errors))
    households = {h.household_code: h for h in analysis.grouping_result.households}
    code = "TNg_QTRI_T18_ 110"
    if len(households) != 38 or previous["pdf_order"] != [c for c in previous["expected_order"] if c != code]:
        raise ValueError("Prior successes do not match the authorized 37 households")
    root = settings.output_directory
    excel = export_excel_batch([households[code]], settings.template_path, root, protected[2], commune_name=settings.commune_name)
    if excel.report.failures:
        raise ValueError(str(excel.report.failures))
    pdf = export_pdf_batch(excel, root)
    if pdf.failures or pdf.merge_error or len(pdf.successes) != 1:
        raise ValueError(str((pdf.failures, pdf.merge_error)))
    previous_directory = Path(previous["merged"]).parent
    items = [PdfResult(c, previous_directory / ("ho_" + c + ".pdf"), previous["counts"][c]) for c in previous["pdf_order"]]
    items += list(pdf.successes)
    items.sort(key=lambda item: household_order(households[item.household_code]))
    mapping = load_template_mapping(protected[2])
    input_hashes = {}
    for item in items:
        household = households[item.household_code]
        if item.pages != math.ceil(len(household.members) / 7):
            raise ValueError("Unexpected page count")
        validate_pdf_content(validate_pdf(item.path, item.pages), household, mapping)
        input_hashes[str(item.path)] = hashlib.sha256(item.path.read_bytes()).hexdigest()
    directory = new_directory(root, "pdf")
    merged = directory / "TongHop.pdf"
    pages = merge_pdfs(items, merged)
    if pages != 38 or before != hashes():
        raise ValueError("Incorrect total pages or changed protected inputs")
    if any(hashlib.sha256(Path(p).read_bytes()).hexdigest() != digest for p, digest in input_hashes.items()):
        raise ValueError("An accepted PDF changed during merge")
    report = {"success_count": len(items), "failures": [], "total_pages": pages, "merged": str(merged), "repaired_household": code, "repaired_excel": str(excel.accepted[0].path), "repaired_pdf": str(pdf.successes[0].path), "order": [i.household_code for i in items], "input_pdf_hashes": input_hashes, "protected_before": before, "protected_after": hashes()}
    (directory / "PDF_REPORT.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
