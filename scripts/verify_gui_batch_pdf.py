"""Persist a synthetic batch using exactly the workflow connected to the GUI."""
from pathlib import Path
import hashlib
import json
import sys
import re
from datetime import datetime
from dataclasses import replace
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pymupdf
from openpyxl import load_workbook
from src.export_workflow import export_excel_batch, export_pdf_batch, ExcelBatch, AcceptedExcel, fingerprint
from tests.test_approved_mapping import full_household
from tests.test_excel_exporter import ROOT, TEMPLATE, MAPPING, COMMUNE
from src.approved_mapping import FIELDS, DISABILITIES


def legible_fixture(code, count):
    household = full_household(code, count)
    sample = dict(zip((item[0] for item in FIELDS.values()), ("SC", "25", "5", "24", "6", "24", "1", "2", "0", "X", "A", "Đến", "KT")))
    sample.update({title: "X" if title == "Khuyết tật khác" else "" for title in DISABILITIES})
    members = []
    for index, member in enumerate(household.members):
        record = member.source_record
        raw = {**record.values, **{column: sample[title] for column, title in record.headers.items()}}
        day, month, year = index % 28 + 1, index % 12 + 1, 2010 + index
        raw.update({"D": day, "E": month, "F": year, "I": "Không", "Q": "10", "R": "10A1", "T": "THPT Kiểm thử", "V": "THCS", "X": "2024"})
        education = {"grade_group": "10", "class_name": "10A1", "school_name": "THPT Kiểm thử", "graduation_level": "THCS", "graduation_year": "2024", "religion": "Không"}
        members.append(replace(member, source_record=replace(record, values=raw), fields={**member.fields, **education}, normalized_fields={**member.normalized_fields, **education}, birth_components=(f"{day:02}", f"{month:02}", str(year))))
    return replace(household, members=tuple(members))


def main(existing_excel=None):
    output = Path(existing_excel).resolve() if existing_excel else ROOT / "output/testing" / ("gui_batch_pdf_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    protected = [ROOT / "dulieumau.xls", TEMPLATE, MAPPING]
    before = [hashlib.sha256(p.read_bytes()).hexdigest() for p in protected]
    households = [legible_fixture(f"DEMO-{count:02}", count) for count in (21, 15, 1, 8, 14, 7)]
    if existing_excel:
        directory = next((output / "excel").iterdir())
        accepted = tuple(AcceptedExcel(h, directory / f"ho_{h.household_code}.xlsx", fingerprint(directory / f"ho_{h.household_code}.xlsx")) for h in sorted(households, key=lambda h: int(h.household_code[-2:])))
        excel = ExcelBatch(SimpleNamespace(success_count=len(accepted), failure_count=0, failures=()), accepted, TEMPLATE, MAPPING, directory)
    else:
        excel = export_excel_batch(households, TEMPLATE, output, MAPPING, commune_name=COMMUNE)
    pdf = export_pdf_batch(excel, output)
    if excel.report.failure_count or pdf.failures or pdf.merge_error or pdf.total_pages != 12:
        raise RuntimeError(str((excel.report.failures, pdf.failures, pdf.merge_error)))
    previews = output / "previews"
    previews.mkdir(exist_ok=True)
    metrics = []
    for item in pdf.successes:
        accepted = next(a for a in excel.accepted if a.household.household_code == item.household_code)
        workbook = load_workbook(accepted.path)
        try:
            source_font_sizes = [sheet["B9"].font.sz for sheet in workbook]
        finally:
            workbook.close()
        with pymupdf.open(item.path) as document:
            for index, page in enumerate(document):
                page.get_pixmap(matrix=pymupdf.Matrix(1.3, 1.3)).save(previews / f"{item.path.stem}_page{index + 1}.png")
                text = page.get_text()
                assert "☒" in text and "☐" in text, "Checkbox glyphs missing"
                assert not re.search(r"(?<!\w)q(?!\w)", text) and "■" not in text and "�" not in text
                bounds = pymupdf.Rect()
                for shape in page.get_drawings():
                    bounds |= shape["rect"]
                ratio = bounds.width / (page.rect.width - 2 * 72 * 0.5 / 2.54)
                center_error = abs((bounds.x0 + bounds.x1) / 2 - page.rect.width / 2)
                names = [span["size"] for block in page.get_text("dict")["blocks"] if "lines" in block for line in block["lines"] for span in line["spans"] if re.search(r"Họ\d+", span["text"])]
                assert ratio > .75 and center_error < 2, "Table width/centering differs"
                metrics.append({"household": item.household_code, "page": index + 1, "width_ratio": ratio, "center_error_pt": center_error, "member_font_pt": names[0] if names else None, "source_font_pt": source_font_sizes[index], "physical_scale": names[0] / source_font_sizes[index] if names else None, "checked": text.count("☒"), "unchecked": text.count("☐")})
    sizes = [m["member_font_pt"] for m in metrics if m["member_font_pt"] is not None]
    assert max(sizes) - min(sizes) < .1, "Inconsistent physical font scale"
    after = [hashlib.sha256(p.read_bytes()).hexdigest() for p in protected]
    assert before == after, "Source, template or mapping changed"
    result = {"synthetic_only": True, "excel_success": excel.report.success_count, "pdf_success": len(pdf.successes), "failures": 0, "pages": pdf.total_pages, "merged": str(pdf.merged_path), "excel_directory": str(excel.directory), "protected_unchanged": True, "pdf_metrics": metrics}
    (output / "VERIFICATION.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(str(output))
    print(json.dumps(result))


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--reuse-excel", help="Revalidate synthetic Excel files from a previous QA run before retrying PDF")
    main(parser.parse_args().reuse_excel)
