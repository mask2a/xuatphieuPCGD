"""Per-source Excel/COM PDF export, optionally parallel, with ordered summaries."""
from dataclasses import dataclass
from pathlib import Path
import hashlib
import json
import math
import os
import re
import tempfile
import uuid
import unicodedata
import logging
from collections import Counter
from types import SimpleNamespace
from settings import export_worker_count

from pypdf import PdfReader, PdfWriter

from src.excel_exporter import (
    export_households, ExportFailure, ExportReport, _verify_template_sheet,
    _template_signature, _validate_saved_workbook,
    _household_code, _safe_filename_component,
)
from src.template_mapper import load_template_mapping
from src.print_layout import page_setup, verify_page_setup, PAGE_CAPACITY
from src.approved_mapping import values as member_print_values, household_address

LOGGER = logging.getLogger(__name__)


def household_order(household):
    code = _household_code(household)
    return tuple((0, int(part)) if part.isdigit() else (1, part.casefold()) for part in re.split(r"(\d+)", code))


def fingerprint(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def new_directory(root, kind):
    directory = Path(root) / kind / ("dot_" + uuid.uuid4().hex)
    directory.mkdir(parents=True, exist_ok=False)
    return directory


@dataclass(frozen=True)
class AcceptedExcel:
    household: object
    path: Path
    sha256: str


@dataclass(frozen=True)
class ExcelBatch:
    report: object
    accepted: tuple[AcceptedExcel, ...]
    template: Path
    mapping: Path
    directory: Path
    source_name: str = ""
    groups: tuple = ()


@dataclass(frozen=True)
class PdfResult:
    household_code: str
    path: Path
    pages: int


@dataclass(frozen=True)
class PdfBatch:
    successes: tuple[PdfResult, ...]
    failures: tuple[ExportFailure, ...]
    merged_path: Path | None
    total_pages: int
    merge_error: str | None = None
    merged_paths: tuple[Path, ...] = ()


def source_directory(root, kind, name):
    """Reserve a readable source folder; never reuse an existing export."""
    root = Path(root)
    base = _safe_filename_component(str(name))[:100]
    number = 1
    while True:
        label = base if number == 1 else f"{base}_{number}"
        directory = root / kind / label
        occupied = directory.exists() or (root / (label + ".pdf")).exists()
        if kind == "excel":
            occupied = occupied or (root / "pdf" / label).exists()
        if not occupied:
            directory.parent.mkdir(parents=True, exist_ok=True)
            try:
                directory.mkdir()
                return directory
            except FileExistsError:
                pass
        number += 1


def group_progress(progress, index, count, label):
    if progress is None:
        return None
    return lambda done, total, message: progress(index * 1000 + min(1000, int(done * 1000 / max(1, total))), count * 1000, f"{label}: {message}")


def export_existing_excel_pdf(paths, output_root, *, progress=None, cancel_event=None):
    """Render existing workbooks using their saved print settings, without editing them."""
    paths = tuple(dict.fromkeys(Path(path).resolve() for path in paths))
    if not paths:
        raise ValueError("Chưa chọn file Excel")
    for path in paths:
        if not path.is_file() or path.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
            raise ValueError(f"File Excel không hợp lệ: {path}")
    sources = {}
    for path in paths:
        sources.setdefault(path.parent, []).append(path)
    groups = []
    for directory, files in sources.items():
        accepted = tuple(AcceptedExcel(SimpleNamespace(household_code=path.stem), path, "") for path in files)
        groups.append(ExcelBatch(ExportReport((), (), ()), accepted, Path(), Path(), directory, source_name=directory.name))
    batch = ExcelBatch(ExportReport((), (), ()), tuple(a for g in groups for a in g.accepted), Path(), Path(), Path(output_root), groups=tuple(groups))
    return export_pdf_batch(batch, output_root, progress=progress, validate=False, cancel_event=cancel_event)


def export_excel_batch(households, template, output_root, mapping, *, commune_name, progress=None, source_name=None, workers=1, _directory=None, cancel_event=None):
    workers = export_worker_count(workers)
    households = tuple(sorted(households, key=household_order))
    sources = {}
    for household in households:
        source = str(Path(getattr(household, "source_file", "dulieu")).resolve()).casefold()
        sources.setdefault(source, []).append(household)
    if len(sources) > 1:
        if workers > 1:
            from src.export_parallel import SourceJob, run_source_jobs
            jobs = []
            for source in sorted(sources):
                members = tuple(sources[source])
                directory = source_directory(output_root, "excel", Path(members[0].source_file).stem)
                jobs.append(SourceJob("excel", members, Path(output_root), directory, Path(template).resolve(), Path(mapping).resolve(), commune_name))
            groups = run_source_jobs(jobs, workers, progress, cancel_event=cancel_event)
        else:
            groups = []
            for index, source in enumerate(sorted(sources)):
                if cancel_event is not None and cancel_event.is_set():
                    break
                members = sources[source]
                label = Path(members[0].source_file).stem
                groups.append(export_excel_batch(members, template, output_root, mapping, commune_name=commune_name, source_name=label, progress=group_progress(progress, index, len(sources), label), cancel_event=cancel_event))
        report = ExportReport(tuple(s for g in groups for s in g.report.successes), tuple(f for g in groups for f in g.report.failures), tuple(w for g in groups for w in g.report.warnings))
        return ExcelBatch(report, tuple(a for g in groups for a in g.accepted), Path(template).resolve(), Path(mapping).resolve(), Path(output_root) / "excel", groups=tuple(groups))
    if source_name is None:
        source_name = Path(getattr(households[0], "source_file", "dulieu")).stem if households else "dulieu"
    counts = Counter(_household_code(h) for h in households)
    duplicates = tuple(ExportFailure(_household_code(h), "Duplicate household code in loaded data; no file was selected or overwritten") for h in households if counts[_household_code(h)] > 1)
    households = tuple(h for h in households if counts[_household_code(h)] == 1)
    directory = Path(_directory) if _directory is not None else source_directory(output_root, "excel", source_name)
    with tempfile.TemporaryDirectory() as temporary:
        settings_path = Path(temporary) / "settings.json"
        settings_path.write_text(json.dumps({"commune_name": commune_name, "school_year": "2025-2026"}), encoding="utf-8")
        report = export_households(households, template, directory, mapping, settings_path=settings_path, progress=progress, **({"cancel_event": cancel_event} if cancel_event is not None else {}))
    report = ExportReport(report.successes, report.failures + duplicates, report.warnings)
    # Exporter preserves household iteration order; duplicate/sanitized names fail safely.
    available = {str(item.household_code): item for item in report.successes}
    accepted = []
    for household in households:
        item = available.pop(_household_code(household), None)
        if item is not None:
            accepted.append(AcceptedExcel(household, item.path, fingerprint(item.path)))
    manifest = {"source_name": source_name, "source_files": sorted({getattr(h, "source_file", "") for h in households}), "success_count": report.success_count, "failures": [vars(f) for f in report.failures], "files": [{"household_code": _household_code(a.household), "file": a.path.name, "sha256": a.sha256, "members": len(a.household.members)} for a in accepted]}
    (directory / "EXPORT_REPORT.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return ExcelBatch(report, tuple(accepted), Path(template).resolve(), Path(mapping).resolve(), directory, source_name=directory.name)


def validate_pdf(path, expected_pages, *, check_layout=True):
    reader = PdfReader(path, strict=True)
    if reader.is_encrypted or (expected_pages is not None and len(reader.pages) != expected_pages):
        raise ValueError("PDF page count/encryption does not match the accepted Excel")
    if not reader.pages:
        raise ValueError("PDF has no pages")
    if not check_layout:
        return reader  # Parsing is still required to safely merge the actual pages.
    for page in reader.pages:
        if not (page.extract_text() or "").strip():
            raise ValueError("PDF contains a blank/unreadable page")
        width, height = float(page.mediabox.width), float(page.mediabox.height)
        if page.rotation % 180 == 90:
            width, height = height, width
        if abs(width - 841.89) > 4 or abs(height - 595.28) > 4:
            raise ValueError("PDF must retain A4 landscape")
    return reader


def validate_pdf_content(reader, household, mapping):
    def compact(value):
        return re.sub(r"\s+", "", unicodedata.normalize("NFC", str(value))).replace("\u00ad", "")
    # Excel may emit text grouped by font/color, rather than visual row order.
    page_texts = [compact(page.extract_text(extraction_mode="layout") or "") for page in reader.pages]
    # Wrapped cells are contiguous in PDF content order but can be interleaved
    # with neighboring cells in layout order. Keep both presentations per page;
    # only visual layout is used to verify member sequence below.
    content_texts = [compact(page.extract_text() or "") for page in reader.pages]
    for text, content in zip(page_texts, content_texts, strict=True):
        for required in ("PHIẾU ĐIỀU TRA", "XÁC NHẬN CỦA UBND", household.commune_name, _household_code(household), household.household_head_name, household.phone, household_address(household)):
            if required and compact(required) not in text and compact(required) not in content:
                raise ValueError("PDF has clipped/missing household header or signature")
    positions = [0] * len(page_texts)
    for index, member in enumerate(household.members, start=1):
        page_index = (index - 1) // PAGE_CAPACITY
        text = page_texts[page_index]
        if member.member_name:
            position = text.find(compact(member.member_name), positions[page_index])
            if position < 0:
                raise ValueError(f"PDF member order/text differs at member {index}")
            positions[page_index] = position + len(compact(member.member_name))
        expected = {"member_name": member.member_name, "relationship": member.relationship_to_head, "parent_guardian": member.normalized_fields.get("parent_guardian"), "birth_date": "/".join(part or "" for part in member.birth_components) if any(member.birth_components) else None}
        expected.update(member_print_values(member)[0])
        expected.update({name: member.fields.get(field["source"]) for name, field in mapping.approved_member_fields.items()})
        expected.update({name: member.normalized_fields.get(name) for name in ("ethnicity", "religion")})
        for name, value in expected.items():
            # Excel General prints an integral numeric 6.0 as 6. Do not alter
            # textual values such as "6.0", or any underlying source/mapping.
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            if value is not None and str(value).strip() and compact(value) not in text and compact(value) not in content_texts[page_index]:
                raise ValueError(f"PDF has clipped/missing text: member {index}, field {name}; review the Excel cell")


def merge_pdfs(items, destination, *, validate=True):
    destination = Path(destination)
    temporary = destination.with_name("." + uuid.uuid4().hex + ".pdf")
    total = sum(item.pages for item in items)
    if not items:
        raise ValueError("No validated PDFs to merge")
    try:
        writer = PdfWriter()
        try:
            for item in items:
                reader = validate_pdf(item.path, item.pages, check_layout=validate)
                for page in reader.pages:
                    writer.add_page(page)
            with temporary.open("xb") as stream:
                writer.write(stream)
        finally:
            writer.close()
        validate_pdf(temporary, total, check_layout=validate)
        # Windows rename refuses replacement; never overwrite a previous summary.
        if destination.exists():
            raise FileExistsError(str(destination))
        os.rename(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
    return total


def export_pdf_batch(batch, output_root, *, progress=None, validate=True, workers=1, _directory=None, cancel_event=None):
    workers = export_worker_count(workers)
    if not batch.accepted:
        raise ValueError("Export and validate Excel before exporting PDF")
    if batch.groups:
        eligible_groups = [g for g in batch.groups if g.accepted]
        if workers > 1 and len(eligible_groups) > 1:
            from src.export_parallel import SourceJob, run_source_jobs
            jobs = [SourceJob("pdf", group, Path(output_root), source_directory(output_root, "pdf", group.source_name or group.directory.name), validate=validate) for group in eligible_groups]
            groups = run_source_jobs(jobs, workers, progress, cancel_event=cancel_event)
        else:
            groups = []
            for index, group in enumerate(eligible_groups):
                if cancel_event is not None and cancel_event.is_set():
                    break
                groups.append(export_pdf_batch(group, output_root, progress=group_progress(progress, index, len(eligible_groups), group.source_name), validate=validate, cancel_event=cancel_event))
        paths = tuple(p for g in groups for p in g.merged_paths)
        errors = [g.merge_error for g in groups if g.merge_error]
        return PdfBatch(tuple(s for g in groups for s in g.successes), tuple(f for g in groups for f in g.failures), paths[0] if len(paths) == 1 else None, sum(g.total_pages for g in groups), "; ".join(errors) or None, paths)
    import pythoncom
    import win32com.client

    directory = Path(_directory) if _directory is not None else source_directory(output_root, "pdf", batch.source_name or batch.directory.name)
    mapping = load_template_mapping(batch.mapping) if validate else None
    successes, failures, eligible = [], [], []
    excel = None
    original_map_paper = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        original_map_paper = excel.MapPaperSize
        excel.MapPaperSize = False
        if validate:
            template = excel.Workbooks.Open(str(batch.template), UpdateLinks=0, ReadOnly=True, AddToMru=False)
            try:
                sheet = _verify_template_sheet(template, mapping)
                signature = _template_signature(batch.template, sheet, mapping)
            finally:
                template.Close(SaveChanges=False)
        ordered = sorted(batch.accepted, key=lambda item: household_order(item.household))
        total = len(ordered) * (2 if validate else 1)
        # Finish validation of the batch before starting any PDF conversion.
        for index, item in enumerate(ordered):
            if cancel_event is not None and cancel_event.is_set():
                break
            if not validate:
                eligible.append(item)
                continue
            if progress:
                progress(index, total, f"Nghiệm thu Excel {index + 1}/{len(ordered)}")
            try:
                if fingerprint(item.path) != item.sha256:
                    raise ValueError("Excel changed after acceptance; export it again")
                _validate_saved_workbook(excel, item.path, item.household, mapping, math.ceil(len(item.household.members) / PAGE_CAPACITY), signature)
                eligible.append(item)
            except Exception as error:
                failures.append(ExportFailure(_household_code(item.household), str(error)))
                if progress:
                    progress(index + 1, total, f"LỖI nghiệm thu Excel hộ {_household_code(item.household)}: {error}")
        for index, item in enumerate(eligible):
            if cancel_event is not None and cancel_event.is_set():
                break
            if progress:
                progress((len(ordered) if validate else 0) + index, total, f"PDF {index + 1}/{len(eligible)}; thành công {len(successes)}, lỗi {len(failures)}")
            workbook = None
            temporary = directory / ("." + uuid.uuid4().hex + ".pdf")
            try:
                workbook = excel.Workbooks.Open(str(item.path), UpdateLinks=0, ReadOnly=True, AddToMru=False)
                expected_pages = workbook.Worksheets.Count
                if validate:
                    for sheet in workbook.Worksheets:
                        verify_page_setup(sheet, mapping.print_area)
                        page_setup(sheet, mapping.print_area)
                workbook.Worksheets.Item(1).Activate()
                workbook.ExportAsFixedFormat(0, str(temporary), 0, True, False, OpenAfterPublish=False)
                # GET.DOCUMENT(50) can report cached pagination before rendering.
                # Check refreshed per-sheet pagination AND actual PDF page count.
                if validate:
                    for sheet in workbook.Worksheets:
                        sheet.Activate()
                        if int(excel.ExecuteExcel4Macro("GET.DOCUMENT(50)")) != 1:
                            raise ValueError("Rendered Excel worksheet does not print on exactly one page")
                    validate_pdf_content(validate_pdf(temporary, expected_pages), item.household, mapping)
                else:
                    expected_pages = len(validate_pdf(temporary, None, check_layout=False).pages)
                workbook.Close(SaveChanges=False)
                workbook = None
                final = directory / (item.path.stem + ".pdf")
                if final.exists():
                    raise FileExistsError(str(final))
                os.rename(temporary, final)
                successes.append(PdfResult(_household_code(item.household), final, expected_pages))
            except Exception as error:
                failures.append(ExportFailure(_household_code(item.household), str(error)))
                if progress:
                    progress((len(ordered) if validate else 0) + index + 1, total, f"LỖI PDF hộ {_household_code(item.household)}: {error}")
            finally:
                try:
                    if workbook is not None:
                        try:
                            workbook.Close(SaveChanges=False)
                        except Exception:
                            LOGGER.exception("Failed closing PDF input workbook")
                finally:
                    temporary.unlink(missing_ok=True)
    finally:
        try:
            if excel is not None:
                try:
                    if original_map_paper is not None:
                        excel.MapPaperSize = original_map_paper
                finally:
                    excel.Quit()
        finally:
            pythoncom.CoUninitialize()
    merged = None
    pages = 0
    merge_error = None
    if successes and not (cancel_event is not None and cancel_event.is_set()):
        try:
            merged = Path(output_root) / (directory.name + ".pdf")
            pages = merge_pdfs(successes, merged, validate=validate)
        except Exception as error:
            merged = None
            merge_error = str(error)
    if progress:
        progress(1, 1, f"PDF: {len(successes)} thành công, {len(failures)} lỗi; tổng hợp {pages} trang")
    report = {"source_name": batch.source_name, "validation_enabled": validate, "successes": [{"household_code": i.household_code, "file": i.path.name, "pages": i.pages} for i in successes], "failures": [vars(f) for f in failures], "merged": str(merged.resolve()) if merged else None, "total_pages": pages, "merge_error": merge_error}
    (directory / "PDF_REPORT.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return PdfBatch(tuple(successes), tuple(failures), merged, pages, merge_error, (merged,) if merged else ())
