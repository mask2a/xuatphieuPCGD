"""Time real export using synthetic data in disposable directories."""
from pathlib import Path
from time import perf_counter
import argparse
import cProfile
import hashlib
import json
import pstats
import sys
import tempfile
from dataclasses import replace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from settings import load_settings
from scripts.verify_gui_batch_pdf import legible_fixture
from src.export_workflow import export_excel_batch, export_pdf_batch


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--max-seconds", type=float)
    parser.add_argument("--max-com-calls", type=int)
    parser.add_argument("--skip-pdf-validation", action="store_true")
    parser.add_argument("--sources", type=int, default=1)
    parser.add_argument("--workers", type=int, default=1)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    settings = load_settings(ROOT / "config/settings.json")
    mapping = ROOT / "config/template_mapping.json"
    protected = [settings.template_path, settings.source_path, mapping]
    hashes = lambda: {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    before = hashes()
    profile = cProfile.Profile()
    profile.enable()
    with tempfile.TemporaryDirectory() as temporary:
        start = perf_counter()
        households = [replace(legible_fixture(f"BENCH-{i}", 1), source_file=str(Path(temporary) / f"Nguon {i}.xlsx")) for i in range(1, args.sources + 1)]
        options = {"workers": args.workers} if args.workers != 1 else {}
        excel = export_excel_batch(households, settings.template_path, temporary, mapping, commune_name="Xã Kiểm thử", **options)
        excel_seconds = perf_counter() - start
        if excel.report.failures:
            raise RuntimeError(str(excel.report.failures))
        start = perf_counter()
        pdf = export_pdf_batch(excel, temporary, validate=not args.skip_pdf_validation, **options)
        pdf_seconds = perf_counter() - start
        if pdf.failures or pdf.merge_error or pdf.total_pages != args.sources:
            raise RuntimeError(str((pdf.failures, pdf.merge_error, pdf.total_pages)))
    profile.disable()
    stats = pstats.Stats(profile)
    functions = sorted([{"function": f"{Path(file).name}:{line}:{name}", "calls": calls, "seconds": cumulative} for (file, line, name), (_, calls, _, cumulative, _) in stats.stats.items() if str(ROOT).casefold() in file.casefold() and "venv" not in file.casefold()], key=lambda row: row["seconds"], reverse=True)[:25]
    com_calls = sum(values[1] for (_, _, name), values in stats.stats.items() if "Invoke" in name)
    result = {"sources": args.sources, "workers": args.workers, "pdf_validation": not args.skip_pdf_validation, "excel_seconds": excel_seconds, "pdf_seconds": pdf_seconds, "total_seconds": excel_seconds + pdf_seconds, "com_invocations_parent_process": com_calls, "functions_parent_process": functions, "protected_unchanged": hashes() == before}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.report:
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return not result["protected_unchanged"] or (args.max_seconds is not None and result["total_seconds"] > args.max_seconds) or (args.max_com_calls is not None and com_calls > args.max_com_calls)


if __name__ == "__main__":
    raise SystemExit(bool(main()))
