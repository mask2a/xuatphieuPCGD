"""Bounded spawn workers. Only plain data crosses process boundaries."""
from concurrent.futures import ProcessPoolExecutor, wait, FIRST_COMPLETED
from contextlib import ExitStack
from dataclasses import dataclass
import json
import logging
import multiprocessing
import os
from pathlib import Path
from queue import Empty
import time

LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class SourceJob:
    kind: str
    payload: object
    output_root: Path
    directory: Path
    template: Path | None = None
    mapping: Path | None = None
    commune_name: str = ""
    validate: bool = True


def _failed(job, error):
    from src.export_workflow import ExcelBatch, PdfBatch
    from src.excel_exporter import ExportReport, ExportFailure, _household_code
    households = job.payload if job.kind == "excel" else [a.household for a in job.payload.accepted]
    failures = tuple(ExportFailure(_household_code(h), f"Nguồn {job.directory.name}: {type(error).__name__}: {error}") for h in households)
    if job.kind == "excel":
        return ExcelBatch(ExportReport((), failures, ()), (), job.template, job.mapping, job.directory, source_name=job.directory.name)
    return PdfBatch((), failures, None, 0)


def _record(job, result, started, pid):
    name = "EXPORT_REPORT.json" if job.kind == "excel" else "PDF_REPORT.json"
    path = job.directory / name
    try:
        if path.exists():
            report = json.loads(path.read_text(encoding="utf-8"))
        else:
            failures = result.report.failures if job.kind == "excel" else result.failures
            report = {"source_name": job.directory.name, "success_count": 0, "files": [], "failures": [vars(f) for f in failures]}
        report["execution"] = {"worker_pid": pid, "started_ns": started, "finished_ns": time.time_ns()}
        path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        LOGGER.exception("Cannot write source execution metadata: %s", path)


def _execute(job, progress_queue, index, cancel_event=None):
    started = time.time_ns()
    try:
        from src.export_workflow import export_excel_batch, export_pdf_batch
        callback = None if progress_queue is None else lambda d, t, m: progress_queue.put((index, d, t, m))
        if job.kind == "excel":
            result = export_excel_batch(job.payload, job.template, job.output_root, job.mapping, commune_name=job.commune_name, progress=callback, source_name=job.directory.name, _directory=job.directory, cancel_event=cancel_event)
        else:
            result = export_pdf_batch(job.payload, job.output_root, validate=job.validate, progress=callback, _directory=job.directory, cancel_event=cancel_event)
    except Exception as error:
        # An ordinary source/COM failure must not poison the worker queue.
        result = _failed(job, error)
    _record(job, result, started, os.getpid())
    return result


def run_source_jobs(jobs, workers, progress=None, cancel_event=None):
    """Keep at most workers jobs queued, return results in source order."""
    context = multiprocessing.get_context("spawn")
    results = [None] * len(jobs)
    fractions = [0] * len(jobs)

    def emit(index, done, total, message):
        fractions[index] = max(fractions[index], min(1000, max(0, int(done * 1000 / max(1, total)))))
        if progress:
            progress(sum(fractions), len(jobs) * 1000, f"{jobs[index].directory.name}: {message}")

    with ExitStack() as stack:
        manager = stack.enter_context(context.Manager()) if progress or cancel_event is not None else None
        messages = manager.Queue() if progress and manager else None
        shared_cancel = manager.Event() if cancel_event is not None else None
        pool = stack.enter_context(ProcessPoolExecutor(max_workers=min(workers, len(jobs)), mp_context=context))
        pending = {}
        next_index = 0
        while next_index < len(jobs) or pending:
            if cancel_event is not None and cancel_event.is_set():
                shared_cancel.set()
                next_index = len(jobs)
            while next_index < len(jobs) and len(pending) < workers:
                index = next_index
                next_index += 1
                try:
                    pending[pool.submit(_execute, jobs[index], messages, index, shared_cancel)] = index
                except Exception as error:
                    results[index] = _failed(jobs[index], error)
                    _record(jobs[index], results[index], time.time_ns(), None)
                    emit(index, 1, 1, "Không khởi chạy được tác vụ")
            completed, _ = wait(pending, timeout=0.1, return_when=FIRST_COMPLETED) if pending else ((), ())
            if messages is not None:
                while True:
                    try:
                        emit(*messages.get_nowait())
                    except Empty:
                        break
            for future in completed:
                index = pending.pop(future)
                try:
                    results[index] = future.result()
                except Exception as error:
                    results[index] = _failed(jobs[index], error)
                    _record(jobs[index], results[index], time.time_ns(), None)
                emit(index, 1, 1, "Hoàn tất nguồn")
    return tuple(result for result in results if result is not None)
