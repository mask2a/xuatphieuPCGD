"""Compose reader, normalizer, and household grouping into one analysis result."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from household_grouper import GroupingResult, group_households
from normalizer import NormalizedMember, normalize_records
from reader import ReadResult, read_sources
from settings import Settings


@dataclass(frozen=True)
class AnalysisResult:
    read_result: ReadResult
    members: tuple[NormalizedMember, ...]
    grouping_result: GroupingResult
    settings: Settings


def analyze_sources(
    paths: Iterable[str | Path], settings: Settings | None = None,
    *, progress: Callable[[int, int, str], None] | None = None,
) -> AnalysisResult:
    effective_settings = settings or Settings()
    read_result = read_sources(paths, effective_settings, progress=progress)
    if progress:
        progress(0, 1, f"Đang chuẩn hóa {len(read_result.records)} bản ghi...")
    members = normalize_records(read_result.records, effective_settings)
    if progress:
        progress(0, 1, f"Đang nhóm {len(members)} nhân khẩu thành hộ...")
    grouping_result = group_households(members)
    if progress:
        progress(1, 1, f"Hoàn tất: {len(members)} nhân khẩu, {len(grouping_result.households)} hộ.")
    return AnalysisResult(
        read_result=read_result,
        members=members,
        grouping_result=grouping_result,
        settings=effective_settings,
    )
