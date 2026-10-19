"""Group normalized members without merging equal codes from separate files."""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from normalizer import NormalizedMember


@dataclass(frozen=True)
class HouseholdWarning:
    code: str
    message: str
    sources: tuple[str, ...] = ()


@dataclass(frozen=True)
class Household:
    household_id: str
    commune_name: str
    household_code: object
    household_code_key: str | None
    source_file: str
    members: tuple[NormalizedMember, ...]
    household_head_name: str | None
    suggested_head_name: str | None
    phone: str | None
    residence: str | None
    warnings: tuple[HouseholdWarning, ...]


@dataclass(frozen=True)
class GroupingResult:
    households: tuple[Household, ...]

    @property
    def warnings(self) -> tuple[HouseholdWarning, ...]:
        return tuple(
            warning for household in self.households for warning in household.warnings
        )


def group_households(
    members: Iterable[NormalizedMember],
) -> GroupingResult:
    """Group by commune, source file and household code; retain every row."""
    buckets: dict[tuple[str, str, str], list[NormalizedMember]] = {}
    duplicate_index: dict[str, set[tuple[str, str, str]]] = {}

    for member in members:
        source_file = str(Path(member.source_record.source.file).resolve()).casefold()
        code_key = member.household_code_key
        group_code = (
            code_key
            if code_key is not None
            else f"__missing__:{member.source_record.source.sheet}:{member.source_record.source.row}"
        )
        key = (member.commune_key, source_file, group_code)
        buckets.setdefault(key, []).append(member)
        if code_key is not None:
            duplicate_index.setdefault(code_key, set()).add(key)

    duplicate_keys = {
        key
        for group_keys in duplicate_index.values()
        if len({group_key[1] for group_key in group_keys}) > 1
        for key in group_keys
    }

    households: list[Household] = []
    for key, group_members in buckets.items():
        commune_key, source_file, group_code = key
        household_code_key = group_members[0].household_code_key
        warnings = _household_warnings(group_members)
        if key in duplicate_keys:
            warnings.append(
                HouseholdWarning(
                    "HOUSEHOLD_CODE_DUPLICATE_ACROSS_FILES",
                    "Mã hộ xuất hiện trong nhiều tệp nguồn; các nhóm được giữ riêng.",
                    tuple(sorted({member.source_record.source.file for member in group_members})),
                )
            )

        head_members = [
            member for member in group_members
            if _is_household_head(member.relationship_to_head)
        ]
        suggested_head_name = (
            head_members[0].member_name or None if len(head_members) == 1 else None
        )
        if len(head_members) == 1:
            household_phone = head_members[0].phone
        else:
            household_phone = None
            if not head_members:
                warnings.append(
                    HouseholdWarning(
                        "HOUSEHOLD_HEAD_NOT_IDENTIFIED",
                        "Không xác định được thành viên chủ hộ từ cột AU; để trống điện thoại hộ.",
                    )
                )
            else:
                warnings.append(
                    HouseholdWarning(
                        "HOUSEHOLD_HEAD_NOT_UNIQUE",
                        "Có nhiều bản ghi được đánh dấu chủ hộ; để trống điện thoại hộ.",
                        tuple(member.source_record.source.file for member in head_members),
                    )
                )
                head_phones = {member.phone for member in head_members if member.phone}
                if len(head_phones) > 1:
                    warnings.append(
                        HouseholdWarning(
                            "HOUSEHOLD_HEAD_PHONE_CONFLICT",
                            "Các bản ghi chủ hộ có số điện thoại khác nhau; không tự chọn số.",
                            tuple(member.source_record.source.file for member in head_members),
                        )
                    )

        source_path = group_members[0].source_record.source.file
        household_id = f"{commune_key}|{source_file}|{group_code}"
        households.append(
            Household(
                household_id=household_id,
                commune_name=group_members[0].commune_name,
                household_code=group_members[0].household_code,
                household_code_key=household_code_key,
                source_file=source_path,
                members=tuple(group_members),
                household_head_name=_single_value(
                    member.household_head_name for member in group_members
                ),
                suggested_head_name=suggested_head_name,
                phone=household_phone,
                residence=_single_value(member.residence for member in group_members),
                warnings=tuple(warnings),
            )
        )

    households.sort(
        key=lambda household: (
            household.commune_name.casefold(),
            household.household_code_key or "",
            household.source_file.casefold(),
            household.members[0].source_record.source.row,
        )
    )
    return GroupingResult(tuple(households))


def _household_warnings(
    members: list[NormalizedMember],
) -> list[HouseholdWarning]:
    warnings: list[HouseholdWarning] = []
    if members[0].household_code_key is None:
        warnings.append(
            HouseholdWarning(
                "HOUSEHOLD_CODE_MISSING",
                "Thiếu mã hộ; bản ghi được giữ thành nhóm riêng.",
            )
        )

    head_names = {
        _declared_head_name_key(member)
        for member in members
        if _declared_head_name_key(member)
    }
    if len(head_names) > 1:
        warnings.append(
            HouseholdWarning(
                "HOUSEHOLD_HEAD_NAME_CONFLICT",
                "Tên chủ hộ khai báo trong các bản ghi không thống nhất.",
                tuple(member.source_record.source.file for member in members),
            )
        )

    residences = {
        member.residence.casefold()
        for member in members
        if member.residence
    }
    if len(residences) > 1:
        warnings.append(
            HouseholdWarning(
                "HOUSEHOLD_RESIDENCE_CONFLICT",
                "Diện cư trú khai báo trong các bản ghi không thống nhất.",
                tuple(member.source_record.source.file for member in members),
            )
        )
    return warnings


def _is_household_head(relationship: str | None) -> bool:
    if not relationship:
        return False
    normalized = unicodedata.normalize("NFD", relationship).casefold()
    normalized = normalized.replace("đ", "d")
    normalized = "".join(
        character
        for character in normalized
        if not unicodedata.combining(character)
    )
    normalized = " ".join(normalized.split())
    return normalized in {"chu ho", "chu ho gia dinh"}


def _single_value(values: Iterable[str | None]) -> str | None:
    distinct = {value for value in values if value}
    return next(iter(distinct)) if len(distinct) == 1 else None


def _declared_head_name_key(member: NormalizedMember) -> str:
    parts = []
    for name in ("head_family_name", "head_given_name"):
        value = member.fields.get(name)
        if value is not None and str(value).strip():
            parts.append(unicodedata.normalize("NFC", str(value)).strip().casefold())
    return " ".join(parts)
