"""Create normalized member records while retaining each untouched source row."""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable, Mapping

from reader import RawRecord
from settings import DEFAULT_SOURCE_COLUMNS, Settings


@dataclass(frozen=True)
class NormalizedMember:
    source_record: RawRecord
    fields: Mapping[str, Any]
    normalized_fields: Mapping[str, str | None]
    member_name: str
    member_name_key: str
    birth_components: tuple[str | None, str | None, str | None]
    household_code: str | None
    household_code_key: str | None
    household_head_name: str | None
    relationship_to_head: str | None
    phone: str | None
    residence: str | None
    commune_name: str
    commune_key: str


def normalize_records(
    records: Iterable[RawRecord], settings: Settings | None = None
) -> tuple[NormalizedMember, ...]:
    """Normalize searchable fields without changing raw source cell values."""
    effective_settings = settings or Settings()
    columns = {**DEFAULT_SOURCE_COLUMNS, **effective_settings.source_columns}
    members: list[NormalizedMember] = []

    for record in records:
        fields = {
            name: record.values.get(column)
            for name, column in columns.items()
        }
        normalized_fields = {
            name: _normalized_text(record.display_values.get(column))
            for name, column in columns.items()
        }
        family_name = normalized_fields["family_name"] or ""
        given_name = normalized_fields["given_name"] or ""
        member_name = " ".join(part for part in (family_name, given_name) if part)
        household_code_key = _key(record.display_values.get(columns["household_code"]))
        commune_name = _clean_text(effective_settings.commune_name)

        members.append(
            NormalizedMember(
                source_record=record,
                fields=fields,
                normalized_fields=normalized_fields,
                member_name=member_name,
                member_name_key=member_name.casefold(),
                birth_components=(
                    _normalized_text(record.display_values.get(columns["birth_day"])),
                    _normalized_text(record.display_values.get(columns["birth_month"])),
                    _normalized_text(record.display_values.get(columns["birth_year"])),
                ),
                household_code=_optional_display(
                    record.display_values.get(columns["household_code"])
                ),
                household_code_key=household_code_key,
                household_head_name=_joined_fields(
                    normalized_fields.get("head_family_name"),
                    normalized_fields.get("head_given_name"),
                ),
                relationship_to_head=normalized_fields.get("relationship_to_head"),
                phone=_optional_display(record.display_values.get(columns["phone"])),
                residence=_normalized_text(record.display_values.get(columns["residence"])),
                commune_name=commune_name,
                commune_key=_key(commune_name) or "",
            )
        )

    return tuple(members)


def _normalized_text(value: Any) -> str | None:
    if value is None:
        return None
    text = unicodedata.normalize("NFC", str(value))
    text = " ".join(text.split())
    return text or None


def _clean_text(value: Any) -> str:
    return _normalized_text(value) or ""


def _key(value: Any) -> str | None:
    text = _normalized_text(value)
    return text.casefold() if text is not None else None


def _optional_display(value: Any) -> str | None:
    if value is None:
        return None
    display = str(value)
    return display if display.strip() else None


def _joined_fields(first: str | None, second: str | None) -> str | None:
    joined = " ".join(part for part in (first, second) if part)
    return joined or None
