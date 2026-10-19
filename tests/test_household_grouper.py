import unittest
from pathlib import Path

from normalizer import normalize_records
from reader import RawRecord, SourceLocation, read_sources
from household_grouper import group_households
from settings import Settings


def member_record(
    source: str,
    row: int,
    name: str,
    code: str = "001",
    relationship: str = "Thành viên",
    phone: str | None = None,
    head_name: str = "Nguyễn An",
    residence: str = "Thường trú",
) -> RawRecord:
    values = {
        "B": name,
        "C": None,
        "D": None,
        "E": None,
        "F": None,
        "K": head_name,
        "L": None,
        "N": code,
        "O": residence,
        "AU": relationship,
        "AW": phone,
    }
    displays = {column: "" if value is None else str(value) for column, value in values.items()}
    return RawRecord(
        values=values,
        display_values=displays,
        source=SourceLocation(source, "MauNhapLieu", row),
    )


class HouseholdGrouperTests(unittest.TestCase):
    def test_selects_only_the_unique_head_phone_and_keeps_duplicate_names(self):
        records = [
            member_record("a.xls", 5, "Nguyễn An", relationship="Chủ hộ", phone="0911"),
            member_record("a.xls", 6, "Nguyễn An", phone="0922"),
        ]

        household = group_households(normalize_records(records)).households[0]

        self.assertEqual(len(household.members), 2)
        self.assertEqual(household.phone, "0911")
        self.assertEqual(household.suggested_head_name, "Nguyễn An")
        self.assertFalse(any(item.code == "HOUSEHOLD_PHONE_CONFLICT" for item in household.warnings))

    def test_missing_or_ambiguous_head_never_uses_another_members_phone(self):
        no_head = group_households(
            normalize_records([
                member_record("a.xls", 5, "A", phone="0911"),
                member_record("a.xls", 6, "B", phone="0922"),
            ])
        ).households[0]
        ambiguous = group_households(
            normalize_records([
                member_record("a.xls", 5, "A", relationship="Chủ hộ", phone="0911"),
                member_record("a.xls", 6, "B", relationship="Chủ hộ", phone="0922"),
            ])
        ).households[0]

        self.assertIsNone(no_head.phone)
        self.assertIn("HOUSEHOLD_HEAD_NOT_IDENTIFIED", {item.code for item in no_head.warnings})
        self.assertIsNone(ambiguous.phone)
        self.assertIn("HOUSEHOLD_HEAD_NOT_UNIQUE", {item.code for item in ambiguous.warnings})
        self.assertIn("HOUSEHOLD_HEAD_PHONE_CONFLICT", {item.code for item in ambiguous.warnings})

    def test_same_code_in_two_files_is_reported_but_never_merged(self):
        members = normalize_records(
            [member_record("one.xls", 5, "A", relationship="Chủ hộ", phone="0900")],
            Settings(commune_name="Ward A"),
        ) + normalize_records(
            [member_record("two.xls", 5, "B", relationship="Chủ hộ", phone="0800")],
            Settings(commune_name="Ward B"),
        )

        result = group_households(members)

        self.assertEqual(len(result.households), 2)
        self.assertEqual([len(household.members) for household in result.households], [1, 1])
        self.assertTrue(
            all(
                "HOUSEHOLD_CODE_DUPLICATE_ACROSS_FILES"
                in {item.code for item in household.warnings}
                for household in result.households
            )
        )

    def test_same_code_is_separated_when_configured_communes_differ(self):
        records = [
            member_record("same.xls", 5, "A", relationship="Chủ hộ", phone="0900"),
            member_record("same.xls", 6, "B", relationship="Chủ hộ", phone="0800"),
        ]
        members = normalize_records(records[:1], Settings(commune_name="Ward A")) + normalize_records(
            records[1:], Settings(commune_name="Ward B")
        )

        result = group_households(members)

        self.assertEqual(len(result.households), 2)
        self.assertEqual(
            {household.commune_name for household in result.households},
            {"Ward A", "Ward B"},
        )

    def test_keeps_households_with_blank_codes_separate(self):
        members = normalize_records([
            member_record("one.xls", 5, "A", code=""),
            member_record("one.xls", 6, "B", code=""),
        ])

        result = group_households(members)

        self.assertEqual(len(result.households), 2)
        self.assertTrue(
            all(
                "HOUSEHOLD_CODE_MISSING" in {item.code for item in household.warnings}
                for household in result.households
            )
        )

    def test_verified_source_has_expected_conflict_counts_and_owner_phone_rule(self):
        fixture = Path(__file__).resolve().parents[1] / "dulieumau.xls"
        if not fixture.exists():
            self.skipTest("Verified source workbook is not present")
        result = group_households(
            normalize_records(read_sources([fixture]).records)
        )
        warning_codes = [
            warning.code
            for household in result.households
            for warning in household.warnings
        ]

        self.assertEqual(len(result.households), 120)
        self.assertEqual(warning_codes.count("HOUSEHOLD_HEAD_NAME_CONFLICT"), 3)
        self.assertEqual(warning_codes.count("HOUSEHOLD_RESIDENCE_CONFLICT"), 5)
        self.assertEqual(
            sum(
                len({member.phone for member in household.members if member.phone}) > 1
                for household in result.households
            ),
            28,
        )
        self.assertEqual(warning_codes.count("HOUSEHOLD_PHONE_CONFLICT"), 0)
        self.assertEqual(warning_codes.count("HOUSEHOLD_HEAD_PHONE_CONFLICT"), 0)
        self.assertEqual(
            sum(
                member.birth_components[2] is None
                for household in result.households
                for member in household.members
            ),
            1,
        )
        first_source_household = next(
            household
            for household in result.households
            if any(member.source_record.source.row == 5 for member in household.members)
        )
        self.assertEqual(first_source_household.phone, "0946236722")


if __name__ == "__main__":
    unittest.main()
