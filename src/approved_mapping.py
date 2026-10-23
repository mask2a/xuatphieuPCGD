"""Additional user-approved mappings, verified against the original Excel form."""
import re
import unicodedata


# Exact composite headers (merged parent / child) verified through Excel COM.
FIELDS = {
    "vocational_graduation_level": ("Bậc TN Nghề", "S"),
    "vocational_graduation_year": ("Năm TN nghề", "T"),
    "completed_education": ("Học xong / Lớp", "V"),
    "completed_education_year": ("Học xong / Năm", "W"),
    "dropped_out_class": ("Bỏ học / Lớp", "X"),
    "dropped_out_year": ("Bỏ học / Năm", "Y"),
    "currently_studying_class": ("Đang học lớp", "Z"),
    "completed_class": ("Hoàn thành lớp", "AA"),
    "reliteracy_level": ("Tái mù chữ mức", "AB"),
    "disability_certified": ("Có chứng nhận khuyết tật", "AD"),
    "learning_capacity": ("Khả năng học tập", "AE"),
    "special_circumstances_details": ("Chi tiết hoàn cảnh đặt biệt", "AF"),
    "notes": ("Ghi chú", "AG"),
}
DISABILITIES = (
    "Khuyết tật vận động", "Khuyết tật nghe, nói", "Khuyết tật nhìn",
    "Khuyết tật thần kinh, tâm thần", "Khuyết tật trí tuệ",
    "Khuyết tật học tập", "Tự kỷ", "Khuyết tật khác",
)


def key(value):
    return " ".join(unicodedata.normalize("NFC", str(value or "")).split()).casefold()


def resolve(record, title):
    matches = [column for column, header in record.headers.items() if key(header) == key(title)]
    if len(matches) != 1:
        return None, "SOURCE_HEADER_UNRESOLVED"
    return record.values.get(matches[0]), None


def values(member):
    result, warnings = {}, set()
    for field, (title, target) in FIELDS.items():
        value, warning = resolve(member.source_record, title)
        result[target] = value
        if warning:
            warnings.add((warning, f"Không xác định duy nhất tiêu đề nguồn: {title}; không ghi trường này."))
    selected = []
    disability_error = False
    for title in DISABILITIES:
        value, warning = resolve(member.source_record, title)
        if warning:
            disability_error = True
            warnings.add((warning, f"Không xác định duy nhất tiêu đề nguồn: {title}; không ghi loại khuyết tật."))
        elif key(value) == "x":
            label = title.removeprefix("Khuyết tật ")
            selected.append(label[0].upper() + label[1:])
    result["AC"] = None if disability_error else ", ".join(selected)
    return result, sorted(warnings)


def address_text(address, commune):
    address = str(address or "").strip().strip(", ")
    commune = str(commune or "").strip().strip(", ")
    if not address:
        return commune
    # Match the whole locality, not a substring of another locality name.
    if commune and not re.search(r"(?<!\w)" + re.escape(key(commune)) + r"(?!\w)", key(address)):
        return f"{address}, {commune}"
    return address


def household_address(household):
    addresses = {str(m.fields.get("address")).strip() for m in household.members if m.fields.get("address") is not None and str(m.fields.get("address")).strip()}
    if len(addresses) > 1:
        return None
    return address_text(next(iter(addresses), ""), household.commune_name)


def template_label(sheet, address):
    text = str(sheet.Range(address).Value2 or "")
    label, separator, _ = text.partition(":")
    if not separator or not label.strip():
        raise ValueError(f"Template label must contain ':' at {address}")
    return label + separator


def address_label(sheet):
    return template_label(sheet, "A3")


def verify_targets(sheet, rows):
    if sheet.Range("A3").MergeArea.Address.replace("$", "") != "A3:G3":
        raise ValueError("Unexpected address merge area")
    address_label(sheet)
    residence = sheet.Range("Z2")
    if residence.MergeArea.Address.replace("$", "") != "Z2:AG2" or str(residence.Value2).count("q") != 2:
        raise ValueError("Unexpected residence checkbox structure")
    for row in rows:
        female = sheet.Range(f"C{row + 3}")
        if female.Value2 != "q" or female.MergeArea.Address.replace("$", "") != f"C{row + 3}" or female.HasFormula:
            raise ValueError("Unexpected female checkbox")
        for column in [item[1] for item in FIELDS.values()] + ["AC"]:
            cell = sheet.Range(f"{column}{row}")
            expected = f"T{row}:U{row}" if column == "T" else f"{column}{row}"
            if cell.MergeArea.Address.replace("$", "") != expected:
                raise ValueError(f"Unexpected approved target merge: {column}{row}")
            if cell.HasFormula or cell.Value2 not in (None, ""):
                raise ValueError(f"Approved target is not empty: {column}{row}")
