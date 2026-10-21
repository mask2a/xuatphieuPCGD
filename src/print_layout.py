"""One A4 landscape geometry for Excel and PDF; never changes the template."""
from datetime import date
from decimal import Decimal, InvalidOperation

CHECKED = "☒"
UNCHECKED = "☐"
CHECKBOX_FONT = "Segoe UI Symbol"
ETHNICITY_FONT_SIZE = 8
LONG_CLASS_FONT_SIZE = 8
PAGE_CAPACITY = 7
PRINT_ZOOM = False
MARGIN = 72 * 0.5 / 2.54


def page_setup(sheet, print_area="$A$1:$AG$87"):
    setup = sheet.PageSetup
    setup.PaperSize = 9
    setup.Orientation = 2
    setup.PrintArea = print_area
    setup.Zoom = PRINT_ZOOM
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = 1
    setup.LeftMargin = setup.RightMargin = MARGIN
    setup.TopMargin = setup.BottomMargin = MARGIN
    setup.HeaderMargin = setup.FooterMargin = 0
    setup.CenterHorizontally = True
    setup.CenterVertically = True
    for name in ("LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter"):
        setattr(setup, name, "")


def normalize_geometry(sheet, member_count, mapping):
    if member_count > PAGE_CAPACITY:
        raise ValueError("A printed sheet may contain at most seven members")
    # Pagination only: retain all seven slots, including unused ones, and the footer.
    # Column widths, row heights, fonts, borders and merges stay as in the template.
    sheet.Rows("44:81").Hidden = True
    sheet.ResetAllPageBreaks()
    page_setup(sheet, mapping.print_area)


def verify_page_setup(sheet, print_area="$A$1:$AG$87"):
    setup = sheet.PageSetup
    if (setup.PaperSize, setup.Orientation, setup.Zoom, setup.FitToPagesWide, setup.FitToPagesTall) != (9, 2, PRINT_ZOOM, 1, 1):
        raise ValueError("A4 landscape fit-to-page settings differ")
    if setup.PrintArea != print_area or not setup.CenterHorizontally or not setup.CenterVertically:
        raise ValueError("Print area/centering differs")
    if any(abs(float(getattr(setup, name)) - MARGIN) > 0.05 for name in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin")):
        raise ValueError("Print margins differ")
    if setup.HeaderMargin != 0 or setup.FooterMargin != 0:
        raise ValueError("Header/footer margins differ")


def checkbox(cell, selected):
    cell.Value2 = CHECKED if selected else UNCHECKED
    cell.Font.Name = CHECKBOX_FONT
    cell.Font.Size = 11


def invalid_birth_date(member):
    parts = member.birth_components
    if not all(parts):
        return False
    try:
        numbers = [Decimal(str(part)) for part in parts]
        if any(not number.is_finite() or number != number.to_integral_value() for number in numbers):
            return True
        day, month, year = (int(number) for number in numbers)
        date(year, month, day)
    except (ValueError, TypeError, OverflowError, InvalidOperation):
        return True
    return False


def class_font_size(cell, original_size):
    """User-approved local reduction; never change the source template."""
    return min(original_size, LONG_CLASS_FONT_SIZE) if fit_text(cell, original_size) else original_size


def fit_text(cell, font_size=None):
    """Detect long text without changing the approved template formatting."""
    value = cell.Value2
    if not isinstance(value, str) or not value:
        return False
    import win32ui
    dc = win32ui.CreateDC().CreateCompatibleDC()
    try:
        font = win32ui.CreateFont({"name": cell.Font.Name or "Arial", "height": -round(float(cell.Font.Size if font_size is None else font_size) * 96 / 72), "weight": 700 if cell.Font.Bold else 400})
        previous = dc.SelectObject(font)
        try:
            width = max(dc.GetTextExtent(line)[0] for line in value.splitlines() or [value]) * 72 / 96
        finally:
            dc.SelectObject(previous)
        available = float(cell.MergeArea.Width) - 4
        if width > available or "\n" in value:
            return True
        return False
    finally:
        dc.DeleteDC()
