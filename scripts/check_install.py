"""Check shipped resources and runtime dependencies without opening Excel."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    if sys.platform != "win32":
        raise RuntimeError("PCGD requires Windows and Microsoft Excel desktop")
    if sys.version_info < (3, 10):
        raise RuntimeError("Python 3.10 or newer is required")

    import tkinter as tk
    import openpyxl
    import xlrd
    import pypdf
    import pythoncom
    import win32com.client
    import winreg
    import gui
    from settings import load_settings
    from src.template_mapper import load_template_mapping

    settings = load_settings(ROOT / "config/settings.json")
    mapping = load_template_mapping(ROOT / "config/template_mapping.json")
    template = ROOT / "templates/maubaocao.xls"
    workbook = xlrd.open_workbook(str(template))
    try:
        workbook.sheet_by_name(mapping.template_sheet)
    finally:
        workbook.release_resources()
    if settings.template_path is None or not settings.template_path.is_file():
        raise FileNotFoundError(f"Configured template not found: {settings.template_path}")
    window = tk.Tk()
    window.withdraw()
    window.destroy()
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
        winreg.QueryValueEx(key, "")
    print("OK: Python, Tk, dependencies, configuration, mapping and bundled template")
    print("OK: Excel COM is registered (actual export must run in your Windows session)")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)
