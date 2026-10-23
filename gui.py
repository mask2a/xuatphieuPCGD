"""PCGD desktop UI: file/folder sources and official Excel/PDF exports."""

from __future__ import annotations

import queue
import threading
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Callable, Iterable

from analysis import AnalysisResult, analyze_sources
from household_grouper import Household
from settings import (
    Settings,
    load_settings,
    save_settings,
    reader_worker_count,
    export_worker_count,
)
from src.school_year_labels import build_school_year_labels
from src.template_mapper import TemplateMapping, load_template_mapping
from src.export_workflow import ExcelBatch, export_excel_batch, export_pdf_batch, export_existing_excel_pdf

BATCH_FEATURES_VERIFIED = True


PROJECT_ROOT = Path(__file__).resolve().parent
SETTINGS_PATH = PROJECT_ROOT / "config" / "settings.json"
MAPPING_PATH = PROJECT_ROOT / "config" / "template_mapping.json"
DEFAULT_TEMPLATE_PATH = PROJECT_ROOT / "templates" / "maubaocao.xls"

FIELD_LABELS = {
    "address": "Địa chỉ",
    "residence": "Diện cư trú",
    "priority_category": "Diện ưu tiên",
    "female": "Giới tính",
    "residence_status": "Tình trạng cư trú",
    "school_district": "Huyện đang học",
    "school_code": "Mã trường",
    "supplementary_education": "Bổ túc",
    "vocational_graduation_level": "Bậc tốt nghiệp nghề",
    "diploma_number": "Số bằng tốt nghiệp",
    "vocational_graduation_year": "Năm tốt nghiệp nghề",
    "completed_education": "Học xong",
    "completed_education_year": "Năm học xong",
    "dropped_out_class": "Lớp bỏ học",
    "dropped_out_year": "Năm bỏ học",
    "currently_studying_class": "Lớp đang học",
    "completed_class": "Lớp đã hoàn thành",
    "reliteracy_level": "Mức tái mù chữ",
    "disability_mobility": "Khuyết tật vận động",
    "disability_hearing_speech": "Khuyết tật nghe, nói",
    "disability_vision": "Khuyết tật nhìn",
    "disability_neurological": "Khuyết tật thần kinh/tâm thần",
    "disability_intellectual": "Khuyết tật trí tuệ",
    "disability_learning": "Khuyết tật học tập",
    "disability_autism": "Tự kỷ",
    "disability_other": "Khuyết tật khác",
    "disability_certified": "Có chứng nhận khuyết tật",
    "learning_capacity": "Khả năng học tập",
    "special_circumstances": "Hoàn cảnh đặc biệt",
    "special_circumstances_details": "Chi tiết hoàn cảnh đặc biệt",
    "notes": "Ghi chú",
}

MEMBER_DETAIL_FIELDS = (
    ("grade_group", "Khối nguồn (Q)"),
    ("class_name", "Lớp (R)"),
    ("school_district", "Huyện (S)"),
    ("school_name", "Trường (T)"),
    ("graduation_level", "Bậc tốt nghiệp (V)"),
    ("supplementary_education", "Bổ túc (W)"),
    ("graduation_year", "Năm tốt nghiệp (X)"),
)


def filter_households(
    households: Iterable[Household],
    query: str,
) -> tuple[Household, ...]:
    normalized_query = query.strip().casefold()
    if not normalized_query:
        return tuple(households)
    return tuple(
        household
        for household in households
        if normalized_query in str(household.household_code or "").casefold()
        or normalized_query
        in str(household.household_head_name or "").casefold()
    )


def export_one_household(
    household: Household,
    template_path: str | Path,
    output_root: str | Path,
    mapping_path: str | Path,
    *,
    settings_path: str | Path,
    commune_name: str | None = None,
    progress=None,
    source_name=None,
    cancel_event=None,
) -> ExcelBatch:
    """Export the selected household through the official, no-overwrite workflow."""
    if commune_name is None:
        commune_name = load_settings(settings_path).commune_name
    return export_excel_batch(
        (household,),
        template_path,
        output_root,
        mapping_path,
        commune_name=commune_name,
        progress=progress,
        source_name=source_name,
        cancel_event=cancel_event,
    )


class PCGDApp:
    def __init__(
        self,
        root: tk.Tk,
        *,
        settings_path: str | Path = SETTINGS_PATH,
        mapping_path: str | Path = MAPPING_PATH,
        show_dialogs: bool = True,
    ) -> None:
        self.root = root
        self.settings_path = Path(settings_path)
        self.mapping_path = Path(mapping_path)
        self.show_dialogs = show_dialogs
        self.analysis_result: AnalysisResult | None = None
        self.households: tuple[Household, ...] = ()
        self._analysis_commune: str | None = None
        self._analysis_source: Path | None = None
        self._household_by_id: dict[str, Household] = {}
        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._busy = False
        self._cancel_event = threading.Event()
        self._closing = False
        self._last_excel_batch = None

        settings = load_settings(self.settings_path)
        self.source_path = tk.StringVar(
            master=root,
            value=str(settings.source_path) if settings.source_path else "",
        )
        self.template_path = tk.StringVar(
            master=root,
            value=(
                str(settings.template_path)
                if settings.template_path
                else str(DEFAULT_TEMPLATE_PATH)
                if DEFAULT_TEMPLATE_PATH.is_file()
                else ""
            ),
        )
        self.output_path = tk.StringVar(master=root, value=str(settings.output_directory))
        self.commune_name = tk.StringVar(master=root, value=settings.commune_name)
        self.school_year = tk.StringVar(
            master=root,
            value="2025-2026",
        )
        self.search_query = tk.StringVar(master=root, value="")
        self.reader_workers = tk.StringVar(master=root, value=str(settings.reader.max_workers))
        self.pdf_validation = tk.BooleanVar(master=root, value=settings.pdf_validation)
        self.export_workers = tk.StringVar(master=root, value=str(settings.export_workers))
        self.household_total = tk.StringVar(master=root, value="0")
        self.member_total = tk.StringVar(master=root, value="0")
        self.warning_total = tk.StringVar(master=root, value="0")
        self.progress_text = tk.StringVar(master=root, value="Sẵn sàng")
        self.export_stats = tk.StringVar(master=root, value="Excel: 0 thành công / 0 lỗi · PDF: 0 thành công / 0 lỗi")

        self._configure_window()
        self._build_interface()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._refresh_export_buttons()

    def _configure_window(self) -> None:
        self.root.title("PCGD HỘ GIA ĐÌNH")
        self.root.geometry("1180x870")
        self.root.minsize(920, 730)
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("AppTitle.TLabel", font=("Segoe UI", 19, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("StatValue.TLabel", font=("Segoe UI", 18, "bold"), foreground="#174a7e")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", rowheight=27, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_interface(self) -> None:
        container = ttk.Frame(self.root, padding=(18, 14))
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text="PCGD HỘ GIA ĐÌNH", style="AppTitle.TLabel").pack(
            side="left"
        )
        ttk.Label(
            header,
            text="Đọc tệp/thư mục · rà soát hộ · xuất Excel/PDF",
            foreground="#566578",
        ).pack(side="left", padx=(16, 0), pady=(7, 0))

        files = ttk.LabelFrame(
            container,
            text="Tệp và cấu hình",
            padding=12,
            style="Section.TLabelframe",
        )
        files.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        files.columnconfigure(1, weight=1)
        files.columnconfigure(4, weight=1)
        ttk.Label(files, text="Tệp / thư mục nguồn").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(files, textvariable=self.source_path).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        source_buttons = ttk.Frame(files)
        source_buttons.grid(row=0, column=2, sticky="e")
        ttk.Button(source_buttons, text="Chọn tệp", command=self._browse_source).pack(side="left")
        ttk.Button(source_buttons, text="Chọn thư mục", command=self._browse_source_folder).pack(side="left", padx=(4, 0))
        self._add_path_row(
            files, 1, "Biểu mẫu Excel", self.template_path, self._browse_template
        )
        self._add_path_row(
            files, 2, "Thư mục kết quả", self.output_path, self._browse_output
        )
        ttk.Label(files, text="Xã/phường").grid(
            row=0, column=3, padx=(16, 6), sticky="w"
        )
        ttk.Entry(files, textvariable=self.commune_name, width=24).grid(
            row=0, column=4, sticky="ew"
        )
        ttk.Label(files, text="Năm học").grid(
            row=1, column=3, padx=(16, 6), sticky="w"
        )
        ttk.Entry(files, textvariable=self.school_year, width=24).grid(
            row=1, column=4, sticky="ew"
        )
        ttk.Button(files, text="Lưu cấu hình", command=self.save_configuration).grid(
            row=2, column=4, sticky="e", pady=(8, 0)
        )
        ttk.Label(files, text="Luồng đọc (1–8)").grid(row=3, column=3, padx=(16, 6), sticky="w")
        ttk.Spinbox(files, from_=1, to=8, textvariable=self.reader_workers, width=6).grid(row=3, column=4, sticky="w")

        stats = ttk.Frame(container)
        stats.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        for column in range(3):
            stats.columnconfigure(column, weight=1)
        self._stat_card(stats, 0, "TỔNG SỐ HỘ", self.household_total)
        self._stat_card(stats, 1, "TỔNG NHÂN KHẨU", self.member_total)
        self._stat_card(stats, 2, "CẢNH BÁO", self.warning_total)

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        self.read_button = ttk.Button(
            actions,
            text="Đọc dữ liệu",
            style="Primary.TButton",
            command=self.read_data,
        )
        self.read_button.pack(side="left")
        self.view_members_button = ttk.Button(
            actions,
            text="Xem thành viên",
            command=self.view_selected_members,
        )
        self.view_members_button.pack(side="left", padx=(8, 0))
        self.mapping_button = ttk.Button(
            actions,
            text="Cấu hình ánh xạ",
            command=self.show_mapping_window,
        )
        self.mapping_button.pack(side="left", padx=(8, 0))
        self.export_selected_button = ttk.Button(
            actions,
            text="Xuất hộ đã chọn",
            style="Primary.TButton",
            command=self.export_selected,
        )
        self.export_selected_button.pack(side="left", padx=(8, 0))
        self.export_all_button = ttk.Button(
            actions,
            text="Xuất tất cả Excel",
            state="disabled",
            command=self.export_all_excel,
        )
        self.export_all_button.pack(side="left", padx=(8, 0))
        self.export_pdf_button = ttk.Button(
            actions,
            text="Xuất PDF và tổng hợp",
            state="disabled",
            command=self.export_all_pdf,
        )
        self.export_pdf_button.pack(side="left", padx=(8, 0))
        self.existing_pdf_button = ttk.Button(
            files, text="Excel có sẵn → PDF", command=self.export_existing_pdf,
        )
        self.existing_pdf_button.grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.open_output_button = ttk.Button(
            actions,
            text="Mở thư mục kết quả",
            command=self.open_output_directory,
        )
        self.open_output_button.pack(side="right")
        self.pdf_validation_check = ttk.Checkbutton(
            files, text="Kiểm tra khi xuất PDF", variable=self.pdf_validation,
        )
        self.pdf_validation_check.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(files, text="Xuất song song (1–4)").grid(row=4, column=3, sticky="w")
        self.export_workers_input = ttk.Spinbox(files, from_=1, to=4, textvariable=self.export_workers, width=6)
        self.export_workers_input.grid(row=4, column=4, sticky="w")

        list_frame = ttk.LabelFrame(
            container,
            text="Danh sách hộ — chọn một hộ để xem hoặc xuất Excel",
            padding=10,
            style="Section.TLabelframe",
        )
        list_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 9))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)
        search_bar = ttk.Frame(list_frame)
        search_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(search_bar, text="Tìm mã hộ hoặc tên chủ hộ:").pack(side="left")
        search_entry = ttk.Entry(search_bar, textvariable=self.search_query, width=38)
        search_entry.pack(side="left", padx=(8, 0))
        self.search_query.trace_add("write", self._on_search_changed)

        columns = ("code", "head", "members", "alerts")
        self.household_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.household_tree.heading("code", text="Mã hộ")
        self.household_tree.heading("head", text="Chủ hộ")
        self.household_tree.heading("members", text="Thành viên")
        self.household_tree.heading("alerts", text="Cảnh báo")
        self.household_tree.column("code", width=170, minwidth=110)
        self.household_tree.column("head", width=380, minwidth=180)
        self.household_tree.column("members", width=110, minwidth=85, anchor="center")
        self.household_tree.column("alerts", width=100, minwidth=75, anchor="center")
        vertical = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.household_tree.yview,
        )
        horizontal = ttk.Scrollbar(
            list_frame,
            orient="horizontal",
            command=self.household_tree.xview,
        )
        self.household_tree.configure(
            yscrollcommand=vertical.set,
            xscrollcommand=horizontal.set,
        )
        self.household_tree.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        self.household_tree.bind(
            "<<TreeviewSelect>>",
            lambda _event: self._refresh_export_buttons(),
        )
        self.household_tree.bind(
            "<Double-1>",
            lambda _event: self.view_selected_members(),
        )

        footer = ttk.LabelFrame(container, text="Tiến trình và nhật ký", padding=8)
        footer.grid(row=5, column=0, sticky="ew")
        footer.columnconfigure(1, weight=1)
        ttk.Label(footer, textvariable=self.progress_text, width=24).grid(
            row=0, column=0, sticky="w"
        )
        self.progress = ttk.Progressbar(
            footer,
            mode="indeterminate",
            length=160,
        )
        self.progress.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.stop_button = ttk.Button(footer, text="Dừng xuất", state="disabled", command=self.stop_export)
        self.stop_button.grid(row=0, column=2, padx=(8, 0))
        self.log_text = tk.Text(
            footer,
            height=10,
            wrap="word",
            font=("Consolas", 9),
            state="disabled",
        )
        self.log_text.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(7, 0))
        log_scroll = ttk.Scrollbar(footer, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=1, column=2, sticky="ns", pady=(7, 0))
        self.log_text.configure(yscrollcommand=log_scroll.set)
        ttk.Label(footer, textvariable=self.export_stats).grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 0))
        self._log("Sẵn sàng. Chọn một tệp hoặc thư mục nguồn; Excel COM chỉ chạy khi xuất Excel/PDF.")

    @staticmethod
    def _add_path_row(
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        browse_command: Callable[[], None],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=variable).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=(0, 6),
            pady=3,
        )
        ttk.Button(parent, text="Chọn...", command=browse_command).grid(
            row=row,
            column=2,
            sticky="e",
            pady=3,
        )

    @staticmethod
    def _stat_card(
        parent: ttk.Frame,
        column: int,
        title: str,
        variable: tk.StringVar,
    ) -> None:
        card = ttk.LabelFrame(parent, text=title, padding=(12, 7))
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, textvariable=variable, style="StatValue.TLabel").pack(
            anchor="w"
        )

    def _browse_source(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Chọn tệp Excel nguồn",
            filetypes=[("Tệp Excel", "*.xls *.xlsx"), ("Tất cả tệp", "*.*")],
        )
        if selected:
            self.source_path.set(selected)

    def _browse_template(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="Chọn biểu mẫu Excel",
            filetypes=[("Biểu mẫu Excel", "*.xls *.xlsx"), ("Tất cả tệp", "*.*")],
        )
        if selected:
            self.template_path.set(selected)

    def _browse_source_folder(self) -> None:
        selected = filedialog.askdirectory(parent=self.root, title="Chọn thư mục Excel nguồn (gồm thư mục con)", mustexist=True)
        if selected:
            self.source_path.set(selected)

    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(
            parent=self.root,
            title="Chọn thư mục lưu kết quả",
            mustexist=True,
        )
        if selected:
            self.output_path.set(selected)

    def save_configuration(self) -> None:
        try:
            build_school_year_labels(self.school_year.get())
            saved = save_settings(
                self.settings_path,
                commune_name=self.commune_name.get(),
                school_year=self.school_year.get(),
                source_path=self.source_path.get(),
                template_path=self.template_path.get(),
                output_directory=self.output_path.get(),
                reader_workers=self.reader_workers.get(),
                pdf_validation=self.pdf_validation.get(),
                export_workers=self.export_workers.get(),
            )
        except (OSError, TypeError, ValueError) as error:
            self._show_error("Không lưu được cấu hình", str(error))
            return
        self.school_year.set(saved.school_year or "")
        self.commune_name.set(saved.commune_name)
        self._log("Đã lưu cấu hình xã/phường, năm học và đường dẫn.")
        if self.show_dialogs:
            messagebox.showinfo(
                "Đã lưu cấu hình",
                "Cấu hình xã/phường, năm học và đường dẫn đã được lưu.",
                parent=self.root,
            )

    def read_data(self) -> None:
        if self._busy:
            return
        source_text = self.source_path.get().strip()
        if not source_text:
            self._show_error("Chưa chọn nguồn", "Hãy chọn một tệp Excel hoặc thư mục nguồn.")
            return
        source = Path(source_text).expanduser()
        if not (source.is_file() or source.is_dir()):
            self._show_error("Chưa chọn nguồn", "Hãy chọn một tệp Excel hoặc thư mục nguồn hợp lệ.")
            return
        if source.is_file() and source.suffix.casefold() not in {".xls", ".xlsx"}:
            self._show_error("Định dạng không hỗ trợ", "Tệp nguồn phải có đuôi .xls hoặc .xlsx.")
            return
        try:
            build_school_year_labels(self.school_year.get())
            settings = self._settings_for_gui()
        except (OSError, ValueError) as error:
            self._show_error("Cấu hình không hợp lệ", str(error))
            return
        if source.resolve().is_relative_to(settings.output_directory.resolve()):
            self._show_error("Nguồn nằm trong thư mục kết quả", "Chọn thư mục kết quả riêng, không chứa nguồn, để tránh đọc lại các file đã xuất.")
            return

        self.analysis_result = None
        self._analysis_source = None
        self._last_excel_batch = None
        self.households = ()
        self._household_by_id.clear()
        self._refresh_household_tree()
        self._refresh_export_buttons()
        self.member_total.set("0")
        self.household_total.set("0")
        self.warning_total.set("0")
        self._log(f"Đang đọc nguồn: {source}; tối đa {settings.reader.max_workers} luồng.")

        def work() -> AnalysisResult:
            self._events.put(("progress", "Đang đọc và nhóm dữ liệu..."))
            return analyze_sources((source,), settings, progress=self._batch_progress)

        self._start_background("Đang đọc dữ liệu", work, self._on_analysis_complete)

    def _on_analysis_complete(self, result: AnalysisResult) -> None:
        self.analysis_result = result
        self.households = result.grouping_result.households
        self._analysis_commune = result.settings.commune_name.strip()
        self._analysis_source = result.settings.source_path
        self.member_total.set(str(len(result.members)))
        self.household_total.set(str(len(self.households)))
        self.warning_total.set(
            str(
                len(result.grouping_result.warnings)
                + len(result.read_result.errors)
            )
        )
        self._refresh_household_tree()
        self._refresh_export_buttons()
        self._log(
            f"Đã đọc {len(result.read_result.records)} nhân khẩu, "
            f"{len(self.households)} hộ; "
            f"{len(result.read_result.errors)} tệp lỗi."
        )
        for path in result.read_result.files_read:
            self._log(f"Đã đọc: {path}")
        for error in result.read_result.errors:
            self._log(f"LỖI {error.file}: {error.message}")
        for skipped in result.read_result.skipped_files:
            self._log(f"Bỏ qua tệp/thư mục không thuộc nguồn dữ liệu: {skipped}")
        if not result.read_result.files_read and not result.read_result.errors:
            self._show_warning("Không có tệp nguồn", "Không tìm thấy Excel nguồn phù hợp. Biểu mẫu, tệp tạm và thư mục kết quả được bỏ qua.")
        if result.read_result.errors:
            self._show_warning(
                "Đọc dữ liệu có lỗi",
                f"Đã đọc được {len(result.members)} nhân khẩu và "
                f"{len(self.households)} hộ. Có {len(result.read_result.errors)} "
                "tệp không đọc được; xem nhật ký để biết chi tiết.",
            )

    def _refresh_household_tree(self) -> None:
        self.household_tree.delete(*self.household_tree.get_children())
        self._household_by_id.clear()
        visible = filter_households(self.households, self.search_query.get())
        for household in visible:
            identifier = household.household_id
            self._household_by_id[identifier] = household
            self.household_tree.insert(
                "",
                "end",
                iid=identifier,
                values=(
                    household.household_code or "(thiếu mã)",
                    household.household_head_name or "(chưa xác định)",
                    len(household.members),
                    len(household.warnings),
                ),
            )

    def _on_search_changed(self, *_args) -> None:
        self._refresh_household_tree()

    def _selected_household(self) -> Household | None:
        selection = self.household_tree.selection()
        if not selection:
            return None
        return self._household_by_id.get(selection[0])

    def _refresh_export_buttons(self) -> None:
        enabled = (
            not self._busy
            and self._selected_household() is not None
            and self.analysis_result is not None
        )
        self.export_selected_button.configure(state="normal" if enabled else "disabled")
        self.view_members_button.configure(state="normal" if enabled else "disabled")
        self.read_button.configure(state="disabled" if self._busy else "normal")
        self.mapping_button.configure(state="disabled" if self._busy else "normal")
        self.open_output_button.configure(state="disabled" if self._busy else "normal")
        loaded = BATCH_FEATURES_VERIFIED and not self._busy and bool(self.households) and self.analysis_result is not None
        self.export_all_button.configure(state="normal" if loaded else "disabled")
        self.export_pdf_button.configure(state="normal" if loaded and self._last_excel_batch is not None and self._last_excel_batch.accepted else "disabled")
        self.pdf_validation_check.configure(state="disabled" if self._busy else "normal")
        self.export_workers_input.configure(state="disabled" if self._busy else "normal")
        self.existing_pdf_button.configure(state="disabled" if self._busy else "normal")

    def _batch_context(self):
        if self._busy or not BATCH_FEATURES_VERIFIED:
            return None
        commune = self.commune_name.get().strip()
        if not commune:
            self._show_error("Thiếu xã/phường", "Nhập xã/phường trước khi xuất.")
            return None
        if self._analysis_source is not None and Path(self.source_path.get()).expanduser().resolve() != self._analysis_source.resolve():
            self._show_warning("Cần đọc lại dữ liệu", "Nguồn đã thay đổi. Hãy đọc lại trước khi xuất.")
            return None
        if not self.households or self.analysis_result is None or not commune or self._analysis_commune != commune:
            self._show_warning("Cần đọc lại dữ liệu", "Nhập xã/phường và đọc dữ liệu trước khi xuất; đọc lại nếu xã/phường đã đổi.")
            return None
        output = Path(self.output_path.get()).expanduser()
        if not output.is_dir():
            self._show_error("Thư mục không hợp lệ", "Chọn thư mục kết quả đã tồn tại.")
            return None
        return commune, output

    def _batch_progress(self, done, total, message):
        self._events.put(("batch_progress", (done, total, message)))

    def _export_source_name(self):
        if len({household.source_file for household in self.households if hasattr(household, "source_file")}) > 1:
            return None  # A separate output group is named after each input file.
        source = Path(self.source_path.get()).expanduser()
        return source.name if source.is_dir() else source.stem

    def export_all_excel(self):
        context = self._batch_context()
        if context is None:
            return
        try:
            workers = export_worker_count(self.export_workers.get())
        except ValueError as error:
            self._show_error("Số tiến trình không hợp lệ", str(error))
            return
        commune, output = context
        template = Path(self.template_path.get()).expanduser()
        if not template.is_file():
            self._show_error("Chưa chọn biểu mẫu", "Chọn biểu mẫu Excel hợp lệ.")
            return
        households = tuple(self.households)
        mapping = self.mapping_path
        source_name = self._export_source_name()
        self._last_excel_batch = None
        self._log(f"Bắt đầu xuất Excel: {len(households)} hộ. Kết quả chia theo tên nguồn; thêm hậu tố khi trùng, không ghi đè.")
        self._log(f"Xuất song song: tối đa {workers} tệp nguồn cùng lúc.")
        self._start_background("Đang xuất Excel", lambda: export_excel_batch(households, template, output, mapping, commune_name=commune, progress=self._batch_progress, source_name=source_name, workers=workers, cancel_event=self._cancel_event), self._on_batch_excel_complete)

    def _on_batch_excel_complete(self, batch):
        self._last_excel_batch = batch
        report = batch.report
        self.export_stats.set(f"Excel: {report.success_count} thành công / {report.failure_count} lỗi · PDF: chưa xuất")
        self._log(f"Excel: {report.success_count} thành công, {report.failure_count} lỗi. Thư mục: {batch.directory}")
        for group in getattr(batch, "groups", ()):
            self._log(f"Nguồn {group.source_name}: {group.directory}")
        for failure in report.failures:
            self._log(f"LỖI Excel hộ {failure.household_code}: {failure.error}")
        for warning in report.warnings:
            self._log(f"CẢNH BÁO hộ {warning.household_code}: {warning.message}")

    def export_all_pdf(self):
        context = self._batch_context()
        if context is None:
            return
        try:
            workers = export_worker_count(self.export_workers.get())
        except ValueError as error:
            self._show_error("Số tiến trình không hợp lệ", str(error))
            return
        if self._last_excel_batch is None or not self._last_excel_batch.accepted:
            self._show_warning("Chưa có Excel hợp lệ", "Xuất Excel trước; chỉ phiếu đã nghiệm thu mới được chuyển PDF.")
            return
        batch = self._last_excel_batch
        output = context[1]
        validate = self.pdf_validation.get()
        mode = "kiểm tra đầy đủ" if validate else "bỏ kiểm tra lại Excel và nội dung PDF"
        self._log(f"Bắt đầu PDF: {len(batch.accepted)} phiếu; {mode}.")
        self._log(f"Xuất song song: tối đa {workers} tệp nguồn cùng lúc.")
        self._start_background("Đang xuất PDF", lambda: export_pdf_batch(batch, output, progress=self._batch_progress, validate=validate, workers=workers, cancel_event=self._cancel_event), self._on_batch_pdf_complete)

    def export_existing_pdf(self):
        if self._busy:
            return
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Chọn các file Excel để xuất PDF",
            filetypes=[("Excel", "*.xls *.xlsx *.xlsm")],
        )
        if not paths:
            return
        output = self.output_path.get().strip()
        if not output:
            output = filedialog.askdirectory(parent=self.root, title="Chọn thư mục kết quả PDF")
        if not output:
            return
        self._log(f"Xuất PDF từ {len(paths)} file Excel có sẵn, dùng vùng in đã lưu trong file.")
        self._start_background(
            "Đang xuất PDF từ Excel có sẵn",
            lambda: export_existing_excel_pdf(paths, output, progress=self._batch_progress, cancel_event=self._cancel_event),
            self._on_batch_pdf_complete,
        )

    def _on_batch_pdf_complete(self, report):
        excel = self._last_excel_batch.report if self._last_excel_batch else None
        self._show_pdf_result(report, excel)

    def _show_pdf_result(self, report, excel):
        from types import SimpleNamespace
        excel = excel or SimpleNamespace(success_count=0, failure_count=0)
        self.export_stats.set(f"Excel: {excel.success_count} thành công / {excel.failure_count} lỗi · PDF: {len(report.successes)} thành công / {len(report.failures)} lỗi · Tổng hợp: {report.total_pages} trang")
        for item in report.successes:
            self._log(f"PDF hộ {item.household_code}: {item.pages} trang — {item.path}")
        for item in report.failures:
            self._log(f"LỖI PDF hộ {item.household_code}: {item.error}")
        paths = getattr(report, "merged_paths", ()) or ((report.merged_path,) if report.merged_path else ())
        for path in paths:
            self._log(f"PDF tổng hợp: {path}; thứ tự mã hộ tăng dần trong từng nguồn.")
        if report.merge_error:
            self._show_error("Không gộp được PDF", report.merge_error)

    def view_selected_members(self) -> None:
        household = self._selected_household()
        if household is None:
            self._show_warning("Chưa chọn hộ", "Hãy chọn một hộ trong danh sách.")
            return
        window = tk.Toplevel(self.root)
        window.title(f"Thành viên hộ {household.household_code or '(thiếu mã)'}")
        window.geometry("950x470")
        window.minsize(760, 360)
        window.transient(self.root)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        columns = ("number", "name", "relation", "class", "school", "graduation")
        tree = ttk.Treeview(window, columns=columns, show="headings")
        headings = {
            "number": "STT",
            "name": "Họ tên",
            "relation": "Quan hệ",
            "class": "Lớp nguồn (R)",
            "school": "Trường nguồn (T)",
            "graduation": "Tốt nghiệp (V/X)",
        }
        widths = {
            "number": 55,
            "name": 180,
            "relation": 120,
            "class": 150,
            "school": 230,
            "graduation": 140,
        }
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=widths[column], minwidth=60)
        vertical = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vertical.grid(row=0, column=1, sticky="ns", pady=12)
        horizontal.grid(row=1, column=0, sticky="ew", padx=(12, 0))
        for index, member in enumerate(household.members, start=1):
            fields = member.fields
            graduation = " / ".join(
                str(fields[name])
                for name in ("graduation_level", "graduation_year")
                if fields.get(name) is not None and str(fields[name]).strip()
            )
            tree.insert(
                "",
                "end",
                values=(
                    index,
                    member.member_name,
                    member.relationship_to_head or "",
                    fields.get("class_name") or "",
                    fields.get("school_name") or "",
                    graduation,
                ),
            )
        warning_frame = ttk.LabelFrame(
            window,
            text=f"Cảnh báo ({len(household.warnings)})",
            padding=8,
        )
        warning_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=8)
        warning_text = tk.Text(warning_frame, height=4, wrap="word")
        warning_text.pack(fill="both", expand=True)
        for warning in household.warnings:
            warning_text.insert("end", f"• {warning.code}: {warning.message}\n")
        warning_text.configure(state="disabled")

    def show_mapping_window(self) -> None:
        try:
            try:
                mapping = load_template_mapping(self.mapping_path)
            except (OSError, ValueError, KeyError) as error:
                self._show_error("Không đọc được cấu hình ánh xạ", str(error))
                return
            settings = load_settings(self.settings_path)
        except (OSError, ValueError, KeyError) as error:
            self._show_error("Không đọc được cấu hình ánh xạ", str(error))
            return
        window = tk.Toplevel(self.root)
        window.title("Cấu hình ánh xạ — chỉ xem các trường chưa duyệt")
        window.geometry("1050x540")
        window.minsize(850, 400)
        window.transient(self.root)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(
            window,
            text=(
                "Các mục dưới đây chưa được duyệt và không được phép tự động ghi. "
                "Cửa sổ này chỉ để xem; chọn ô không làm thay đổi mapping."
            ),
            wraplength=1000,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=10)
        columns = ("column", "field", "target", "status", "notes")
        tree = ttk.Treeview(window, columns=columns, show="headings")
        headings = {
            "column": "Cột nguồn",
            "field": "Tên trường nguồn",
            "target": "Ô đích / quy tắc",
            "status": "Trạng thái",
            "notes": "Ghi chú",
        }
        widths = {
            "column": 100,
            "field": 230,
            "target": 185,
            "status": 110,
            "notes": 420,
        }
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=widths[column], minwidth=75)
        vertical = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(window, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(0, 12))
        vertical.grid(row=1, column=1, sticky="ns", pady=(0, 12))
        horizontal.grid(row=2, column=0, sticky="ew", padx=(12, 0))
        rows = _unapproved_mapping_rows(mapping, settings.source_columns)
        for index, row in enumerate(rows):
            tree.insert("", "end", iid=f"unapproved-{index}", values=row)

    def export_selected(self) -> None:
        if self._batch_context() is None:
            return
        household = self._selected_household()
        if household is None:
            self._show_warning("Chưa chọn hộ", "Hãy chọn đúng một hộ để xuất Excel.")
            return
        template = Path(self.template_path.get()).expanduser()
        if not template.is_file():
            self._show_error("Chưa chọn biểu mẫu", "Hãy chọn tệp biểu mẫu Excel hợp lệ.")
            return
        output_root = Path(self.output_path.get()).expanduser()
        if not output_root.is_dir():
            self._show_error(
                "Thư mục kết quả không hợp lệ",
                "Hãy chọn một thư mục kết quả đã tồn tại.",
            )
            return
        commune = self.commune_name.get().strip()
        if not commune:
            self._show_error(
                "Thiếu xã/phường",
                "Cần nhập và lưu xã/phường trước khi xuất.",
            )
            return
        try:
            build_school_year_labels(self.school_year.get())
        except ValueError as error:
            self._show_error("Năm học không hợp lệ", str(error))
            return
        if self.analysis_result is None or self._analysis_commune != commune:
            self._show_warning(
                "Cần đọc lại dữ liệu",
                "Xã/phường đã thay đổi sau lần đọc dữ liệu. Hãy đọc lại để gắn "
                "đúng địa bàn cho các hộ trước khi xuất.",
            )
            return
        try:
            save_settings(
                self.settings_path,
                commune_name=commune,
                school_year=self.school_year.get(),
            )
        except (OSError, TypeError, ValueError) as error:
            self._show_error("Không lưu được cấu hình", str(error))
            return
        self._last_excel_batch = None
        self._log(f"Bắt đầu xuất hộ {household.household_code or '(thiếu mã)'} vào thư mục kết quả chính thức.")
        source_name = self._export_source_name()

        def work() -> ExcelBatch:
            return export_one_household(
                household,
                template,
                output_root,
                self.mapping_path,
                settings_path=self.settings_path,
                commune_name=commune,
                progress=self._batch_progress,
                source_name=source_name,
                cancel_event=self._cancel_event,
            )

        self._start_background("Đang xuất hộ đã chọn", work, self._on_batch_excel_complete)

    def open_output_directory(self) -> None:
        directory = Path(self.output_path.get()).expanduser()
        if not directory.is_dir():
            self._show_error(
                "Thư mục không tồn tại",
                "Hãy chọn một thư mục kết quả đã tồn tại.",
            )
            return
        try:
            import os
            os.startfile(str(directory))
        except OSError as error:
            self._show_error("Không mở được thư mục", str(error))

    def _settings_for_gui(self) -> Settings:
        settings = load_settings(self.settings_path)
        return replace(
            settings,
            commune_name=self.commune_name.get().strip(),
            school_year=self.school_year.get().strip(),
            output_directory=Path(self.output_path.get()).expanduser(),
            source_path=Path(self.source_path.get()).expanduser().resolve(),
            template_path=Path(self.template_path.get()).expanduser().resolve(),
            reader=replace(settings.reader, max_workers=reader_worker_count(self.reader_workers.get())),
        )

    def _start_background(
        self,
        description: str,
        operation: Callable[[], object],
        on_success: Callable[[object], None],
    ) -> None:
        if self._busy:
            return
        self._cancel_event.clear()
        self._busy = True
        self.stop_button.configure(state="normal" if on_success == self._on_batch_excel_complete or on_success == self._on_batch_pdf_complete else "disabled")
        self.progress.configure(mode="indeterminate")
        self.progress_text.set(description)
        self.progress.start(12)
        self._refresh_export_buttons()

        def run() -> None:
            try:
                result = operation()
            except Exception as error:
                self._events.put(("error", error))
            else:
                self._events.put(("success", (on_success, result)))
            finally:
                self._events.put(("finished", None))

        threading.Thread(target=run, name="pcgd-background-task", daemon=False).start()
        self.root.after(100, self._drain_events)

    def stop_export(self) -> None:
        if self._busy:
            self._cancel_event.set()
            self.stop_button.configure(state="disabled")
            self.progress_text.set("Đang dừng sau hộ hiện tại…")
            self._log("Đã yêu cầu dừng xuất; đang đóng workbook và giữ các file đã hoàn tất.")

    def _drain_events(self) -> None:
        if self._closing:
            return
        # Yield to Tk regularly when a large folder produces many log events.
        for _ in range(100):
            try:
                kind, payload = self._events.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                self.progress_text.set(str(payload))
            elif kind == "batch_progress":
                done, total, message = payload
                self.progress.stop()
                self.progress.configure(mode="determinate", maximum=max(1, total), value=done)
                self.progress_text.set(message)
                if self._cancel_event.is_set():
                    self.progress_text.set("Đang dừng sau hộ hiện tại…")
                self._log(message)
            elif kind == "success":
                callback, result = payload
                callback(result)
            elif kind == "error":
                error = payload
                self._log(f"LỖI: {type(error).__name__}: {error}")
                self._show_error("Không thể hoàn tất tác vụ", str(error))
            elif kind == "finished":
                self._busy = False
                self.stop_button.configure(state="disabled")
                self.progress.stop()
                self.progress_text.set("Sẵn sàng")
                if self._cancel_event.is_set():
                    self.progress_text.set("Đã dừng xuất")
                    self._log("Đã dừng xuất an toàn. Các file đã hoàn tất được giữ lại.")
                self._refresh_export_buttons()
        if self._busy or not self._events.empty():
            self.root.after(100, self._drain_events)

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _show_error(self, title: str, message: str) -> None:
        self._log(f"{title}: {message}")
        if self.show_dialogs:
            messagebox.showerror(title, message, parent=self.root)

    def _show_warning(self, title: str, message: str) -> None:
        self._log(f"{title}: {message}")
        if self.show_dialogs:
            messagebox.showwarning(title, message, parent=self.root)

    def _on_close(self) -> None:
        if self._busy:
            self._show_warning(
                "Tác vụ đang chạy",
                "Nếu đang xuất, hãy bấm Dừng xuất và đợi hộ hiện tại hoàn tất "
                "để đóng workbook và giải phóng Excel an toàn.",
            )
            return
        self._closing = True
        self.root.destroy()


def _unapproved_mapping_rows(
    mapping: TemplateMapping,
    source_columns: dict[str, str],
) -> tuple[tuple[str, str, str, str, str], ...]:
    rows = []
    groups = (
        (mapping.unmapped_household_fields, "Dữ liệu hộ"),
        (mapping.unmapped_member_fields, "Dữ liệu thành viên"),
    )
    for fields, category in groups:
        for field_name, notes in fields.items():
            source_column = source_columns.get(field_name, "—")
            field_label = FIELD_LABELS.get(field_name, field_name.replace("_", " ").title())
            rows.append(
                (
                    f"{source_column} · {category}",
                    field_label,
                    "Chưa xác định / chưa được duyệt",
                    "Chưa duyệt",
                    notes,
                )
            )
    return tuple(rows)


def main() -> None:
    root = tk.Tk()
    PCGDApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
