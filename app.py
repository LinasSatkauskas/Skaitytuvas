import os
import calendar
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import ttkbootstrap as ttk

from taskworkbook.constants import (
    ACCENT,
    APP_BG,
    CHECKBOX_COLUMNS,
    FILTERABLE_SHEETS,
    FILTER_ALL,
    MUTED,
    SHEET_HEADERS,
)
from taskworkbook.workbook_service import WorkbookService

FIXED_SHEETS = ("Tasks", "CompletedTasks", "TaskHistory")
TEXT = "#0f172a"


class TaskWorkbookApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Task Workbook")
        self.geometry("1280x800")
        self.minsize(1100, 680)
        self.configure(bg=APP_BG)

        self.workbook = None
        self.worksheet = None
        self.excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tasks_data.xlsx")

        self.current_sheet_name = ""
        self.headers = []
        self.all_rows = []
        self.filtered_rows = []
        self.current_row = None
        self.current_cell = None
        self.current_edit = None
        self.checkbox_widgets = {}
        self.checkbox_vars = {}
        self.worker_options = [FILTER_ALL]
        self.day_options = [FILTER_ALL]

        self.worker_filter_var = tk.StringVar(value=FILTER_ALL)
        self.day_filter_var = tk.StringVar(value=FILTER_ALL)
        self.created_on_filter_var = tk.StringVar(value=FILTER_ALL)
        self.created_on_filter_display_var = tk.StringVar(value="Any date")
        self.sheet_var = tk.StringVar(value="Sheet: -")
        self.count_var = tk.StringVar(value="Rows: 0")
        self.status_var = tk.StringVar(value="Loading workbook...")
        self.row_pos_var = tk.StringVar(value="Row: -")
        self.created_on_calendar_window = None
        self.service = WorkbookService(self.excel_path, self._weekday_from_value, self._stringify)

        self._suppress_select = False

        self._configure_style()
        self._build_ui()
        self._load_or_create_workbook()

    def _configure_style(self):
        style = ttk.Style()
        style.configure("TLabel", font=("Segoe UI", 10), foreground=TEXT)
        style.configure("Muted.TLabel", foreground=MUTED)
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10), borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))
        style.map(
            "Treeview",
            background=[("selected", ACCENT)],
            foreground=[("selected", "white")],
        )
        style.configure("Task.TLabelframe", padding=14)
        style.configure("Task.TLabelframe.Label", font=("Segoe UI Semibold", 10), foreground=TEXT)

    def _build_ui(self):
        hero = tk.Frame(self, bg=ACCENT, height=104)
        hero.pack(fill="x")
        hero.pack_propagate(False)

        hero_inner = tk.Frame(hero, bg=ACCENT)
        hero_inner.pack(fill="both", expand=True)

        tk.Label(hero_inner, text="Task Workbook", font=("Segoe UI Semibold", 20), fg="white", bg=ACCENT).pack(
            anchor="w", padx=22, pady=(16, 0)
        )
        tk.Label(
            hero_inner,
            text="Edit the table directly. Double-click any cell to change it in place.",
            font=("Segoe UI", 10),
            fg="#dbeafe",
            bg=ACCENT,
        ).pack(anchor="w", padx=22, pady=(4, 0))

        toolbar = ttk.Frame(self, padding=(16, 14, 16, 10))
        toolbar.pack(fill="x")

        toolbar_top = ttk.Frame(toolbar)
        toolbar_top.pack(fill="x")

        action_buttons = ttk.Frame(toolbar_top)
        action_buttons.pack(side="left")

        ttk.Button(action_buttons, text="Save", command=self.save_workbook, bootstyle="primary").pack(side="left")
        ttk.Button(action_buttons, text="New Row", command=self.add_row, bootstyle="success").pack(side="left", padx=(8, 0))
        ttk.Button(action_buttons, text="Delete Selected", command=self.delete_selected_rows, bootstyle="danger").pack(side="left", padx=(8, 0))
        self.submit_button = ttk.Button(action_buttons, text="Submit", command=self.submit_completed_rows, bootstyle="warning")
        self.submit_button.pack(side="left", padx=(8, 0))

        toolbar_top.columnconfigure(1, weight=1)
        toolbar_top.columnconfigure(3, weight=1)

        ttk.Frame(toolbar_top).pack(side="left", fill="x", expand=True)

        review_buttons = ttk.Frame(toolbar_top)
        review_buttons.pack(side="left")
        self.approve_button = ttk.Button(review_buttons, text="Approve", command=self.approve_completed_rows, bootstyle="success")
        self.approve_button.pack(side="left")
        self.deny_button = ttk.Button(review_buttons, text="Deny", command=self.deny_completed_rows, bootstyle="danger")
        self.deny_button.pack(side="left", padx=(8, 0))

        ttk.Frame(toolbar_top).pack(side="left", fill="x", expand=True)

        sheet_buttons = ttk.Frame(toolbar_top)
        sheet_buttons.pack(side="left")
        ttk.Button(sheet_buttons, text="Tasks", command=lambda: self.open_sheet("Tasks"), bootstyle="secondary").pack(side="left")
        ttk.Button(sheet_buttons, text="CompletedTasks", command=lambda: self.open_sheet("CompletedTasks"), bootstyle="secondary").pack(side="left", padx=(8, 0))
        ttk.Button(sheet_buttons, text="TaskHistory", command=lambda: self.open_sheet("TaskHistory"), bootstyle="secondary").pack(side="left", padx=(8, 0))

        info = ttk.Frame(toolbar)
        info.pack(fill="x", pady=(10, 0))
        ttk.Label(info, textvariable=self.sheet_var).pack(anchor="w")
        ttk.Label(info, textvariable=self.count_var, style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Label(info, textvariable=self.row_pos_var).pack(anchor="e")

        main = ttk.Frame(self, padding=(16, 0, 16, 16))
        main.pack(fill="both", expand=True)

        card = ttk.Labelframe(main, text="Rows", style="Task.TLabelframe", padding=14)
        card.pack(fill="both", expand=True)

        filter_row = ttk.Frame(card)
        filter_row.pack(fill="x")
        filter_row.columnconfigure(1, weight=1)

        filters_frame = ttk.Frame(filter_row)
        filters_frame.pack(side="left", fill="x", expand=True)

        worker_frame = ttk.Frame(filters_frame)
        worker_frame.pack(side="left", padx=(0, 10))
        ttk.Label(worker_frame, text="Worker", style="Muted.TLabel").pack(anchor="w")
        self.worker_filter = ttk.Combobox(worker_frame, textvariable=self.worker_filter_var, state="readonly", width=24)
        self.worker_filter.pack(anchor="w")

        day_frame = ttk.Frame(filters_frame)
        day_frame.pack(side="left", padx=(0, 10))
        ttk.Label(day_frame, text="DayOfWeek", style="Muted.TLabel").pack(anchor="w")
        self.day_filter = ttk.Combobox(day_frame, textvariable=self.day_filter_var, state="readonly", width=14)
        self.day_filter.pack(anchor="w")

        created_on_frame = ttk.Frame(filters_frame)
        created_on_frame.pack(side="left", padx=(0, 10))
        ttk.Label(created_on_frame, text="CreatedOn", style="Muted.TLabel").pack(anchor="w")
        self.created_on_filter_button = ttk.Button(
            created_on_frame,
            textvariable=self.created_on_filter_display_var,
            command=self.open_created_on_calendar,
            bootstyle="secondary",
        )
        self.created_on_filter_button.pack(anchor="w")

        buttons_frame = ttk.Frame(filter_row)
        buttons_frame.pack(side="right")
        ttk.Button(buttons_frame, text="Create Filter", command=self.apply_filter, bootstyle="secondary").pack(side="left")
        ttk.Button(buttons_frame, text="Clear Filters", command=self.clear_filter, bootstyle="light").pack(side="left", padx=(8, 0))

        ttk.Label(card, text="Double-click a cell to edit it in place.", style="Muted.TLabel").pack(fill="x", pady=(8, 10))

        table_wrap = ttk.Frame(card)
        table_wrap.pack(fill="both", expand=True)
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, columns=("row_no",), show="headings", height=20, selectmode="extended")
        self.tree.heading("row_no", text="Row")
        self.tree.column("row_no", width=72, anchor="center", stretch=False)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.begin_cell_edit)
        self.tree.bind("<Escape>", self.cancel_cell_edit)
        self.tree.bind("<Return>", self.begin_selected_edit)
        self.tree.bind("<Delete>", self.delete_selected_rows)

        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=lambda first, last: self._on_tree_yview(y_scroll, first, last), xscrollcommand=lambda first, last: self._on_tree_xview(x_scroll, first, last))

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("odd", background="#f8fbff")
        self.tree.bind("<Configure>", lambda _event: self._sync_checkbox_overlays())

        status_bar = ttk.Frame(self, padding=(16, 0, 16, 12))
        status_bar.pack(fill="x")
        ttk.Label(status_bar, textvariable=self.status_var, style="Muted.TLabel", anchor="w").pack(fill="x")

    def _load_or_create_workbook(self):
        try:
            self.workbook = self.service.load_or_create()

            self.open_sheet(FIXED_SHEETS[0])
            self.status_var.set(f"Ready: {os.path.basename(self.excel_path)}")
        except Exception as exc:
            messagebox.showerror("Startup failed", f"Could not prepare workbook.\n\n{exc}")
            self.status_var.set("Workbook initialization failed.")

    def _ensure_fixed_sheets(self):
        existing = set(self.workbook.sheetnames)
        for name in FIXED_SHEETS:
            if name not in existing:
                self.workbook.create_sheet(name)

    def _order_fixed_sheets(self):
        ordered = [self.workbook[name] for name in FIXED_SHEETS if name in self.workbook.sheetnames]
        extras = [sheet for sheet in self.workbook.worksheets if sheet.title not in FIXED_SHEETS]
        self.workbook._sheets = ordered + extras

    def _ensure_headers(self):
        for sheet_name in FIXED_SHEETS:
            sheet = self.workbook[sheet_name]
            headers = SHEET_HEADERS[sheet_name]
            current_headers = [sheet.cell(row=1, column=index + 1).value for index in range(sheet.max_column)]

            if sheet_name == "Tasks" and "DayOfWeek" not in current_headers:
                sheet.insert_cols(4, 1)
                current_headers = [sheet.cell(row=1, column=index + 1).value for index in range(sheet.max_column)]

            if sheet_name == "CompletedTasks" and "DayOfWeek" not in current_headers:
                sheet.insert_cols(5, 1)
                current_headers = [sheet.cell(row=1, column=index + 1).value for index in range(sheet.max_column)]

            if sheet_name == "TaskHistory" and "DayOfWeek" not in current_headers:
                sheet.insert_cols(5, 1)
                current_headers = [sheet.cell(row=1, column=index + 1).value for index in range(sheet.max_column)]

            if sheet_name == "CompletedTasks" and tuple(current_headers[: len(headers)]) != headers:
                self._reorder_sheet_columns(sheet, headers)

            if sheet_name == "TaskHistory" and tuple(current_headers[: len(headers)]) != headers:
                self._reorder_sheet_columns(sheet, headers)

            for index, header in enumerate(headers, start=1):
                sheet.cell(row=1, column=index, value=header)

            if sheet_name == "Tasks":
                for row_number in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row_number, column=4).value is None:
                        sheet.cell(row=row_number, column=4, value=None)

            if sheet_name == "CompletedTasks":
                for row_number in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row_number, column=5).value is None:
                        sheet.cell(row=row_number, column=5, value=None)

            if sheet_name == "TaskHistory":
                for row_number in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row_number, column=5).value is None:
                        sheet.cell(row=row_number, column=5, value=None)

    def _reorder_sheet_columns(self, sheet, desired_headers):
        current_headers = [sheet.cell(row=1, column=index + 1).value for index in range(sheet.max_column)]
        source_lookup = {header: index + 1 for index, header in enumerate(current_headers) if header}
        max_row = sheet.max_row
        max_col = max(sheet.max_column, len(desired_headers))

        for row_number in range(1, max_row + 1):
            row_values = {}
            for header in desired_headers:
                source_column = source_lookup.get(header)
                row_values[header] = sheet.cell(row=row_number, column=source_column).value if source_column else None

            for column_number, header in enumerate(desired_headers, start=1):
                sheet.cell(row=row_number, column=column_number, value=row_values[header])

            for column_number in range(len(desired_headers) + 1, max_col + 1):
                sheet.cell(row=row_number, column=column_number, value=None)

    def _weekday_from_value(self, value):
        if isinstance(value, datetime):
            return value.strftime("%A")

        text = self._stringify(value)
        if not text:
            return ""

        for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(text, pattern).strftime("%A")
            except ValueError:
                continue
        return ""

    def _export_fixed_sheet_files(self):
        self.service.export_fixed_sheet_files()

    def open_sheet(self, sheet_name):
        if not self.workbook or sheet_name not in FIXED_SHEETS:
            return

        self._commit_cell_edit()
        self.current_sheet_name = sheet_name
        self.worksheet = self.workbook[sheet_name]
        self.headers = self._read_headers()
        self._configure_tree_columns()
        self._refresh_rows()
        self._update_action_buttons()
        self.sheet_var.set(f"Sheet: {sheet_name}")
        self.status_var.set(f"Loaded {sheet_name}")

        if self.filtered_rows:
            self.select_row(self.filtered_rows[0])
        else:
            self.current_row = None
            self.row_pos_var.set("Row: -")

    def _read_headers(self):
        return list(SHEET_HEADERS[self.current_sheet_name])

    def _configure_tree_columns(self):
        columns = ["row_no"] + [f"col_{index}" for index in range(1, len(self.headers) + 1)]
        self.tree["columns"] = columns
        self.tree.heading("row_no", text="Row")
        self.tree.column("row_no", width=72, anchor="center", stretch=False)

        for index, header in enumerate(self.headers, start=1):
            column_id = f"col_{index}"
            self.tree.heading(column_id, text=header)
            if header in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
                self.tree.column(column_id, width=110, anchor="center", stretch=False)
            else:
                width = max(120, min(220, len(header) * 11))
                self.tree.column(column_id, width=width, anchor="w", stretch=True)

    def _on_tree_yview(self, scrollbar, first, last):
        scrollbar.set(first, last)
        self.after_idle(self._sync_checkbox_overlays)

    def _on_tree_xview(self, scrollbar, first, last):
        scrollbar.set(first, last)
        self.after_idle(self._sync_checkbox_overlays)

    def _refresh_rows(self):
        self.all_rows = list(range(2, self.worksheet.max_row + 1)) if self.worksheet.max_row >= 2 else []
        self._refresh_filter_options()
        self.filtered_rows = [row_number for row_number in self.all_rows if self._row_matches_filters(row_number)]
        self._render_rows()

    def _row_matches_filters(self, row_number):
        if self.current_sheet_name not in FILTERABLE_SHEETS:
            return True

        worker_filter = self.worker_filter_var.get().strip()
        day_filter = self.day_filter_var.get().strip()
        created_on_filter = self.created_on_filter_var.get().strip()

        if worker_filter != FILTER_ALL:
            worker_value = self._stringify(self.worksheet.cell(row=row_number, column=self._column_index("Worker")).value)
            if worker_value != worker_filter:
                return False

        if day_filter != FILTER_ALL:
            day_value = self._day_of_week_for_row(row_number)
            if day_value != day_filter:
                return False

        if created_on_filter != FILTER_ALL:
            created_on_value = self._created_on_for_row(row_number)
            if created_on_value != created_on_filter:
                return False

        return True

    def _column_index(self, header_name):
        for index, existing_header in enumerate(self.headers, start=1):
            if existing_header == header_name:
                return index
        return None

    def _refresh_filter_options(self):
        if self.current_sheet_name not in FILTERABLE_SHEETS:
            self.worker_options = [FILTER_ALL]
            self.day_options = [FILTER_ALL]
            self.worker_filter.configure(values=self.worker_options, state="disabled")
            self.day_filter.configure(values=self.day_options, state="disabled")
            self.worker_filter_var.set(FILTER_ALL)
            self.day_filter_var.set(FILTER_ALL)
            self.created_on_filter_var.set(FILTER_ALL)
            self._update_created_on_filter_display()
            self.created_on_filter_button.configure(state="disabled")
            return

        worker_index = self._column_index("Worker")

        worker_values = sorted({self._stringify(self.worksheet.cell(row=row_number, column=worker_index).value) for row_number in self.all_rows if worker_index and self.worksheet.cell(row=row_number, column=worker_index).value not in (None, "")}) if worker_index else []
        day_values = sorted({self._day_of_week_for_row(row_number) for row_number in self.all_rows if self._day_of_week_for_row(row_number) != ""})

        self.worker_options = [FILTER_ALL] + worker_values
        self.day_options = [FILTER_ALL] + day_values

        self.worker_filter.configure(values=self.worker_options, state="readonly")
        self.day_filter.configure(values=self.day_options, state="readonly")
        self.created_on_filter_button.configure(state="normal")
        self._update_created_on_filter_display()

        if self.worker_filter_var.get() not in self.worker_options:
            self.worker_filter_var.set(FILTER_ALL)
        if self.day_filter_var.get() not in self.day_options:
            self.day_filter_var.set(FILTER_ALL)

    def _row_matches(self, row_number, query):
        for column in range(1, len(self.headers) + 1):
            value = self.worksheet.cell(row=row_number, column=column).value
            if value is not None and query in self._stringify(value).lower():
                return True
        return False

    def _render_rows(self):
        self.tree.delete(*self.tree.get_children())
        self._clear_checkbox_overlays()

        for index, row_number in enumerate(self.filtered_rows):
            values = [row_number]
            for column in range(1, len(self.headers) + 1):
                values.append(self._display_cell_value(column, self.worksheet.cell(row=row_number, column=column).value))

            tag = "even" if index % 2 == 0 else "odd"
            self.tree.insert("", "end", iid=str(row_number), values=values, tags=(tag,))

        self.count_var.set(f"Rows: {len(self.filtered_rows)} / {len(self.all_rows)}")
        self.after_idle(self._sync_checkbox_overlays)

    def _display_value(self, value):
        text = self._stringify(value)
        if len(text) > 48:
            return text[:45] + "..."
        return text

    def _day_of_week_index(self):
        return self._column_index("DayOfWeek")

    def _created_on_for_row(self, row_number):
        created_on_index = self._column_index("CreatedOn")
        if not created_on_index:
            return ""

        value = self.worksheet.cell(row=row_number, column=created_on_index).value
        if value is None or value == "":
            return ""

        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")

        text = self._stringify(value)
        for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(text, pattern).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return text

    def _day_of_week_for_row(self, row_number):
        day_index = self._day_of_week_index()
        if day_index:
            return self._stringify(self.worksheet.cell(row=row_number, column=day_index).value)
        return ""

    def _is_read_only_header(self, header_name):
        return header_name == "CreatedOn"

    def _first_editable_tree_column(self):
        for column_number, header_name in enumerate(self.headers, start=2):
            if not self._is_read_only_header(header_name) and header_name not in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
                return column_number
        return None

    def _default_row_values(self):
        values = [None] * len(self.headers)
        created_on_column = self._column_index("CreatedOn")
        if created_on_column:
            values[created_on_column - 1] = datetime.now()
        return values

    def add_row(self):
        if not self.worksheet:
            return

        self._commit_cell_edit()

        new_row_number = self.worksheet.max_row + 1
        values = self._default_row_values()
        for column_number, value in enumerate(values, start=1):
            self.worksheet.cell(row=new_row_number, column=column_number, value=value)

        self.worker_filter_var.set(FILTER_ALL)
        self.day_filter_var.set(FILTER_ALL)
        self.created_on_filter_var.set(FILTER_ALL)
        self._refresh_rows()
        self.select_row(new_row_number)
        self.status_var.set(f"Added row {new_row_number} to {self.current_sheet_name}.")

    def _update_action_buttons(self):
        for button in (self.submit_button, self.approve_button, self.deny_button):
            button.pack_forget()

        if self.current_sheet_name == "Tasks":
            self.submit_button.pack(side="left", padx=(8, 0))
        elif self.current_sheet_name == "CompletedTasks":
            self.approve_button.pack(side="left", padx=(0, 8))
            self.deny_button.pack(side="left")

    def _selected_completed_rows(self):
        if not self.worksheet or self.current_sheet_name != "CompletedTasks":
            return []

        audited_column = self._column_index("Audited")
        if not audited_column:
            return []

        selected_rows = []
        for row_number in self.all_rows:
            if self._completion_symbol(self.worksheet.cell(row=row_number, column=audited_column).value) == "☑":
                selected_rows.append(row_number)
        return selected_rows

    def _append_completed_row_to_sheet(self, target_sheet_name, source_row_number):
        target_sheet = self.workbook[target_sheet_name]
        target_headers = SHEET_HEADERS[target_sheet_name]
        target_columns = {header: index + 1 for index, header in enumerate(target_headers)}

        created_on_value = datetime.now()
        worker_value = self.worksheet.cell(row=source_row_number, column=self._column_index("Worker")).value
        task_value = self.worksheet.cell(row=source_row_number, column=self._column_index("Task")).value
        description_value = self.worksheet.cell(row=source_row_number, column=self._column_index("Description")).value
        day_value = self.worksheet.cell(row=source_row_number, column=self._column_index("DayOfWeek")).value

        new_row_number = target_sheet.max_row + 1
        if target_sheet_name == "CompletedTasks":
            target_sheet.cell(row=new_row_number, column=target_columns["CreatedOn"], value=created_on_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Worker"], value=worker_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Task"], value=task_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Description"], value=description_value)
            target_sheet.cell(row=new_row_number, column=target_columns["DayOfWeek"], value=day_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Audited"], value=None)
        elif target_sheet_name == "TaskHistory":
            target_sheet.cell(row=new_row_number, column=target_columns["CreatedOn"], value=created_on_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Worker"], value=worker_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Task"], value=task_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Description"], value=description_value)
            target_sheet.cell(row=new_row_number, column=target_columns["DayOfWeek"], value=day_value)
        elif target_sheet_name == "Tasks":
            target_sheet.cell(row=new_row_number, column=target_columns["CreatedOn"], value=created_on_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Task"], value=task_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Description"], value=description_value)
            target_sheet.cell(row=new_row_number, column=target_columns["DayOfWeek"], value=day_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Worker"], value=worker_value)
            target_sheet.cell(row=new_row_number, column=target_columns["Completion"], value=None)

    def _move_completed_rows(self, destination_sheet_name):
        if self.current_sheet_name != "CompletedTasks":
            self.status_var.set("This action is only available on CompletedTasks.")
            return

        self._commit_cell_edit()
        selected_rows = self._selected_completed_rows()
        if not selected_rows:
            self.status_var.set("Select checked rows in Audited first.")
            return

        for row_number in selected_rows:
            self._append_completed_row_to_sheet(destination_sheet_name, row_number)

        for row_number in sorted(selected_rows, reverse=True):
            self.worksheet.delete_rows(row_number, 1)

        self.workbook.save(self.excel_path)
        self._export_fixed_sheet_files()
        self._refresh_rows()

        if self.filtered_rows:
            self.select_row(self.filtered_rows[0])
        else:
            self.current_row = None
            self.row_pos_var.set("Row: -")

        self.status_var.set(f"Moved {len(selected_rows)} row(s) to {destination_sheet_name}.")

    def approve_completed_rows(self):
        self._move_completed_rows("TaskHistory")

    def deny_completed_rows(self):
        self._move_completed_rows("Tasks")

    def submit_completed_rows(self):
        if not self.worksheet or self.current_sheet_name != "Tasks":
            self.status_var.set("Submit is only available on Tasks.")
            return

        self._commit_cell_edit()

        completion_column = self._column_index("Completion")
        if not completion_column:
            self.status_var.set("Tasks sheet is missing the Completion column.")
            return

        completed_sheet = self.workbook["CompletedTasks"]
        submitted_rows = []
        completed_columns = {header: index + 1 for index, header in enumerate(SHEET_HEADERS["CompletedTasks"])}
        tasks_columns = {header: index + 1 for index, header in enumerate(SHEET_HEADERS["Tasks"])}
        created_on_value = datetime.now()

        for row_number in self.all_rows:
            if self._completion_symbol(self.worksheet.cell(row=row_number, column=completion_column).value) != "☑":
                continue

            task_value = self.worksheet.cell(row=row_number, column=tasks_columns["Task"]).value
            description_value = self.worksheet.cell(row=row_number, column=tasks_columns["Description"]).value
            day_value = self.worksheet.cell(row=row_number, column=tasks_columns["DayOfWeek"]).value
            worker_value = self.worksheet.cell(row=row_number, column=tasks_columns["Worker"]).value

            new_row_number = completed_sheet.max_row + 1
            completed_sheet.cell(row=new_row_number, column=completed_columns["CreatedOn"], value=created_on_value)
            completed_sheet.cell(row=new_row_number, column=completed_columns["Worker"], value=worker_value)
            completed_sheet.cell(row=new_row_number, column=completed_columns["Task"], value=task_value)
            completed_sheet.cell(row=new_row_number, column=completed_columns["Description"], value=description_value)
            completed_sheet.cell(row=new_row_number, column=completed_columns["DayOfWeek"], value=day_value)
            completed_sheet.cell(row=new_row_number, column=completed_columns["Audited"], value=None)
            submitted_rows.append(row_number)

        if not submitted_rows:
            self.status_var.set("No checked rows to submit.")
            return

        for row_number in sorted(submitted_rows, reverse=True):
            self.worksheet.delete_rows(row_number, 1)

        self.workbook.save(self.excel_path)
        self._export_fixed_sheet_files()
        self._refresh_rows()

        if self.filtered_rows:
            self.select_row(self.filtered_rows[0])
        else:
            self.current_row = None
            self.row_pos_var.set("Row: -")

        self.status_var.set(f"Submitted {len(submitted_rows)} row(s) to CompletedTasks.")

    def delete_selected_rows(self, _event=None):
        if not self.worksheet:
            return

        self._commit_cell_edit()
        selected_ids = self.tree.selection()
        if not selected_ids:
            self.status_var.set("Select one or more rows to delete.")
            return "break" if _event is not None else None

        row_numbers = sorted((int(row_id) for row_id in selected_ids), reverse=True)
        focus_row = min(row_numbers)

        for row_number in row_numbers:
            if row_number > 1 and row_number <= self.worksheet.max_row:
                self.worksheet.delete_rows(row_number, 1)

        self.worker_filter_var.set(FILTER_ALL)
        self.day_filter_var.set(FILTER_ALL)
        self.created_on_filter_var.set(FILTER_ALL)
        self._refresh_rows()

        if self.worksheet.max_row >= 2:
            next_row = min(focus_row, self.worksheet.max_row)
            if next_row < 2:
                next_row = 2
            if next_row in self.filtered_rows:
                self.select_row(next_row)
            elif self.filtered_rows:
                self.select_row(self.filtered_rows[0])
        else:
            self.current_row = None
            self.row_pos_var.set("Row: -")

        try:
            self.workbook.save(self.excel_path)
            self._export_fixed_sheet_files()
        except Exception as exc:
            messagebox.showerror("Delete failed", f"Could not save deleted rows.\n\n{exc}")
            return "break" if _event is not None else None

        self.status_var.set(f"Deleted {len(row_numbers)} row(s) from {self.current_sheet_name}.")
        return "break" if _event is not None else None

    def _display_cell_value(self, column_index, value):
        header_name = self.headers[column_index - 1] if 0 <= column_index - 1 < len(self.headers) else ""
        if header_name in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
            return ""
        return self._display_value(value)

    def _header_for_tree_column(self, column_number):
        header_index = column_number - 2
        if 0 <= header_index < len(self.headers):
            return self.headers[header_index]
        return None

    def _completion_symbol(self, value):
        text = self._stringify(value).lower()
        if text in {"☑", "check", "checked", "yes", "true", "1", "x"}:
            return "☑"
        return "☐"

    def _toggle_completion(self, row_number):
        completion_column = len(self.headers)
        current_value = self.worksheet.cell(row=row_number, column=completion_column).value
        next_value = "" if self._completion_symbol(current_value) == "☑" else "☑"
        self.worksheet.cell(row=row_number, column=completion_column, value=next_value)

        tree_values = list(self.tree.item(str(row_number), "values"))
        tree_values[completion_column] = ""
        self.tree.item(str(row_number), values=tree_values)
        if row_number in self.checkbox_vars:
            self.checkbox_vars[row_number].set(next_value == "☑")
        self.status_var.set(f"Updated {self.current_sheet_name} row {row_number} completion.")

    def _checkbox_column_index(self):
        checkbox_headers = CHECKBOX_COLUMNS.get(self.current_sheet_name, set())
        for index, header in enumerate(self.headers, start=1):
            if header in checkbox_headers:
                return index
        return None

    def _clear_checkbox_overlays(self):
        for widget in self.checkbox_widgets.values():
            widget.destroy()
        self.checkbox_widgets.clear()
        self.checkbox_vars.clear()

    def _sync_checkbox_overlays(self):
        self._clear_checkbox_overlays()

        checkbox_column = self._checkbox_column_index()
        if checkbox_column is None:
            return

        for row_number in self.filtered_rows:
            bbox = self.tree.bbox(str(row_number), f"#{checkbox_column + 1}")
            if not bbox:
                continue

            x, y, width, height = bbox
            cell_value = self.worksheet.cell(row=row_number, column=checkbox_column).value
            var = tk.BooleanVar(value=self._completion_symbol(cell_value) == "☑")

            def on_toggle(row=row_number, variable=var):
                self.worksheet.cell(row=row, column=checkbox_column, value="☑" if variable.get() else "")
                self.status_var.set(f"Updated {self.current_sheet_name} row {row} completion.")

            checkbox = ttk.Checkbutton(self.tree, variable=var, command=on_toggle)
            checkbox.place(x=x + (width - 20) // 2, y=y + 2, width=20, height=max(18, height - 4))
            self.checkbox_widgets[row_number] = checkbox
            self.checkbox_vars[row_number] = var

    def _stringify(self, value):
        if value is None:
            return ""
        if hasattr(value, "strftime"):
            try:
                return value.strftime("%Y-%m-%d")
            except Exception:
                pass
        return str(value).strip()

    def on_row_select(self, _event=None):
        if self._suppress_select:
            return
        selection = self.tree.selection()
        if not selection:
            return
        self.current_row = int(selection[0])
        self.row_pos_var.set(f"Row: {self.current_row} (Excel row)")
        self.after_idle(self._sync_checkbox_overlays)

    def select_row(self, row_number):
        if row_number is None or str(row_number) not in self.tree.get_children():
            return

        self._suppress_select = True
        self.tree.selection_set(str(row_number))
        self.tree.focus(str(row_number))
        self.tree.see(str(row_number))
        self._suppress_select = False

        self.current_row = row_number
        self.row_pos_var.set(f"Row: {row_number} (Excel row)")
        self.after_idle(self._sync_checkbox_overlays)

    def apply_filter(self, _event=None):
        self._commit_cell_edit()
        previous_row = self.current_row
        self._refresh_rows()

        if previous_row in self.filtered_rows:
            self.select_row(previous_row)
        elif self.filtered_rows:
            self.select_row(self.filtered_rows[0])
        else:
            self.current_row = None
            self.row_pos_var.set("Row: -")
            self.status_var.set("No rows match the filter.")

    def clear_filter(self):
        self.worker_filter_var.set(FILTER_ALL)
        self.day_filter_var.set(FILTER_ALL)
        self.created_on_filter_var.set(FILTER_ALL)
        self._update_created_on_filter_display()
        self.apply_filter()

    def _update_created_on_filter_display(self):
        selected_value = self.created_on_filter_var.get().strip()
        if selected_value and selected_value != FILTER_ALL:
            self.created_on_filter_display_var.set(selected_value)
        else:
            self.created_on_filter_display_var.set("Any date")

    def _parse_created_on_date(self, value):
        if not value or value == FILTER_ALL:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None

    def open_created_on_calendar(self):
        if self.current_sheet_name not in FILTERABLE_SHEETS or self.created_on_filter_button.instate(["disabled"]):
            return

        if self.created_on_calendar_window and self.created_on_calendar_window.winfo_exists():
            self.created_on_calendar_window.lift()
            self.created_on_calendar_window.focus_force()
            return

        window = tk.Toplevel(self)
        window.title("Choose CreatedOn date")
        window.transient(self)
        window.grab_set()
        window.resizable(False, False)
        window.configure(bg=APP_BG)
        self.created_on_calendar_window = window

        selected_value = self.created_on_filter_var.get().strip()
        initial_date = datetime.now().date()
        parsed_date = self._parse_created_on_date(selected_value)
        if parsed_date:
            initial_date = parsed_date

        state = {
            "year": initial_date.year,
            "month": initial_date.month,
            "selected": initial_date,
        }

        content = ttk.Frame(window, padding=14)
        content.pack(fill="both", expand=True)

        header = ttk.Frame(content)
        header.pack(fill="x", pady=(0, 10))

        month_label = ttk.Label(header, text="", font=("Segoe UI Semibold", 13))
        month_label.pack(side="left")

        nav = ttk.Frame(header)
        nav.pack(side="right")

        days_frame = ttk.Frame(content)
        days_frame.pack(fill="both", expand=True)

        def refresh_calendar():
            for widget in days_frame.winfo_children():
                widget.destroy()

            month_label.configure(text=f"{calendar.month_name[state['month']]} {state['year']}")

            weekday_names = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
            for column_index, weekday_name in enumerate(weekday_names):
                ttk.Label(days_frame, text=weekday_name, style="Muted.TLabel", anchor="center", width=5).grid(
                    row=0, column=column_index, padx=2, pady=(0, 4)
                )

            weeks = calendar.monthcalendar(state["year"], state["month"])
            for row_index, week in enumerate(weeks, start=1):
                for column_index, day in enumerate(week):
                    if day == 0:
                        ttk.Label(days_frame, text="", width=5).grid(row=row_index, column=column_index, padx=2, pady=2)
                        continue

                    is_selected = (
                        state["selected"]
                        and state["selected"].year == state["year"]
                        and state["selected"].month == state["month"]
                        and state["selected"].day == day
                    )

                    def choose_day(chosen_day=day):
                        chosen_date = datetime(state["year"], state["month"], chosen_day).date().isoformat()
                        self.created_on_filter_var.set(chosen_date)
                        self._update_created_on_filter_display()
                        window.destroy()

                    ttk.Button(
                        days_frame,
                        text=str(day),
                        width=5,
                        command=choose_day,
                        bootstyle="primary" if is_selected else "light",
                    ).grid(row=row_index, column=column_index, padx=2, pady=2)

        def previous_month():
            if state["month"] == 1:
                state["month"] = 12
                state["year"] -= 1
            else:
                state["month"] -= 1
            refresh_calendar()

        def next_month():
            if state["month"] == 12:
                state["month"] = 1
                state["year"] += 1
            else:
                state["month"] += 1
            refresh_calendar()

        ttk.Button(nav, text="<", command=previous_month, bootstyle="secondary").pack(side="left", padx=(0, 6))
        ttk.Button(nav, text=">", command=next_month, bootstyle="secondary").pack(side="left")

        footer = ttk.Frame(content)
        footer.pack(fill="x", pady=(10, 0))

        ttk.Button(
            footer,
            text="Clear",
            command=lambda: (self.created_on_filter_var.set(FILTER_ALL), self._update_created_on_filter_display(), window.destroy()),
            bootstyle="light",
        ).pack(side="left")
        ttk.Button(footer, text="Close", command=window.destroy, bootstyle="secondary").pack(side="right")

        window.protocol("WM_DELETE_WINDOW", window.destroy)
        window.bind("<Escape>", lambda _event: window.destroy())
        refresh_calendar()
        window.wait_window()
        self.created_on_calendar_window = None

    def begin_selected_edit(self, _event=None):
        if self.current_cell:
            row_number, column_number = self.current_cell
            header_name = self._header_for_tree_column(column_number)
            if self._is_read_only_header(header_name):
                return
            bbox = self.tree.bbox(str(row_number), f"#{column_number}")
            if bbox:
                self._begin_edit(str(row_number), column_number, bbox)
                return

        selection = self.tree.selection()
        if not selection:
            return
        row_number = selection[0]
        column_number = self._first_editable_tree_column()
        if not column_number:
            return
        bbox = self.tree.bbox(row_number, f"#{column_number}")
        if bbox:
            self._begin_edit(row_number, column_number, bbox)

    def on_tree_click(self, event):
        if self.current_sheet_name not in CHECKBOX_COLUMNS:
            return

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        try:
            column_number = int(column_id.replace("#", ""))
        except ValueError:
            return

        self.current_cell = (int(row_id), column_number)

        header_name = self._header_for_tree_column(column_number)
        if self._is_read_only_header(header_name):
            return
        if header_name in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
            self._toggle_completion(int(row_id))
            return "break"

    def begin_cell_edit(self, event):
        if not self.current_sheet_name:
            return

        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id or column_id == "#1":
            return

        try:
            column_number = int(column_id.replace("#", ""))
        except ValueError:
            return

        self.current_cell = (int(row_id), column_number)

        header_name = self._header_for_tree_column(column_number)
        if self._is_read_only_header(header_name):
            return
        if header_name in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
            self._toggle_completion(int(row_id))
            return

        bbox = self.tree.bbox(row_id, column_id)
        if not bbox:
            return

        self._begin_edit(row_id, column_number, bbox)

    def _begin_edit(self, row_id, column_number, bbox):
        self._commit_cell_edit()

        row_number = int(row_id)
        workbook_column = column_number - 1
        current_value = self.worksheet.cell(row=row_number, column=workbook_column).value
        x, y, width, height = bbox

        editor = ttk.Entry(self.tree)
        if 1 <= workbook_column <= len(self.headers) and self.headers[workbook_column - 1] in CHECKBOX_COLUMNS.get(self.current_sheet_name, set()):
            editor.insert(0, self._completion_symbol(current_value))
        else:
            editor.insert(0, self._stringify(current_value))
        editor.select_range(0, tk.END)
        editor.focus_set()
        editor.place(x=x, y=y, width=width, height=height)

        self.current_edit = {
            "row_number": row_number,
            "workbook_column": workbook_column,
            "tree_column": column_number,
            "editor": editor,
        }

        editor.bind("<Return>", self._commit_cell_edit)
        editor.bind("<FocusOut>", self._commit_cell_edit)
        editor.bind("<Escape>", self.cancel_cell_edit)

    def _commit_cell_edit(self, _event=None):
        if not self.current_edit:
            return

        editor = self.current_edit["editor"]
        row_number = self.current_edit["row_number"]
        workbook_column = self.current_edit["workbook_column"]
        tree_column = self.current_edit["tree_column"]
        new_value = editor.get().strip()

        self.worksheet.cell(row=row_number, column=workbook_column, value=new_value if new_value else None)

        tree_values = list(self.tree.item(str(row_number), "values"))
        tree_values[tree_column - 1] = self._display_value(new_value)
        self.tree.item(str(row_number), values=tree_values)

        editor.destroy()
        self.current_edit = None
        self.status_var.set(f"Updated {self.current_sheet_name} row {row_number}.")
        self.after_idle(self._sync_checkbox_overlays)

        if _event is not None:
            return "break"

    def cancel_cell_edit(self, _event=None):
        if self.current_edit:
            self.current_edit["editor"].destroy()
            self.current_edit = None
        if _event is not None:
            return "break"

    def save_workbook(self):
        if not self.workbook:
            messagebox.showinfo("No workbook", "Workbook is not ready yet.")
            return

        self._commit_cell_edit()
        try:
            self.service.save()
            self.status_var.set(f"Saved: {self.excel_path}")
            messagebox.showinfo("Saved", "Workbook saved successfully.")
        except Exception as exc:
            messagebox.showerror("Save failed", f"Could not save workbook.\n\n{exc}")

    def prev_row(self):
        if not self.filtered_rows or self.current_row not in self.filtered_rows:
            return
        index = self.filtered_rows.index(self.current_row)
        if index > 0:
            self.select_row(self.filtered_rows[index - 1])

    def next_row(self):
        if not self.filtered_rows or self.current_row not in self.filtered_rows:
            return
        index = self.filtered_rows.index(self.current_row)
        if index < len(self.filtered_rows) - 1:
            self.select_row(self.filtered_rows[index + 1])


if __name__ == "__main__":
    app = TaskWorkbookApp()
    app.mainloop()
