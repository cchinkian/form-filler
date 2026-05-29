"""
Offline Windows PDF Procedure Automation Tool.

The app is built around:
Excel customer workbook -> Source Forms -> coordinate mapping -> Procedures
-> combined PDF output -> Excel history log.
"""
from __future__ import annotations

import os
import platform
import re
import shutil
import subprocess
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, str(Path(__file__).parent))

import catalog
import config_loader
import excel_reader
import package_engine
from excel_reader import ExcelLockedError


LABEL_W = 18


class ProcedureAutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Procedure Automation")
        self.geometry("1120x760")
        self.minsize(980, 640)
        self._set_window_icon()

        self.settings: dict = {}
        self.forms_config: dict = {}
        self.procedures: list[dict] = []
        self.source_forms: list[dict] = []
        self.procedure_items: list[dict] = []
        self.procedure_by_code: dict[str, dict] = {}
        self.source_by_code: dict[str, dict] = {}
        self.customers: list[dict] = []
        self.accounts: list[dict] = []
        self.account_rows: list[dict] = []
        self.workbook_schema: list[dict] = []
        self.default_fields: dict[str, set[str]] = {}
        self.staff_profile: dict = {}
        self.rm_profile: dict = {}

        self.selected_customer: dict | None = None
        self.selected_account: dict | None = None
        self.search_results: list[dict] = []
        self.manual_entries: dict[str, tk.StringVar] = {}
        self.bulk_rows: list[dict] = []
        self.builder_items: list[dict] = []
        self.selected_source_code: str | None = None
        self.recent_history_rows: list[dict] = []

        try:
            config_loader.backup_settings_on_launch()
        except Exception:
            pass

        self._build_ui()
        self._load_all()

    def _set_window_icon(self):
        bases = []
        if getattr(sys, "frozen", False):
            bases.append(Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)))
            bases.append(Path(sys.executable).parent)
        bases.append(Path(__file__).resolve().parents[1])
        for base in bases:
            icon = base / "assets" / "app.ico"
            if icon.exists():
                try:
                    self.iconbitmap(default=str(icon))
                    return
                except tk.TclError:
                    continue

    # UI -----------------------------------------------------------------

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)
        ttk.Label(top, text="PDF Procedure Automation", font=("", 14, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="Reload", command=self._load_all).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Settings", command=self._open_settings_dialog).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Mapping Editor", command=self._show_coordinate_tab).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Output", command=lambda: self._open_path(config_loader.output_folder_path(self.settings))).pack(side=tk.RIGHT, padx=(4, 0))

        session = ttk.Frame(self, padding=(10, 0, 10, 8))
        session.pack(fill=tk.X)
        ttk.Label(session, text="RM Code").pack(side=tk.LEFT)
        self.var_rm_code = tk.StringVar()
        self.cmb_rm_code = ttk.Combobox(session, textvariable=self.var_rm_code, width=14, state="normal")
        self.cmb_rm_code.pack(side=tk.LEFT, padx=(4, 14))
        ttk.Label(session, text="Branch").pack(side=tk.LEFT)
        self.var_branch = tk.StringVar()
        self.cmb_branch = ttk.Combobox(session, textvariable=self.var_branch, width=18, state="normal")
        self.cmb_branch.pack(side=tk.LEFT, padx=(4, 14))
        ttk.Label(session, text="Date").pack(side=tk.LEFT)
        self.var_date = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.var_date.trace_add("write", lambda *_: self._refresh_generate_structure() if hasattr(self, "lst_gen_structure") else None)
        ttk.Entry(session, textvariable=self.var_date, width=14).pack(side=tk.LEFT, padx=(4, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))

        self.tab_generate = ttk.Frame(self.nb, padding=10)
        self.tab_bulk = ttk.Frame(self.nb, padding=10)
        self.tab_builder = ttk.Frame(self.nb, padding=10)
        self.tab_sources = ttk.Frame(self.nb, padding=10)
        self.tab_coordinate = ttk.Frame(self.nb, padding=0)
        self.tab_history = ttk.Frame(self.nb, padding=10)

        self.nb.add(self.tab_generate, text="Generate Package")
        self.nb.add(self.tab_bulk, text="Bulk Export")
        self.nb.add(self.tab_builder, text="Procedure Builder")
        self.nb.add(self.tab_sources, text="Source Forms")
        self.nb.add(self.tab_coordinate, text="Coordinate Pointer")
        self.nb.add(self.tab_history, text="History / Settings")

        self._build_generate_tab()
        self._build_bulk_tab()
        self._build_builder_tab()
        self._build_sources_tab()
        self._build_coordinate_tab()
        self._build_history_tab()

        self.lbl_status = ttk.Label(self, text="Loading...", anchor="w")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 8))

    def _build_generate_tab(self):
        left = ttk.Frame(self.tab_generate)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        right = ttk.Frame(self.tab_generate)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        search_row = ttk.Frame(left)
        search_row.pack(fill=tk.X)
        ttk.Label(search_row, text="Search customer", width=LABEL_W).pack(side=tk.LEFT)
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self._refresh_search_results())
        ttk.Entry(search_row, textvariable=self.var_search, width=28).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.var_customer_choice = tk.StringVar()
        self.cmb_customer = ttk.Combobox(search_row, textvariable=self.var_customer_choice, state="readonly", width=42)
        self.cmb_customer.pack(side=tk.LEFT, padx=(6, 0))
        self.cmb_customer.bind("<<ComboboxSelected>>", lambda _: self._pick_customer())

        self.lbl_customer = ttk.Label(left, text="No customer selected.", anchor="w", font=("", 9, "bold"))
        self.lbl_customer.pack(fill=tk.X, pady=(0, 8))

        proc_box = ttk.LabelFrame(left, text="Procedure")
        proc_box.pack(fill=tk.X, pady=(0, 8))
        row = ttk.Frame(proc_box, padding=6)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Search", width=LABEL_W).pack(side=tk.LEFT)
        self.var_gen_proc_search = tk.StringVar()
        self.var_gen_proc_search.trace_add("write", lambda *_: self._refresh_generate_procedures())
        ttk.Entry(row, textvariable=self.var_gen_proc_search).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tag_row = ttk.Frame(proc_box, padding=(6, 0, 6, 6))
        tag_row.pack(fill=tk.X)
        ttk.Label(tag_row, text="Tags", width=LABEL_W).pack(side=tk.LEFT)
        self.frm_gen_tags = ttk.Frame(tag_row)
        self.frm_gen_tags.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.gen_tag_vars: dict[str, tk.BooleanVar] = {}

        row = ttk.Frame(proc_box, padding=(6, 0, 6, 6))
        row.pack(fill=tk.X)
        ttk.Label(row, text="Procedure", width=LABEL_W).pack(side=tk.LEFT)
        self.var_gen_proc = tk.StringVar()
        self.cmb_gen_proc = ttk.Combobox(row, textvariable=self.var_gen_proc, state="readonly")
        self.cmb_gen_proc.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cmb_gen_proc.bind("<<ComboboxSelected>>", lambda _: self._on_generate_proc_change())
        ttk.Button(row, text="Add Procedure", command=self._new_procedure_from_generate).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(row, text="Edit", command=self._open_selected_procedure_in_builder).pack(side=tk.LEFT, padx=(4, 0))

        session_box = ttk.LabelFrame(left, text="Session")
        session_box.pack(fill=tk.X, pady=(0, 8))
        row = ttk.Frame(session_box, padding=6)
        row.pack(fill=tk.X)
        ttk.Label(row, text="RM Code", width=LABEL_W).pack(side=tk.LEFT)
        self.cmb_gen_rm_code = ttk.Combobox(row, textvariable=self.var_rm_code, width=16, state="normal")
        self.cmb_gen_rm_code.pack(side=tk.LEFT)
        ttk.Label(row, text="Branch").pack(side=tk.LEFT, padx=(12, 4))
        self.cmb_gen_branch = ttk.Combobox(row, textvariable=self.var_branch, width=18, state="normal")
        self.cmb_gen_branch.pack(side=tk.LEFT)
        ttk.Label(row, text="Date").pack(side=tk.LEFT, padx=(12, 4))
        ttk.Entry(row, textvariable=self.var_date, width=14).pack(side=tk.LEFT)

        account_box = ttk.LabelFrame(left, text="Related Account")
        account_box.pack(fill=tk.X, pady=(0, 8))
        row = ttk.Frame(account_box, padding=6)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Account", width=LABEL_W).pack(side=tk.LEFT)
        self.var_account = tk.StringVar()
        self.cmb_account = ttk.Combobox(row, textvariable=self.var_account, state="readonly")
        self.cmb_account.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cmb_account.bind("<<ComboboxSelected>>", lambda _: self._pick_account())

        ttk.Button(left, text="Generate PDF Package", command=self._generate_single).pack(anchor="e", pady=(4, 0))

        structure_box = ttk.LabelFrame(right, text="Procedure Structure")
        structure_box.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.lst_gen_structure = tk.Listbox(structure_box, height=10, font=("Courier", 9))
        self.lst_gen_structure.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        field_box = ttk.LabelFrame(right, text="Client Default Information / Form Inputs")
        field_box.pack(fill=tk.BOTH, expand=True)
        self.canvas_manual = tk.Canvas(field_box, highlightthickness=0)
        self.frm_manual = ttk.Frame(self.canvas_manual)
        sb = ttk.Scrollbar(field_box, orient="vertical", command=self.canvas_manual.yview)
        self.canvas_manual.configure(yscrollcommand=sb.set)
        self.canvas_manual.create_window((0, 0), window=self.frm_manual, anchor="nw")
        self.frm_manual.bind("<Configure>", lambda e: self.canvas_manual.configure(scrollregion=self.canvas_manual.bbox("all")))
        self.canvas_manual.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=6)
        sb.pack(side=tk.RIGHT, fill=tk.Y, pady=6)

    def _build_bulk_tab(self):
        top = ttk.Frame(self.tab_bulk)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Search", width=LABEL_W).pack(side=tk.LEFT)
        self.var_bulk_proc_search = tk.StringVar()
        self.var_bulk_proc_search.trace_add("write", lambda *_: self._refresh_bulk_procedures())
        ttk.Entry(top, textvariable=self.var_bulk_proc_search, width=28).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(top, text="Procedure").pack(side=tk.LEFT)
        self.var_bulk_proc = tk.StringVar()
        self.cmb_bulk_proc = ttk.Combobox(top, textvariable=self.var_bulk_proc, state="readonly", width=36)
        self.cmb_bulk_proc.pack(side=tk.LEFT, padx=(4, 8))

        ttk.Button(top, text="Import CIS List", command=self._import_cis_list).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Analyze CIS", command=self._analyze_bulk).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Generate Bulk", command=self._generate_bulk).pack(side=tk.RIGHT, padx=(4, 0))

        tag_row = ttk.Frame(self.tab_bulk)
        tag_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(tag_row, text="Tags", width=LABEL_W).pack(side=tk.LEFT)
        self.frm_bulk_tags = ttk.Frame(tag_row)
        self.frm_bulk_tags.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bulk_tag_vars: dict[str, tk.BooleanVar] = {}

        body = ttk.Frame(self.tab_bulk)
        body.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        left = ttk.LabelFrame(body, text="CIS List")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        self.txt_cis = tk.Text(left, width=32, height=18)
        self.txt_cis.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        right = ttk.LabelFrame(body, text="Match Result")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_bulk = ttk.Treeview(right, columns=("cis", "name", "status"), show="headings", height=18)
        self.tree_bulk.heading("cis", text="CIS")
        self.tree_bulk.heading("name", text="Customer")
        self.tree_bulk.heading("status", text="Status")
        self.tree_bulk.column("cis", width=120)
        self.tree_bulk.column("name", width=260)
        self.tree_bulk.column("status", width=160)
        self.tree_bulk.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

    def _build_builder_tab(self):
        top = ttk.Frame(self.tab_builder)
        top.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(top, text="Procedure", width=LABEL_W).pack(side=tk.LEFT)
        self.var_builder_proc = tk.StringVar()
        self.cmb_builder_proc = ttk.Combobox(top, textvariable=self.var_builder_proc, state="readonly", width=46)
        self.cmb_builder_proc.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cmb_builder_proc.bind("<<ComboboxSelected>>", lambda _: self._load_builder_procedure())
        ttk.Button(top, text="New", command=self._new_procedure).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Save", command=self._save_builder_procedure).pack(side=tk.RIGHT, padx=(4, 0))

        split = ttk.Frame(self.tab_builder)
        split.pack(fill=tk.BOTH, expand=True)

        details = ttk.LabelFrame(split, text="Procedure Details")
        details.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        self.builder_vars: dict[str, tk.StringVar] = {}
        for key in ["ProcedureCode", "DisplayName", "Tags", "Version", "DefaultOutputName", "Description", "Remarks"]:
            row = ttk.Frame(details, padding=(6, 4))
            row.pack(fill=tk.X)
            ttk.Label(row, text=key, width=LABEL_W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self.builder_vars[key] = var
            widget = ttk.Entry(row, textvariable=var)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.var_builder_active = tk.BooleanVar(value=True)
        ttk.Checkbutton(details, text="Active", variable=self.var_builder_active).pack(anchor="w", padx=8, pady=(4, 0))
        self.var_builder_auto_blank = tk.BooleanVar(value=True)
        ttk.Checkbutton(details, text="Auto insert blank page after odd-page PDFs", variable=self.var_builder_auto_blank).pack(anchor="w", padx=8, pady=(2, 0))

        items = ttk.LabelFrame(split, text="Package Items")
        items.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        builder_split = ttk.Frame(items)
        builder_split.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        available = ttk.LabelFrame(builder_split, text="Available Source Forms")
        available.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        seq = ttk.LabelFrame(builder_split, text="Procedure Sequence")
        seq.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.var_source_filter = tk.StringVar()
        self.var_source_filter.trace_add("write", lambda *_: self._refresh_available_sources())
        ttk.Entry(available, textvariable=self.var_source_filter).pack(fill=tk.X, padx=4, pady=(4, 2))
        self.lst_available_sources = tk.Listbox(available, height=12, font=("Courier", 9))
        self.lst_available_sources.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        self.lst_available_sources.bind("<Double-Button-1>", lambda _: self._builder_add_source())

        self.lst_builder_items = tk.Listbox(seq, height=12, font=("Courier", 9))
        self.lst_builder_items.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        controls = ttk.Frame(items)
        controls.pack(fill=tk.X, padx=6, pady=(0, 6))
        ttk.Button(controls, text="Add →", command=self._builder_add_source).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(controls, text="Remove", command=self._builder_remove).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(controls, text="Up", command=lambda: self._builder_move(-1)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(controls, text="Down", command=lambda: self._builder_move(1)).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(controls, text="Advanced: Add Manual Blank", command=self._builder_add_blank).pack(side=tk.RIGHT)

    def _build_sources_tab(self):
        split = ttk.Frame(self.tab_sources)
        split.pack(fill=tk.BOTH, expand=True)
        left = ttk.Frame(split)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        right = ttk.LabelFrame(split, text="Source Form Details")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scan_box = ttk.LabelFrame(left, text="Form Root Folder Scan")
        scan_box.pack(fill=tk.X, pady=(0, 8))
        scan_row = ttk.Frame(scan_box, padding=6)
        scan_row.pack(fill=tk.X)
        ttk.Label(scan_row, text="Folder", width=LABEL_W).pack(side=tk.LEFT)
        self.var_form_root = tk.StringVar()
        ttk.Entry(scan_row, textvariable=self.var_form_root).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(scan_row, text="Choose Folder", command=self._choose_form_root_folder).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(scan_row, text="Scan", command=self._scan_form_root).pack(side=tk.LEFT, padx=(4, 0))
        self.tree_form_scan = ttk.Treeview(scan_box, columns=("folder", "pdf", "mapping", "updated"), show="headings", height=7)
        self.tree_form_scan.heading("folder", text="Form Folder")
        self.tree_form_scan.heading("pdf", text="PDF Status")
        self.tree_form_scan.heading("mapping", text="Mapping")
        self.tree_form_scan.heading("updated", text="Last Mapping Edit")
        self.tree_form_scan.column("folder", width=190)
        self.tree_form_scan.column("pdf", width=130)
        self.tree_form_scan.column("mapping", width=120)
        self.tree_form_scan.column("updated", width=150)
        self.tree_form_scan.pack(fill=tk.X, padx=6, pady=(0, 6))
        self.tree_form_scan.bind("<<TreeviewSelect>>", lambda _: self._pick_scanned_form_folder())

        self.lst_sources = tk.Listbox(left, height=12, font=("Courier", 9))
        self.lst_sources.pack(fill=tk.BOTH, expand=True)
        self.lst_sources.bind("<<ListboxSelect>>", lambda _: self._pick_source_form())
        src_btns = ttk.Frame(left)
        src_btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(src_btns, text="Auto Detect Folders", command=self._auto_detect_source_folders).pack(side=tk.LEFT)
        ttk.Button(src_btns, text="New Source Form", command=self._new_source_form).pack(side=tk.RIGHT)

        self.source_vars: dict[str, tk.StringVar] = {}
        for key in ["SourceFormCode", "DisplayName", "Version", "PDFFilePath", "MappingKey", "EffectiveDate", "ExpiryDate", "Remarks"]:
            row = ttk.Frame(right, padding=(6, 4))
            row.pack(fill=tk.X)
            ttk.Label(row, text=key, width=LABEL_W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self.source_vars[key] = var
            widget = ttk.Entry(row, textvariable=var)
            widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
            if key == "PDFFilePath":
                ttk.Button(row, text="Browse PDF", command=self._browse_source_pdf).pack(side=tk.LEFT, padx=(4, 0))
                ttk.Button(row, text="Browse Folder", command=self._browse_source_folder).pack(side=tk.LEFT, padx=(4, 0))
        self.var_source_active = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="Active", variable=self.var_source_active).pack(anchor="w", padx=8, pady=(4, 0))
        btns = ttk.Frame(right, padding=6)
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Save Source Form", command=self._save_source_form).pack(side=tk.LEFT)
        ttk.Button(btns, text="Open Mapping Editor", command=self._open_mapping_for_source).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="Open PDF", command=self._open_selected_source_pdf).pack(side=tk.LEFT, padx=(6, 0))

    def _build_history_tab(self):
        paths = ttk.LabelFrame(self.tab_history, text="Local Files")
        paths.pack(fill=tk.X, pady=(0, 8))
        self.lbl_paths = ttk.Label(paths, text="", justify="left", anchor="w")
        self.lbl_paths.pack(fill=tk.X, padx=8, pady=8)

        buttons = ttk.Frame(self.tab_history)
        buttons.pack(fill=tk.X)
        for text, action in [
            ("Open Customer Workbook", lambda: self._open_path(config_loader.customer_workbook_path(self.settings))),
            ("Open History Log", lambda: self._open_path(config_loader.history_log_path(self.settings))),
            ("Open Procedures", lambda: self._open_path(config_loader._config_path("procedures.json"))),
            ("Open Source Forms", lambda: self._open_path(config_loader._config_path("source_forms.json"))),
            ("Open Procedure Items", lambda: self._open_path(config_loader._config_path("procedure_items.json"))),
            ("Open Mapping JSON", lambda: self._open_path(config_loader._config_path("forms.json"))),
            ("Open Output Folder", lambda: self._open_path(config_loader.output_folder_path(self.settings))),
        ]:
            ttk.Button(buttons, text=text, command=action).pack(side=tk.LEFT, padx=(0, 6), pady=4)

    def _build_coordinate_tab(self):
        try:
            import coord_picker
            self.coord_frame = coord_picker.CoordPickerFrame(self.tab_coordinate)
            self.coord_frame.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            self.coord_frame = None
            body = ttk.Frame(self.tab_coordinate, padding=16)
            body.pack(fill=tk.BOTH, expand=True)
            ttk.Label(body, text="Coordinate Pointer could not be loaded.", font=("", 10, "bold")).pack(anchor="w")
            ttk.Label(body, text=str(e), wraplength=760).pack(anchor="w", pady=(6, 0))

    # Loading -------------------------------------------------------------

    def _load_all(self):
        try:
            self._ensure_settings_exists()
            self.settings = config_loader.load_settings()
            config_loader.ensure_runtime_dirs(self.settings)
            config_loader.ensure_clients_xlsx()
            self.forms_config = config_loader.load_forms()
            self.procedures = catalog.load_procedures()
            self.source_forms = catalog.load_source_forms()
            self.procedure_items = catalog.load_procedure_items()
            self.procedure_by_code = catalog.procedure_map()
            self.source_by_code = catalog.source_form_map()
            workbook = config_loader.customer_workbook_path(self.settings)
            self.customers = excel_reader.load_customer_records(workbook)
            self.accounts = excel_reader.load_accounts(workbook)
            self.workbook_schema = excel_reader.workbook_schema(workbook)
            self.default_fields = excel_reader.sheet_field_defaults(workbook)
            self.staff_profile = excel_reader.load_staff_profile(workbook)
            self.rm_profile = self.staff_profile
            self.settings["_staff_profile"] = self.staff_profile
            self.settings["_rm_profile"] = self.staff_profile
        except ExcelLockedError as e:
            messagebox.showerror("Excel is open", str(e))
            return
        except Exception as e:
            messagebox.showerror("Load error", str(e))
            return

        self._refresh_all_controls()
        self._set_status(f"Loaded {len(self.customers)} customer(s), {len(self.accounts)} account(s), {len(self.procedures)} procedure(s), {len(self.source_forms)} source form(s).")

    def _ensure_settings_exists(self):
        path = config_loader._config_path("settings.json")
        if path.exists():
            return
        example = config_loader._config_path("settings.example.json")
        if example.exists():
            shutil.copy(example, path)

    def _refresh_all_controls(self):
        self._rebuild_tag_filters()
        self._refresh_generate_procedures()
        self._refresh_bulk_procedures()
        self._refresh_search_results()
        self._refresh_builder_procedures()
        self._refresh_source_forms()
        self._refresh_paths()
        self._refresh_recent_history()
        if hasattr(self, "var_form_root"):
            self.var_form_root.set(str(config_loader.forms_folder_path(self.settings)))
            self._scan_form_root(silent=True)
        rm_codes = self.staff_profile.get("staff_rm_codes") or self.staff_profile.get("rm_codes") or []
        branches = self.staff_profile.get("staff_branches") or self.staff_profile.get("branches") or []
        for cmb in [self.cmb_rm_code, getattr(self, "cmb_gen_rm_code", None)]:
            if cmb:
                cmb["values"] = rm_codes
        for cmb in [self.cmb_branch, getattr(self, "cmb_gen_branch", None)]:
            if cmb:
                cmb["values"] = [""] + branches
        self.var_rm_code.set(self.settings.get("default_rm_code", "") or self._first_rm_code())
        self.var_branch.set(self.settings.get("default_branch", ""))

    def _first_rm_code(self) -> str:
        codes = self.staff_profile.get("staff_rm_codes") or self.staff_profile.get("rm_codes") or []
        return codes[0] if codes else ""

    def _staff_name(self) -> str:
        return self.staff_profile.get("staff_name") or self.staff_profile.get("rm_name") or ""

    # Procedure helpers ---------------------------------------------------

    def _procedure_tags(self, proc: dict) -> list[str]:
        raw_tags = str(catalog.get_value(proc, "Tags", "") or "").strip()
        if raw_tags:
            tags = [t.strip() for t in raw_tags.replace(";", ",").split(",") if t.strip()]
            return tags[:10]

        text = " ".join([
            str(catalog.get_value(proc, "ProcedureCode", "")),
            str(catalog.get_value(proc, "DisplayName", "")),
            str(catalog.get_value(proc, "Description", "")),
            str(catalog.get_value(proc, "Category", "")),
        ]).lower()
        rules = [
            ("FD", ["fd", "fixed deposit"]),
            ("UT", ["ut", "unit trust"]),
            ("SI", ["si", "structured"]),
            ("Bond", ["bond", "sukuk"]),
            ("Account", ["account opening", "account"]),
            ("Risk", ["risk profile", "risk"]),
            ("Insurance", ["insurance", "takaful", "tokio", "banca"]),
            ("Maintenance", ["maintenance", "update customer", "customer maintenance"]),
        ]
        tags = [tag for tag, needles in rules if any(needle in text for needle in needles)]
        return tags[:10] or ["General"]

    def _all_procedure_tags(self) -> list[str]:
        seen = []
        for proc in self.procedures:
            if not catalog.is_active(proc):
                continue
            for tag in self._procedure_tags(proc):
                if tag not in seen:
                    seen.append(tag)
        preferred = ["FD", "UT", "SI", "Bond", "Account", "Risk", "Insurance", "Maintenance", "General"]
        ordered = [tag for tag in preferred if tag in seen]
        ordered.extend(tag for tag in seen if tag not in ordered)
        return ordered[:10]

    def _rebuild_tag_filters(self):
        for frame, vars_, refresh in [
            (getattr(self, "frm_gen_tags", None), getattr(self, "gen_tag_vars", {}), self._refresh_generate_procedures),
            (getattr(self, "frm_bulk_tags", None), getattr(self, "bulk_tag_vars", {}), self._refresh_bulk_procedures),
        ]:
            if frame is None:
                continue
            selected = {tag for tag, var in vars_.items() if var.get()}
            for child in frame.winfo_children():
                child.destroy()
            vars_.clear()
            for tag in self._all_procedure_tags():
                var = tk.BooleanVar(value=tag in selected)
                var.trace_add("write", lambda *_args, cb=refresh: cb())
                vars_[tag] = var
                ttk.Checkbutton(frame, text=tag, variable=var).pack(side=tk.LEFT, padx=(0, 8))

    @staticmethod
    def _selected_tags(tag_vars: dict[str, tk.BooleanVar]) -> set[str]:
        return {tag for tag, var in tag_vars.items() if var.get()}

    def _procedure_labels(self, search: str = "", tags: set[str] | None = None) -> list[str]:
        rows = [p for p in self.procedures if catalog.is_active(p)]
        query = search.strip().lower()
        if query:
            rows = [
                p for p in rows
                if query in " ".join([
                    str(catalog.get_value(p, "ProcedureCode", "")),
                    str(catalog.get_value(p, "DisplayName", "")),
                    str(catalog.get_value(p, "Description", "")),
                    ",".join(self._procedure_tags(p)),
                ]).lower()
            ]
        if tags:
            rows = [p for p in rows if tags.intersection(self._procedure_tags(p))]
        return [catalog.procedure_label(p) for p in rows]

    def _procedure_from_label(self, label: str) -> dict | None:
        code = str(label).split(" - ", 1)[0].strip()
        return self.procedure_by_code.get(code)

    def _auto_blank_enabled_for_proc(self, proc: dict | None) -> bool:
        if not proc:
            return bool(self.settings.get("auto_blank_after_odd", True))
        raw = catalog.get_value(proc, "AutoBlankAfterOdd", None)
        if raw in ("", None):
            raw = self.settings.get("auto_blank_after_odd", True)
        if isinstance(raw, bool):
            return raw
        return str(raw).strip().lower() in {"1", "yes", "y", "true", "on"}

    def _source_labels(self) -> list[str]:
        labels = []
        for source in self.source_forms:
            label = catalog.source_form_label(source)
            if not catalog.is_active(source):
                label += " [inactive]"
            labels.append(label)
        return labels

    def _source_from_label(self, label: str) -> dict | None:
        code = str(label).split(" - ", 1)[0].strip()
        return self.source_by_code.get(code)

    def _refresh_generate_procedures(self):
        labels = self._procedure_labels(
            self.var_gen_proc_search.get() if hasattr(self, "var_gen_proc_search") else "",
            self._selected_tags(getattr(self, "gen_tag_vars", {})),
        )
        self.cmb_gen_proc["values"] = labels
        if labels and self.var_gen_proc.get() not in labels:
            self.var_gen_proc.set(labels[0])
        elif not labels:
            self.var_gen_proc.set("")
        self._on_generate_proc_change()

    def _refresh_bulk_procedures(self):
        labels = self._procedure_labels(
            self.var_bulk_proc_search.get() if hasattr(self, "var_bulk_proc_search") else "",
            self._selected_tags(getattr(self, "bulk_tag_vars", {})),
        )
        self.cmb_bulk_proc["values"] = labels
        if labels and self.var_bulk_proc.get() not in labels:
            self.var_bulk_proc.set(labels[0])
        elif not labels:
            self.var_bulk_proc.set("")

    def _on_generate_proc_change(self):
        self._refresh_generate_structure()
        self._refresh_account_choices()
        self._refresh_manual_fields()

    def _open_selected_procedure_in_builder(self):
        proc = self._selected_generate_procedure()
        if not proc:
            self.nb.select(self.tab_builder)
            return
        label = catalog.procedure_label(proc)
        self.var_builder_proc.set(label)
        self.nb.select(self.tab_builder)
        self._load_builder_procedure()

    def _new_procedure_from_generate(self):
        self.nb.select(self.tab_builder)
        self._new_procedure()

    def _refresh_generate_structure(self):
        self.lst_gen_structure.delete(0, tk.END)
        proc = self._procedure_from_label(self.var_gen_proc.get())
        if not proc:
            return
        try:
            as_of = catalog.parse_catalog_date(self._normalize_date(self.var_date.get().strip())) or date.today()
        except ValueError:
            as_of = date.today()
        issues = catalog.source_form_date_issues(
            catalog.get_value(proc, "ProcedureCode"),
            self.source_by_code,
            self.procedure_items,
            as_of=as_of,
        )
        issue_by_code = {item["code"]: item for item in issues}
        if self._auto_blank_enabled_for_proc(proc):
            self.lst_gen_structure.insert(tk.END, "Auto blank page: ON after odd-page source PDFs")
        for item in catalog.procedure_items_for(catalog.get_value(proc, "ProcedureCode")):
            step = catalog.get_value(item, "StepNo")
            item_type = catalog.get_value(item, "ItemType")
            if item_type == "BlankPage":
                self.lst_gen_structure.insert(tk.END, f"{step:>2}. Blank page x {catalog.get_value(item, 'BlankPageCount', 1)}")
            else:
                code = catalog.get_value(item, "SourceFormCode")
                src = self.source_by_code.get(code, {})
                name = catalog.get_value(src, "DisplayName", code)
                issue = issue_by_code.get(code)
                if issue:
                    name = f"{name} [{issue['level'].upper()}]"
                self.lst_gen_structure.insert(tk.END, f"{step:>2}. {code:<6} {name}")

    def _refresh_manual_fields(self):
        for child in self.frm_manual.winfo_children():
            child.destroy()
        self.manual_entries = {}
        proc = self._procedure_from_label(self.var_gen_proc.get())
        if not proc:
            return
        all_fields = package_engine.data_fields_for_procedure(
            catalog.get_value(proc, "ProcedureCode"),
            self.forms_config,
            self.source_by_code,
            self.procedure_items,
            include_common=True,
        )
        selected = self._client_with_selected_account()
        default_rows = []
        editable_rows = []
        seen = set()
        for field in all_fields:
            key = field["key"]
            sheet = field.get("excel_sheet", "")
            identity = package_engine.field_identity(field)
            value = package_engine.field_value(selected, field)
            is_default = key in self.default_fields.get(sheet, set()) or key in self.default_fields.get(excel_reader.normalize_header(sheet), set())
            show_required_missing = field["required"] and value in ("", None)
            show_procedure_specific = key not in package_engine.COMMON_DATA_FIELDS
            if is_default and identity not in seen:
                seen.add(identity)
                default_rows.append(field)
            elif (show_required_missing or show_procedure_specific) and identity not in seen:
                seen.add(identity)
                editable_rows.append(field)
        if not default_rows and not editable_rows:
            ttk.Label(self.frm_manual, text="No mapped procedure-specific fields.").pack(anchor="w", padx=4, pady=4)
            return

        if default_rows:
            ttk.Label(self.frm_manual, text="Default Section (from * Excel headers)", font=("", 9, "bold")).pack(anchor="w", padx=4, pady=(4, 2))
        for field in default_rows:
            row = ttk.Frame(self.frm_manual)
            row.pack(fill=tk.X, padx=4, pady=3)
            label = field["label"] + (" *" if field["required"] else "")
            ttk.Label(row, text=label, width=24).pack(side=tk.LEFT)
            ttk.Label(row, text=str(package_engine.field_value(selected, field) or ""), relief="sunken", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(row, text="Edit", command=lambda f=field: self._edit_default_field(f)).pack(side=tk.LEFT, padx=(4, 0))

        if editable_rows:
            ttk.Label(self.frm_manual, text="Transaction / Form Inputs", font=("", 9, "bold")).pack(anchor="w", padx=4, pady=(8, 2))
        for field in editable_rows:
            row = ttk.Frame(self.frm_manual)
            row.pack(fill=tk.X, padx=4, pady=3)
            label = field["label"] + (" *" if field["required"] else "")
            ttk.Label(row, text=label, width=24).pack(side=tk.LEFT)
            var = tk.StringVar(value=str(package_engine.field_value(selected, field) or ""))
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.manual_entries[package_engine.field_identity(field)] = var

    def _edit_default_field(self, field: dict):
        if not self.selected_customer:
            messagebox.showwarning("No customer", "Select a customer first.")
            return
        sheet = field.get("excel_sheet") or ""
        if not sheet:
            messagebox.showwarning("No sheet", "This default field is not linked to a clients.xlsx sheet.")
            return
        client = self._client_with_selected_account()
        old_value = package_engine.field_value(client, field)

        dlg = tk.Toplevel(self)
        dlg.title("Edit Default Field")
        dlg.geometry("520x170")
        dlg.grab_set()
        var = tk.StringVar(value=str(old_value or ""))
        body = ttk.Frame(dlg, padding=12)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text=f"{sheet} / {field['key']}").pack(anchor="w")
        ttk.Entry(body, textvariable=var).pack(fill=tk.X, pady=(6, 8))
        ttk.Label(body, text="This updates clients.xlsx after creating a backup.").pack(anchor="w")

        def save():
            try:
                excel_reader.update_customer_field(
                    config_loader.customer_workbook_path(self.settings),
                    client,
                    sheet,
                    field["key"],
                    var.get().strip(),
                )
            except ExcelLockedError as e:
                messagebox.showerror("Excel is open", str(e))
                return
            except Exception as e:
                messagebox.showerror("Update failed", str(e))
                return
            dlg.destroy()
            self._load_all()
            messagebox.showinfo("Updated", "clients.xlsx updated. Data reloaded.")

        btns = ttk.Frame(body)
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Update Excel", command=save).pack(side=tk.LEFT)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=(6, 0))

    # Customer search -----------------------------------------------------

    def _refresh_search_results(self):
        if not hasattr(self, "cmb_customer"):
            return
        current = self.var_customer_choice.get()
        self.search_results = excel_reader.search_customers(self.customers, self.var_search.get(), limit=80)
        labels = [self._customer_label(c) for c in self.search_results]
        self.cmb_customer["values"] = labels
        if current in labels:
            self.var_customer_choice.set(current)
        else:
            self.var_customer_choice.set("")

    def _customer_label(self, c: dict) -> str:
        name = str(c.get("name") or c.get("client_name") or "-")[:30]
        cis = str(c.get("cis") or c.get("cif_no") or "-")[:14]
        ic = str(c.get("ic_number") or "-")[:14]
        pol = str(c.get("policy_number") or "-")[:14]
        return f"{name:<30}  CIS:{cis:<14}  IC:{ic:<14}  Policy:{pol}"

    def _pick_customer(self):
        if not hasattr(self, "cmb_customer"):
            return
        idx = self.cmb_customer.current()
        if idx < 0 or idx >= len(self.search_results):
            label = self.var_customer_choice.get()
            labels = [self._customer_label(c) for c in self.search_results]
            idx = labels.index(label) if label in labels else -1
        if idx < 0:
            return
        self.selected_customer = dict(self.search_results[idx])
        self.lbl_customer.config(text=f"Selected: {self.selected_customer.get('name') or self.selected_customer.get('client_name') or '-'}")
        self._refresh_account_choices()
        self._refresh_manual_fields()
        self._refresh_recent_history()

    def _account_type_for_procedure(self, proc: dict | None) -> str | None:
        if not proc:
            return None
        text = " ".join([
            str(catalog.get_value(proc, "ProcedureCode", "")),
            str(catalog.get_value(proc, "DisplayName", "")),
            str(catalog.get_value(proc, "Description", "")),
        ]).lower()
        matches = []
        if "bond" in text:
            matches.append("BOND")
        if "ut" in text or "unit trust" in text:
            matches.append("UT")
        if "si" in text or "structured" in text:
            matches.append("SI")
        return matches[0] if len(matches) == 1 else None

    def _refresh_account_choices(self):
        if not hasattr(self, "cmb_account"):
            return
        self.account_rows = []
        self.selected_account = None
        self.var_account.set("")
        if not self.selected_customer:
            self.cmb_account["values"] = []
            return
        proc = self._selected_generate_procedure()
        account_type = self._account_type_for_procedure(proc)
        self.account_rows = excel_reader.accounts_for_customer(self.accounts, self.selected_customer, account_type)
        if not self.account_rows and account_type:
            self.account_rows = excel_reader.accounts_for_customer(self.accounts, self.selected_customer)
        labels = [row.get("_label") or row.get("common_name") or row.get("account_number") for row in self.account_rows]
        self.cmb_account["values"] = labels
        if len(labels) == 1:
            self.var_account.set(labels[0])
            self.selected_account = dict(self.account_rows[0])
        elif labels:
            self.var_account.set("")
            self.selected_account = None

    def _pick_account(self):
        label = self.var_account.get()
        for row in self.account_rows:
            if (row.get("_label") or row.get("common_name") or row.get("account_number")) == label:
                self.selected_account = dict(row)
                break
        self._refresh_manual_fields()

    def _client_with_selected_account(self) -> dict:
        client = dict(self.selected_customer or {})
        sheet_data = dict(client.get("_sheet_data", {}))
        if self.selected_account:
            account = dict(self.selected_account)
            sheet_data["default_accounts"] = account
            if account.get("_source_sheet"):
                sheet_data[account["_source_sheet"]] = account
            client.update({k: v for k, v in account.items() if k not in {"_label", "_source_sheet"} and v not in ("", None)})
        client["_sheet_data"] = sheet_data
        return client

    # Recent history restore ------------------------------------------------

    def _refresh_recent_history(self):
        if not hasattr(self, "lst_recent_history"):
            return
        self.recent_history_rows = []
        self.lst_recent_history.delete(0, tk.END)
        if not self.selected_customer:
            self.lst_recent_history.insert(tk.END, "No customer selected")
            return
        try:
            self.recent_history_rows = excel_reader.load_recent_history(
                config_loader.history_log_path(self.settings),
                self.selected_customer,
                limit=10,
            )
        except ExcelLockedError:
            self.lst_recent_history.insert(tk.END, "HistoryLog.xlsx is open")
            return
        except Exception:
            self.lst_recent_history.insert(tk.END, "History unavailable")
            return
        if not self.recent_history_rows:
            self.lst_recent_history.insert(tk.END, "No recent history")
            return
        for row in self.recent_history_rows:
            self.lst_recent_history.insert(tk.END, self._history_label(row))

    def _history_label(self, row: dict) -> str:
        stamp = str(row.get("GeneratedDateTime", ""))[:16]
        proc = str(row.get("ProcedureName") or row.get("ProcedureCode") or "-")[:28]
        amount = str(row.get("Amount") or "")[:14]
        status = str(row.get("Status") or "")[:18]
        return f"{stamp:<16}  {proc:<28}  {amount:<14}  {status}"

    def _restore_selected_history(self):
        sel = self.lst_recent_history.curselection()
        if not sel or sel[0] >= len(self.recent_history_rows):
            return
        row = self.recent_history_rows[sel[0]]
        payload = row.get("_payload", {}) if isinstance(row.get("_payload"), dict) else {}
        proc_code = payload.get("procedure_code") or row.get("ProcedureCode")
        proc = self.procedure_by_code.get(proc_code)
        if proc:
            label = catalog.procedure_label(proc)
            if label not in self.cmb_gen_proc["values"]:
                self.var_gen_proc_search.set("")
                self._refresh_generate_procedures()
            self.var_gen_proc.set(label)

        session = payload.get("session", {}) if isinstance(payload.get("session"), dict) else {}
        self._restore_session_from_history(session)
        self._on_generate_proc_change()
        self._restore_account_from_history(payload.get("account", {}))
        self._refresh_manual_fields()
        self._restore_manual_values_from_history(payload, row)
        self._refresh_generate_structure()
        self._set_status(f"Restored history: {row.get('ProcedureName') or row.get('ProcedureCode') or 'procedure'}")

    def _restore_session_from_history(self, session: dict):
        rm_code = session.get("staff_rm_code") or session.get("rm_code")
        branch = session.get("rm_branch") or session.get("branch_code")
        date_text = session.get("date")
        if rm_code:
            self.var_rm_code.set(str(rm_code))
        if branch is not None:
            self.var_branch.set(str(branch))
        if date_text:
            self.var_date.set(str(date_text))

    def _restore_account_from_history(self, account_payload):
        if not isinstance(account_payload, dict) or not account_payload:
            return
        label = account_payload.get("_label") or account_payload.get("common_name") or account_payload.get("account_number")
        target_number = excel_reader.normalize_lookup_key(account_payload.get("account_number", ""))
        target_common = excel_reader.normalize_lookup_key(account_payload.get("common_name", ""))
        for row in self.account_rows:
            row_label = row.get("_label") or row.get("common_name") or row.get("account_number")
            if label and row_label == label:
                self.var_account.set(row_label)
                self.selected_account = dict(row)
                return
            if target_number and excel_reader.normalize_lookup_key(row.get("account_number", "")) == target_number:
                self.var_account.set(row_label)
                self.selected_account = dict(row)
                return
            if target_common and excel_reader.normalize_lookup_key(row.get("common_name", "")) == target_common:
                self.var_account.set(row_label)
                self.selected_account = dict(row)
                return

    def _restore_manual_values_from_history(self, payload: dict, row: dict):
        manual = payload.get("manual_values", {}) if isinstance(payload.get("manual_values"), dict) else {}
        if not manual:
            manual = {
                "amount": row.get("Amount", ""),
                "fd_details": row.get("FDDetails", ""),
                "product_type": row.get("ProductType", ""),
                "action_purpose": row.get("ActionPurpose", ""),
                "follow_up_note": row.get("FollowUpNote", ""),
            }
        normalized = {
            excel_reader.normalize_header(k): v
            for k, v in manual.items()
            if v not in ("", None)
        }
        for identity, var in self.manual_entries.items():
            if identity in manual:
                var.set(str(manual[identity]))
                continue
            _, key = package_engine.split_field_identity(identity)
            norm_key = excel_reader.normalize_header(key)
            if norm_key in normalized:
                var.set(str(normalized[norm_key]))

    # Generation ---------------------------------------------------------

    def _session_context(self) -> dict:
        return {
            "staff_rm_code": self.var_rm_code.get().strip(),
            "rm_code": self.var_rm_code.get().strip(),
            "rm_branch": self.var_branch.get().strip(),
            "date": self._normalize_date(self.var_date.get().strip()),
        }

    @staticmethod
    def _normalize_date(raw: str) -> str:
        import datetime as _dt
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y%m%d"):
            try:
                return _dt.datetime.strptime(raw, fmt).strftime("%d/%m/%Y")
            except ValueError:
                pass
        raise ValueError("Date must be dd/mm/yyyy.")

    def _selected_generate_procedure(self) -> dict | None:
        return self._procedure_from_label(self.var_gen_proc.get())

    def _confirm_source_form_dates(self, proc: dict, session: dict) -> bool:
        as_of = catalog.parse_catalog_date(session.get("date")) or date.today()
        issues = catalog.source_form_date_issues(
            catalog.get_value(proc, "ProcedureCode"),
            self.source_by_code,
            self.procedure_items,
            as_of=as_of,
        )
        blockers = [i["message"] for i in issues if i["level"] == "block"]
        warnings = [i["message"] for i in issues if i["level"] == "warn"]
        if blockers:
            messagebox.showerror("Expired or not-yet-effective forms", "\n".join(blockers))
            return False
        if warnings:
            return messagebox.askyesno(
                "Forms near expiry",
                "\n".join(warnings) + "\n\nContinue generating?",
            )
        return True

    def _generate_single(self):
        if not self.selected_customer:
            messagebox.showwarning("No customer", "Select a customer first.")
            return
        proc = self._selected_generate_procedure()
        if not proc:
            messagebox.showwarning("No procedure", "Select a procedure first.")
            return
        try:
            session = self._session_context()
        except ValueError as e:
            messagebox.showerror("Bad date", str(e))
            return
        if not self._confirm_source_form_dates(proc, session):
            return

        manual = {k: v.get().strip() for k, v in self.manual_entries.items()}
        client = package_engine.merge_manual_values(self._client_with_selected_account(), manual)
        missing = package_engine.missing_required_fields(
            catalog.get_value(proc, "ProcedureCode"),
            client,
            self.forms_config,
            self.source_by_code,
            self.procedure_items,
        )
        if missing:
            messagebox.showerror("Missing required fields", "\n".join(f["label"] for f in missing))
            return

        try:
            result = package_engine.generate_package(
                procedure=proc,
                source_forms=self.source_by_code,
                procedure_items=self.procedure_items,
                forms_config=self.forms_config,
                client=client,
                settings=self.settings,
                output_root=config_loader.output_folder_path(self.settings),
                session=session,
                manual_values=manual,
            )
            result["account"] = dict(self.selected_account or {})
        except Exception as e:
            try:
                excel_reader.append_history_rows(
                    config_loader.history_log_path(self.settings),
                    [package_engine.error_history_row(client, proc, "Failed", str(e), self._staff_name())],
                )
            except ExcelLockedError:
                pass
            messagebox.showerror("Generate failed", str(e))
            return

        history_warning = ""
        try:
            excel_reader.append_history_rows(
                config_loader.history_log_path(self.settings),
                [package_engine.history_row(result, generated_by=self._staff_name())],
            )
        except ExcelLockedError:
            history_warning = "\n\nHistoryLog.xlsx is open, so this generation was not logged."

        message = f"Saved:\n{result['output_path']}"
        if result["warnings"]:
            message += "\n\nReview needed:\n" + "\n".join(result["warnings"])
        message += history_warning
        messagebox.showinfo("Generated", message)
        self._set_status(f"Generated {result['output_path'].name}")
        self._refresh_recent_history()
        if self.settings.get("auto_open_output", True):
            self._open_path(result["output_path"].parent)

    # Bulk ---------------------------------------------------------------

    def _parse_cis_text(self) -> list[str]:
        import re
        text = self.txt_cis.get("1.0", tk.END)
        return [v.strip() for v in re.split(r"[\n,;]+", text) if v.strip()]

    def _import_cis_list(self):
        path = filedialog.askopenfilename(
            title="Import CIS list",
            filetypes=[("Excel/CSV/Text", "*.xlsx *.xlsm *.csv *.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            rows = excel_reader.read_cis_list(Path(path))
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return
        self.txt_cis.delete("1.0", tk.END)
        self.txt_cis.insert(tk.END, "\n".join(rows))

    def _analyze_bulk(self):
        self.bulk_rows = []
        self.tree_bulk.delete(*self.tree_bulk.get_children())
        seen = set()
        for cis in self._parse_cis_text():
            key = excel_reader.normalize_lookup_key(cis)
            if key in seen:
                row = {"cis": cis, "customer": {}, "status": "Duplicate"}
            else:
                seen.add(key)
                customer = excel_reader.find_customer_by_cis(self.customers, cis)
                if customer:
                    row = {"cis": cis, "customer": dict(customer), "status": "Matched"}
                else:
                    row = {"cis": cis, "customer": {}, "status": "CIS Not Found"}
            self.bulk_rows.append(row)
            name = row["customer"].get("name") or row["customer"].get("client_name") or ""
            self.tree_bulk.insert("", tk.END, values=(cis, name, row["status"]))
        self._set_status(f"Analyzed {len(self.bulk_rows)} CIS row(s).")

    def _generate_bulk(self):
        proc = self._procedure_from_label(self.var_bulk_proc.get())
        if not proc:
            messagebox.showwarning("No procedure", "Select a procedure first.")
            return
        # Always re-analyze so edits made after a previous analysis are honored.
        self._analyze_bulk()
        try:
            session = self._session_context()
        except ValueError as e:
            messagebox.showerror("Bad date", str(e))
            return
        if not self._confirm_source_form_dates(proc, session):
            return

        batch_root = config_loader.output_folder_path(self.settings) / f"BulkExport_{package_engine.yyyymmdd(session['date'])}"
        history = []
        saved = 0
        review = 0
        failed = 0

        for row in self.bulk_rows:
            if row["status"] != "Matched":
                failed += 1
                history.append(package_engine.error_history_row(
                    {"cis": row["cis"]}, proc, row["status"], row["status"], self._staff_name()
                ))
                continue
            client = row["customer"]
            try:
                result = package_engine.generate_package(
                    procedure=proc,
                    source_forms=self.source_by_code,
                    procedure_items=self.procedure_items,
                    forms_config=self.forms_config,
                    client=client,
                    settings=self.settings,
                    output_root=config_loader.output_folder_path(self.settings),
                    session=session,
                    bulk_root=batch_root,
                    client_folder=bool(self.settings.get("bulk_create_client_folders", True)),
                )
                history.append(package_engine.history_row(result, self._staff_name()))
                if result["status"] == "Success":
                    saved += 1
                else:
                    review += 1
            except Exception as e:
                failed += 1
                history.append(package_engine.error_history_row(client, proc, "Failed", str(e), self._staff_name()))

        try:
            excel_reader.append_history_rows(config_loader.history_log_path(self.settings), history)
        except ExcelLockedError as e:
            messagebox.showerror("History log locked", str(e))
            return
        self._set_status(f"Bulk export complete. Success {saved}; review {review}; failed {failed}; log rows {len(history)}.")
        if saved or review:
            self._open_path(batch_root)

    # Procedure builder ---------------------------------------------------

    def _refresh_builder_procedures(self):
        labels = [catalog.procedure_label(p) for p in self.procedures]
        self.cmb_builder_proc["values"] = labels
        if labels and self.var_builder_proc.get() not in labels:
            self.var_builder_proc.set(labels[0])
        self._refresh_available_sources()
        self._load_builder_procedure()

    def _refresh_available_sources(self):
        if not hasattr(self, "lst_available_sources"):
            return
        query = self.var_source_filter.get().strip().lower()
        self.lst_available_sources.delete(0, tk.END)
        for label in self._source_labels():
            if not query or query in label.lower():
                self.lst_available_sources.insert(tk.END, label)

    def _load_builder_procedure(self):
        proc = self._procedure_from_label(self.var_builder_proc.get())
        if not proc:
            return
        for key, var in self.builder_vars.items():
            value = catalog.get_value(proc, key, "")
            if key == "Tags" and not value:
                value = ", ".join(self._procedure_tags(proc))
            if key == "DefaultOutputName" and not value:
                value = "{client_name}_{date}_{procedure}"
            var.set(str(value))
        self.var_builder_active.set(catalog.is_active(proc))
        self.var_builder_auto_blank.set(self._auto_blank_enabled_for_proc(proc))
        self.builder_items = [dict(i) for i in catalog.procedure_items_for(catalog.get_value(proc, "ProcedureCode"))]
        self._render_builder_items()

    def _render_builder_items(self):
        self.lst_builder_items.delete(0, tk.END)
        for item in self.builder_items:
            step = catalog.get_value(item, "StepNo")
            typ = catalog.get_value(item, "ItemType")
            if typ == "BlankPage":
                text = f"{step:>2}. Blank page x {catalog.get_value(item, 'BlankPageCount', 1)}"
            else:
                code = catalog.get_value(item, "SourceFormCode")
                src = self.source_by_code.get(code, {})
                text = f"{step:>2}. {code:<6} {catalog.get_value(src, 'DisplayName', code)}"
            self.lst_builder_items.insert(tk.END, text)

    def _new_procedure(self):
        next_num = 1
        existing = {catalog.get_value(p, "ProcedureCode") for p in self.procedures}
        while f"P{next_num:03d}" in existing:
            next_num += 1
        for key, var in self.builder_vars.items():
            var.set("")
        self.builder_vars["ProcedureCode"].set(f"P{next_num:03d}")
        self.builder_vars["Version"].set("V01")
        self.builder_vars["DefaultOutputName"].set("{client_name}_{date}_{procedure}")
        self.var_builder_active.set(True)
        self.var_builder_auto_blank.set(bool(self.settings.get("auto_blank_after_odd", True)))
        self.builder_items = []
        self._render_builder_items()

    def _save_builder_procedure(self):
        code = self.builder_vars["ProcedureCode"].get().strip()
        if not code:
            messagebox.showerror("Missing", "ProcedureCode is required.")
            return
        row = {k: v.get().strip() for k, v in self.builder_vars.items()}
        row["Active"] = self.var_builder_active.get()
        row["AutoBlankAfterOdd"] = self.var_builder_auto_blank.get()
        found = False
        for idx, proc in enumerate(self.procedures):
            if catalog.get_value(proc, "ProcedureCode") == code:
                self.procedures[idx] = row
                found = True
                break
        if not found:
            self.procedures.append(row)

        remaining = [i for i in self.procedure_items if catalog.get_value(i, "ProcedureCode") != code]
        for idx, item in enumerate(self.builder_items, 1):
            item["ProcedureCode"] = code
            item["StepNo"] = idx
            remaining.append(item)
        catalog.save_procedures(self.procedures)
        catalog.save_procedure_items(remaining)
        saved_label = catalog.procedure_label(row)
        self._load_all()
        self.var_builder_proc.set(saved_label)
        self._refresh_generate_procedures()
        self.var_gen_proc.set(saved_label)
        self._load_builder_procedure()
        self._set_status(f"Saved procedure {code}.")

    def _builder_add_source(self):
        sel = self.lst_available_sources.curselection()
        if not sel:
            return
        source = self._source_from_label(self.lst_available_sources.get(sel[0]))
        if not source:
            return
        if not catalog.is_active(source):
            messagebox.showwarning("Inactive source form", "Activate the source form before adding it to a procedure.")
            return
        self.builder_items.append({
            "ProcedureCode": self.builder_vars["ProcedureCode"].get(),
            "StepNo": len(self.builder_items) + 1,
            "ItemType": "SourceForm",
            "SourceFormCode": catalog.get_value(source, "SourceFormCode"),
            "BlankPageCount": 0,
            "Remarks": "",
        })
        self._render_builder_items()

    def _builder_add_blank(self):
        self.builder_items.append({
            "ProcedureCode": self.builder_vars["ProcedureCode"].get(),
            "StepNo": len(self.builder_items) + 1,
            "ItemType": "BlankPage",
            "SourceFormCode": "",
            "BlankPageCount": 1,
            "Remarks": "",
        })
        self._render_builder_items()

    def _builder_remove(self):
        sel = self.lst_builder_items.curselection()
        if not sel:
            return
        self.builder_items.pop(sel[0])
        self._renumber_builder_items()

    def _builder_move(self, delta: int):
        sel = self.lst_builder_items.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= len(self.builder_items):
            return
        self.builder_items[i], self.builder_items[j] = self.builder_items[j], self.builder_items[i]
        self._renumber_builder_items()
        self.lst_builder_items.selection_set(j)

    def _renumber_builder_items(self):
        for idx, item in enumerate(self.builder_items, 1):
            item["StepNo"] = idx
        self._render_builder_items()

    # Source form library -------------------------------------------------

    def _refresh_source_forms(self):
        self.lst_sources.delete(0, tk.END)
        for src in self.source_forms:
            self.lst_sources.insert(tk.END, catalog.source_form_label(src))

    def _form_root_path(self) -> Path:
        raw = self.var_form_root.get().strip() if hasattr(self, "var_form_root") else ""
        return config_loader.resolve_path(raw) if raw else config_loader.forms_folder_path(self.settings)

    @staticmethod
    def _display_from_folder(folder_name: str, code: str = "") -> str:
        display = folder_name
        if code and folder_name.lower().startswith(code.lower()):
            display = folder_name[len(code):]
        display = display.strip(" -_")
        display = display.replace("_", " ").replace("-", " ")
        display = " ".join(display.split())
        return display.title() if display else (code or folder_name)

    @staticmethod
    def _source_code_for_folder(folder_name: str, existing: set[str]) -> str:
        code = catalog.source_code_from_folder(folder_name) or folder_name.strip()
        code = code.replace(" ", "_")
        if code and code not in existing:
            return code
        base = code or "SF"
        idx = 2
        while f"{base}_{idx}" in existing:
            idx += 1
        return f"{base}_{idx}"

    def _choose_form_root_folder(self):
        path = filedialog.askdirectory(title="Select form root folder")
        if not path:
            return
        self.var_form_root.set(path)
        self.settings["forms_folder"] = path
        config_loader.save_settings(self.settings)
        self._scan_form_root()

    def _scan_form_root(self, silent: bool = False):
        if not hasattr(self, "tree_form_scan"):
            return
        root = self._form_root_path()
        self.tree_form_scan.delete(*self.tree_form_scan.get_children())
        if not root.exists():
            if not silent:
                messagebox.showerror("Folder not found", str(root))
            return
        self.settings["forms_folder"] = str(root)
        if not silent:
            config_loader.save_settings(self.settings)
        rows = config_loader.scan_form_subfolders(self.settings, root)
        for row in rows:
            if row["pdf_count"] == 0:
                pdf_status = "Missing PDF"
            elif row["pdf_count"] == 1:
                pdf_status = row["pdf_files"][0]
            else:
                pdf_status = f"Multiple PDFs ({row['pdf_count']})"
            if row["mapping_exists"]:
                mapping_status = f"{row['mapping_key'] or 'mapped'} ({row['field_count']} fields)"
            else:
                mapping_status = "No mapping.json"
            self.tree_form_scan.insert(
                "",
                tk.END,
                iid=row["path"],
                values=(row["folder"], pdf_status, mapping_status, row["mapping_updated"]),
            )
        if not silent:
            empty = sum(1 for row in rows if row["pdf_count"] == 0)
            multiple = sum(1 for row in rows if row["pdf_count"] > 1)
            self._set_status(f"Scanned {len(rows)} form folder(s). Missing PDF {empty}; multiple PDF {multiple}.")

    def _pick_scanned_form_folder(self):
        sel = self.tree_form_scan.selection() if hasattr(self, "tree_form_scan") else ()
        if not sel:
            return
        folder = Path(sel[0])
        self._new_source_form()
        self._set_source_pdf_path(folder)

    def _pick_source_form(self):
        sel = self.lst_sources.curselection()
        if not sel:
            return
        src = self.source_forms[sel[0]]
        self.selected_source_code = catalog.get_value(src, "SourceFormCode")
        for key, var in self.source_vars.items():
            var.set(str(catalog.get_value(src, key, "")))
        self.var_source_active.set(catalog.is_active(src))

    def _new_source_form(self):
        self.selected_source_code = None
        next_num = 1
        existing = {catalog.get_value(s, "SourceFormCode") for s in self.source_forms}
        while f"SF{next_num:03d}" in existing:
            next_num += 1
        for var in self.source_vars.values():
            var.set("")
        code = f"SF{next_num:03d}"
        self.source_vars["SourceFormCode"].set(code)
        self.source_vars["MappingKey"].set(code)
        self.source_vars["Version"].set("V01")
        self.var_source_active.set(True)

    def _auto_detect_source_folders(self):
        root = self._form_root_path()
        if not root.exists():
            messagebox.showerror("Folder not found", str(root))
            return
        self.settings["forms_folder"] = str(root)
        config_loader.save_settings(self.settings)
        folders = config_loader.scan_form_subfolders(self.settings, root)
        existing = {catalog.get_value(s, "SourceFormCode") for s in self.source_forms}
        existing_paths = {
            Path(str(catalog.get_value(s, "PDFFilePath", "") or "")).name
            for s in self.source_forms
            if str(catalog.get_value(s, "PDFFilePath", "") or "").strip()
        }
        added = 0
        skipped = 0
        for folder in folders:
            folder_name = folder["folder"]
            suggested_code = catalog.source_code_from_folder(folder_name) or folder_name.replace(" ", "_")
            if folder_name in existing_paths or suggested_code in existing:
                skipped += 1
                continue
            code = self._source_code_for_folder(folder_name, existing)
            if folder["pdf_count"] != 1:
                skipped += 1
                continue
            display = self._display_from_folder(folder_name, code)
            self.source_forms.append({
                "SourceFormCode": code,
                "DisplayName": display,
                "Version": "V01",
                "PDFFilePath": folder_name,
                "MappingKey": code,
                "Active": True,
                "EffectiveDate": "",
                "ExpiryDate": "",
                "Remarks": "Auto-detected from forms folder.",
            })
            existing.add(code)
            added += 1
        if added:
            catalog.save_source_forms(self.source_forms)
            self._load_all()
        messagebox.showinfo("Auto detect complete", f"Added {added} source form(s).\nSkipped {skipped} folder(s).")

    def _set_source_pdf_path(self, path: str | Path):
        p = Path(path)
        base = config_loader.forms_folder_path(self.settings)
        try:
            rel = p.relative_to(base)
            self.source_vars["PDFFilePath"].set(str(rel))
        except ValueError:
            self.source_vars["PDFFilePath"].set(str(p))
        self._autofill_source_identity_from_path(p, force_code=self.selected_source_code is None)

    def _autofill_source_identity_from_path(self, path: Path, force_code: bool = False):
        folder_name = path.name if path.is_dir() else path.parent.name
        if not folder_name:
            return
        code = catalog.source_code_from_folder(folder_name) or re.sub(r"[^A-Za-z0-9]+", "_", folder_name).strip("_")
        display = folder_name
        if code:
            display = self._display_from_folder(folder_name, code)
        if code and (force_code or not self.source_vars["SourceFormCode"].get().strip()):
            self.source_vars["SourceFormCode"].set(code)
            if force_code or not self.source_vars["MappingKey"].get().strip():
                self.source_vars["MappingKey"].set(code)
        if display and (
            not self.source_vars["DisplayName"].get().strip()
            or self.source_vars["DisplayName"].get().strip() == self.source_vars["SourceFormCode"].get().strip()
        ):
            self.source_vars["DisplayName"].set(display)

    def _browse_source_pdf(self):
        path = filedialog.askopenfilename(title="Select source PDF", filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
        if path:
            self._set_source_pdf_path(path)

    def _browse_source_folder(self):
        path = filedialog.askdirectory(title="Select source form folder")
        if path:
            self._set_source_pdf_path(path)

    def _save_source_form(self):
        code = self.source_vars["SourceFormCode"].get().strip()
        if not code:
            messagebox.showerror("Missing", "SourceFormCode is required.")
            return
        row = {k: v.get().strip() for k, v in self.source_vars.items()}
        row["Active"] = self.var_source_active.get()
        if not row.get("MappingKey"):
            row["MappingKey"] = code
        if row["Active"]:
            problem = catalog.source_pdf_path_problem(row, self.settings)
            if problem:
                messagebox.showerror("PDF path issue", problem)
                return
            for key in ("EffectiveDate", "ExpiryDate"):
                try:
                    catalog.parse_catalog_date(row.get(key, ""))
                except ValueError as e:
                    messagebox.showerror("Bad date", f"{key}: {e}")
                    return
        found = False
        for idx, src in enumerate(self.source_forms):
            if catalog.get_value(src, "SourceFormCode") == code:
                self.source_forms[idx] = row
                found = True
                break
        if not found:
            self.source_forms.append(row)
        catalog.save_source_forms(self.source_forms)
        self._load_all()
        self._set_status(f"Saved source form {code}.")

    def _open_selected_source_pdf(self):
        code = self.source_vars["SourceFormCode"].get().strip()
        src = self.source_by_code.get(code)
        if not src:
            return
        path = catalog.resolve_source_pdf_path(src, self.settings)
        if path.exists():
            self._open_path(path)
        else:
            messagebox.showwarning("PDF not found", str(path))

    # Settings/history ----------------------------------------------------

    def _refresh_paths(self):
        text = "\n".join([
            f"Customer workbook: {config_loader.customer_workbook_path(self.settings)}",
            f"Source forms folder: {config_loader.forms_folder_path(self.settings)}",
            f"Output folder: {config_loader.output_folder_path(self.settings)}",
            f"History log: {config_loader.history_log_path(self.settings)}",
            f"Procedures: {config_loader._config_path('procedures.json')}",
            f"Source forms: {config_loader._config_path('source_forms.json')}",
            f"Mappings: {config_loader._config_path('forms.json')}",
        ])
        self.lbl_paths.config(text=text)

    def _open_settings_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("680x430")
        dlg.grab_set()
        vars_: dict[str, tk.StringVar] = {}

        rows = [
            ("Customer workbook", "customer_workbook", "file"),
            ("Source form folder", "forms_folder", "dir"),
            ("Output folder", "output_folder", "dir"),
            ("History log", "history_log_path", "file"),
            ("Default RM code", "default_rm_code", ""),
            ("Default branch", "default_branch", ""),
            ("Default font", "default_font", ""),
            ("Default font size", "default_font_size", ""),
        ]
        for label, key, browse in rows:
            row = ttk.Frame(dlg, padding=(10, 5))
            row.pack(fill=tk.X)
            ttk.Label(row, text=label, width=20).pack(side=tk.LEFT)
            var = tk.StringVar(value=str(self.settings.get(key, "")))
            vars_[key] = var
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            if browse == "dir":
                ttk.Button(row, text="Browse", command=lambda v=var: v.set(filedialog.askdirectory() or v.get())).pack(side=tk.LEFT, padx=(4, 0))
            elif browse == "file":
                ttk.Button(row, text="Browse", command=lambda v=var: v.set(filedialog.askopenfilename() or v.get())).pack(side=tk.LEFT, padx=(4, 0))

        auto_open = tk.BooleanVar(value=bool(self.settings.get("auto_open_output", True)))
        client_folders = tk.BooleanVar(value=bool(self.settings.get("bulk_create_client_folders", True)))
        auto_blank = tk.BooleanVar(value=bool(self.settings.get("auto_blank_after_odd", True)))
        ttk.Checkbutton(dlg, text="Auto-open output folder after generation", variable=auto_open).pack(anchor="w", padx=12, pady=(8, 0))
        ttk.Checkbutton(dlg, text="Bulk export creates one folder per client", variable=client_folders).pack(anchor="w", padx=12)
        ttk.Checkbutton(dlg, text="Auto insert blank page after odd-page PDFs", variable=auto_blank).pack(anchor="w", padx=12)

        btns = ttk.Frame(dlg, padding=10)
        btns.pack(fill=tk.X)

        def save():
            new_settings = dict(self.settings)
            for key, var in vars_.items():
                val = var.get().strip()
                if key == "default_font_size":
                    try:
                        val = int(val)
                    except ValueError:
                        val = 10
                new_settings[key] = val
            new_settings["auto_open_output"] = auto_open.get()
            new_settings["bulk_create_client_folders"] = client_folders.get()
            new_settings["auto_blank_after_odd"] = auto_blank.get()
            config_loader.save_settings(new_settings)
            dlg.destroy()
            self._load_all()

        ttk.Button(btns, text="Save", command=save).pack(side=tk.LEFT)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=(6, 0))

    def _open_mapping_for_source(self):
        code = self.source_vars["SourceFormCode"].get().strip()
        src = self.source_by_code.get(code)
        if not src:
            self._show_coordinate_tab()
            return
        self._show_coordinate_tab(src)

    def _show_coordinate_tab(self, source_form: dict | None = None):
        self.nb.select(self.tab_coordinate)
        if source_form and getattr(self, "coord_frame", None):
            try:
                self.coord_frame.load_source_form(source_form, self.settings)
                self._set_status(f"Loaded mapping editor for {catalog.source_form_label(source_form)}.")
            except Exception as e:
                messagebox.showerror("Cannot load mapping editor", str(e))

    def _open_coord_picker(self, args: list[str] | None = None):
        self._show_coordinate_tab()
        args = args or []
        if getattr(self, "coord_frame", None) and not args:
            return
        if getattr(sys, "frozen", False):
            exe = Path(sys.executable).parent / "CoordPicker.exe"
            if exe.exists():
                try:
                    subprocess.Popen([str(exe), *args])
                except Exception as e:
                    messagebox.showerror("Cannot open", str(e))
                return
        script = Path(__file__).parent / "coord_picker.py"
        if script.exists() and not getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable, str(script), *args])
            return
        messagebox.showwarning("Mapping editor not found", "CoordPicker.exe must be next to FormFiller.exe.")

    def _open_path(self, path: Path):
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Cannot open", f"{path}\n\n{e}")

    def _set_status(self, text: str):
        self.lbl_status.config(text=text)


if __name__ == "__main__":
    ProcedureAutomationApp().mainloop()
