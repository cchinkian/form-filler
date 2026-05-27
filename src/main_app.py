"""
Offline Windows PDF Procedure Automation Tool.

The app is built around:
Excel customer workbook -> Source Forms -> coordinate mapping -> Procedures
-> combined PDF output -> Excel history log.
"""
from __future__ import annotations

import os
import platform
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

        try:
            config_loader.backup_settings_on_launch()
        except Exception:
            pass

        self._build_ui()
        self._load_all()

    # UI -----------------------------------------------------------------

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)
        ttk.Label(top, text="PDF Procedure Automation", font=("", 14, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="Reload", command=self._load_all).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Settings", command=self._open_settings_dialog).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Mapping Editor", command=self._open_coord_picker).pack(side=tk.RIGHT, padx=(4, 0))
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
        ttk.Entry(session, textvariable=self.var_date, width=14).pack(side=tk.LEFT, padx=(4, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))

        self.tab_generate = ttk.Frame(self.nb, padding=10)
        self.tab_bulk = ttk.Frame(self.nb, padding=10)
        self.tab_builder = ttk.Frame(self.nb, padding=10)
        self.tab_sources = ttk.Frame(self.nb, padding=10)
        self.tab_history = ttk.Frame(self.nb, padding=10)

        self.nb.add(self.tab_generate, text="Generate Package")
        self.nb.add(self.tab_bulk, text="Bulk Export")
        self.nb.add(self.tab_builder, text="Procedure Builder")
        self.nb.add(self.tab_sources, text="Source Forms")
        self.nb.add(self.tab_history, text="History / Settings")

        self._build_generate_tab()
        self._build_bulk_tab()
        self._build_builder_tab()
        self._build_sources_tab()
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
        ttk.Entry(search_row, textvariable=self.var_search, width=34).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.lst_customers = tk.Listbox(left, height=9, font=("Courier", 9), activestyle="dotbox")
        self.lst_customers.pack(fill=tk.X, pady=(6, 8))
        self.lst_customers.bind("<<ListboxSelect>>", lambda _: self._pick_customer())

        self.lbl_customer = ttk.Label(left, text="No customer selected.", anchor="w", font=("", 9, "bold"))
        self.lbl_customer.pack(fill=tk.X, pady=(0, 8))

        proc_box = ttk.LabelFrame(left, text="Procedure")
        proc_box.pack(fill=tk.X, pady=(0, 8))
        row = ttk.Frame(proc_box, padding=6)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Category", width=LABEL_W).pack(side=tk.LEFT)
        self.var_gen_category = tk.StringVar()
        self.cmb_gen_category = ttk.Combobox(row, textvariable=self.var_gen_category, state="readonly")
        self.cmb_gen_category.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cmb_gen_category.bind("<<ComboboxSelected>>", lambda _: self._refresh_generate_procedures())
        ttk.Button(row, text="Add Category", command=self._add_category_dialog).pack(side=tk.LEFT, padx=(4, 0))

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

        field_box = ttk.LabelFrame(right, text="Missing / Procedure Fields")
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

        ttk.Label(top, text="Category", width=LABEL_W).pack(side=tk.LEFT)
        self.var_bulk_category = tk.StringVar()
        self.cmb_bulk_category = ttk.Combobox(top, textvariable=self.var_bulk_category, state="readonly", width=28)
        self.cmb_bulk_category.pack(side=tk.LEFT, padx=(0, 8))
        self.cmb_bulk_category.bind("<<ComboboxSelected>>", lambda _: self._refresh_bulk_procedures())

        ttk.Label(top, text="Procedure").pack(side=tk.LEFT)
        self.var_bulk_proc = tk.StringVar()
        self.cmb_bulk_proc = ttk.Combobox(top, textvariable=self.var_bulk_proc, state="readonly", width=36)
        self.cmb_bulk_proc.pack(side=tk.LEFT, padx=(4, 8))

        ttk.Button(top, text="Import CIS List", command=self._import_cis_list).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Analyze CIS", command=self._analyze_bulk).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(top, text="Generate Bulk", command=self._generate_bulk).pack(side=tk.RIGHT, padx=(4, 0))

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
        for key in ["ProcedureCode", "Category", "DisplayName", "Version", "DefaultOutputName", "Description", "Remarks"]:
            row = ttk.Frame(details, padding=(6, 4))
            row.pack(fill=tk.X)
            ttk.Label(row, text=key, width=LABEL_W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self.builder_vars[key] = var
            if key == "Category":
                widget = ttk.Combobox(row, textvariable=var, values=catalog.PROCEDURE_CATEGORIES)
            else:
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

        self.lst_sources = tk.Listbox(left, height=20, font=("Courier", 9))
        self.lst_sources.pack(fill=tk.BOTH, expand=True)
        self.lst_sources.bind("<<ListboxSelect>>", lambda _: self._pick_source_form())
        src_btns = ttk.Frame(left)
        src_btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(src_btns, text="Auto Detect Folders", command=self._auto_detect_source_folders).pack(side=tk.LEFT)
        ttk.Button(src_btns, text="New Source Form", command=self._new_source_form).pack(side=tk.RIGHT)

        self.source_vars: dict[str, tk.StringVar] = {}
        for key in ["SourceFormCode", "Category", "DisplayName", "Version", "PDFFilePath", "MappingKey", "EffectiveDate", "ExpiryDate", "Remarks"]:
            row = ttk.Frame(right, padding=(6, 4))
            row.pack(fill=tk.X)
            ttk.Label(row, text=key, width=LABEL_W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self.source_vars[key] = var
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            if key == "PDFFilePath":
                ttk.Button(row, text="Browse", command=self._browse_source_pdf).pack(side=tk.LEFT, padx=(4, 0))
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
        categories = sorted({catalog.get_value(p, "Category") for p in self.procedures if catalog.is_active(p)})
        if not categories:
            categories = catalog.PROCEDURE_CATEGORIES
        for cmb, var in [
            (self.cmb_gen_category, self.var_gen_category),
            (self.cmb_bulk_category, self.var_bulk_category),
        ]:
            cmb["values"] = categories
            if not var.get() and categories:
                var.set(categories[0])

        self._refresh_generate_procedures()
        self._refresh_bulk_procedures()
        self._refresh_search_results()
        self._refresh_builder_procedures()
        self._refresh_source_forms()
        self._refresh_paths()
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

    def _procedure_labels(self, category: str | None = None) -> list[str]:
        rows = [p for p in self.procedures if catalog.is_active(p)]
        if category:
            rows = [p for p in rows if catalog.get_value(p, "Category") == category]
        return [catalog.procedure_label(p) for p in rows]

    def _procedure_from_label(self, label: str) -> dict | None:
        code = str(label).split(" - ", 1)[0].strip()
        return self.procedure_by_code.get(code)

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
        labels = self._procedure_labels(self.var_gen_category.get())
        self.cmb_gen_proc["values"] = labels
        if labels and self.var_gen_proc.get() not in labels:
            self.var_gen_proc.set(labels[0])
        self._on_generate_proc_change()

    def _refresh_bulk_procedures(self):
        labels = self._procedure_labels(self.var_bulk_category.get())
        self.cmb_bulk_proc["values"] = labels
        if labels and self.var_bulk_proc.get() not in labels:
            self.var_bulk_proc.set(labels[0])

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

    def _add_category_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Add Procedure Category")
        dlg.geometry("420x130")
        dlg.grab_set()
        var = tk.StringVar()
        row = ttk.Frame(dlg, padding=12)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Category", width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        def save():
            name = var.get().strip()
            if not name:
                return
            values = sorted(set(list(self.cmb_gen_category["values"]) + [name]))
            self.cmb_gen_category["values"] = values
            self.cmb_bulk_category["values"] = values
            self.builder_vars["Category"].set(name)
            self.var_gen_category.set(name)
            self.var_bulk_category.set(name)
            dlg.destroy()
            self._new_procedure_from_generate(category=name)

        btns = ttk.Frame(dlg, padding=(12, 0, 12, 12))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="Create Procedure In This Category", command=save).pack(side=tk.LEFT)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT, padx=(6, 0))

    def _new_procedure_from_generate(self, category: str | None = None):
        self.nb.select(self.tab_builder)
        self._new_procedure()
        self.builder_vars["Category"].set(category or self.var_gen_category.get() or catalog.PROCEDURE_CATEGORIES[0])

    def _refresh_generate_structure(self):
        self.lst_gen_structure.delete(0, tk.END)
        proc = self._procedure_from_label(self.var_gen_proc.get())
        if not proc:
            return
        if str(catalog.get_value(proc, "AutoBlankAfterOdd", True)).strip().lower() not in {"false", "0", "no"}:
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
            is_default = key in self.default_fields.get(sheet, set()) or key in self.default_fields.get(excel_reader.normalize_header(sheet), set())
            show_required_missing = field["required"] and selected.get(key) in ("", None)
            show_procedure_specific = key not in package_engine.COMMON_DATA_FIELDS
            if is_default and selected.get(key) not in ("", None) and key not in seen:
                seen.add(key)
                default_rows.append(field)
            elif (show_required_missing or show_procedure_specific) and key not in seen:
                seen.add(key)
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
            ttk.Label(row, text=str(selected.get(field["key"], "") or ""), relief="sunken", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(row, text="Edit", command=lambda f=field: self._edit_default_field(f)).pack(side=tk.LEFT, padx=(4, 0))

        if editable_rows:
            ttk.Label(self.frm_manual, text="Non-default Information", font=("", 9, "bold")).pack(anchor="w", padx=4, pady=(8, 2))
        for field in editable_rows:
            row = ttk.Frame(self.frm_manual)
            row.pack(fill=tk.X, padx=4, pady=3)
            label = field["label"] + (" *" if field["required"] else "")
            ttk.Label(row, text=label, width=24).pack(side=tk.LEFT)
            var = tk.StringVar(value=str(selected.get(field["key"], "") or ""))
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.manual_entries[field["key"]] = var

    def _edit_default_field(self, field: dict):
        messagebox.showinfo(
            "Default field edit",
            "Default-field write-back to clients.xlsx will be enabled with backup, confirmation, and change log.\n\n"
            "For now, edit this default value in clients.xlsx and click Reload."
        )

    # Customer search -----------------------------------------------------

    def _refresh_search_results(self):
        if not hasattr(self, "lst_customers"):
            return
        self.search_results = excel_reader.search_customers(self.customers, self.var_search.get(), limit=80)
        self.lst_customers.delete(0, tk.END)
        for c in self.search_results:
            self.lst_customers.insert(tk.END, self._customer_label(c))

    def _customer_label(self, c: dict) -> str:
        name = str(c.get("name") or c.get("client_name") or "-")[:30]
        cis = str(c.get("cis") or c.get("cif_no") or "-")[:14]
        ic = str(c.get("ic_number") or "-")[:14]
        pol = str(c.get("policy_number") or "-")[:14]
        return f"{name:<30}  CIS:{cis:<14}  IC:{ic:<14}  Policy:{pol}"

    def _pick_customer(self):
        sel = self.lst_customers.curselection()
        if not sel:
            return
        self.selected_customer = dict(self.search_results[sel[0]])
        self.lbl_customer.config(text=f"Selected: {self.selected_customer.get('name') or self.selected_customer.get('client_name') or '-'}")
        self._refresh_account_choices()
        self._refresh_manual_fields()

    def _account_type_for_procedure(self, proc: dict | None) -> str | None:
        if not proc:
            return None
        text = " ".join([
            str(catalog.get_value(proc, "ProcedureCode", "")),
            str(catalog.get_value(proc, "DisplayName", "")),
            str(catalog.get_value(proc, "Description", "")),
        ]).lower()
        if "bond" in text:
            return "BOND"
        if "ut" in text or "unit trust" in text:
            return "UT"
        if "si" in text or "structured" in text:
            return "SI"
        return None

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
            self.var_account.set(labels[0])
            self.selected_account = dict(self.account_rows[0])

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

        manual = {k: v.get().strip() for k, v in self.manual_entries.items()}
        client = self._client_with_selected_account()
        client.update({k: v for k, v in manual.items() if v})
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
            var.set(str(catalog.get_value(proc, key, "")))
        self.var_builder_active.set(catalog.is_active(proc))
        self.var_builder_auto_blank.set(str(catalog.get_value(proc, "AutoBlankAfterOdd", True)).strip().lower() not in {"false", "0", "no"})
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
        self.builder_vars["Category"].set(catalog.PROCEDURE_CATEGORIES[0])
        self.builder_vars["Version"].set("V01")
        self.var_builder_active.set(True)
        self.var_builder_auto_blank.set(True)
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
        self._load_all()
        self._set_status(f"Saved procedure {code}.")

    def _builder_add_source(self):
        sel = self.lst_available_sources.curselection()
        if not sel:
            return
        source = self._source_from_label(self.lst_available_sources.get(sel[0]))
        if not source:
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
        folders = config_loader.scan_forms_folder(self.settings)
        existing = {catalog.get_value(s, "SourceFormCode") for s in self.source_forms}
        added = 0
        skipped = 0
        for folder in folders:
            code = catalog.source_code_from_folder(folder)
            if not code or code in existing:
                skipped += 1
                continue
            folder_path = config_loader.forms_folder_path(self.settings) / folder
            pdfs = [p for p in folder_path.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
            if len(pdfs) != 1:
                skipped += 1
                continue
            display = folder.replace(code, "", 1).strip(" -_") or folder
            self.source_forms.append({
                "SourceFormCode": code,
                "Category": "",
                "DisplayName": display,
                "Version": "V01",
                "PDFFilePath": folder,
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

    def _browse_source_pdf(self):
        path = filedialog.askopenfilename(title="Select source PDF", filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
        if not path:
            return
        p = Path(path)
        base = config_loader.forms_folder_path(self.settings)
        try:
            rel = p.relative_to(base)
            self.source_vars["PDFFilePath"].set(str(rel))
        except ValueError:
            self.source_vars["PDFFilePath"].set(str(p))

    def _save_source_form(self):
        code = self.source_vars["SourceFormCode"].get().strip()
        if not code:
            messagebox.showerror("Missing", "SourceFormCode is required.")
            return
        row = {k: v.get().strip() for k, v in self.source_vars.items()}
        row["Active"] = self.var_source_active.get()
        if not row.get("MappingKey"):
            row["MappingKey"] = code
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
            self._open_coord_picker()
            return
        pdf_path = catalog.resolve_source_pdf_path(src, self.settings)
        args = [
            "--form-id", catalog.mapping_key(src),
            "--display-name", catalog.get_value(src, "DisplayName", code),
        ]
        if pdf_path.exists():
            args.extend(["--pdf", str(pdf_path)])
        self._open_coord_picker(args)

    def _open_coord_picker(self, args: list[str] | None = None):
        args = args or []
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
