"""
FormFiller — data-driven PDF form filler.

P1-2: Settings dialog with Browse buttons for forms_folder / output_folder.
P1-3: Startup health check — green/red form subfolder status.
P1-4: Blank required fields → _REVIEW_ filename prefix + blocking warning.
P2-5: "Execute All" single button in bulk mode.
P2-6: Auto-open output folder after fill (settings: auto_open_output).
P2-7: Session memory — restores last application + mode on launch.
P2-8: "Open CoordPicker" button launches CoordPicker.exe.
P3-1: Top bar Branch + Date pickers — values flow into form fields via 'session' source.
P3-2: RM profile sourced from clients.xlsx -> RM_Profile sheet, not settings.json.
P3-3: One-click open buttons for clients.xlsx, settings.json, forms.json, config folder.
P3-4: settings.json auto-backup on every launch (config/backups/, last 100 retained).
"""
import os
import sys
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from datetime import date, datetime

sys.path.insert(0, str(Path(__file__).parent))

import client_db
import config_loader
import excel_reader
import pdf_engine
from excel_reader import ExcelLockedError
from config_loader import TemplateChangedWarning, TemplateSurrenderedError

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

_MASTER_COLS = {
    "name", "ic_number", "phone", "email",
    "address_line1", "address_line2", "city",
    "state", "postcode", "dob", "occupation",
}
LABEL_W = 22


class FormFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Form Filler")
        self.geometry("820x620")
        self.resizable(True, True)
        self.minsize(720, 500)

        self.settings     = {}
        self.forms        = {}
        self.applications = []
        self.app_map      = {}
        self.master_data  = {}
        self.rm_profile   = {}
        self.xlsx_path    = None
        self._mode        = tk.StringVar(value="bulk")
        self._bulk_vars   = []
        self._single_entries = {}
        self._first_run_msgs = []

        # Backup settings.json on every launch (P3-4)
        try:
            config_loader.backup_settings_on_launch()
        except Exception:
            pass

        self._build_ui()
        self._load_all()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Row 1 — Application + Branch + Date + Mode
        top = tk.Frame(self, pady=6)
        top.pack(fill=tk.X, padx=12)

        tk.Label(top, text="Application:").pack(side=tk.LEFT)
        self.var_app = tk.StringVar()
        self.cmb_app = ttk.Combobox(top, textvariable=self.var_app,
                                    state="readonly", width=26)
        self.cmb_app.pack(side=tk.LEFT, padx=(4, 8))
        self.cmb_app.bind("<<ComboboxSelected>>", self._on_app_change)

        tk.Label(top, text="Branch:").pack(side=tk.LEFT)
        self.var_branch = tk.StringVar()
        self.cmb_branch = ttk.Combobox(top, textvariable=self.var_branch,
                                       state="readonly", width=8)
        self.cmb_branch.pack(side=tk.LEFT, padx=(4, 8))
        self.cmb_branch.bind("<<ComboboxSelected>>", lambda _: self._save_session())

        tk.Label(top, text="Date:").pack(side=tk.LEFT)
        if HAS_TKCALENDAR:
            self.dat_picker = DateEntry(top, width=11, date_pattern="dd/mm/yyyy",
                                        firstweekday="monday")
            self.dat_picker.pack(side=tk.LEFT, padx=(4, 8))
            self.dat_picker.bind("<<DateEntrySelected>>",
                                 lambda _: self._save_session())
        else:
            # Fallback: plain entry pre-filled with today
            self.var_date = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
            tk.Entry(top, textvariable=self.var_date, width=12
                     ).pack(side=tk.LEFT, padx=(4, 8))

        tk.Label(top, text="Mode:").pack(side=tk.LEFT, padx=(4, 4))
        tk.Radiobutton(top, text="Bulk",   variable=self._mode,
                       value="bulk",   command=self._on_mode_change).pack(side=tk.LEFT)
        tk.Radiobutton(top, text="Single", variable=self._mode,
                       value="single", command=self._on_mode_change).pack(side=tk.LEFT)

        tk.Button(top, text="↻", width=3,
                  command=self._load_all).pack(side=tk.LEFT, padx=(8, 0))

        # Row 2 — Utility buttons
        util = tk.Frame(self, pady=2)
        util.pack(fill=tk.X, padx=12)
        tk.Button(util, text="📊 Open Excel",
                  command=self._open_excel).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(util, text="🗂 CoordPicker",
                  command=self._open_coord_picker).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(util, text="⚙ Settings",
                  command=self._open_settings).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(util, text="📁 Open Config",
                  command=self._open_config_folder).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=(4, 4))

        # Health check banner (hidden until needed)
        self.frm_health = tk.Frame(self, bg="#fff3cd")
        self.lbl_health = tk.Label(self.frm_health, text="", bg="#fff3cd",
                                   anchor="w", justify="left",
                                   font=("", 8), wraplength=780)
        self.lbl_health.pack(fill=tk.X, padx=8, pady=4)

        # Dynamic panel
        self._panel = tk.Frame(self)
        self._panel.pack(fill=tk.BOTH, expand=True, padx=12)

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12, pady=4)

        # Bottom bar
        bot = tk.Frame(self)
        bot.pack(fill=tk.X, padx=12, pady=(0, 4))

        self.btn_exec_all = tk.Button(
            bot, text="▶ Execute All", width=14, height=2,
            bg="#155724", fg="white", activebackground="#0d3d18",
            command=self._execute_all)
        self.btn_exec_all.pack(side=tk.LEFT, padx=(0, 4))

        self.btn_fill = tk.Button(
            bot, text="Fill Selected", width=14, height=2,
            bg="#28a745", fg="white", activebackground="#1e7e34",
            command=self._on_fill)
        self.btn_fill.pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(bot, text="Open Output Folder", width=18, height=2,
                  command=self._open_output).pack(side=tk.LEFT)

        self.lbl_status = tk.Label(self, text="  Loading...",
                                   anchor="w", fg="gray")
        self.lbl_status.pack(fill=tk.X, padx=12, pady=(0, 6))

    # ── Data loading ──────────────────────────────────────────────────────────

    def _load_all(self):
        self._first_run_msgs = []

        try:
            self.settings = config_loader.load_settings()
        except FileNotFoundError:
            self._set_status(
                "settings.json missing. Open ⚙ Settings to create it.", "red")
            return

        # First-run scaffolding: create folders + clone clients template
        try:
            dir_info = config_loader.ensure_runtime_dirs(self.settings)
            xlsx_info = config_loader.ensure_clients_xlsx()
            if dir_info.get("forms_created"):
                self._first_run_msgs.append(
                    f"Created forms folder: {config_loader.forms_folder_path(self.settings)}")
            if dir_info.get("output_created"):
                self._first_run_msgs.append(
                    f"Created output folder: {config_loader.output_folder_path(self.settings)}")
            if xlsx_info.get("copied"):
                self._first_run_msgs.append(
                    "Created data/clients.xlsx from clients_template.xlsx — "
                    "open it (📊 Open Excel) and fill in RM_Profile + your clients.")
            if xlsx_info.get("source_missing"):
                self._first_run_msgs.append(
                    "Neither clients.xlsx nor clients_template.xlsx exists in data/. "
                    "Cannot proceed without client data.")
        except Exception as e:
            self._set_status(f"First-run setup error: {e}", "red")
            return

        try:
            self.forms        = config_loader.load_forms()
            self.applications = config_loader.load_applications()
            self.app_map      = {a["name"]: a for a in self.applications}
            self.xlsx_path    = config_loader.data_path("clients.xlsx")

            # Phase 1A: client master data comes from client_db.db.
            # Master sheet is the read-only fallback if DB is empty (legacy mode).
            client_db.init_db()
            self.master_data = self._load_master_data()

            # RM profile still lives in clients.xlsx -> RM_Profile sheet
            if self.xlsx_path.exists():
                self.rm_profile = excel_reader.load_rm_profile(self.xlsx_path)
            else:
                self.rm_profile = {"branches": []}
        except ExcelLockedError as e:
            messagebox.showerror("Excel is open", str(e))
            return
        except Exception as e:
            self._set_status(f"Load error: {e}", "red")
            return

        # Populate dropdowns
        self.cmb_app["values"] = [a["name"] for a in self.applications]
        branches = self.rm_profile.get("branches") or []
        self.cmb_branch["values"] = branches
        if branches and not self.var_branch.get():
            self.cmb_branch.set(branches[0])

        # P2-7: restore session
        state = config_loader.load_state()
        last_app    = state.get("last_app", "")
        last_mode   = state.get("last_mode", "bulk")
        last_branch = state.get("last_branch", "")
        self._mode.set(last_mode)
        names = self.cmb_app["values"]
        if last_app in names:
            self.cmb_app.set(last_app)
        elif names:
            self.cmb_app.current(0)
        if last_branch and last_branch in branches:
            self.cmb_branch.set(last_branch)

        self._on_app_change()
        self._run_health_check()

        n_clients = len(self.master_data)
        rm_name = self.rm_profile.get("rm_name") or "(rm_name not set)"
        self._set_status(
            f"Loaded {n_clients} client(s), {len(self.applications)} application(s). "
            f"RM: {rm_name}", "green" if n_clients else "orange")

    def _load_master_data(self) -> dict:
        """Returns {ic_normalized: client_dict}.
        Prefers client_db.db; falls back to clients.xlsx Master sheet if DB empty.
        """
        # Try DB first
        try:
            rows = client_db.list_all(active_only=True)
        except Exception as e:
            rows = []
            print(f"client_db error (falling back to Master sheet): {e}")
        if rows:
            return {r["ic_number"]: r for r in rows}

        # Fallback: Master sheet (legacy mode — pre-Phase 1A data)
        if self.xlsx_path and self.xlsx_path.exists():
            try:
                return excel_reader.load_master(self.xlsx_path)
            except Exception:
                return {}
        return {}

    def _run_health_check(self):
        """P1-3: Show banner if any form subfolders are missing or first-run."""
        results = config_loader.health_check(self.settings, self.forms)

        # First-run messages take priority
        if self._first_run_msgs:
            self.lbl_health.config(
                text="  First-run setup:\n  • " + "\n  • ".join(self._first_run_msgs))
            self.frm_health.pack(fill=tk.X, padx=12, pady=(0, 4),
                                 before=self._panel)
            return

        errors      = [r for r in results if r["status"] == "error"]
        warns       = [r for r in results if r["status"] == "warn"]
        surrendered = [r for r in results if r["status"] == "surrendered"]
        unmapped    = [r for r in results if r["status"] == "unmapped"]
        empty       = [r for r in results if r["status"] == "empty"]
        ok          = [r for r in results if r["status"] == "ok"]

        if not results:
            self.frm_health.pack_forget()
            return

        if empty:
            self.lbl_health.config(text=f"  {empty[0]['message']}")
            self.frm_health.pack(fill=tk.X, padx=12, pady=(0, 4),
                                 before=self._panel)
            return

        parts = []
        if ok:          parts.append(f"✓ {len(ok)} ready")
        if surrendered: parts.append(f"🔒 {len(surrendered)} surrendered")
        if unmapped:    parts.append(f"📂 {len(unmapped)} unmapped")
        if warns:       parts.append(f"⚠ {len(warns)} warning(s)")
        if errors:
            names = ", ".join(r["name"] for r in errors)
            parts.append(f"✗ {len(errors)} missing: {names}")

        msg = "  |  ".join(parts)
        if errors or warns:
            self.lbl_health.config(text=f"  Forms health: {msg}")
            self.frm_health.pack(fill=tk.X, padx=12, pady=(0, 4),
                                 before=self._panel)
        else:
            self.frm_health.pack_forget()

    # ── Panel switching ───────────────────────────────────────────────────────

    def _on_app_change(self, _=None):
        self._save_session()
        self._rebuild_panel()

    def _on_mode_change(self):
        self._save_session()
        self._rebuild_panel()

    def _save_session(self):
        config_loader.save_state({
            "last_app":    self.var_app.get(),
            "last_mode":   self._mode.get(),
            "last_branch": self.var_branch.get(),
        })

    def _rebuild_panel(self):
        for w in self._panel.winfo_children():
            w.destroy()
        self._bulk_vars      = []
        self._single_entries = {}

        app = self.app_map.get(self.var_app.get(), {})
        if not app:
            return

        is_bulk = self._mode.get() == "bulk"
        self.btn_exec_all.config(state=tk.NORMAL if is_bulk else tk.DISABLED)
        self.btn_fill.config(text="Fill Selected" if is_bulk else "Fill & Save")

        if is_bulk:
            self._build_bulk_panel(app)
        else:
            self._build_single_panel(app)

    # ── Bulk panel ────────────────────────────────────────────────────────────

    def _build_bulk_panel(self, app: dict):
        if not self.xlsx_path or not self.xlsx_path.exists():
            tk.Label(self._panel,
                     text="data/clients.xlsx not found. Click 📊 Open Excel after creating it.",
                     fg="red").pack(anchor="w", pady=8)
            return

        sheet = app.get("data_sheet", "Master")
        try:
            batch = excel_reader.load_batch(
                self.xlsx_path, sheet, self.master_data)
        except ExcelLockedError as e:
            messagebox.showerror("Excel is open", str(e))
            return
        except ValueError as e:
            tk.Label(self._panel, text=str(e), fg="red",
                     wraplength=760, justify="left").pack(anchor="w", pady=8)
            return

        ctrl = tk.Frame(self._panel)
        ctrl.pack(fill=tk.X, pady=(4, 4))
        tk.Label(ctrl,
                 text=f"{len(batch)} client(s) in '{sheet}' sheet",
                 fg="gray").pack(side=tk.LEFT)
        tk.Button(ctrl, text="Select All",
                  command=self._bulk_select_all).pack(side=tk.RIGHT, padx=(4, 0))
        tk.Button(ctrl, text="Clear All",
                  command=self._bulk_clear_all).pack(side=tk.RIGHT)

        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(
                            int(-1*(e.delta/120)), "units"))

        tx_cols = [k for k in (batch[0].keys() if batch else [])
                   if k not in _MASTER_COLS and k != "ic_number"]

        for client in batch:
            var = tk.BooleanVar(value=True)
            parts = [f"{k}: {client[k]}" for k in tx_cols[:4]
                     if client.get(k)]
            label = f"  {client.get('name', '?'):<28}  " + "   |   ".join(parts)
            tk.Checkbutton(inner, text=label, variable=var,
                           anchor="w", font=("Courier", 9)
                           ).pack(fill=tk.X, padx=4, pady=1)
            self._bulk_vars.append((var, client))

    def _bulk_select_all(self):
        for var, _ in self._bulk_vars:
            var.set(True)

    def _bulk_clear_all(self):
        for var, _ in self._bulk_vars:
            var.set(False)

    # ── Single panel ──────────────────────────────────────────────────────────

    def _build_single_panel(self, app: dict):
        sheet = app.get("data_sheet", "Master")
        self._single_sheet = sheet
        self._selected_ic: str | None = None

        # Search row
        srch = tk.Frame(self._panel)
        srch.pack(fill=tk.X, pady=(4, 2))
        tk.Label(srch, text="Search:", width=8, anchor="w").pack(side=tk.LEFT)
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self._on_search_changed())
        ent = tk.Entry(srch, textvariable=self.var_search, width=32)
        ent.pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(srch, text="(full IC or partial name)",
                 fg="gray", font=("", 8)).pack(side=tk.LEFT)

        tk.Button(srch, text="➕ Add",
                  command=self._open_add_client).pack(side=tk.RIGHT, padx=(2, 0))
        tk.Button(srch, text="🗑 Delete",
                  command=self._delete_selected_client
                  ).pack(side=tk.RIGHT, padx=(2, 0))
        tk.Button(srch, text="✏️ Edit",
                  command=self._open_edit_client).pack(side=tk.RIGHT, padx=(0, 0))

        # Results listbox (limited height)
        list_frm = tk.Frame(self._panel)
        list_frm.pack(fill=tk.X, pady=(2, 4))
        self.lst_clients = tk.Listbox(list_frm, height=5, font=("Courier", 9),
                                       activestyle="dotbox")
        self.lst_clients.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb = ttk.Scrollbar(list_frm, orient="vertical",
                           command=self.lst_clients.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lst_clients.config(yscrollcommand=sb.set)
        self.lst_clients.bind("<<ListboxSelect>>",
                               lambda _: self._on_client_pick())
        self._render_search_results(client_db.list_all(active_only=True))

        # Selected client banner
        self.lbl_selected = tk.Label(self._panel,
                                      text="No client selected.",
                                      anchor="w", fg="gray",
                                      font=("", 9, "bold"))
        self.lbl_selected.pack(fill=tk.X, padx=4, pady=(0, 2))

        ttk.Separator(self._panel, orient="horizontal").pack(
            fill=tk.X, pady=(2, 4))

        try:
            headers = excel_reader.get_sheet_headers(self.xlsx_path, sheet)
        except ExcelLockedError:
            headers = []

        outer = tk.Frame(self._panel)
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(
                       scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        if not headers:
            tk.Label(inner,
                     text=f"No transaction columns in '{sheet}' sheet.\n"
                          "All data from Master sheet.",
                     fg="gray", justify="left").pack(anchor="w", pady=8, padx=4)
        else:
            tk.Label(inner,
                     text=f"{len(headers)} transaction field(s) from '{sheet}'. "
                          "Auto-filled if client already has a row.",
                     fg="gray", font=("", 8),
                     justify="left").pack(anchor="w", padx=4, pady=(2, 6))
            for col in headers:
                r = tk.Frame(inner)
                r.pack(fill=tk.X, pady=2, padx=4)
                tk.Label(r, text=col.replace("_", " ").title() + ":",
                         width=LABEL_W, anchor="w").pack(side=tk.LEFT)
                var = tk.StringVar()
                tk.Entry(r, textvariable=var, width=40).pack(side=tk.LEFT)
                self._single_entries[col] = var

        self._autofill_single(sheet)

    def _autofill_single(self, sheet: str):
        ic = self._selected_ic if hasattr(self, "_selected_ic") else None
        if not ic:
            return
        if not self.xlsx_path or not self.xlsx_path.exists():
            return
        try:
            existing = excel_reader.find_client_in_batch(
                self.xlsx_path, sheet, ic)
        except ExcelLockedError:
            return
        for col, var in self._single_entries.items():
            if existing.get(col) not in ("", None):
                var.set(str(existing[col]))

    def _ic_for_name(self, name: str) -> str:
        """Legacy helper used by some callers; prefer _selected_ic."""
        for ic, rec in self.master_data.items():
            if rec.get("name") == name:
                return ic
        return ""

    # ── Phase 1A: client search + CRUD helpers ────────────────────────────────

    def _on_search_changed(self):
        q = self.var_search.get().strip()
        if not q:
            results = client_db.list_all(active_only=True, limit=200)
        elif q.replace("-", "").replace(" ", "").isdigit():
            # Looks like an IC — full match
            hit = client_db.by_ic(q)
            results = [hit] if hit else []
        else:
            results = client_db.by_name(q, limit=50)
        self._render_search_results(results)

    def _render_search_results(self, results: list[dict]):
        if not hasattr(self, "lst_clients"):
            return
        self.lst_clients.delete(0, tk.END)
        self._search_results = results
        for c in results:
            star = " ⭐" if c.get("permanent") else ""
            label = f"  {c['name']:<32s}  {c['ic_number']}{star}"
            self.lst_clients.insert(tk.END, label)

    def _on_client_pick(self):
        sel = self.lst_clients.curselection()
        if not sel:
            return
        client = self._search_results[sel[0]]
        self._selected_ic = client["ic_number"]
        star = " ⭐ permanent" if client.get("permanent") else ""
        self.lbl_selected.config(
            text=f"Selected: {client['name']}  ({client['ic_number']}){star}",
            fg="#155724")
        self._autofill_single(getattr(self, "_single_sheet", "Master"))

    def _open_add_client(self):
        dlg = ClientFormDialog(self, mode="add")
        self.wait_window(dlg)
        if dlg.result:
            try:
                client_db.add(dlg.result)
                self._after_client_change(f"Added {dlg.result['name']}.")
            except ValueError as e:
                messagebox.showerror("Cannot add", str(e))

    def _open_edit_client(self):
        if not self._selected_ic:
            messagebox.showwarning("Pick a client",
                                   "Tick a client in the list first.")
            return
        client = client_db.by_ic(self._selected_ic, include_inactive=True)
        if not client:
            return
        dlg = ClientFormDialog(self, mode="edit", client=client)
        self.wait_window(dlg)
        if dlg.result:
            try:
                client_db.update(self._selected_ic, dlg.result)
                self._after_client_change(f"Updated {dlg.result.get('name', '')}.")
            except ValueError as e:
                messagebox.showerror("Cannot update", str(e))

    def _delete_selected_client(self):
        if not self._selected_ic:
            messagebox.showwarning("Pick a client",
                                   "Tick a client in the list first.")
            return
        client = client_db.by_ic(self._selected_ic, include_inactive=True)
        if not client:
            return
        # Two-step: soft delete by default, hard delete if user picks "永久删除"
        choice = messagebox.askyesnocancel(
            "Delete client",
            f"Delete '{client['name']}' ({client['ic_number']})?\n\n"
            "Yes = soft delete (hidden from search, can restore later)\n"
            "No  = permanently delete (cannot undo)\n"
            "Cancel = abort",
            icon="warning")
        if choice is None:
            return
        if choice:
            client_db.soft_delete(self._selected_ic)
            self._after_client_change(f"Soft-deleted {client['name']}.")
        else:
            confirm = simpledialog.askstring(
                "Hard delete confirm",
                f"This is permanent. Type the client's name to confirm:\n\n{client['name']}",
                parent=self)
            if confirm and confirm.strip().upper() == client["name"].strip().upper():
                client_db.hard_delete(self._selected_ic)
                self._after_client_change(f"Hard-deleted {client['name']}.")
            else:
                messagebox.showinfo("Cancelled",
                                    "Name didn't match — nothing deleted.")

    def _after_client_change(self, status_msg: str):
        """Refresh master_data + search list after Add/Edit/Delete."""
        self.master_data = self._load_master_data()
        self._selected_ic = None
        self.lbl_selected.config(text="No client selected.", fg="gray")
        self._on_search_changed()
        self._set_status(status_msg, "green")

    # ── Session context ───────────────────────────────────────────────────────

    def _current_session_context(self) -> dict:
        if HAS_TKCALENDAR and hasattr(self, "dat_picker"):
            d = self.dat_picker.get_date()
            date_str = d.strftime("%d/%m/%Y")
        else:
            raw = self.var_date.get().strip() if hasattr(self, "var_date") \
                else date.today().strftime("%d/%m/%Y")
            # Validate dd/mm/yyyy or normalize from common alternates;
            # if user typed garbage, fall back to today rather than corrupt forms.
            date_str = self._normalize_date(raw)
        return {
            "rm_branch": self.var_branch.get(),
            "date":      date_str,
        }

    @staticmethod
    def _normalize_date(raw: str) -> str:
        """Accept dd/mm/yyyy, d/m/yyyy, dd-mm-yyyy, yyyy-mm-dd. Output dd/mm/yyyy.
        Bad input → raises ValueError so caller surfaces it."""
        from datetime import datetime as _dt
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return _dt.strptime(raw, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        raise ValueError(
            f"Date '{raw}' is not in dd/mm/yyyy format. "
            "Use the calendar picker or type like 05/05/2026.")

    # ── Fill actions ──────────────────────────────────────────────────────────

    def _execute_all(self):
        """P2-5: Select all + fill in one click."""
        self._bulk_select_all()
        self._fill_bulk(self.app_map.get(self.var_app.get(), {}))

    def _on_fill(self):
        app = self.app_map.get(self.var_app.get(), {})
        if not app:
            return
        if self._mode.get() == "bulk":
            self._fill_bulk(app)
        else:
            self._fill_single(app)

    def _fill_bulk(self, app: dict):
        selected = [c for var, c in self._bulk_vars if var.get()]
        if not selected:
            messagebox.showwarning("Nothing selected",
                                   "Tick at least one client.")
            return
        self._run_fill(selected, app, label=f"{len(selected)} client(s)")

    def _fill_single(self, app: dict):
        ic = getattr(self, "_selected_ic", None)
        if not ic or ic not in self.master_data:
            messagebox.showwarning("No client",
                                   "Tick a client in the search list first.")
            return
        client = dict(self.master_data[ic])
        for col, var in self._single_entries.items():
            v = var.get().strip()
            if v:
                client[col] = v
        self._run_fill([client], app, label=client.get("name", ""))

    def _run_fill(self, clients: list[dict], app: dict, label: str = ""):
        try:
            session_ctx = self._current_session_context()
        except ValueError as e:
            messagebox.showerror("Bad date", str(e))
            return
        for c in clients:
            c.setdefault("date", session_ctx["date"])

        output_folder = config_loader.output_folder_path(self.settings)
        log = config_loader.log_path()

        n = len(clients)
        saved, all_warnings, errors = [], [], []

        for i, client in enumerate(clients, 1):
            self._set_status(
                f"Filling {i}/{n}: {client.get('name', '?')}...", "blue")
            self.update_idletasks()
            try:
                results, warnings = pdf_engine.fill_bundle(
                    application=app,
                    forms_config=self.forms,
                    client=client,
                    output_folder=output_folder,
                    settings=self.settings,
                    find_template_fn=config_loader.find_template,
                    log_path=log,
                    rm_profile=self.rm_profile,
                    session=session_ctx,
                )
                saved.extend(results)
                all_warnings.extend(warnings)
            except TemplateSurrenderedError as e:
                if messagebox.askyesno(
                    "Form Surrendered — Test Fill?",
                    f"{e}\n\nRun TEST FILL with old coordinates?\n"
                    "(Output will be prefixed _TEST_ — do not submit)",
                    icon="warning"
                ):
                    def _test_finder(settings, subfolder, form_cfg=None):
                        return config_loader.find_template(
                            settings, subfolder, form_cfg, test_mode=True)
                    try:
                        results, warnings = pdf_engine.fill_bundle(
                            application=app,
                            forms_config=self.forms,
                            client=client,
                            output_folder=output_folder,
                            settings=self.settings,
                            find_template_fn=_test_finder,
                            log_path=log,
                            test_mode=True,
                            rm_profile=self.rm_profile,
                            session=session_ctx,
                        )
                        saved.extend(results)
                        all_warnings.extend(warnings)
                    except Exception as e2:
                        errors.append(f"{client.get('name', '?')}: {e2}")
                else:
                    errors.append(
                        f"{client.get('name', '?')}: skipped — form surrendered. "
                        "Re-map in CoordPicker.")

            except TemplateChangedWarning as e:
                if messagebox.askyesno(
                    "Form content changed",
                    f"{e}\n\nFill anyway with current coordinates?",
                    icon="warning"
                ):
                    def _no_hash(settings, subfolder, form_cfg=None):
                        return config_loader.find_template(
                            settings, subfolder, None)
                    try:
                        results, warnings = pdf_engine.fill_bundle(
                            application=app,
                            forms_config=self.forms,
                            client=client,
                            output_folder=output_folder,
                            settings=self.settings,
                            find_template_fn=_no_hash,
                            log_path=log,
                            rm_profile=self.rm_profile,
                            session=session_ctx,
                        )
                        saved.extend(results)
                        all_warnings.extend(warnings)
                    except Exception as e2:
                        errors.append(f"{client.get('name', '?')}: {e2}")
                else:
                    errors.append(f"{client.get('name', '?')}: skipped (content changed)")
            except Exception as e:
                errors.append(f"{client.get('name', '?')}: {e}")

        if all_warnings:
            messagebox.showwarning(
                "Required fields blank",
                f"{len(all_warnings)} form(s) had blank required fields.\n"
                f"Files prefixed with _REVIEW_ need checking before submission.\n\n"
                + "\n".join(all_warnings)
            )

        if errors:
            messagebox.showerror("Errors", "\n".join(errors))

        parts = [f"Saved {len(saved)} PDF(s)."]
        if errors:
            parts.append(f"{len(errors)} error(s).")
        if all_warnings:
            parts.append("Check _REVIEW_ files.")
        self._set_status("  ".join(parts),
                         "green" if not errors and not all_warnings else "orange")

        if saved and self.settings.get("auto_open_output", True):
            self._open_output()

    # ── Settings dialog ───────────────────────────────────────────────────────

    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.geometry("560x320")
        dlg.resizable(False, False)
        dlg.grab_set()

        s = self.settings
        vars_ = {}

        def row(parent, label, key, default="", browse=None):
            frm = tk.Frame(parent)
            frm.pack(fill=tk.X, padx=12, pady=4)
            tk.Label(frm, text=label, width=18, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar(value=s.get(key, default))
            vars_[key] = var
            entry = tk.Entry(frm, textvariable=var, width=36)
            entry.pack(side=tk.LEFT)
            if browse:
                tk.Button(frm, text="Browse…",
                          command=lambda v=var: v.set(
                              filedialog.askdirectory() or v.get())
                          ).pack(side=tk.LEFT, padx=4)
            return var

        tk.Label(dlg, text="Paths (relative = next to FormFiller.exe)",
                 font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        row(dlg, "Forms folder:", "forms_folder", "forms", browse=True)
        row(dlg, "Output folder:", "output_folder", "filled", browse=True)

        tk.Label(dlg, text="Options", font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        frm_opt = tk.Frame(dlg)
        frm_opt.pack(fill=tk.X, padx=12)
        auto_open = tk.BooleanVar(value=s.get("auto_open_output", True))
        tk.Checkbutton(frm_opt, text="Auto-open output folder after fill",
                       variable=auto_open).pack(anchor="w")

        tk.Label(dlg, text="Edit config files",
                 font=("", 10, "bold"),
                 anchor="w").pack(fill=tk.X, padx=12, pady=(10, 2))
        edit_frm = tk.Frame(dlg)
        edit_frm.pack(fill=tk.X, padx=12)
        tk.Button(edit_frm, text="Open settings.json",
                  command=lambda: self._open_path(
                      config_loader._config_path("settings.json"))
                  ).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(edit_frm, text="Open forms.json",
                  command=lambda: self._open_path(
                      config_loader._config_path("forms.json"))
                  ).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(edit_frm, text="Open backups",
                  command=lambda: self._open_path(
                      config_loader._config_path("backups"))
                  ).pack(side=tk.LEFT, padx=(0, 4))

        tk.Label(dlg, text="Tip: RM name, staff_id, FIMM, IPPC, branches "
                          "live in data/clients.xlsx → RM_Profile sheet.",
                 fg="gray", wraplength=520, justify="left", anchor="w"
                 ).pack(fill=tk.X, padx=12, pady=(8, 0))

        def _save():
            new_settings = dict(s)
            for k, v in vars_.items():
                new_settings[k] = v.get().strip()
            new_settings["auto_open_output"] = auto_open.get()
            config_loader.save_settings(new_settings)
            self.settings = new_settings
            self._load_all()
            dlg.destroy()
            self._set_status("Settings saved.", "green")

        btn_frm = tk.Frame(dlg)
        btn_frm.pack(fill=tk.X, padx=12, pady=12)
        tk.Button(btn_frm, text="Save", width=10, bg="#28a745",
                  fg="white", command=_save).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_frm, text="Cancel", width=10,
                  command=dlg.destroy).pack(side=tk.LEFT)

    # ── CoordPicker launcher ──────────────────────────────────────────────────

    def _open_coord_picker(self):
        if getattr(sys, "frozen", False):
            exe = Path(sys.executable).parent / "CoordPicker.exe"
        else:
            exe = Path(__file__).parent.parent / "dist" / "CoordPicker.exe"

        if exe.exists():
            try:
                self._open_path(exe)
            except Exception as e:
                messagebox.showerror("Cannot open", str(e))
        else:
            messagebox.showinfo(
                "CoordPicker not found",
                f"CoordPicker.exe not found at:\n{exe}\n\n"
                "Make sure CoordPicker.exe is in the same folder as FormFiller.exe."
            )

    # ── One-click open helpers (P3-3) ─────────────────────────────────────────

    def _open_excel(self):
        target = config_loader.data_path("clients.xlsx")
        if not target.exists():
            messagebox.showinfo(
                "Not found",
                f"clients.xlsx not in {target.parent}\n"
                "Will be auto-created from clients_template.xlsx on next launch.")
            return
        self._open_path(target)

    def _open_config_folder(self):
        self._open_path(config_loader._config_path(""))

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

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _open_output(self):
        self._open_path(config_loader.output_folder_path(self.settings))

    def _set_status(self, msg: str, color: str = "gray"):
        self.lbl_status.config(text=f"  {msg}", fg=color)


class ClientFormDialog(tk.Toplevel):
    """Add or Edit a client. mode = 'add' | 'edit'."""

    FIELDS = [
        ("ic_number",      "IC number",      True),   # (key, label, required)
        ("name",           "Name",           True),
        ("cif_no",         "CIF no",         False),
        ("phone",          "Phone",          False),
        ("email",          "Email",          False),
        ("address_line1",  "Address line 1", False),
        ("address_line2",  "Address line 2", False),
        ("city",           "City",           False),
        ("state",          "State",          False),
        ("postcode",       "Postcode",       False),
        ("dob",            "DOB (dd/mm/yyyy)", False),
        ("occupation",     "Occupation",     False),
        ("notes",          "Notes",          False),
    ]

    def __init__(self, parent, mode: str = "add", client: dict | None = None):
        super().__init__(parent)
        self.title("Add Client" if mode == "add" else "Edit Client")
        self.geometry("440x540")
        self.resizable(False, False)
        self.grab_set()
        self.result: dict | None = None
        self._mode = mode
        self._original = client or {}

        self._vars: dict[str, tk.StringVar] = {}
        for key, label, required in self.FIELDS:
            row = tk.Frame(self)
            row.pack(fill=tk.X, padx=12, pady=2)
            tag = label + (" *" if required else "")
            tk.Label(row, text=tag, width=18, anchor="w"
                     ).pack(side=tk.LEFT)
            v = tk.StringVar(value=str(self._original.get(key) or ""))
            self._vars[key] = v
            entry = tk.Entry(row, textvariable=v, width=28)
            entry.pack(side=tk.LEFT)
            if mode == "edit" and key == "ic_number":
                entry.config(state="disabled")  # IC is the PK; can't change

        # Permanent flag
        flag_row = tk.Frame(self)
        flag_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        self.var_permanent = tk.BooleanVar(
            value=bool(self._original.get("permanent")))
        tk.Checkbutton(flag_row,
                       text="⭐ Mark Permanent (survives future monthly imports)",
                       variable=self.var_permanent).pack(anchor="w")

        # Action buttons
        btn_row = tk.Frame(self); btn_row.pack(fill=tk.X, padx=12, pady=12)
        tk.Button(btn_row, text="Save", width=10, bg="#28a745", fg="white",
                  command=self._save).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Cancel", width=10,
                  command=self.destroy).pack(side=tk.LEFT)

    def _save(self):
        data: dict = {}
        for key, _, required in self.FIELDS:
            val = self._vars[key].get().strip()
            if required and not val:
                messagebox.showerror("Missing field",
                                     f"{key.replace('_', ' ').title()} is required.")
                return
            data[key] = val
        data["permanent"] = self.var_permanent.get()
        # In edit mode we don't need to send ic_number through update()
        if self._mode == "edit":
            data.pop("ic_number", None)
        self.result = data
        self.destroy()


if __name__ == "__main__":
    FormFillerApp().mainloop()
