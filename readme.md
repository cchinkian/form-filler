# Form Filler

Portable Windows green app (pen drive, no admin) that auto-fills flat/scanned PDF bank forms from client data using coordinate-based text overlay.

## Project Status
**Sales Flow build (2026-05-19)** — daily workflow is now:

`Search customer → choose form → fill missing fields → generate PDF`

Excel-first/bulk concepts are deferred out of the main UI. See `DEFERRED_FEATURES.md`.

## Key Files
```
form_filler/
├── src/
│   ├── main_app.py          # FormFiller GUI — sales flow, settings dialog, client search + CRUD
│   ├── client_db.py         # SQLite client master (Phase 1A) — CRUD, search, lifecycle flags
│   ├── pdf_engine.py        # reportlab overlay + pypdf merge, 6 source types (data/rm_profile/session/settings/fixed/auto)
│   ├── excel_reader.py      # Multi-sheet Excel, IC normalization, RM_Profile loader, native type preservation
│   ├── config_loader.py     # Path resolver (relative→sibling-of-EXE), health check, settings backup, single-PDF folder lookup
│   └── coord_picker.py      # Visual PDF field mapper — click → records x/y → forms.json (preserves _shared_fields)
├── config/
│   ├── forms.json           # Field coordinate config + _shared_fields library (rm_name/staff_id/fimm_id/ippc_id/rm_branch/date)
│   ├── applications.json    # Form bundles (6 defined: UT sub/redeem/switch, FD, account opening)
│   ├── settings.json        # Paths + UI prefs (gitignored — RM identity moved to Excel RM_Profile sheet in Phase 1A)
│   ├── settings.example.json
│   └── backups/             # Auto-backup of settings.json on every launch (last 100 retained, gitignored)
├── data/
│   ├── client_db.db           # NEW Phase 1A: SQLite client master (gitignored, PDPA)
│   ├── clients.xlsx           # Master sheet (legacy fallback) + RM_Profile + per-product batch sheets (gitignored)
│   ├── clients_template.xlsx  # Template — RM_Profile + 5 batch sheets (committed, no PII)
│   └── fill_log.csv           # Audit trail (auto-created, gitignored)
├── green_test.py            # Environment verification app (4 tests incl. PDF autofill)
├── tests/sales_flow_smoke.py # Core sales-flow PDF generation smoke test
├── zzz_migrate_master_to_db.py  # One-shot Master sheet → client_db.db (gitignored, if present)
├── PLAN.md                  # Architecture optimization decisions
├── HANDOFF.md               # Session handoff
├── DEFERRED_FEATURES.md     # Old/complex concepts intentionally not in main UI
├── lesson.md                # Lessons learned across sessions
└── .github/workflows/build.yml  # CI: builds all 3 EXEs + bundles config/data into FormFiller-portable
```

## Architecture

### Pen Drive Layout (deployed)
Paths are now relative to the EXE — pen drive plugs into any PC, no edits.
```
[pen drive]/
├── FormFiller.exe       # Main daily-use app
├── CoordPicker.exe      # Setup tool — map form fields once per form
├── GreenTest.exe        # One-time environment check
├── forms/               # ← drop blank PDF subfolders here (sibling of EXE)
│   ├── kyc_form/
│   │   └── kyc_form.pdf
│   └── fd_renewal/
│       └── fd_renewal.pdf
├── filled/              # ← output PDFs land here
├── config/
│   ├── settings.json    # paths + UI prefs (RM identity is in clients.xlsx → RM_Profile sheet)
│   ├── forms.json       # field coordinates — populated by CoordPicker
│   ├── applications.json
│   └── backups/         # auto-backup of settings.json on every launch
└── data/
    ├── client_db.db     # Phase 1A: SQLite client master
    └── clients.xlsx     # RM_Profile sheet + transaction batch sheets (UT_Subscription, FD_Renewal, etc.)
```

## Data Model

| Layer | Where | Contains |
|---|---|---|
| **Client master** | `data/client_db.db` | Name, IC, passport, CIF, phone, email, address, dob, occupation, permanent flag, active flag |
| **Legacy fallback** | `clients.xlsx` Master sheet | Same columns — read only if DB empty |
| **RM identity** | `clients.xlsx` → RM_Profile sheet | rm_name, staff_id, fimm_id, ippc_id, branches CSV |
| **Transaction** | `clients.xlsx` batch sheets (FD_Renewal, UT_Subscription, etc.) | Deal-specific data, joined by IC |
| **Per-fill session** | Top bar pickers | Branch (dropdown from RM_Profile), Date (dd/mm/yyyy calendar) |
| **Form field defs** | `config/forms.json` | Coordinates + source bindings |

## Field Source Types (forms.json)

| source | Value comes from |
|---|---|
| `data` | Excel column (master via DB, or batch sheet) |
| `rm_profile` | `clients.xlsx` RM_Profile sheet (rm_name, staff_id, fimm_id, ippc_id) |
| `session` | Top bar pickers (rm_branch, date) — chosen at fill time |
| `settings` | `settings.json` — legacy, kept for backward compat |
| `fixed` | Hardcoded in forms.json (currency, bank name) |
| `auto` | Generated at runtime (date, year, month) — date honours session override |

## Shared Fields
`_shared_fields` in forms.json defines common fields once (client_name, ic_number, date, rm_name, etc.). Per-form entries only need coordinates — no source/format repeat.

## PDF Folder Logic
- Each form folder should contain exactly one PDF at the top level.
- The PDF filename is not used for routing. Rename the merged PDF freely; if it is the only PDF in that folder, FormFiller uses it.
- Keep old PDFs inside a subfolder such as `old forms/` so the top level stays unambiguous.
- Coordinates still belong to that folder/form mapping, so replacing the PDF with a different layout may require re-mapping in CoordPicker.

## How to Build EXEs
```bash
# Push to main → GitHub Actions builds automatically
git push

# Download: Actions tab → latest run → FormFiller-portable artifact
# Contains: FormFiller.exe + CoordPicker.exe + GreenTest.exe + config/ + data/
```

## Running Tests (Mac development)
```bash
cd ~/Documents/tools/form_filler
python3 -m py_compile src/client_db.py src/main_app.py src/config_loader.py src/pdf_engine.py src/coord_picker.py src/excel_reader.py tests/sales_flow_smoke.py
python3 tests/sales_flow_smoke.py
```

## Phase 1A — Client master DB (SQLite)

Replaces the Excel Master sheet as canonical client demographics store.

**One-shot migration** (run once after pulling Phase 1A):
```bash
python3 zzz_migrate_master_to_db.py --dry-run  # preview
python3 zzz_migrate_master_to_db.py            # actually migrate
```

**Schema** — `data/client_db.db` table `clients`:
- `ic_number` PK (normalized 12-digit), `passport_number`, `name`, `cif_no`
- User-entered sticky fields: phone, email, address_line1/2, city, state, postcode, dob, occupation, notes
- Bank fields (JSON, future Phase 1B): `bank_fields TEXT`
- Lifecycle: `permanent` (1=never auto-removed), `active` (0=soft-deleted), `source`, `last_seen_in_import`

**GUI search** (single mode):
- Type name / IC / passport / CIF into the same search box
- Listbox shows matches with ⭐ next to permanent ones
- ➕ Add / ✏️ Edit / 🗑 Delete buttons (Delete dialog: Yes=soft, No=hard with name-typing confirm, Cancel=abort)

**Phase 1B (deferred)** — monthly bank Excel import. `client_db.upsert_from_import()` + `soft_delete_dropped()` already implemented; awaiting actual Excel schema before wiring an import script.

## GitHub Repo
`cchinkian/form-filler` (public — clients.xlsx + client_db.db gitignored)

## Key Constraints
- Target PC: Windows 10, NO admin privileges, runs from pen drive
- PDFs are flat/scanned — coordinate overlay via reportlab → pypdf merge
- No Excel file watcher (FAT32/USB unreliable) — manual Reload button
- CoordPicker stays separate from FormFiller (blast radius + bundle size)
- SmartScreen warning: right-click → Properties → Unblock (one-time fix)

## Formatters Available
`currency_myr`, `currency_no_symbol`, `date_dmy`, `date_dmy_long`, `ic_dashed`, `phone_dashed`, `uppercase`, `integer`
