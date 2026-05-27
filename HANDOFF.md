## Current Handoff — 2026-05-27
### Current Product Direction
The active product is now the **Offline Windows PDF Procedure Automation Tool**.

Core flow:

`Search customer -> choose procedure/package -> choose related account if needed -> review default/procedure inputs -> generate PDF package -> log to Excel`

The app is local-only and offline. Do not add cloud sync, login, API validation, CRM, manager approval, OCR, email/WhatsApp sending, SQLite/server database, or online form update features for Version 1.

### Latest Release
- Release tag: `procedure-automation-source-history-update-2026-05-27`
- Release page: https://github.com/cchinkian/form-filler/releases/tag/procedure-automation-source-history-update-2026-05-27
- Latest app ZIP: https://github.com/cchinkian/form-filler/releases/latest/download/FormFiller-portable.zip
- Latest synthetic test work folder ZIP: https://github.com/cchinkian/form-filler/releases/latest/download/FormFiller_Test_WorkFolder_20260527.zip
- GitHub Windows build: `26522180335`
- Latest release assets:
  - `FormFiller-portable.zip` — contains `FormFiller.exe`, `CoordPicker.exe`, `config/`, `data/`
  - `FormFiller_Test_WorkFolder_20260527.zip` — refreshed sample work folder with 57 synthetic form folders
- `GreenTest.exe` is no longer required and is not included in the latest portable ZIP.

### Latest Commits
- `f15527f` — `Improve source form workflow and history restore`
- `d0a65b6` — `Improve mapping and procedure workflows`
- `a6417c6` — `Fix review findings for workflow update`

### Verification Completed
- `python3 -m py_compile src/catalog.py src/config_loader.py src/excel_reader.py src/pdf_engine.py src/package_engine.py src/coord_picker.py src/main_app.py`
- `python3 tests/procedure_package_smoke.py`
- `python3 tests/sales_flow_smoke.py`
- `git diff --check`
- GitHub Actions Windows build passed for run `26522180335`.
- Downloaded and inspected `FormFiller-portable.zip`:
  - `FormFiller.exe` present.
  - `CoordPicker.exe` present.
  - `GreenTest.exe` absent.
  - `config/settings.json` and `data/clients_template.xlsx` present.

### Latest Release Changes
- Source Forms now has a `Category` dropdown populated from existing procedure/source-form categories.
- Generate tab now has `Edit Category`; renaming a category updates matching procedures and source forms.
- Source Forms PDF path UI now has separate `Browse PDF` and `Browse Folder` buttons.
- Choosing a Source Form folder now auto-fills internal SourceFormCode/MappingKey/DisplayName from the folder name where possible, so those fields are less manual.
- Active source forms validate `PDFFilePath` on save. Folder paths must contain exactly one top-level PDF; duplicate-PDF warnings include the PDF filenames.
- Package generation now refuses invalid source folder paths instead of trying to process the folder as a PDF.
- `EffectiveDate` / `ExpiryDate` are now checked before single and bulk generation:
  - Not-yet-effective and expired active forms block generation.
  - Forms expiring within 30 days show a confirmation warning.
  - Procedure Structure marks date issues with `[BLOCK]` or `[WARN]`.
- Generate screen wording changed from `Non-default Information` to `Transaction / Form Inputs`.
- Generate screen now shows `Recent Client History` for the selected customer, limited to 10 rows.
- Clicking a recent history row restores the previous procedure, account, session values, and mapped procedure inputs when available.
- `HistoryLog.xlsx` keeps existing visible columns and adds a hidden `RestorePayload` column for restoreable state. Older history rows still restore basic values where possible.
- CoordPicker is now folder-first: selecting or refreshing `Template subfolder` derives the mapping key/display name from the folder and auto-loads the single top-level PDF using `config_loader.find_template`.
- Opening Mapping Editor from Source Forms passes the source folder to CoordPicker, so the Template subfolder dropdown and the loaded PDF stay linked.

### Local Verification For Current Changes
- `python3 -m py_compile src/catalog.py src/config_loader.py src/excel_reader.py src/pdf_engine.py src/package_engine.py src/coord_picker.py src/main_app.py`
- `python3 tests/procedure_package_smoke.py`
- `python3 tests/sales_flow_smoke.py`
- `git diff --check`

### Current Architecture
- Procedures are not stored as pre-combined PDFs.
- A Procedure is an ordered list of Source Forms.
- A Source Form is the real PDF/form folder plus one mapping key.
- The generation engine dynamically combines latest active source forms.
- Source form PDF paths can point to a folder; the app auto-picks exactly one top-level PDF in that folder.
- Old form versions should be placed inside an `old/` subfolder inside that form folder.
- Source form folder paths are valid only when exactly one top-level PDF is present.
- Source Form category is now a dropdown, and category rename is supported from the Generate tab.
- SourceFormCode/MappingKey/DisplayName remain internal identifiers, but the UI now derives them from the selected folder when possible.
- `ExpiryDate` / `EffectiveDate` are validated before generation; expired/not-yet-effective forms block generation and near-expiry forms warn.
- Coordinate overlay remains the main PDF strategy, so scanned/flat PDFs are supported by placing text over the PDF background.

### Current Excel Workbook Structure
The current default workbook sheets are:

- `default_clients`
- `default_accounts`
- `default_investment`
- `default_insurance`
- `default_staff`
- `bulk_cis_template`
- `history_log`

Important details:
- Headers starting with `*` are default/locked fields, e.g. `*cis`, `*name`, `*ic_number`.
- Default/locked values show in the Default Section.
- The Default Section now has an Edit button that writes back to `clients.xlsx` after making a backup.
- `.xlsm` write-back preserves VBA by opening with `keep_vba=True`.
- `default_accounts` supports `common_name`, `account_type`, `account_number`, primary holder, and joint holders.
- Related Account dropdown uses `common_name` so the user can choose the correct UT/Bond/SI/account relationship.
- CIS is used for matching and logs, but should not appear in output filenames.
- `HistoryLog.xlsx` now keeps a hidden `RestorePayload` column for richer restore/reuse from recent client history while preserving the visible history columns.

### Mapping Editor Decisions
Current format choices:

- `text`
- `ic_dashed`
- `currency_2_decimals`
- `currency_4_decimals`
- `currency_no_decimals`
- `date_ddmmyyyy`
- `date_27_May_2026`
- `all_uppercase`
- `1st_letter_case`

Removed/renamed old format names:
- `currency_myr`
- `currency_no_symbol`
- `date_dmy`
- `date_dmy_long`
- `phone_dashed`
- `uppercase`
- `integer`

User explicitly preferred not to keep old internal aliases. Existing old mappings should be updated to the new format names.

Mapping Editor current behavior:
- Source dropdown uses real sheet names from `clients.xlsx` plus `Session Input`, `Fixed Text`, and `Auto Date`.
- Excel fields can be reloaded after workbook header changes.
- Field ID / Excel field entry auto-suggests fields from the selected sheet.
- Format auto-suggests from the field name.
- Phone/mobile/contact/WhatsApp defaults to `text`.
- Red markers can be selected.
- Selected fields can be updated/deleted.
- Marker/box can be dragged.
- Box width can be resized visually by dragging the right-side handle.
- Field list can select the mapped field.
- Clear Page and Clear All are available with confirmation.
- Opening Mapping Editor from Source Forms now loads existing mapping first, so saving will not wipe previous mappings.

### Procedure Builder Decisions
- User wants a left list of available source forms and a right list of procedure sequence.
- Implemented searchable left source-form list, Add button, double-click add, Remove, Up, Down, Save.
- Inactive source forms cannot be added to a procedure.
- Adding/editing category/procedure from the Generate tab is supported.
- True drag-and-drop reorder is not implemented yet; current reorder uses Up/Down buttons.
- Blank pages are now mainly auto-inserted after odd-page source PDFs.
- Auto blank follows the previous page size/orientation.
- Manual blank page remains available as an advanced option.
- Auto blank avoids double-inserting immediately before a manual BlankPage item.
- Auto blank defaults ON for missing old settings/procedure values.

### Branch/Staff Decisions
- Branch should stay blank when the user leaves Branch blank.
- Do not auto-fill branch from customer branch in Excel.
- Staff info lives in `default_staff`:
  - `staff_name`
  - `staff_ic`
  - `staff_id`
  - `fimm_id`
  - `ippc_id`
  - `staff_position`
  - `staff_rm_codes`
  - `staff_branches`
- Leader/approval information should not be auto-filled for Version 1.
- Approval/signature areas are left for handwriting.

### Test Work Folder
Local refreshed folder:

`/Users/ck_macmini/Downloads/FormFiller_Test_WorkFolder_20260527`

ZIP:

`/Users/ck_macmini/Downloads/FormFiller_Test_WorkFolder_20260527.zip`

Inside the ZIP:

`FormFiller_Test_WorkFolder_20260527/5.SOP Folder/All About Series/All About Series - Investment Basic Forms/Automate/`

The refreshed test folder:
- Uses the current workbook sheet names.
- Contains `default_accounts` with account `common_name` examples.
- Has no old mapping format names.
- Has no manual blank rows in sample procedure items.
- Uses folder-based source form paths.
- Includes sample generated PDFs.

### Known Gaps / Next Work
- Recent history restore depends on new `RestorePayload` for full fidelity; older log rows restore only basic fields such as procedure and common transaction inputs.
- Bulk export currently prioritizes one same Procedure for the whole batch.
- Mixed procedure per client in one bulk batch remains deferred.
- Account selection is implemented for the single-customer Generate tab; bulk export does not yet choose per-client account rows.
- True drag-and-drop in Procedure Builder is still a future improvement.
- Search primarily loads customer rows from customer sheets; account-only joint holder search may still need improvement if a joint holder exists only in `default_accounts` and not `default_clients`.
- Default-field write-back updates one existing matched row; it is not a full Excel data-entry system.
- UI is still Tkinter/admin-tool style. It is functional but can be redesigned further if the user wants a more polished interface.

### Local Repo Notes
- Current repo path: `/Users/ck_macmini/Documents/tools/form_filler`
- GitHub repo: `cchinkian/form-filler`
- Current branch: `main`
- Untracked local files exist and should not be touched unless user asks:
  - `docs/business_opportunity.md`
  - `review by claude and codex.md`

---

## Session 2026-05-19
### Goal
Simplify FormFiller into the clean daily sales workflow:

`Search customer → choose form → fill missing fields → generate PDF`

### What Changed
- Main UI now defaults to Sales Flow. Excel-first bulk controls are no longer exposed in the daily UI.
- Customer search now uses one box for name / IC / passport / CIF via `client_db.search()`.
- `client_db.py` gained `passport_number`, passport/CIF lookup helpers, and a lightweight migration for existing `client_db.db`.
- The form-specific input panel is now derived from the selected form's `forms.json` mapping, not from an Excel product sheet.
- PDF routing is folder-based: each form folder must contain exactly one top-level PDF; PDF filename no longer controls fill eligibility.
- CoordPicker now generates/updates `data/<form_id>_template.xlsx` from mapped `source=data` labels.
- Legacy/deferred product ideas are recorded in `DEFERRED_FEATURES.md`.
- Added `tests/sales_flow_smoke.py` to prove the core PDF fill path with a temporary PDF/form/customer.

### Validation
- `python3 -m py_compile src/client_db.py src/main_app.py src/config_loader.py src/pdf_engine.py src/coord_picker.py src/excel_reader.py tests/sales_flow_smoke.py` ✅
- `python3 tests/sales_flow_smoke.py` ✅
  - Creates temp PDF with arbitrary filename
  - Uses single-PDF folder lookup
  - Fills customer name, IC, CIF, amount, date
  - Verifies generated PDF text contains expected values
- GitHub Actions Windows build ✅
  - Run: `26045873626`
  - Commit: `010fa0e`
  - Artifact: `FormFiller-portable`
  - Local downloaded copy: `build_artifacts/FormFiller-portable-26045873626/`
  - Bundle contains `FormFiller.exe`, `CoordPicker.exe`, `GreenTest.exe`, `config/`, `data/`
- GitHub Release download ✅
  - https://github.com/cchinkian/form-filler/releases/tag/sales-flow-2026-05-19
  - Simple latest URL: https://github.com/cchinkian/form-filler/releases/latest
  - Direct download: https://github.com/cchinkian/form-filler/releases/latest/download/FormFiller-portable.zip
  - Asset: `FormFiller-portable-26045873626.zip`
  - Simple asset: `FormFiller-portable.zip`
  - Unzip result: full folder `FormFiller-portable-26045873626/`; double-click `FormFiller.exe`

### Product Decision
- The active product should stay clean and focused on the 4-step sales flow.
- Excel bulk workflow, dynamic PDF module composition, preview-first flow, and filename/hash surrender are deferred/removed from the product path.
- CoordPicker remains a separate setup tool because it uses PyMuPDF/Pillow and should not bloat or destabilize FormFiller.exe.

### Portable App Notes
- No install required.
- No admin permission required for normal use.
- Keep the whole portable folder together and double-click `FormFiller.exe`.
- The app does not install services or registry entries. It does keep working files inside its own folder: `data/client_db.db`, `filled/`, `data/fill_log.csv`, and `config/backups/`.

### Next Step
- Copy `build_artifacts/FormFiller-portable-26045873626/` to the user's working folder or pen drive, then test on the Windows/bank PC by double-clicking `FormFiller.exe`.

---

## Backlog
- [ ] **Push 2026-05-19 Sales Flow build to GitHub** — triggers CI rebuild with simplified UI + folder-based PDF logic
- [ ] **Run `zzz_migrate_master_to_db.py`** on the actual data after deploying (one-shot, idempotent)  <!-- 2026-05-18: PENDING-USER — must run on bank PC -->
- [ ] Map first real form via CoordPicker on Windows PC (suggested: FD Renewal — simplest, 1 page)  <!-- 2026-05-18: PENDING-USER — Windows-only -->
- [ ] Verify new top-bar Branch + Date pickers work on Windows EXE (Babel locale data)  <!-- 2026-05-18: PENDING-USER — Windows-only -->
- [ ] Fill RM_Profile sheet in `data/clients.xlsx` with real values (rm_name, staff_id, fimm_id, ippc_id, branches)  <!-- 2026-05-18: PENDING-USER — needs RM's own IDs/branch CSV -->
- [ ] Test end-to-end fill: CoordPicker → forms.json → FormFiller → PDF output  <!-- 2026-05-18: PENDING-USER — Windows test -->
- [ ] **Phase 4 — UT repeatable fields (max 4 slots)** — see Session 2026-05-06 for design notes  <!-- 2026-05-18: ASPIRATIONAL — gated on Phase 1A having 1 week of real-world use -->
- [ ] **Phase 1B — Monthly bank Excel import** — `client_db.upsert_from_import()` already implemented; needs an import script + Excel schema (deferred until user can share columns)  <!-- 2026-05-18: PENDING-USER — needs Excel schema from bank -->
- [ ] Move `zzz_test_engine.py` out of zzz_ prefix into `tests/` so CI can run it (Codex flagged: tests gitignored, regressions ship undetected)  <!-- 2026-05-18: verified no zzz_test_engine.py exists in /tools/form_filler/ — file was either renamed or already removed (only green_test.py present). Still PENDING-USER if test file actually lives in /works/form_filler/ or was renamed silently -->
- [ ] Add signature/initial placeholder: `type: rect` field to draw an "X sign here" box  <!-- 2026-05-18: ASPIRATIONAL — feature backlog -->
- [ ] Multi-line text wrap for long address fields (max_width + auto-wrap in pdf_engine)  <!-- 2026-05-18: ASPIRATIONAL -->
- [ ] Add print integration (auto-print after fill as optional step)  <!-- 2026-05-18: ASPIRATIONAL -->
- [ ] Consider form version numbering in UI (e.g. "KYC Form v2 — mapped 2026-05-01")  <!-- 2026-05-18: ASPIRATIONAL — "Consider…" -->

---

## Session 2026-05-06
### Goal
Phase 1A — replace Excel Master sheet with SQLite client master DB and add search-based GUI for IC/Name lookup + CRUD. Build user's vision for an integrated Service Request system, starting with the most painful workflow item (Q1=a: typing IC/address every form).

### What Was Done — Phase 1A (commit `73c04c7`, LOCAL only — not pushed)
- **`src/client_db.py`** (new, 341 lines) — SQLite store at `data/client_db.db`. `clients` table:
  - `ic_number` PK (normalized 12-digit), `name`, `cif_no`
  - User-entered sticky fields: phone, email, address_line1/2, city, state, postcode, dob, occupation, notes
  - `bank_fields TEXT` (JSON, future Phase 1B — preserves any unknown columns from a future monthly Excel import without schema changes)
  - Lifecycle flags: `permanent` (1=never auto-removed), `active` (0=soft-deleted), `source`, `last_seen_in_import`
  - CRUD: add, update, soft_delete, hard_delete, restore, set_permanent
  - Search: `by_ic` (full normalized), `by_name` (case-insensitive substring)
  - Phase 1B helpers (not yet wired): `upsert_from_import()`, `soft_delete_dropped()`
- **`src/main_app.py`** — single-mode panel rebuilt around DB:
  - Search box with live typeahead: digits-only → full IC exact match; otherwise partial-name match
  - Results listbox with permanent ⭐ marker
  - ➕ Add / ✏️ Edit / 🗑 Delete buttons; Delete dialog: Yes=soft, No=hard with name-typing confirm, Cancel=abort
  - `_load_master_data()` helper: DB first, Master sheet fallback (zero-touch upgrade — old setups still work)
  - Bulk mode unchanged in this phase
- **`zzz_migrate_master_to_db.py`** (gitignored, zzz_ prefix) — one-shot reads Master sheet → upserts into `client_db.db`. Idempotent (skips IC already in DB). Tested: 3 sample clients migrated cleanly.
- **`.gitignore`** — `client_db.db` + journal added (PDPA: never commit)

### Design Decisions Locked This Session
- **Two-DB separation** (vs extending `~/Documents/works/client_master.db`): pen-drive deployment forces local DB. Zero column overlap with central DB so no duplication concern. Documented in lesson.md.
- **Field-source tagging**: bank fields (refresh on import) vs user-entered sticky fields (never overwritten). Combined with client-level `permanent` flag — two orthogonal axes for hybrid data lifecycle.
- **Phase 4 design** (next big feature): max 4 fund slots per UT form (user confirmed). `forms.json` to gain `slot_index` + `slot_group` per field. CoordPicker will get an "Add slot" button. FormFiller GUI will detect repeatable groups → ask "How many funds?" → render N rows → fill first N coordinate slots, skip rest. Estimated 2-3 weeks once Phase 1A is in production.
- **Search UX**: full-IC exact match (12 digits, no partial), partial-name substring match. Listbox style, not dropdown — scales better when DB grows past ~50 clients.

### Codex Review (this session)
Both projects passed Codex review. Form_filler had 5 medium issues (CoordPicker `_shared_fields` strip, Babel missing in CI, fallback date validation, health_check subdir count, tests gitignored). All fixed except the last (architectural, deferred).

### What Worked
- Smoke tests pass: 10/10 client_db CRUD tests, migration idempotent, search by full IC + partial name returns correct results.
- All 5 source files compile clean on Mac (tkinter local issue irrelevant — CI Windows handles it).
- Two-DB design holds up — no data overlap with `~/Documents/works/client_master.db`.

### What Didn't Work
- Mac M-series Macmini has only 16GB free disk → Windows VM via VMware Fusion was infeasible (needs ~40GB). User accepted Mac-Python source testing or bank-PC testing as the path. Documented for next session.

### Next Steps
1. **User decides on push timing** — Phase 1A is committed locally (`73c04c7`), ready to push. They were going to test the previous bundle on Windows in parallel; once that's done, push triggers CI rebuild.
2. After push → download new `FormFiller-portable` artifact → deploy to pen drive
3. Run `zzz_migrate_master_to_db.py` once on bank PC to seed DB from existing Master sheet
4. Use search-based GUI for daily work; report any rough edges
5. Once Phase 1A has real-world use for ~1 week, kick off Phase 4 (UT repeatable fields)

---

## Session 2026-05-05
### Goal
Make FormFiller fully portable and refactor RM identity out of `settings.json` into Excel.

### What Was Done — portable rebuild (commit `5127514`)
- **Paths now relative to EXE by default** — `forms_folder="forms"`, `output_folder="filled"`. `_base_dir()` resolves via `resolve_path()`. Pen drive plugs into any PC, no path edits.
- **Auto-create folders + clone clients_template.xlsx on first launch** — no more crash on fresh deploy. Friendly first-run banner.
- **RM identity moved to Excel** — new `RM_Profile` sheet in `clients.xlsx` (rm_name, staff_id, fimm_id, ippc_id, branches CSV). New `excel_reader.load_rm_profile()`.
- **New PDF source types** — `rm_profile` and `session`. `pdf_engine._resolve()` and `fill_bundle()` updated. Backward-compat with `source: "settings"` preserved.
- **Top bar gains Branch dropdown + Date picker** (`tkcalendar.DateEntry`, dd/mm/yyyy). Pick once per session, applies to all fills.
- **One-click open buttons** — 📊 Open Excel, ⚙ Settings (with Open settings.json / forms.json / backups), 📁 Open Config.
- **`settings.json` auto-backup** every launch → `config/backups/settings_YYYYMMDD_HHMMSS.json`. Last 100 retained.
- **CI rebuilt** to bundle FormFiller + CoordPicker + GreenTest + config + data into `FormFiller-portable` artifact.

### What Was Done — Codex review fixes (commit `5fb48ea`)
- 🔴 **CoordPicker no longer strips `_shared_fields`** — was deleting the new staff_id/fimm_id/ippc_id/rm_branch/date library from forms.json on every save. Critical fix — without it the new RM_Profile flow silently breaks on first form save.
- 🔴 **CI now bundles Babel** (`--collect-all babel --hidden-import babel.numbers`) — tkcalendar.DateEntry depends on Babel locale data; without it the frozen EXE crashes on Windows runtime.
- 🟡 **Fallback date Entry validates dd/mm/yyyy** — if tkcalendar unavailable, accept dd/mm/yyyy or normalize from common alternates; bad input shows error rather than corrupting form fields.
- 🟡 **`health_check` empty detection counts subdirs only** — a stray loose file was hiding the first-run banner.
- 🟢 Deferred: move tests out of `zzz_` prefix so CI can run them (Codex finding, architectural).

### What Worked
- All 5 Python files compile clean on Mac (tkinter local issue, but py_compile passes).
- Both CI builds succeed (`run 25351512968` is green).
- Backward compat preserved — old `forms.json` files with `source: "settings"` still resolve correctly.

### Next Steps (for next session — starting on Windows)
1. **Download** the new bundle: https://github.com/cchinkian/form-filler/actions/runs/25351512968 → Artifacts → `FormFiller-portable`
2. **Drop on pen drive**, plug into bank PC.
3. **Run GreenTest.exe** first to verify env (one-time).
4. **Open `data/clients.xlsx` → RM_Profile sheet** — fill in your name, staff ID, FIMM ID, IPPC ID, branches (CSV like `KJG,SUJ,KEL`).
5. **Drop a blank form PDF** into `forms/<form_name>/<form>.pdf`.
6. **Run CoordPicker.exe** → map fields → save (verify `_shared_fields` survives in forms.json).
7. **Run FormFiller.exe** → top bar should show Application + Branch + Date + Mode. Pick branch + date, fill a form.
8. Verify the filled PDF has correct branch + date + RM identity.

---

## Session 2026-05-02
### Goal
Deploy and verify FormFiller.exe on the target Windows PC.

### What Was Done
- **GreenTest.exe verified** — all 4 tests passed on Windows 10 (Python 3.11.9, non-admin):
  - Read Folder: OK (found SOP folder PDFs)
  - Write App Dir: OK (C:\Users\460391\Downloads)
  - Output Folder: OK
  - PDF Autofill: OK — generated test PDF with correct coordinate-based text, currency formatter working
- **FormFiller.exe crash diagnosed** — `AttributeError: module 'tkinter' has no attribute 'Separator'` at line 80
  - Root cause: committed EXE in repo was from an old build before the ttk.Separator fix
  - The artifact from run 25174618031 had the correct fixed EXE, but it was never committed to repo
- **Fixed** — downloaded fixed artifact, committed all 3 EXEs (FormFiller.exe, CoordPicker.exe, GreenTest.exe) to repo
  - Commit: `d5896eb` — "Update all 3 EXEs to latest build (fix tk.Separator crash in FormFiller)"
  - Raw download URL now works: `https://github.com/cchinkian/green-test-build/raw/main/FormFiller.exe`
- **SmartScreen Unblock** — explained full steps (right-click → Properties → tick Unblock → Apply → OK)

### What Worked
- GreenTest.exe PDF autofill engine confirmed working on the actual target Windows machine
- Artifact had the correct fixed EXE — just needed to be committed to repo

### What Didn't Work
- Stale committed EXE caused crash — lesson: always commit EXE after each successful build

### Next Steps
1. **Download fixed FormFiller.exe** from: `https://github.com/cchinkian/green-test-build/raw/main/FormFiller.exe`
2. **Unblock**: right-click → Properties → tick Unblock → Apply → OK
3. **Run FormFiller.exe** — should launch without crash now
4. **Settings dialog**: click ⚙ Settings → set forms_folder to wherever your forms are (e.g. `C:\Forms` or the SOP folder path seen in GreenTest)
5. **Create form folder**: e.g. `C:\Forms\fd_renewal\` → place blank FD Renewal PDF inside
6. **Run CoordPicker.exe** → select `fd_renewal` subfolder → open PDF → click each field → Save
7. **Add test row** to FD_Renewal sheet in clients.xlsx → run FormFiller → Execute All → verify PDF output

---

## Session 2026-04-30 / 2026-05-02
### Goal
Build full production-ready Windows portable form filler app from scratch.

### What Was Done
- All 5 source files built (main_app.py, pdf_engine.py, excel_reader.py, config_loader.py, coord_picker.py)
- 64-test suite passing (3 consecutive runs)
- Key features: 4 source types, 8 formatters, _shared_fields, surrender logic, auto-detect C:\Forms, health check, session memory, ExcelLockedError handling, forms.json.bak
- GitHub Actions CI building 3 EXEs + portable ZIP artifact

### Next Steps (remaining from this session)
- Map first real form via CoordPicker (FD Renewal suggested)
- Test full end-to-end fill on Windows
