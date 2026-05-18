### 2026-04-01: PyInstaller cannot cross-compile
- PyInstaller builds .exe only on Windows — cannot build Windows .exe from macOS
- Solution: use GitHub Actions with `runs-on: windows-latest` to build in the cloud
- Workflow: push code → Actions builds → download artifact
- Build takes ~55 seconds for a simple tkinter app (11MB output)

### 2026-04-01: tkinter not available on macOS Python 3.14
- `import tkinter` fails on Homebrew Python 3.14 (no _tkinter module)
- This is fine — the app targets Windows where tkinter is bundled with Python
- Use `py_compile.compile()` for syntax verification on Mac instead of running

### 2026-04-30: tk.Separator does not exist — use ttk.Separator
- `tk.Separator` throws AttributeError immediately on Windows (GUI never opens)
- All separators must use `ttk.Separator(parent, orient="horizontal")`
- Always grep for `tk.Separator` before building — caught by code review, missed in initial dev

### 2026-04-30: Excel float IC numbers lose precision as strings
- Excel stores IC like `880101145678` as float `880101145678.0`
- `str(880101145678.0)` → `"880101145678.0"` — has `.0` suffix
- Fix: check `isinstance(raw, float)` → cast to `int` first, then `str()`
- normalize_ic() must handle float input explicitly

### 2026-04-30: FAT32 pen drives are hostile to file watchers
- Decided against 3-second mtime polling for Excel auto-reload
- FAT32 mtime resolution is 2 seconds — causes phantom "changed" events
- Excel holds exclusive lock during save — openpyxl gets PermissionError on concurrent open
- Solution: manual Reload button with timestamp display; simpler and more reliable

### 2026-04-30: PyMuPDF + Pillow must NOT be bundled into FormFiller.exe
- CoordPicker uses PyMuPDF (fitz) + Pillow for PDF rendering — heavy native libs (~30-40MB)
- Merging into FormFiller would balloon .exe from 22MB to ~55MB and risk MuPDF segfaults
- Keep CoordPicker as a separate .exe; FormFiller has an "Open CoordPicker" button instead

### 2026-04-30: forms.json.bak must be written BEFORE overwrite, not after
- If save_forms() crashes mid-write (disk full, USB yanked), forms.json is corrupted
- shutil.copy(forms.json → forms.json.bak) must happen BEFORE open(forms.json, "w")
- Always test backup creation in unit tests

### 2026-04-30: Hash check must run even when no template_filename stored
- Initial design: hash check only ran when `stored_name == pdf_path.name`
- Bug: if `template_filename` is empty (old format), hash check was skipped entirely
- Fix: hash check runs when (a) filename matches, OR (b) no filename stored (backwards compat)
- Caught by unit test T7: "hash mismatch raises" — failed because cfg had no filename

### 2026-04-30: CoordPicker docstring raw string escape warning
- `"""Scan C:\Forms..."""` → SyntaxWarning: `\F` is invalid escape sequence
- Fix: use `\\F` in regular string or `r"""..."""` raw docstring
- py_compile catches this as a SyntaxWarning but does not fail — easy to miss

### 2026-04-30: Excel field names do NOT need to match PDF label text
- Common confusion: RM assumes Excel column header must match the label printed on the PDF form
- Reality: Excel column name must match forms.json field "name" only
- PDF label ("Full Name as per NRIC:") is pre-printed on the form; app writes VALUE at coordinates
- The bridge is forms.json, not the PDF visual text

### 2026-04-30: forms_folder path drift is the #1 real-world failure
- Hardcoded path `C:\Users\460391\Documents\5. SOP Folder\...` breaks on any other machine
- Solution: default to `C:\Forms` (simple, memorable), add Browse button in Settings dialog
- Settings dialog writes directly to settings.json — RM never needs to edit JSON manually

### 2026-05-02: Committed EXE in repo can be stale vs artifact — always commit after each build
- FormFiller.exe in repo root was from an old build (before ttk.Separator fix)
- Later builds fixed the source and produced correct artifacts, but exe was never re-committed
- Result: user downloaded raw URL exe → got old broken version → crash on launch
- Fix: after every successful GitHub Actions build, download artifact and commit all 3 EXEs to repo
- Rule: raw URL (raw/main/FormFiller.exe) and artifact must always be in sync

### 2026-05-02: GreenTest.exe confirmed working on Windows 10 — PDF autofill engine verified
- All 4 GreenTest tests passed: Read Folder, Write (App Dir), Output Folder, PDF Autofill
- PDF autofill test generated correct output with all fields at correct coordinates
- Currency formatter (RM 10,000.00), date (2026-04-30), IC (880101-14-5678) all correct
- This confirms reportlab + pypdf overlay approach works on the target Windows environment

### 2026-04-30: Surrender vs TemplateChangedWarning distinction matters
- Two separate failure modes for form templates:
  - Filename changed (v1.pdf → v2.pdf) → TemplateSurrenderedError (strong: block real fills)
  - Same filename, content changed → TemplateChangedWarning (soft: offer to fill anyway)
- Keep them as separate exception types — different UX responses needed

### 2026-05-05: Hardcoded Windows paths kill pen-drive portability
- Original `settings.json` had `forms_folder = "C:\Users\460391\Documents\..."` — broke on any other PC
- Fix: default paths are now relative (`"forms"`, `"filled"`) — `_base_dir()` resolves them to siblings of the EXE
- Lesson: for portable apps, NEVER hardcode user-specific Windows paths in committed defaults; always use relative + base-dir resolution

### 2026-05-05: PowerShell `if exist` — only works in cmd, not pwsh
- GitHub Actions Windows runners default to PowerShell. `if exist file del file` fails with "Missing '(' after 'if'"
- Fix: `shell: cmd` directive on the step, OR rewrite to PowerShell syntax `if (Test-Path file) { del file }`
- Lesson: when the step uses cmd-style batch syntax, force `shell: cmd` explicitly

### 2026-05-05: tkcalendar.DateEntry needs Babel locale data — not auto-bundled
- `--collect-all tkcalendar` does NOT pull `babel` — frozen EXE crashes at runtime when DateEntry tries to format dates
- Fix: `--collect-all babel --hidden-import babel.numbers` in PyInstaller args + add `babel` to requirements.txt
- Lesson: when bundling C-extension or locale-aware libs, audit transitive deps; collect-all the package is not enough

### 2026-05-05: CoordPicker silently stripped `_shared_fields` from forms.json
- `coord_picker._save()` had: `clean = {k: v for k, v in existing.items() if not k.startswith("_")}`
- This was meant to drop the `_TEMPLATE` help block — but `_shared_fields` got dropped too
- After our Phase-3 refactor that put rm_name/staff_id/fimm_id/ippc_id/rm_branch/date INTO `_shared_fields`, this silently broke the new flow on first form save
- Fix: preserve all internal blocks, only overwrite the form being saved
- Lesson: be VERY careful with `startswith("_")` filters — any future internal block needs to be preserved, not assumed disposable. Use exact-name allowlists when removing.

### 2026-05-05: pikepdf `is_encrypted` detection misses owner-password-only PDFs
- Initial `is_encrypted()` opened with no password and caught `pikepdf.PasswordError` — return True
- This MISSED PDFs with only an owner password (restrictions but no user password) — those open fine but are still encrypted
- Fix: open the PDF AND check `bool(pdf.is_encrypted)` after open; both signals matter
- Lesson: pikepdf's `is_encrypted` attribute is the canonical check; don't infer from open-success alone

### 2026-05-05: Reorder duplicates produce copies, but rotate cumulates angles
- Same `parse_pages()` output `[3, 3, 3]` from spec `"3,3,3"` has DIFFERENT meaning per op
  - Reorder: page 3 appears 3 times in output — intentional (allows duplication)
  - Rotate: page 3 gets rotated 3× cumulatively (90 → 180 → 270) — surprise bug
- Fix: each op decides whether to dedupe before iterating; not a parser concern
- Lesson: the page-list parser is a primitive — semantics are owner-specific. Don't assume one consumer's needs apply to all.

### 2026-05-06: Daughter DBs vs central DB — separation of concerns wins for portability
- Considered extending `~/Documents/works/client_master.db` (existing central DB used by many projects) with form-fill demographics
- BLOCKED by pen-drive constraint: form_filler runs on bank PC, can't reach Mac filesystem
- Decision: keep `data/client_db.db` separate, ship on pen drive. No data overlap with central DB (different concerns: identity-mapping vs form-demographics)
- Lesson: when a portable subsystem and a central system have ZERO column overlap, separate DBs is correct — don't force a single source of truth across deployment boundaries

### 2026-05-06: Field-source tagging beats client-level lifecycle for hybrid data
- For Phase 1B (monthly bank Excel import), naive design: client-level "permanent vs monthly" flag
- Problem: even non-permanent clients have FIELDS that should never be overwritten by import (address, phone — typed by RM, not in bank export)
- Better: split fields into "bank-managed" (overwritten on import) and "user-entered sticky" (never touched). PLUS client-level permanent flag for the lifecycle question
- Lesson: hybrid data sources need TWO orthogonal axes: (1) which fields refresh, (2) which clients persist. Don't conflate them.
### 2026-05-19: Keep the product workflow clean; move old concepts out of the UI
- The daily workflow is `Search customer → choose form → fill missing fields → generate PDF`.
- Excel-first bulk processing is useful as a reference, but it should not be visible in the main product unless real usage proves it is needed.
- Dynamic PDF module composition is overkill for the actual use case. A folder with one complete merged PDF is easier to maintain.
- PDF filename/hash surrender conflicts with the user's real workflow. Folder identity + exactly one top-level PDF is the right rule.
- Keep `CoordPicker` separate because PyMuPDF/Pillow are heavy; the main FormFiller should stay small and stable.
