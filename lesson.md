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
