## Backlog
- [ ] Map first real form via CoordPicker on Windows PC (suggested: FD Renewal — simplest, 1 page)
- [ ] Verify new top-bar Branch + Date pickers work on Windows EXE (Babel locale data)
- [ ] Fill RM_Profile sheet in `data/clients.xlsx` with real values (rm_name, staff_id, fimm_id, ippc_id, branches)
- [ ] Test end-to-end fill: CoordPicker → forms.json → FormFiller → PDF output
- [ ] Move `zzz_test_engine.py` out of zzz_ prefix into `tests/` so CI can run it (Codex flagged: tests gitignored, regressions ship undetected)
- [ ] Add signature/initial placeholder: `type: rect` field to draw an "X sign here" box
- [ ] Multi-line text wrap for long address fields (max_width + auto-wrap in pdf_engine)
- [ ] Add print integration (auto-print after fill as optional step)
- [ ] Consider form version numbering in UI (e.g. "KYC Form v2 — mapped 2026-05-01")

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
