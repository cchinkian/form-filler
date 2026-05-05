## Backlog
- [ ] Map first real form via CoordPicker on Windows PC (suggested: FD Renewal — simplest, 1 page)
- [ ] Test end-to-end fill: CoordPicker → forms.json → FormFiller → PDF output
- [ ] Add signature/initial placeholder: `type: rect` field to draw an "X sign here" box
- [ ] Multi-line text wrap for long address fields (max_width + auto-wrap in pdf_engine)
- [ ] Add print integration (auto-print after fill as optional step)
- [ ] Consider form version numbering in UI (e.g. "KYC Form v2 — mapped 2026-05-01")

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
