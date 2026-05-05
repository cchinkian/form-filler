# FormFiller Optimization Plan
# Synthesized from: Codex review + user direction ("C:\Forms folder")
# Execution: autonomous, self-tested, no approval needed

## Decisions (from Codex review)

### DROP these 3 items
- CoordPicker merge into FormFiller (blast radius + bundle weight)
- Excel file watcher 3s poll (FAT32/USB unreliable)
- "requires" evaluator (scope creep, wrong abstraction layer)

### BUILD these 11 items (priority order)

## Phase 1 — Stabilize (real RM failure modes first)

### P1-1: Excel PermissionError handler
File: src/excel_reader.py
- Wrap all load_workbook() calls in try/except PermissionError
- Return friendly message "Close clients.xlsx in Excel first"
- Main app shows dialog, not stack trace

### P1-2: Settings dialog + Browse buttons
File: src/main_app.py
- Add Settings button (gear icon) → opens dialog
- Dialog: forms_folder (Browse), output_folder (Browse), RM name/branch/staff, auto_open_output
- On save: rewrites settings.json in-place

### P1-3: Startup health check
File: src/main_app.py + src/config_loader.py
- On launch: verify forms_folder exists
- For each form in forms.json: check subfolder exists (skip _TEMPLATE/_FIELD_REFERENCE)
- Show status: "✓ 4 forms ready | ⚠ 2 subfolders missing"
- Never block launch — just warn

### P1-4: Blank field → _REVIEW_ prefix + blocking warning
File: src/pdf_engine.py + src/main_app.py
- If blanks on required fields: prefix output filename with _REVIEW_
- Warning dialog is blocking (must click OK before next client processes)
- RM sees _REVIEW_ prefix in output folder → knows to check

## Phase 2 — UX wins

### P2-5: Execute All single button
File: src/main_app.py
- Add "▶ Execute All" button in bulk mode = Select All + Bulk Fill in one click
- Keep checkbox list for selective processing
- "Execute All" uses green background, prominent position

### P2-6: Auto-open output folder after fill
File: src/main_app.py + config/settings.example.json
- Add "auto_open_output": true to settings
- After successful fill: if enabled, os.startfile(output_folder)
- Opt-out via Settings dialog checkbox

### P2-7: Session memory (last app + mode)
File: src/main_app.py
- On any change: write data/state.json {"last_app": "...", "last_mode": "bulk"}
- On startup: restore from state.json if exists
- Silent — no UI for this

### P2-8: Open Setup Tool button
File: src/main_app.py
- Small button in corner: "Open CoordPicker"
- Does os.startfile("CoordPicker.exe") if file exists next to FormFiller.exe
- Falls back to message if not found

## Phase 3 — Power features

### P3-9: _shared_fields resolver
File: src/pdf_engine.py + config/forms.json
- Add _expand_shared(fields, shared) function: 10 lines, resolution-time expansion
- Form field with "shared": "client_name" inherits source/format/font_size
- Per-form entry only specifies x, y, page (coordinates only)
- _resolve() unchanged

### P3-10: Template PDF hash check
File: src/config_loader.py
- find_template() computes MD5 of template PDF
- Compare against "template_hash" in forms.json
- If different: warn "PDF changed since last mapping. Re-run CoordPicker for this form."
- CoordPicker saves hash when saving forms.json

### P3-11: forms.json.bak before every save
File: src/config_loader.py
- save_forms(): shutil.copy forms.json → forms.json.bak before overwrite
- Prevents total loss if save crashes mid-write

## User direction: C:\Forms

- Default forms_folder = "C:\\Forms" (not the long SOP path)
- RM creates C:\Forms\ on target PC, puts subfolders there
- settings.example.json updated to C:\Forms default
- Browse button makes it easy to point anywhere else

## Files changed

| File | Changes |
|---|---|
| src/main_app.py | P1-2, P1-3, P2-5, P2-6, P2-7, P2-8 |
| src/pdf_engine.py | P1-4, P3-9 |
| src/config_loader.py | P1-3, P3-10, P3-11 |
| src/excel_reader.py | P1-1 |
| config/settings.example.json | C:\Forms default, auto_open_output |
| config/forms.json | _shared_fields template |

## Test plan (automated)

File: zzz_test_engine.py
1. IC normalization (dashes, zero-pad, already-clean)
2. _resolve() all 4 source types
3. _format_value() all formatters (currency, dates, ic_dashed, phone)
4. _expand_shared() with and without shared refs
5. Session memory save/load round-trip
6. forms.json.bak creation
7. find_template() single PDF, zero PDF, multiple PDFs
8. Excel PermissionError handling (mock)

Run: python3 zzz_test_engine.py → must pass ALL before build
