# Form Filler

Portable Windows green app (pen drive, no admin) that auto-fills flat/scanned PDF bank forms from Excel client data using coordinate-based text overlay.

## Project Status
**Production-ready engine.** All 64 unit tests pass. Pending: first real form mapped via CoordPicker on Windows PC.

## Key Files
```
form_filler/
├── src/
│   ├── main_app.py          # FormFiller GUI — bulk/single mode, settings dialog
│   ├── pdf_engine.py        # reportlab overlay + pypdf merge, 4 source types, formatters
│   ├── excel_reader.py      # Multi-sheet Excel, IC normalization, native type preservation
│   ├── config_loader.py     # Path resolver, health check, surrender logic, auto-detect
│   └── coord_picker.py      # Visual PDF field mapper — click → records x/y → forms.json
├── config/
│   ├── forms.json           # Field coordinate config + _shared_fields library
│   ├── applications.json    # Form bundles (6 defined: UT sub/redeem/switch, FD, account opening)
│   ├── settings.json        # RM profile + paths (gitignored — stays on pen drive)
│   └── settings.example.json
├── data/
│   ├── clients_template.xlsx  # 6 sheets: Master + FD_Renewal/UT_Subscription/UT_Redemption/UT_Switch/FD_New
│   └── fill_log.csv           # Audit trail (auto-created, gitignored)
├── green_test.py            # Environment verification app (4 tests incl. PDF autofill)
├── zzz_test_engine.py       # 64 automated tests — must all pass before building
├── PLAN.md                  # Architecture optimization decisions
├── HANDOFF.md               # Session handoff
└── .github/workflows/build.yml  # CI: builds all 3 EXEs on windows-latest
```

## Architecture

### Pen Drive Layout (deployed)
```
[pen drive]/
├── FormFiller.exe       # Main daily-use app
├── CoordPicker.exe      # Setup tool — map form fields once
├── GreenTest.exe        # One-time environment check
├── config/
│   ├── settings.json    # Edit: set forms_folder + RM details
│   ├── forms.json       # Auto-populated by CoordPicker
│   └── applications.json
└── data/
    └── clients.xlsx     # Client data (Master sheet + per-product batch sheets)
```

### C:\Forms\ on Target PC
```
C:\Forms\
├── kyc_form_principal\    ← one subfolder per form
│   └── kyc_form.pdf       ← exactly ONE blank PDF at top level
├── fd_renewal\
│   └── fd_renewal.pdf
└── ...                    ← 30-40 form folders
```

## Data Model (3 layers)

| Layer | Where | Contains |
|---|---|---|
| Client static | `clients.xlsx` Master sheet | Name, IC, address, phone, DOB |
| Transaction | `clients.xlsx` batch sheets (FD_Renewal, UT_Subscription, etc.) | Deal-specific data per transaction |
| Auto/fixed | `forms.json` field definitions | Today's date, RM name, bank name |

## Field Source Types (forms.json)

| source | Value comes from |
|---|---|
| `data` | Excel column (Master or batch sheet) |
| `settings` | settings.json (rm_name, rm_branch, rm_staff_id) |
| `fixed` | Hardcoded in forms.json (currency, bank name) |
| `auto` | Generated at runtime (date, year, month) |

## Shared Fields
`_shared_fields` in forms.json defines common fields once (client_name, ic_number, date, rm_name, etc.). Per-form entries only need coordinates — no source/format repeat.

## Surrender Logic
- `template_filename` + `template_hash` stored per form in forms.json
- If PDF filename changes → form **surrendered** (real fills blocked)
- Test fills still allowed → output prefixed `_TEST_`
- Re-mapping in CoordPicker clears surrender

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
python3 zzz_test_engine.py   # Must show: 64 passed, 0 failed
```

## GitHub Repo
`cchinkian/green-test-build` (public — clients.xlsx gitignored)

## Key Constraints
- Target PC: Windows 10, NO admin privileges, runs from pen drive
- PDFs are flat/scanned — coordinate overlay via reportlab → pypdf merge
- No Excel file watcher (FAT32/USB unreliable) — manual Reload button
- CoordPicker stays separate from FormFiller (blast radius + bundle size)
- SmartScreen warning: right-click → Properties → Unblock (one-time fix)

## Formatters Available
`currency_myr`, `currency_no_symbol`, `date_dmy`, `date_dmy_long`, `ic_dashed`, `phone_dashed`, `uppercase`, `integer`
