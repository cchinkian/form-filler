# PDF Procedure Automation

Offline Windows desktop software for preparing PDF procedure packages from Excel customer data.

The app is designed for admin/personal use where speed, offline operation, and easy maintenance of changing forms matter more than CRM-style features.

## Current Product Model

```text
Excel customer workbook
-> Source Form Library
-> Coordinate mapping
-> Procedure / Package Builder
-> Combined PDF generation
-> Local output folder
-> Excel history log
```

The PDF engine treats each PDF as a background and overlays text using saved coordinates. This works for flat, scanned, and non-fillable PDFs.

## Included Apps

Portable bundle:

```text
FormFiller.exe      Main daily-use app
CoordPicker.exe     Mapping editor for source PDF coordinates
config/             Procedures, source forms, mapping JSON, settings
data/               Customer workbook template and history log location
docs/               User manual with screenshots
forms/              User source PDF folder, created on launch
Output/             Generated PDFs, created on launch
```

`GreenTest.exe` is no longer included in the normal release bundle. It remains only as a development/troubleshooting script.

## Main Screens

- `Generate Package`: search one customer, choose a procedure, fill missing fields, generate one combined PDF.
- `Bulk Export`: paste/import CIS, IC, passport, or name list; review match status; create drafts or generate selected rows for one procedure.
- `Procedure Builder`: edit procedure names, tags, output filename tokens, package order, source forms, duplicate/sunset versions, and enable automatic blank pages after odd-page PDFs.
- `Source Forms`: manage multiple form root folders, scan source form subfolders, check PDF status, old forms folder, effective/expiry dates, mapping status, and last mapping edit.
- `Coordinate Pointer`: embedded mapping editor for placing Excel/session fields onto PDF coordinates.
- `History / Settings`: search drafts/generated history, restore previous work, and open local workbook/config/history/output files.

## User Manual

Open the manual from the app using:

```text
User Manual
```

The same guide is included in the portable bundle:

```text
docs/user_manual.html
docs/user_manual.md
```

## Local Editable Files

```text
config/procedures.json       P001-P016 procedure master list
config/source_forms.json     SF001+ source form master list
config/procedure_items.json  Procedure source-form composition/order
config/forms.json            Coordinate mappings saved by CoordPicker
config/settings.json         Local path settings, copied from settings.example.json
data/clients.xlsx            User customer workbook, created from clients_template.xlsx
data/HistoryLog.xlsx         Editable generation history log
```

Nothing is cloud-based. There is no login, network check, API call, server database, or approval workflow.

## Customer Workbook

Default workbook path:

```text
data/clients.xlsx
```

The app reads saved/calculated values from `.xlsx` or `.xlsm` workbooks. Formula results must already be saved by Excel.

The shipped template contains:

- `default_clients`
- `default_accounts`
- `default_investment`
- `default_insurance`
- `default_staff`
- `bulk_cis_template`
- `history_log`

Search supports customer name, CIS, IC, passport, and policy number. Bulk matching accepts CIS, IC, passport, or name.

Column headers starting with `*` are treated as default/locked fields in the app, for example `*cis`, `*name`, `*ic_number`.

`default_accounts` holds selectable UT/BOND/SI accounts. Use `common_name` as the human-friendly selector and keep holder CIS/IC values in `holder_1_*`, `holder_2_*`, and `holder_3_*` columns.

`default_staff` holds default staff/RM information used across forms:
staff name, staff IC, staff ID, FIMM ID, IPPC ID, position, RM code options, and branch options.
Leader/approval details are not maintained in the app; approval columns can be signed by hand.

## Output Rules

Single output:

```text
Output/
  ClientName_YYYYMMDD_ProcedureName.pdf
```

Bulk output:

```text
Output/
  BulkExport_YYYYMMDD/
    Client Name/
      Client Name_YYYYMMDD_ProcedureName.pdf
```

CIS is used for matching and history only. It is not included in generated PDF filenames.

## Build

GitHub Actions builds the Windows portable artifact on every push to `main`.

```bash
git push
```

Download from:

```text
https://github.com/cchinkian/form-filler/releases/latest/download/FormFiller-portable.zip
```

## Local Tests

```bash
python3 -m py_compile src/catalog.py src/config_loader.py src/excel_reader.py src/pdf_engine.py src/package_engine.py src/coord_picker.py src/main_app.py
python3 tests/sales_flow_smoke.py
python3 tests/procedure_package_smoke.py
```
