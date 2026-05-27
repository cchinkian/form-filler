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
forms/              User source PDF folder, created on launch
Output/             Generated PDFs, created on launch
```

`GreenTest.exe` is no longer included in the normal release bundle. It remains only as a development/troubleshooting script.

## Main Screens

- `Generate Package`: search one customer, choose a procedure, fill missing fields, generate one combined PDF.
- `Bulk Export`: paste/import CIS list, review matched/not found/duplicate rows, generate the same procedure for matched customers.
- `Procedure Builder`: edit procedure names and package order, add/remove/reorder source forms, insert blank pages.
- `Source Forms`: maintain SourceFormCode, display name, version, PDF path, mapping key, active flag, and expiry remarks.
- `History / Settings`: open local workbook/config/history/output files and configure paths.

## Local Editable Files

```text
config/procedures.json       P001-P016 procedure master list
config/source_forms.json     SF001+ source form master list
config/procedure_items.json  Procedure composition/order/blank pages
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

- `Basic Information`
- `Investor Information`
- `ProcedureSpecificData`
- `InsuranceData`
- `Staff_Profile`
- `Bulk_CIS_Template`

Search supports customer name, CIS, IC, and policy number. Bulk matching uses CIS only.

`Staff_Profile` holds default staff/RM information used across forms:
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
