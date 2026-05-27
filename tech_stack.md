# Tech Stack
_Updated: 2026-05-27_

## Runtime
- Offline Windows desktop app
- Python standard library + `tkinter` / `ttk`
- Portable PyInstaller EXE, no admin permission expected

## Data
- Customer workbook: `.xlsx` / `.xlsm` read through `openpyxl`
- Procedure master: `config/procedures.json`
- Source form master: `config/source_forms.json`
- Procedure composition: `config/procedure_items.json`
- Mapping store: `config/forms.json`
- History log: `data/HistoryLog.xlsx`

## PDF
- Overlay engine: `src/pdf_engine.py`
- Package engine: `src/package_engine.py`
- PDF merge/blank pages: `pypdf`
- Text overlay: `reportlab`
- Mapping editor: `src/coord_picker.py`
- Mapping preview dependencies: `PyMuPDF`, `Pillow`

## Build
- Windows build: GitHub Actions on `windows-latest`
- Bundler: `pyinstaller`
- Release artifact: `FormFiller-portable.zip`

## Tests
- Legacy overlay smoke: `tests/sales_flow_smoke.py`
- Procedure package smoke: `tests/procedure_package_smoke.py`
