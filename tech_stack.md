# Tech Stack
_Auto-generated: 2026-05-18 16:00 MYT_

## Database
- `data/client_db.db`
- SQLite usage: `src/client_db.py`

## PDF
- Fill engine: `src/pdf_engine.py`
- PDF library: `pypdf`
- Overlay text: `reportlab`
- Coordinate mapping: `src/coord_picker.py` / `CoordPicker.exe`
- Mapping preview dependencies: `PyMuPDF`, `Pillow` (kept out of main FormFiller runtime)

## UI
- Desktop GUI: `tkinter` / `ttk`
- Date picker: `tkcalendar` + `babel`

## Build
- Windows exe build: GitHub Actions on `windows-latest`
- Bundler: `pyinstaller`

## Tests
- Core smoke test: `tests/sales_flow_smoke.py`
