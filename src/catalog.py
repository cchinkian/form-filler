"""
Local catalog files for procedures, source forms, and package items.

The app intentionally keeps these as editable JSON files instead of code so
procedures and source forms can change without rebuilding the EXE.
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path
import re

import config_loader


PROCEDURES_FILE = "procedures.json"
SOURCE_FORMS_FILE = "source_forms.json"
PROCEDURE_ITEMS_FILE = "procedure_items.json"


PROCEDURE_CATEGORIES = [
    "Customer Maintenance",
    "FD / Special Product Bundle",
    "Investment / Account Opening",
    "Insurance",
]


DEFAULT_PROCEDURES = [
    {"ProcedureCode": "P001", "Category": "Customer Maintenance", "DisplayName": "Customer Maintenance", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_CustomerMaintenance", "Remarks": ""},
    {"ProcedureCode": "P002", "Category": "FD / Special Product Bundle", "DisplayName": "FD Related", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_FDRelated", "Remarks": ""},
    {"ProcedureCode": "P003", "Category": "FD / Special Product Bundle", "DisplayName": "FD Bundle + Bond", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_FDBundleBond", "Remarks": ""},
    {"ProcedureCode": "P004", "Category": "FD / Special Product Bundle", "DisplayName": "FD Bundle + UT", "Version": "V01", "Active": True, "Description": "FD bundled with UT investment", "DefaultOutputName": "ClientName_Date_FDBundleUT", "Remarks": ""},
    {"ProcedureCode": "P005", "Category": "FD / Special Product Bundle", "DisplayName": "FD Bundle + SI", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_FDBundleSI", "Remarks": ""},
    {"ProcedureCode": "P006", "Category": "FD / Special Product Bundle", "DisplayName": "FD Bundle + Banca", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_FDBundleBanca", "Remarks": ""},
    {"ProcedureCode": "P007", "Category": "FD / Special Product Bundle", "DisplayName": "FD Bundle + ASNB", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_FDBundleASNB", "Remarks": ""},
    {"ProcedureCode": "P008", "Category": "Investment / Account Opening", "DisplayName": "Risk Profile + Indemnity", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_RiskProfileIndemnity", "Remarks": ""},
    {"ProcedureCode": "P009", "Category": "Investment / Account Opening", "DisplayName": "Account Opening - Bond / SI / UT", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_AccountOpening", "Remarks": ""},
    {"ProcedureCode": "P010", "Category": "Investment / Account Opening", "DisplayName": "UT Investment", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_UTInvestment", "Remarks": ""},
    {"ProcedureCode": "P011", "Category": "Investment / Account Opening", "DisplayName": "SI Investment", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_SIInvestment", "Remarks": ""},
    {"ProcedureCode": "P012", "Category": "Investment / Account Opening", "DisplayName": "Bond Investment", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_BondInvestment", "Remarks": ""},
    {"ProcedureCode": "P013", "Category": "Insurance", "DisplayName": "Tokio Marine Subscription", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_TokioMarineSubscription", "Remarks": ""},
    {"ProcedureCode": "P014", "Category": "Insurance", "DisplayName": "Takaful Subscription", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_TakafulSubscription", "Remarks": ""},
    {"ProcedureCode": "P015", "Category": "Insurance", "DisplayName": "Tokio Marine Fund Switching", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_TokioMarineFundSwitching", "Remarks": ""},
    {"ProcedureCode": "P016", "Category": "Insurance", "DisplayName": "Takaful Fund Switching", "Version": "V01", "Active": True, "Description": "", "DefaultOutputName": "ClientName_Date_TakafulFundSwitching", "Remarks": ""},
]


DEFAULT_SOURCE_FORMS = [
    {"SourceFormCode": "SF001", "Category": "Customer Maintenance", "DisplayName": "Customer Maintenance Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF001", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF002", "Category": "FD", "DisplayName": "FD Instruction Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF002", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF003", "Category": "UT", "DisplayName": "UT Investment Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF003", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF004", "Category": "Investment", "DisplayName": "Risk Profile Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF004", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF005", "Category": "Investment", "DisplayName": "Indemnity Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF005", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF006", "Category": "Account Opening", "DisplayName": "Account Opening Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF006", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF007", "Category": "Bond", "DisplayName": "Bond Investment Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF007", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF008", "Category": "SI", "DisplayName": "SI Investment Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF008", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF009", "Category": "Banca", "DisplayName": "Banca Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF009", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF010", "Category": "ASNB", "DisplayName": "ASNB Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF010", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF011", "Category": "Insurance", "DisplayName": "Tokio Marine Subscription Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF011", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF012", "Category": "Insurance", "DisplayName": "Takaful Subscription Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF012", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF013", "Category": "Insurance", "DisplayName": "Tokio Marine Fund Switch Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF013", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
    {"SourceFormCode": "SF014", "Category": "Insurance", "DisplayName": "Takaful Fund Switch Form", "Version": "V01", "PDFFilePath": "", "MappingKey": "SF014", "Active": False, "EffectiveDate": "", "ExpiryDate": "", "Remarks": ""},
]


DEFAULT_PROCEDURE_ITEMS = [
    {"ProcedureCode": "P001", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF001", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P002", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF002", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P004", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF002", "BlankPageCount": 0, "Remarks": "FD form"},
    {"ProcedureCode": "P004", "StepNo": 2, "ItemType": "SourceForm", "SourceFormCode": "SF003", "BlankPageCount": 0, "Remarks": "UT form"},
    {"ProcedureCode": "P004", "StepNo": 3, "ItemType": "SourceForm", "SourceFormCode": "SF004", "BlankPageCount": 0, "Remarks": "Risk profile"},
    {"ProcedureCode": "P004", "StepNo": 4, "ItemType": "SourceForm", "SourceFormCode": "SF005", "BlankPageCount": 0, "Remarks": "Indemnity"},
    {"ProcedureCode": "P008", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF004", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P008", "StepNo": 2, "ItemType": "SourceForm", "SourceFormCode": "SF005", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P009", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF006", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P010", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF003", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P010", "StepNo": 2, "ItemType": "SourceForm", "SourceFormCode": "SF004", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P010", "StepNo": 3, "ItemType": "SourceForm", "SourceFormCode": "SF005", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P011", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF008", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P012", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF007", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P013", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF011", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P014", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF012", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P015", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF013", "BlankPageCount": 0, "Remarks": ""},
    {"ProcedureCode": "P016", "StepNo": 1, "ItemType": "SourceForm", "SourceFormCode": "SF014", "BlankPageCount": 0, "Remarks": ""},
]


def _path(filename: str) -> Path:
    return config_loader._config_path(filename)


def _load(filename: str, default):
    path = _path(filename)
    if not path.exists():
        save_json(filename, default)
        return list(default)
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, list) else list(default)


def save_json(filename: str, data) -> None:
    path = _path(filename)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        backup = path.with_suffix(path.suffix + ".bak")
        shutil.copy(path, backup)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_procedures(active_only: bool = False) -> list[dict]:
    rows = _load(PROCEDURES_FILE, DEFAULT_PROCEDURES)
    return [r for r in rows if is_active(r)] if active_only else rows


def save_procedures(rows: list[dict]) -> None:
    save_json(PROCEDURES_FILE, rows)


def load_source_forms(active_only: bool = False) -> list[dict]:
    rows = _load(SOURCE_FORMS_FILE, DEFAULT_SOURCE_FORMS)
    return [r for r in rows if is_active(r)] if active_only else rows


def save_source_forms(rows: list[dict]) -> None:
    save_json(SOURCE_FORMS_FILE, rows)


def load_procedure_items() -> list[dict]:
    rows = _load(PROCEDURE_ITEMS_FILE, DEFAULT_PROCEDURE_ITEMS)
    return sorted(rows, key=lambda r: (str(get_value(r, "ProcedureCode")), int(get_value(r, "StepNo", 0) or 0)))


def save_procedure_items(rows: list[dict]) -> None:
    for proc_code in sorted({get_value(r, "ProcedureCode") for r in rows}):
        proc_rows = [r for r in rows if get_value(r, "ProcedureCode") == proc_code]
        for idx, row in enumerate(sorted(proc_rows, key=lambda r: int(get_value(r, "StepNo", 0) or 0)), 1):
            row["StepNo"] = idx
    save_json(PROCEDURE_ITEMS_FILE, rows)


def get_value(row: dict, key: str, default=""):
    if key in row:
        return row.get(key, default)
    snake = key[:1].lower() + "".join(
        ("_" + c.lower() if c.isupper() else c) for c in key[1:]
    )
    return row.get(snake, default)


def set_value(row: dict, key: str, value) -> None:
    row[key] = value


def is_active(row: dict) -> bool:
    val = get_value(row, "Active", True)
    if isinstance(val, bool):
        return val
    return str(val).strip().lower() in {"yes", "y", "true", "1", "active"}


def procedure_map(active_only: bool = False) -> dict[str, dict]:
    return {get_value(r, "ProcedureCode"): r for r in load_procedures(active_only)}


def source_form_map(active_only: bool = False) -> dict[str, dict]:
    return {get_value(r, "SourceFormCode"): r for r in load_source_forms(active_only)}


def procedure_items_for(procedure_code: str) -> list[dict]:
    return [
        r for r in load_procedure_items()
        if str(get_value(r, "ProcedureCode")) == str(procedure_code)
    ]


def procedure_label(row: dict) -> str:
    return f"{get_value(row, 'ProcedureCode')} - {get_value(row, 'DisplayName')}"


def source_form_label(row: dict) -> str:
    return f"{get_value(row, 'SourceFormCode')} - {get_value(row, 'DisplayName')}"


def resolve_source_pdf_path(source_form: dict, settings: dict) -> Path:
    raw = str(get_value(source_form, "PDFFilePath", "") or "").strip()
    if not raw:
        return Path()
    path = Path(raw)
    if path.is_absolute():
        candidate = path
    else:
        base = config_loader.forms_folder_path(settings)
        candidate = base / path
    if candidate.is_dir():
        pdfs = sorted(
            p for p in candidate.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"
        )
        if len(pdfs) == 1:
            return pdfs[0]
    return candidate


def mapping_key(source_form: dict) -> str:
    return str(get_value(source_form, "MappingKey", "") or get_value(source_form, "SourceFormCode", "")).strip()


def source_code_from_folder(folder_name: str) -> str:
    match = re.match(r"^\s*(SF\d{3,}[A-Z]?)\b", folder_name, flags=re.I)
    return match.group(1).upper() if match else ""
