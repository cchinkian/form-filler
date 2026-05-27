"""
Procedure/package generation built on top of the existing coordinate overlay.

Procedure = ordered source forms + blank pages. Each source form has its own
PDF path and mapping key. The final output is one combined PDF per client.
"""
from __future__ import annotations

import datetime
import re
import tempfile
from pathlib import Path

import pypdf

import catalog
import config_loader
import pdf_engine


COMMON_DATA_FIELDS = {
    "name", "client_name", "customer_name",
    "cis", "cis_no", "cis_number", "cif_no",
    "ic", "ic_number", "nric", "new_ic",
    "policy", "policy_no", "policy_number",
    "phone", "mobile", "email",
    "address", "address_line1", "address_line2", "city", "state", "postcode",
    "dob", "occupation",
    "branch", "branch_code", "rm_branch", "date",
    "staff_name", "staff_ic", "staff_id", "fimm_id", "ippc_id",
    "staff_position", "staff_rm_code", "rm_code",
}


def safe_filename(value: str, fallback: str = "Client") -> str:
    text = str(value or fallback).strip() or fallback
    text = re.sub(r'[<>:"/\\|?*\n\r\t]+', " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:120] or fallback


def yyyymmdd(date_text: str | None = None) -> str:
    if date_text:
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y%m%d"):
            try:
                return datetime.datetime.strptime(str(date_text), fmt).strftime("%Y%m%d")
            except ValueError:
                pass
    return datetime.date.today().strftime("%Y%m%d")


def procedure_display_name(procedure: dict) -> str:
    return str(catalog.get_value(procedure, "DisplayName", "Procedure"))


def output_path_for_client(
    output_root: Path,
    client: dict,
    procedure: dict,
    session: dict,
    bulk_root: Path | None = None,
    client_folder: bool = False,
) -> Path:
    client_name = safe_filename(client.get("name") or client.get("client_name") or "Client")
    proc_name = safe_filename(procedure_display_name(procedure), "Procedure")
    filename = f"{client_name}_{yyyymmdd(session.get('date'))}_{proc_name}.pdf"
    root = bulk_root or output_root
    if client_folder:
        root = root / client_name
    return root / filename


def _expand_fields(form_config: dict, shared: dict) -> list[dict]:
    fields = []
    for raw in form_config.get("fields", []):
        field = dict(shared.get(raw.get("shared"), {}))
        field.update(raw)
        if "shared" in field:
            field.pop("shared", None)
        fields.append(field)
    return fields


def _normalize_sheet_key(raw) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(raw or "").strip().lower()).strip("_")


def field_identity(field: dict) -> str:
    key = str(field.get("key") or "").strip()
    sheet = str(field.get("excel_sheet") or "").strip()
    return f"{sheet}::{key}" if sheet else key


def split_field_identity(identity: str) -> tuple[str, str]:
    text = str(identity or "")
    if "::" in text:
        sheet, key = text.split("::", 1)
        return sheet, key
    return "", text


def field_value(client: dict, field: dict):
    key = field.get("key")
    sheet = field.get("excel_sheet") or ""
    if not key:
        return ""
    sheet_data = client.get("_sheet_data", {})
    if sheet and isinstance(sheet_data, dict):
        found_sheet = False
        for candidate in (sheet, _normalize_sheet_key(sheet), str(sheet).lower()):
            bucket = sheet_data.get(candidate, {})
            if isinstance(bucket, dict):
                found_sheet = True
                if bucket.get(key) not in ("", None):
                    return bucket.get(key)
        if found_sheet:
            return ""
    return client.get(key, "")


def merge_manual_values(client: dict, manual_values: dict | None) -> dict:
    merged = dict(client or {})
    sheet_data = {
        k: dict(v)
        for k, v in (merged.get("_sheet_data", {}) or {}).items()
        if isinstance(v, dict)
    }
    for identity, value in (manual_values or {}).items():
        if value in ("", None):
            continue
        sheet, key = split_field_identity(identity)
        if not key:
            continue
        if sheet:
            bucket = sheet_data.setdefault(sheet, {})
            bucket[key] = value
            normalized = _normalize_sheet_key(sheet)
            if normalized and normalized != sheet:
                sheet_data.setdefault(normalized, bucket)
        merged[key] = value
    if sheet_data:
        merged["_sheet_data"] = sheet_data
    return merged


def data_fields_for_procedure(
    procedure_code: str,
    forms_config: dict,
    source_forms: dict[str, dict],
    procedure_items: list[dict],
    include_common: bool = False,
) -> list[dict]:
    shared = forms_config.get("_shared_fields", {})
    rows: list[dict] = []
    seen = set()
    for item in procedure_items:
        if catalog.get_value(item, "ProcedureCode") != procedure_code:
            continue
        if catalog.get_value(item, "ItemType") != "SourceForm":
            continue
        source_code = catalog.get_value(item, "SourceFormCode")
        source = source_forms.get(source_code, {})
        mapping = forms_config.get(catalog.mapping_key(source), {})
        for field in _expand_fields(mapping, shared):
            source_type = field.get("source", "data")
            if source_type != "data":
                continue
            key = field.get("data_key") or field.get("ExcelColumnOrField") or field.get("name")
            if not key:
                continue
            if not include_common and key in COMMON_DATA_FIELDS:
                continue
            sheet = field.get("excel_sheet") or field.get("ExcelSheet") or ""
            identity = f"{sheet}::{key}" if sheet else key
            if identity in seen:
                continue
            seen.add(identity)
            rows.append({
                "id": identity,
                "key": key,
                "label": field.get("DisplayLabel") or key.replace("_", " ").title(),
                "required": bool(field.get("required") or field.get("Required")),
                "source_form": source_code,
                "excel_sheet": sheet,
            })
    return rows


def missing_required_fields(
    procedure_code: str,
    client: dict,
    forms_config: dict,
    source_forms: dict[str, dict],
    procedure_items: list[dict],
) -> list[dict]:
    missing = []
    for field in data_fields_for_procedure(
        procedure_code, forms_config, source_forms, procedure_items, include_common=True
    ):
        if not field["required"]:
            continue
        if field_value(client, field) in ("", None):
            missing.append(field)
    return missing


def _append_pdf(writer: pypdf.PdfWriter, path: Path) -> tuple[int, float, float]:
    reader = pypdf.PdfReader(str(path))
    count = len(reader.pages)
    last_width, last_height = 595.0, 842.0
    for page in reader.pages:
        box = page.mediabox
        last_width, last_height = float(box.width), float(box.height)
        writer.add_page(page)
    return count, last_width, last_height


def _add_blank_pages(writer: pypdf.PdfWriter, count: int, width: float = 595, height: float = 842) -> None:
    for _ in range(max(0, int(count or 0))):
        writer.add_blank_page(width=width, height=height)


def _source_mapping(forms_config: dict, source_form: dict) -> dict:
    key = catalog.mapping_key(source_form)
    return forms_config.get(key, {
        "name": catalog.get_value(source_form, "DisplayName", key),
        "fields": [],
    })


def _truthy(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"1", "yes", "y", "true", "on"}


def generate_package(
    procedure: dict,
    source_forms: dict[str, dict],
    procedure_items: list[dict],
    forms_config: dict,
    client: dict,
    settings: dict,
    output_root: Path,
    session: dict | None = None,
    manual_values: dict | None = None,
    bulk_root: Path | None = None,
    client_folder: bool = False,
) -> dict:
    session = dict(session or {})
    manual_values = dict(manual_values or {})

    branch = (
        session.get("rm_branch")
        or settings.get("default_branch")
        or ""
    )
    rm_code = (
        session.get("staff_rm_code")
        or session.get("rm_code")
        or settings.get("default_rm_code")
        or ""
    )
    session["rm_branch"] = branch
    session["staff_rm_code"] = rm_code
    session["rm_code"] = rm_code
    session.setdefault("date", datetime.date.today().strftime("%d/%m/%Y"))

    fill_data = merge_manual_values(client, manual_values)
    fill_data.setdefault("date", session["date"])
    fill_data.setdefault("rm_branch", branch)
    fill_data.setdefault("branch_code", branch)
    fill_data.setdefault("staff_rm_code", rm_code)
    fill_data.setdefault("rm_code", rm_code)

    procedure_code = catalog.get_value(procedure, "ProcedureCode")
    proc_auto_blank = catalog.get_value(procedure, "AutoBlankAfterOdd", None)
    if proc_auto_blank in ("", None):
        auto_blank_after_odd = _truthy(settings.get("auto_blank_after_odd", True))
    else:
        auto_blank_after_odd = _truthy(proc_auto_blank)
    output_path = output_path_for_client(
        output_root, fill_data, procedure, session,
        bulk_root=bulk_root, client_folder=client_folder
    )
    writer = pypdf.PdfWriter()
    warnings: list[str] = []
    source_count = 0

    settings_for_overlay = {
        **settings,
        "_shared_fields": forms_config.get("_shared_fields", {}),
        "_session": session,
        "_rm_profile": settings.get("_rm_profile", {}),
        "_staff_profile": settings.get("_staff_profile", settings.get("_rm_profile", {})),
    }

    with tempfile.TemporaryDirectory() as td:
        tmp_dir = Path(td)
        sorted_items = sorted(procedure_items, key=lambda r: int(catalog.get_value(r, "StepNo", 0) or 0))
        proc_items = [
            item for item in sorted_items
            if catalog.get_value(item, "ProcedureCode") == procedure_code
        ]
        for idx, item in enumerate(proc_items):
            item_type = catalog.get_value(item, "ItemType")
            if item_type == "BlankPage":
                _add_blank_pages(writer, int(catalog.get_value(item, "BlankPageCount", 1) or 1))
                continue

            if item_type != "SourceForm":
                continue

            source_code = catalog.get_value(item, "SourceFormCode")
            source = source_forms.get(source_code)
            if not source:
                raise ValueError(f"Source form '{source_code}' not found in source_forms.json.")
            if not catalog.is_active(source):
                raise ValueError(f"{source_code} is inactive. Activate it only after the PDF path and mapping are ready.")

            pdf_path = catalog.resolve_source_pdf_path(source, settings)
            if not pdf_path or not pdf_path.exists():
                raise FileNotFoundError(
                    f"{source_code} PDF not found. Set PDFFilePath in source_forms.json."
                )

            mapping_key = catalog.mapping_key(source)
            form_config = _source_mapping(forms_config, source)
            if mapping_key not in forms_config:
                warnings.append(f"{source_code}: no mapping found for MappingKey '{mapping_key}'")
            tmp_pdf = tmp_dir / f"{source_code}.pdf"
            blanks, final_tmp = pdf_engine.fill_form(
                pdf_path, form_config, fill_data, settings_for_overlay, tmp_pdf
            )
            if blanks:
                warnings.append(f"{source_code}: missing required fields - {', '.join(blanks)}")
            page_count, last_width, last_height = _append_pdf(writer, final_tmp)
            source_count += 1

            extra_blank = int(catalog.get_value(item, "BlankPageCount", 0) or 0)
            next_item = proc_items[idx + 1] if idx + 1 < len(proc_items) else {}
            next_is_manual_blank = catalog.get_value(next_item, "ItemType") == "BlankPage"
            if extra_blank:
                _add_blank_pages(writer, extra_blank, last_width, last_height)
            elif auto_blank_after_odd and page_count % 2 == 1 and not next_is_manual_blank:
                _add_blank_pages(writer, 1, last_width, last_height)

    if source_count == 0 and not writer.pages:
        raise ValueError(f"{procedure_code} has no source forms or blank pages.")

    if warnings:
        output_path = output_path.parent / f"_REVIEW_{output_path.name}"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    return {
        "output_path": output_path,
        "warnings": warnings,
        "status": "Success" if not warnings else "Missing Required Data",
        "client": fill_data,
        "procedure_code": procedure_code,
        "procedure_name": procedure_display_name(procedure),
    }


def history_row(result: dict, generated_by: str = "") -> dict:
    client = result.get("client", {})
    return {
        "GeneratedDateTime": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ClientName": client.get("name") or client.get("client_name") or "",
        "CIS": client.get("cis") or client.get("cif_no") or "",
        "ProcedureCode": result.get("procedure_code", ""),
        "ProcedureName": result.get("procedure_name", ""),
        "OutputFilePath": str(result.get("output_path", "")),
        "Amount": client.get("amount", ""),
        "FDDetails": client.get("fd_details", ""),
        "ProductType": client.get("product_type", ""),
        "ActionPurpose": client.get("action_purpose", ""),
        "FollowUpNote": client.get("follow_up_note", ""),
        "Status": result.get("status", ""),
        "ErrorMessage": " | ".join(result.get("warnings", [])),
        "GeneratedBy": generated_by,
    }


def error_history_row(client: dict, procedure: dict, status: str, error: str, generated_by: str = "") -> dict:
    return {
        "GeneratedDateTime": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ClientName": client.get("name") or client.get("client_name") or "",
        "CIS": client.get("cis") or client.get("cif_no") or "",
        "ProcedureCode": catalog.get_value(procedure, "ProcedureCode", ""),
        "ProcedureName": procedure_display_name(procedure) if procedure else "",
        "OutputFilePath": "",
        "Amount": client.get("amount", ""),
        "FDDetails": client.get("fd_details", ""),
        "ProductType": client.get("product_type", ""),
        "ActionPurpose": client.get("action_purpose", ""),
        "FollowUpNote": client.get("follow_up_note", ""),
        "Status": status,
        "ErrorMessage": error,
        "GeneratedBy": generated_by,
    }
