"""
Excel reader — multi-sheet, fully dynamic.
Master sheet: client static info keyed by ic_number.
Batch sheets: transaction-specific data joined to Master by ic_number.

P1-1: All load_workbook() calls protected against PermissionError
       (Excel holds exclusive lock on Windows when file is open).
Fix-8: IC numbers normalized both sides of join.
Fix-9: Native Python types preserved through reader.
"""
import re
import openpyxl
from pathlib import Path


# ── IC normalization ──────────────────────────────────────────────────────────

def normalize_ic(raw) -> str:
    """Strip non-digits, zero-pad to 12. Empty → ''.
    Handles float from Excel (880101145678.0 → '880101145678').
    """
    # Convert float to int string first to avoid '880101145678.0'
    if isinstance(raw, float):
        raw = int(raw)
    cleaned = re.sub(r"[^0-9]", "", str(raw))
    return cleaned.zfill(12) if cleaned else ""


# ── Safe workbook open ────────────────────────────────────────────────────────

class ExcelLockedError(Exception):
    """Raised when Excel holds an exclusive lock on the file."""


def _open_wb(path: Path):
    """Open workbook, converting Windows PermissionError to ExcelLockedError."""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        raise ExcelLockedError(
            f"Cannot open '{path.name}' — it is locked by Excel.\n"
            "Close the file in Excel, then click ↻ Reload."
        )


def _open_wb_edit(path: Path):
    try:
        return openpyxl.load_workbook(path)
    except PermissionError:
        raise ExcelLockedError(
            f"Cannot write '{path.name}' — it is locked by Excel.\n"
            "Close the file in Excel, then retry."
        )


# ── Sheet parsing ─────────────────────────────────────────────────────────────

def _sheet_to_dicts(ws) -> list[dict]:
    """
    Parse a worksheet. Native types (datetime, int, float) preserved.
    None cells → "". Strings stripped. ic_number auto-normalized.
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [
        str(h).strip().lower().replace(" ", "_") if h else f"col_{i}"
        for i, h in enumerate(rows[0])
    ]
    result = []
    for row in rows[1:]:
        if not any(row):
            continue
        record = {}
        for k, v in zip(headers, row):
            if v is None:
                record[k] = ""
            elif isinstance(v, str):
                record[k] = v.strip()
            else:
                record[k] = v  # int, float, datetime preserved for formatters
        if "ic_number" in record:
            record["ic_number"] = normalize_ic(record["ic_number"])
        result.append(record)
    return result


# ── Generic customer workbook parsing ────────────────────────────────────────

_KEY_HEADERS = {
    "cis", "cis_no", "cis_number", "cif_no", "cif_number",
    "name", "client_name", "customer_name",
    "ic", "ic_number", "nric", "new_ic",
    "policy", "policy_no", "policy_number",
}

_CIS_ALIASES = ("cis", "cis_no", "cis_number", "cif_no", "cif_number")
_NAME_ALIASES = ("name", "client_name", "customer_name", "customer")
_IC_ALIASES = ("ic_number", "ic", "nric", "new_ic")
_POLICY_ALIASES = ("policy_number", "policy_no", "policy")


def normalize_header(raw) -> str:
    text = str(raw or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def normalize_lookup_key(raw) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "", str(raw or "")).upper()


def _detect_header_row(ws, max_scan: int = 20) -> int | None:
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        headers = [normalize_header(v) for v in row if v not in (None, "")]
        if len(headers) < 2:
            continue
        if any(h in _KEY_HEADERS for h in headers):
            return idx
    return None


def _canonicalize_record(record: dict) -> dict:
    out = dict(record)

    for key in _CIS_ALIASES:
        if out.get(key) not in ("", None):
            out.setdefault("cis", str(out[key]).strip())
            if key == "cif_no":
                out.setdefault("cif_no", str(out[key]).strip())
            break

    for key in _NAME_ALIASES:
        if out.get(key) not in ("", None):
            out.setdefault("name", str(out[key]).strip())
            out.setdefault("client_name", str(out[key]).strip())
            break

    for key in _IC_ALIASES:
        if out.get(key) not in ("", None):
            out.setdefault("ic_number", normalize_ic(out[key]))
            break

    for key in _POLICY_ALIASES:
        if out.get(key) not in ("", None):
            out.setdefault("policy_number", str(out[key]).strip())
            break

    return out


def _sheet_records_generic(ws) -> list[dict]:
    header_row = _detect_header_row(ws)
    if not header_row:
        return []
    headers = [
        normalize_header(v) if v not in (None, "") else f"col_{i}"
        for i, v in enumerate(next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)))
    ]
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(row):
            continue
        rec = {}
        for key, value in zip(headers, row):
            if value is None:
                rec[key] = ""
            elif isinstance(value, str):
                rec[key] = value.strip()
            else:
                rec[key] = value
        rows.append(_canonicalize_record(rec))
    return rows


def load_customer_records(path: Path) -> list[dict]:
    """
    Read a saved Excel workbook and merge customer rows across sheets.

    Matching priority is CIS, then IC, then policy number. Formula cells are read
    as their saved/calculated values through openpyxl data_only mode.
    """
    if not path.exists():
        return []
    wb = _open_wb(path)
    merged: dict[str, dict] = {}
    for sheet_name in wb.sheetnames:
        if sheet_name.lower() == "rm_profile":
            continue
        ws = wb[sheet_name]
        for rec in _sheet_records_generic(ws):
            key = ""
            if rec.get("cis"):
                key = "CIS:" + normalize_lookup_key(rec["cis"])
            elif rec.get("ic_number"):
                key = "IC:" + normalize_ic(rec["ic_number"])
            elif rec.get("policy_number"):
                key = "POL:" + normalize_lookup_key(rec["policy_number"])
            if not key:
                key = f"{sheet_name}:{len(merged) + 1}"
            existing = merged.setdefault(key, {})
            existing.update({k: v for k, v in rec.items() if v not in ("", None)})
            sheets = set(str(existing.get("_sheets", "")).split("|")) if existing.get("_sheets") else set()
            sheets.add(sheet_name)
            existing["_sheets"] = "|".join(sorted(sheets))
    wb.close()
    return sorted(merged.values(), key=lambda r: str(r.get("name") or r.get("cis") or ""))


def search_customers(records: list[dict], query: str, limit: int = 80) -> list[dict]:
    q = str(query or "").strip().lower()
    if not q:
        return records[:limit]
    q_key = normalize_lookup_key(q)
    out = []
    for rec in records:
        haystacks = [
            str(rec.get("name", "")).lower(),
            str(rec.get("client_name", "")).lower(),
            normalize_lookup_key(rec.get("cis", "")),
            normalize_lookup_key(rec.get("cif_no", "")),
            normalize_lookup_key(rec.get("ic_number", "")),
            normalize_lookup_key(rec.get("policy_number", "")),
        ]
        if any(q in h for h in haystacks[:2]) or any(q_key and q_key in h for h in haystacks[2:]):
            out.append(rec)
        if len(out) >= limit:
            break
    return out


def find_customer_by_cis(records: list[dict], cis: str) -> dict | None:
    target = normalize_lookup_key(cis)
    if not target:
        return None
    for rec in records:
        if normalize_lookup_key(rec.get("cis", "")) == target:
            return rec
        if normalize_lookup_key(rec.get("cif_no", "")) == target:
            return rec
    return None


def read_cis_list(path: Path) -> list[str]:
    if not path.exists():
        return []
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = _open_wb(path)
        ws = wb[wb.sheetnames[0]]
        vals = []
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell not in ("", None):
                    vals.append(str(cell).strip())
                    break
        wb.close()
        if vals and normalize_header(vals[0]) in {"cis", "cis_no", "cis_number"}:
            vals = vals[1:]
        return [v for v in vals if v]
    text = path.read_text(encoding="utf-8-sig")
    return [v.strip() for v in re.split(r"[\n,;]+", text) if v.strip()]


HISTORY_COLUMNS = [
    "GeneratedDateTime",
    "ClientName",
    "CIS",
    "ProcedureCode",
    "ProcedureName",
    "OutputFilePath",
    "Amount",
    "FDDetails",
    "ProductType",
    "ActionPurpose",
    "FollowUpNote",
    "Status",
    "ErrorMessage",
    "GeneratedBy",
]


def append_history_rows(path: Path, rows: list[dict]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        wb = _open_wb_edit(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "History"
        ws.append(HISTORY_COLUMNS)

    existing = [
        str(c.value).strip() if c.value else ""
        for c in ws[1]
    ]
    if existing[:len(HISTORY_COLUMNS)] != HISTORY_COLUMNS:
        ws.insert_rows(1)
        for idx, col in enumerate(HISTORY_COLUMNS, 1):
            ws.cell(row=1, column=idx, value=col)

    for row in rows:
        ws.append([row.get(col, "") for col in HISTORY_COLUMNS])
    wb.save(path)
    wb.close()


# ── Public loaders ────────────────────────────────────────────────────────────

def load_master(path: Path) -> dict[str, dict]:
    """Returns {normalized_ic: client_dict} from Master sheet."""
    wb = _open_wb(path)
    clients = _sheet_to_dicts(wb["Master"])
    wb.close()
    return {c["ic_number"]: c for c in clients if c.get("ic_number")}


def get_sheet_headers(path: Path, sheet_name: str) -> list[str]:
    """Column names from a batch sheet, excluding ic_number."""
    wb = _open_wb(path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[sheet_name]
    first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    wb.close()
    return [
        str(h).strip().lower().replace(" ", "_")
        for h in first_row
        if h and str(h).strip().lower().replace(" ", "_") != "ic_number"
    ]


def load_batch(path: Path, sheet_name: str,
               master: dict[str, dict]) -> list[dict]:
    """Master + batch data merged by ic_number. Missing Master rows skipped."""
    wb = _open_wb(path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}"
        )
    rows = _sheet_to_dicts(wb[sheet_name])
    wb.close()
    merged = []
    for row in rows:
        ic = row.get("ic_number", "")
        master_rec = master.get(ic)
        if not master_rec:
            continue
        merged.append({**master_rec, **row})
    return merged


def find_client_in_batch(path: Path, sheet_name: str, ic_number: str) -> dict:
    """Return batch row for one client (normalized IC), or {} if not found."""
    ic_norm = normalize_ic(ic_number)
    try:
        wb = _open_wb(path)
    except ExcelLockedError:
        return {}
    if sheet_name not in wb.sheetnames:
        wb.close()
        return {}
    rows = _sheet_to_dicts(wb[sheet_name])
    wb.close()
    for row in rows:
        if row.get("ic_number") == ic_norm:
            return row
    return {}


def get_master_names(master: dict[str, dict]) -> list[str]:
    return sorted(c.get("name", "") for c in master.values() if c.get("name"))


def sheet_names(path: Path) -> list[str]:
    wb = openpyxl.load_workbook(path, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


# ── RM profile ────────────────────────────────────────────────────────────────

_RM_PROFILE_KEYS = {"rm_name", "staff_id", "fimm_id", "ippc_id", "branches"}


def load_rm_profile(path: Path) -> dict:
    """
    Read RM_Profile sheet (key-value: column A=field, column B=value).
    Returns {rm_name, staff_id, fimm_id, ippc_id, branches: list[str]}.
    Missing sheet → all empty. Missing keys → empty string / empty list.
    'branches' string is split on commas and trimmed.
    """
    profile = {k: "" for k in _RM_PROFILE_KEYS}
    profile["branches"] = []
    if not path.exists():
        return profile

    wb = _open_wb(path)
    if "RM_Profile" not in wb.sheetnames:
        wb.close()
        return profile

    ws = wb["RM_Profile"]
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip().lower().replace(" ", "_")
        if key not in _RM_PROFILE_KEYS:
            continue
        val = "" if len(row) < 2 or row[1] is None else row[1]
        if isinstance(val, float) and val.is_integer():
            val = str(int(val))
        else:
            val = str(val).strip()
        if key == "branches":
            profile["branches"] = [b.strip() for b in val.split(",") if b.strip()]
        else:
            profile[key] = val
    wb.close()
    return profile
