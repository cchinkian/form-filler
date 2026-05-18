"""
Client master DB — SQLite store for client demographics + lifecycle.

Replaces the old Master sheet in clients.xlsx as the canonical source for:
  ic_number (PK), passport_number, name, cif_no, phone, email, address, dob, etc.

Two-tier lifecycle (Phase 1A — manual only; Phase 1B will add monthly bank import):
  permanent = 1 → row never auto-removed by future monthly imports
  active = 0    → soft-deleted (hidden from search), history preserved
  active = 1    → live

Field-source design (forward-compat for Phase 1B):
  bank_fields TEXT (JSON) holds any extra columns from a future monthly import.
  User-entered fields (phone, address, etc.) live in dedicated columns and
  are never overwritten by the import.

Search:
  by_ic(full_ic)        — full 12-digit IC, exact match (normalized)
  by_name(substring)    — case-insensitive partial match
  list_all(active_only) — for browse/admin

CRUD:
  add, update, soft_delete, hard_delete, restore, set_permanent
"""
from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path

DB_FILENAME = "client_db.db"


# ── Path resolution (matches config_loader._base_dir) ────────────────────────

def _base_dir() -> Path:
    import sys
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def db_path() -> Path:
    return _base_dir() / "data" / DB_FILENAME


# ── Connection ────────────────────────────────────────────────────────────────

def connect() -> sqlite3.Connection:
    """Returns a connection with foreign keys + row factory enabled."""
    p = db_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(p)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


# ── Schema ────────────────────────────────────────────────────────────────────

SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
    ic_number              TEXT PRIMARY KEY,
    passport_number        TEXT,
    name                   TEXT NOT NULL,
    cif_no                 TEXT,

    -- User-entered sticky fields (form-required; never touched by future imports)
    phone                  TEXT,
    email                  TEXT,
    address_line1          TEXT,
    address_line2          TEXT,
    city                   TEXT,
    state                  TEXT,
    postcode               TEXT,
    dob                    TEXT,
    occupation             TEXT,
    notes                  TEXT,

    -- Free-form bank fields from future monthly import (JSON)
    bank_fields            TEXT,

    -- Lifecycle flags
    permanent              INTEGER NOT NULL DEFAULT 0,
    active                 INTEGER NOT NULL DEFAULT 1,
    source                 TEXT NOT NULL DEFAULT 'manual',
    last_seen_in_import    TEXT,

    created_at             TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at             TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_clients_name   ON clients(name);
CREATE INDEX IF NOT EXISTS idx_clients_cif    ON clients(cif_no);
CREATE INDEX IF NOT EXISTS idx_clients_active ON clients(active);
"""


def init_db() -> None:
    """Create tables + indexes if not exist."""
    with connect() as conn:
        conn.executescript(SCHEMA)
        _ensure_columns(conn)


def _ensure_columns(conn: sqlite3.Connection) -> None:
    """Lightweight migrations for existing pen-drive databases."""
    existing = {
        row["name"]
        for row in conn.execute("PRAGMA table_info(clients)").fetchall()
    }
    if "passport_number" not in existing:
        conn.execute("ALTER TABLE clients ADD COLUMN passport_number TEXT")
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_clients_passport ON clients(passport_number)"
    )


# ── IC normalization (mirrors excel_reader.normalize_ic) ──────────────────────

def normalize_ic(raw) -> str:
    """Strip non-digits, zero-pad to 12. Empty → ''."""
    if isinstance(raw, float):
        raw = int(raw)
    cleaned = re.sub(r"[^0-9]", "", str(raw or ""))
    return cleaned.zfill(12) if cleaned else ""


# ── Public column list (used for forms.json data resolution) ──────────────────

USER_FIELDS = (
    "ic_number", "passport_number", "name", "cif_no",
    "phone", "email",
    "address_line1", "address_line2", "city", "state", "postcode",
    "dob", "occupation", "notes",
)


# ── CRUD ──────────────────────────────────────────────────────────────────────

def add(client: dict) -> str:
    """Insert a new client. Returns the normalized ic_number.
    Raises ValueError if IC missing/duplicate."""
    ic = normalize_ic(client.get("ic_number"))
    if not ic:
        raise ValueError("IC number is required.")
    name = (client.get("name") or "").strip()
    if not name:
        raise ValueError("Name is required.")

    init_db()
    with connect() as conn:
        existing = conn.execute(
            "SELECT 1 FROM clients WHERE ic_number = ?", (ic,)).fetchone()
        if existing:
            raise ValueError(f"Client with IC {ic} already exists.")
        bank_fields = client.get("bank_fields")
        if bank_fields and not isinstance(bank_fields, str):
            bank_fields = json.dumps(bank_fields, ensure_ascii=False)

        conn.execute("""
            INSERT INTO clients (
                ic_number, name, cif_no,
                passport_number,
                phone, email,
                address_line1, address_line2, city, state, postcode,
                dob, occupation, notes,
                bank_fields, permanent, active, source
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            ic, name, client.get("cif_no"),
            client.get("passport_number"),
            client.get("phone"), client.get("email"),
            client.get("address_line1"), client.get("address_line2"),
            client.get("city"), client.get("state"), client.get("postcode"),
            client.get("dob"), client.get("occupation"), client.get("notes"),
            bank_fields,
            1 if client.get("permanent") else 0,
            1 if client.get("active", 1) else 0,
            client.get("source", "manual"),
        ))
    return ic


def update(ic: str, fields: dict) -> int:
    """Update specific columns by IC. Returns rowcount.
    Only USER_FIELDS + permanent + active + notes can be updated this way.
    bank_fields is updated separately via set_bank_fields()."""
    ic = normalize_ic(ic)
    allowed = set(USER_FIELDS) | {"permanent", "active", "name", "cif_no"}
    sets = []
    params = []
    for k, v in fields.items():
        if k not in allowed or k == "ic_number":
            continue
        sets.append(f"{k} = ?")
        params.append(v)
    if not sets:
        return 0
    sets.append("updated_at = datetime('now')")
    params.append(ic)
    init_db()
    with connect() as conn:
        cur = conn.execute(
            f"UPDATE clients SET {', '.join(sets)} WHERE ic_number = ?",
            params)
        return cur.rowcount


def soft_delete(ic: str) -> int:
    return update(ic, {"active": 0})


def hard_delete(ic: str) -> int:
    """Permanently remove a client row. Caller is responsible for confirmation."""
    ic = normalize_ic(ic)
    init_db()
    with connect() as conn:
        cur = conn.execute("DELETE FROM clients WHERE ic_number = ?", (ic,))
        return cur.rowcount


def restore(ic: str) -> int:
    return update(ic, {"active": 1})


def set_permanent(ic: str, value: bool = True) -> int:
    return update(ic, {"permanent": 1 if value else 0})


# ── Search / fetch ────────────────────────────────────────────────────────────

def by_ic(ic: str, include_inactive: bool = False) -> dict | None:
    ic = normalize_ic(ic)
    if not ic:
        return None
    init_db()
    with connect() as conn:
        clause = "" if include_inactive else " AND active = 1"
        row = conn.execute(
            f"SELECT * FROM clients WHERE ic_number = ?{clause}",
            (ic,)).fetchone()
        return _row_to_dict(row) if row else None


def by_name(substring: str, limit: int = 30,
            include_inactive: bool = False) -> list[dict]:
    sub = (substring or "").strip()
    if not sub:
        return []
    init_db()
    with connect() as conn:
        clause = "" if include_inactive else " AND active = 1"
        rows = conn.execute(
            f"""SELECT * FROM clients
                WHERE name LIKE ? COLLATE NOCASE{clause}
                ORDER BY name COLLATE NOCASE
                LIMIT ?""",
            (f"%{sub}%", limit)).fetchall()
        return [_row_to_dict(r) for r in rows]


def by_cif(cif_no: str, include_inactive: bool = False) -> dict | None:
    cif = (cif_no or "").strip()
    if not cif:
        return None
    init_db()
    with connect() as conn:
        clause = "" if include_inactive else " AND active = 1"
        row = conn.execute(
            f"SELECT * FROM clients WHERE cif_no = ?{clause}",
            (cif,)).fetchone()
        return _row_to_dict(row) if row else None


def by_passport(passport_number: str, include_inactive: bool = False) -> dict | None:
    passport = (passport_number or "").strip()
    if not passport:
        return None
    init_db()
    with connect() as conn:
        clause = "" if include_inactive else " AND active = 1"
        row = conn.execute(
            f"""SELECT * FROM clients
                WHERE passport_number = ? COLLATE NOCASE{clause}""",
            (passport,)).fetchone()
        return _row_to_dict(row) if row else None


def search(query: str, limit: int = 50,
           include_inactive: bool = False) -> list[dict]:
    """Search by full/partial name, IC, passport, or CIF."""
    q = (query or "").strip()
    if not q:
        return list_all(active_only=not include_inactive, limit=limit)
    digits = re.sub(r"[^0-9]", "", q)
    init_db()
    with connect() as conn:
        clause = "" if include_inactive else " AND active = 1"
        rows = conn.execute(
            f"""SELECT * FROM clients
                WHERE (
                    name LIKE ? COLLATE NOCASE
                    OR ic_number LIKE ?
                    OR passport_number LIKE ? COLLATE NOCASE
                    OR cif_no LIKE ? COLLATE NOCASE
                ){clause}
                ORDER BY
                    CASE
                        WHEN ic_number = ? THEN 0
                        WHEN passport_number = ? COLLATE NOCASE THEN 1
                        WHEN cif_no = ? COLLATE NOCASE THEN 2
                        ELSE 3
                    END,
                    name COLLATE NOCASE
                LIMIT ?""",
            (f"%{q}%", f"%{digits or q}%", f"%{q}%", f"%{q}%",
             digits or q, q, q, limit)).fetchall()
        return [_row_to_dict(r) for r in rows]


def list_all(active_only: bool = True, limit: int = 1000) -> list[dict]:
    init_db()
    with connect() as conn:
        clause = "WHERE active = 1" if active_only else ""
        rows = conn.execute(
            f"""SELECT * FROM clients {clause}
                ORDER BY name COLLATE NOCASE
                LIMIT ?""",
            (limit,)).fetchall()
        return [_row_to_dict(r) for r in rows]


def count() -> dict:
    init_db()
    with connect() as conn:
        active   = conn.execute(
            "SELECT COUNT(*) FROM clients WHERE active = 1").fetchone()[0]
        inactive = conn.execute(
            "SELECT COUNT(*) FROM clients WHERE active = 0").fetchone()[0]
        permanent = conn.execute(
            "SELECT COUNT(*) FROM clients WHERE permanent = 1").fetchone()[0]
    return {"active": active, "inactive": inactive, "permanent": permanent}


def _row_to_dict(row: sqlite3.Row) -> dict:
    d = dict(row)
    # Decode bank_fields JSON for easy consumption
    bf = d.get("bank_fields")
    if bf:
        try:
            d["bank_fields"] = json.loads(bf)
        except (json.JSONDecodeError, TypeError):
            pass
    return d


# ── Upsert (used by future monthly import) ────────────────────────────────────

def upsert_from_import(client: dict, import_date: str) -> str:
    """
    Insert or refresh bank-managed fields. NEVER touches user-entered fields
    (phone, email, address_*, dob, occupation, notes).
    Used by Phase 1B monthly bank Excel import.
    """
    ic = normalize_ic(client.get("ic_number"))
    if not ic:
        raise ValueError("Import row missing IC.")
    name = (client.get("name") or "").strip()
    if not name:
        raise ValueError(f"Import row {ic} missing name.")

    bank_fields = client.get("bank_fields") or {}
    if isinstance(bank_fields, dict):
        bank_fields_json = json.dumps(bank_fields, ensure_ascii=False)
    else:
        bank_fields_json = str(bank_fields)

    init_db()
    with connect() as conn:
        existing = conn.execute(
            "SELECT ic_number FROM clients WHERE ic_number = ?",
            (ic,)).fetchone()
        if existing:
            conn.execute("""
                UPDATE clients SET
                    name = ?, cif_no = ?, passport_number = ?, bank_fields = ?,
                    active = 1, last_seen_in_import = ?, updated_at = datetime('now')
                WHERE ic_number = ?
            """, (name, client.get("cif_no"), client.get("passport_number"), bank_fields_json,
                  import_date, ic))
        else:
            conn.execute("""
                INSERT INTO clients (
                    ic_number, name, cif_no, passport_number, bank_fields,
                    active, source, last_seen_in_import
                ) VALUES (?, ?, ?, ?, ?, 1, 'monthly_import', ?)
            """, (ic, name, client.get("cif_no"),
                  client.get("passport_number"), bank_fields_json, import_date))
    return ic


def soft_delete_dropped(import_date: str) -> int:
    """After a monthly import, soft-delete clients NOT seen in this import
    AND not flagged permanent. Returns rowcount.
    Phase 1B helper — Phase 1A doesn't call this."""
    init_db()
    with connect() as conn:
        cur = conn.execute("""
            UPDATE clients SET active = 0, updated_at = datetime('now')
            WHERE active = 1
              AND permanent = 0
              AND (last_seen_in_import IS NULL OR last_seen_in_import != ?)
        """, (import_date,))
        return cur.rowcount
