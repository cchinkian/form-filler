"""
Config loader — pen-drive portable path resolution.

Path policy:
  - forms_folder / output_folder are RELATIVE to the EXE folder unless absolute.
    Default settings: forms_folder="forms", output_folder="filled" → siblings of FormFiller.exe.
  - Missing folders are auto-created on launch (no crash).
  - On every launch, settings.json is backed up to config/backups/settings_YYYYMMDD_HHMMSS.json.
    Last 100 backups retained.

P3-10: find_template() uses the single PDF in each form folder.
P3-11: save_forms() backs up forms.json before overwriting.
P1-3:  health_check() verifies forms_folder + mapped subfolders.
"""
import datetime
import hashlib
import json
import shutil
import sys
from pathlib import Path


_BACKUP_RETENTION = 100
FOLDER_MAPPING_FILENAME = "mapping.json"


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def _config_path(filename: str) -> Path:
    return _base_dir() / "config" / filename


def resolve_path(p: str) -> Path:
    """Resolve a settings path: absolute → as-is, relative → under _base_dir()."""
    if not p:
        return _base_dir()
    path = Path(p)
    return path if path.is_absolute() else _base_dir() / path


def forms_folder_path(settings: dict) -> Path:
    return resolve_path(settings.get("forms_folder", "forms"))


def forms_folder_paths(settings: dict) -> list[Path]:
    folders = settings.get("forms_folders") or []
    if isinstance(folders, str):
        folders = [folders]
    paths = []
    for raw in folders:
        if raw:
            path = resolve_path(str(raw))
            if path not in paths:
                paths.append(path)
    primary = forms_folder_path(settings)
    if primary not in paths:
        paths.insert(0, primary)
    return paths


def output_folder_path(settings: dict) -> Path:
    return resolve_path(settings.get("output_folder", "Output"))


def customer_workbook_path(settings: dict) -> Path:
    return resolve_path(settings.get("customer_workbook", "data/clients.xlsx"))


def history_log_path(settings: dict) -> Path:
    return resolve_path(settings.get("history_log_path", "data/HistoryLog.xlsx"))


# ── Settings ──────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    with open(_config_path("settings.json"), encoding="utf-8") as f:
        return json.load(f)


def save_settings(data: dict):
    _backup_settings(reason="save")
    with open(_config_path("settings.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _backup_settings(reason: str = "launch"):
    """Copy current settings.json to config/backups/settings_YYYYMMDD_HHMMSS.json.
    No-op if settings.json doesn't exist. Prunes oldest if over _BACKUP_RETENTION."""
    src = _config_path("settings.json")
    if not src.exists():
        return
    backup_dir = _base_dir() / "config" / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = backup_dir / f"settings_{stamp}.json"
    if not dst.exists():
        shutil.copy(src, dst)
    # Prune
    backups = sorted(backup_dir.glob("settings_*.json"))
    while len(backups) > _BACKUP_RETENTION:
        backups[0].unlink(missing_ok=True)
        backups.pop(0)


def backup_settings_on_launch():
    """Public entry point — main_app calls this once at startup."""
    _backup_settings(reason="launch")


# ── First-run scaffolding ─────────────────────────────────────────────────────

def ensure_runtime_dirs(settings: dict) -> dict:
    """Create forms_folder + output_folder if missing.
    Returns dict with keys: forms_created, output_created, forms_empty."""
    info = {
        "forms_created": False,
        "output_created": False,
        "forms_empty": False,
        "data_created": False,
    }

    ff = forms_folder_path(settings)
    if not ff.exists():
        ff.mkdir(parents=True, exist_ok=True)
        info["forms_created"] = True
    if not any(p.is_dir() for p in ff.iterdir()):
        info["forms_empty"] = True

    of = output_folder_path(settings)
    if not of.exists():
        of.mkdir(parents=True, exist_ok=True)
        info["output_created"] = True

    data_dir = customer_workbook_path(settings).parent
    if not data_dir.exists():
        data_dir.mkdir(parents=True, exist_ok=True)
        info["data_created"] = True

    return info


def ensure_clients_xlsx() -> dict:
    """If customer workbook missing, copy data/clients_template.xlsx to it.
    Returns dict with keys: copied, source_missing."""
    info = {"copied": False, "source_missing": False}
    try:
        settings = load_settings()
        target = customer_workbook_path(settings)
    except Exception:
        target = data_path("clients.xlsx")
    if target.exists():
        return info
    template = data_path("clients_template.xlsx")
    if not template.exists():
        info["source_missing"] = True
        return info
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(template, target)
    info["copied"] = True
    return info


# ── Forms ─────────────────────────────────────────────────────────────────────

def _mapping_entry_from_folder_json(folder: Path, data: dict) -> tuple[str, dict] | None:
    if not isinstance(data, dict):
        return None
    mapping_key = (
        data.get("mapping_key")
        or data.get("MappingKey")
        or data.get("form_id")
        or data.get("id")
        or folder.name
    )
    mapping_key = str(mapping_key or "").strip()
    if not mapping_key:
        return None
    entry = {
        "name": data.get("display_name") or data.get("name") or data.get("DisplayName") or mapping_key,
        "template_subfolder": data.get("template_subfolder") or folder.name,
        "template_filename": data.get("template_filename") or data.get("pdf_file") or "",
        "template_hash": data.get("template_hash", ""),
        "last_updated": data.get("last_updated", ""),
        "fields": data.get("fields", []),
    }
    return mapping_key, entry


def load_folder_mapping(folder: Path) -> tuple[str, dict] | None:
    path = folder / FOLDER_MAPPING_FILENAME
    if not path.exists():
        return None
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return _mapping_entry_from_folder_json(folder, data)


def load_folder_mappings(settings: dict | None = None) -> dict:
    try:
        settings = settings or load_settings()
        root = forms_folder_path(settings)
    except Exception:
        return {}
    if not root.exists():
        return {}
    mappings = {}
    for folder in sorted(p for p in root.iterdir() if p.is_dir()):
        try:
            loaded = load_folder_mapping(folder)
        except Exception:
            loaded = None
        if loaded:
            key, entry = loaded
            mappings[key] = entry
    return mappings


def load_forms() -> dict:
    with open(_config_path("forms.json"), encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        data = {}
    merged = dict(data)
    for key, entry in load_folder_mappings().items():
        merged[key] = entry
    return merged


def save_forms(data: dict):
    """P3-11: Back up before overwriting."""
    forms_path = _config_path("forms.json")
    bak_path   = _config_path("forms.json.bak")
    if forms_path.exists():
        shutil.copy(forms_path, bak_path)
    with open(forms_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def save_folder_mapping(folder: Path, mapping_key: str, form_config: dict) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / FOLDER_MAPPING_FILENAME
    pdf_file = form_config.get("template_filename", "")
    payload = {
        "mapping_key": mapping_key,
        "display_name": form_config.get("name") or mapping_key,
        "template_subfolder": form_config.get("template_subfolder") or folder.name,
        "pdf_file": pdf_file,
        "template_hash": form_config.get("template_hash", ""),
        "last_updated": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "fields": form_config.get("fields", []),
    }
    if path.exists():
        shutil.copy(path, path.with_suffix(path.suffix + ".bak"))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    return path


# ── Template PDF ──────────────────────────────────────────────────────────────

def _md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def scan_forms_folder(settings: dict) -> list[str]:
    """
    Auto-detect: list all subdirectory names inside forms_folder.
    Returns sorted list of folder names (not full paths).
    Returns [] if forms_folder doesn't exist.
    """
    forms_folder = forms_folder_path(settings)
    if not forms_folder.exists():
        return []
    return sorted(p.name for p in forms_folder.iterdir() if p.is_dir())


def scan_all_form_folders(settings: dict) -> list[dict]:
    rows = []
    for root in forms_folder_paths(settings):
        if not root.exists():
            continue
        for folder in sorted(p for p in root.iterdir() if p.is_dir()):
            rows.append({
                "label": f"{folder.name}  [{root}]",
                "folder": folder.name,
                "path": str(folder),
                "root": str(root),
            })
    return rows


def scan_form_subfolders(settings: dict, root: Path | None = None) -> list[dict]:
    """Inspect immediate form subfolders for direct PDF and mapping.json status."""
    forms_folder = root or forms_folder_path(settings)
    if not forms_folder.exists():
        return []
    rows = []
    for folder in sorted(p for p in forms_folder.iterdir() if p.is_dir()):
        pdfs = sorted(p.name for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")
        old_forms = folder / "old forms"
        nested_pdfs = sorted(
            str(p.relative_to(folder))
            for p in folder.rglob("*.pdf")
            if p.is_file() and p.parent != folder
        )
        mapping_path = folder / FOLDER_MAPPING_FILENAME
        updated = ""
        mapping_key = ""
        field_count = 0
        if mapping_path.exists():
            try:
                with open(mapping_path, encoding="utf-8") as f:
                    mapping = json.load(f)
                mapping_key = str(mapping.get("mapping_key") or mapping.get("MappingKey") or folder.name)
                updated = str(mapping.get("last_updated") or "")
                field_count = len(mapping.get("fields", []) or [])
            except Exception:
                updated = "unreadable"
        rows.append({
            "folder": folder.name,
            "path": str(folder),
            "pdf_count": len(pdfs),
            "pdf_files": pdfs,
            "old_forms_exists": old_forms.exists() and old_forms.is_dir(),
            "nested_pdf_count": len(nested_pdfs),
            "nested_pdf_files": nested_pdfs,
            "mapping_exists": mapping_path.exists(),
            "mapping_path": str(mapping_path),
            "mapping_key": mapping_key,
            "mapping_updated": updated,
            "field_count": field_count,
        })
    return rows


def find_template(settings: dict, template_subfolder: str,
                  form_config: dict | None = None,
                  test_mode: bool = False) -> Path:
    """
    Find the single PDF at the top level of a form subfolder.

    Compliance guard:
      - Zero PDFs → FileNotFoundError
      - >1 PDF → ValueError (must keep exactly one at top level)

    Filename policy:
      - The PDF filename is intentionally not part of routing.
      - If the folder has exactly one PDF, that PDF is the template.
      - This lets users rename the merged PDF freely without reconfiguring.
    """
    forms_folder = forms_folder_path(settings)
    subfolder    = forms_folder / template_subfolder
    if not subfolder.exists():
        raise FileNotFoundError(
            f"Template subfolder not found:\n  {subfolder}\n"
            "Choose an existing form folder or click refresh after creating it."
        )
    if not subfolder.is_dir():
        raise NotADirectoryError(
            f"Template subfolder is not a folder:\n  {subfolder}\n"
            "Choose a folder that contains exactly one top-level PDF."
        )
    pdfs = [p for p in subfolder.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"]

    if not pdfs:
        raise FileNotFoundError(
            f"No PDF found in:\n  {subfolder}\n"
            f"Place the blank form PDF in {forms_folder}\\{template_subfolder}\\"
        )
    if len(pdfs) > 1:
        names = ", ".join(p.name for p in sorted(pdfs))
        raise ValueError(
            f"Multiple PDFs in {subfolder}:\n  {names}\n"
            "Keep exactly ONE PDF at the top level. "
            "Move old versions into a subfolder (e.g. 'old forms\\')."
        )

    pdf_path = pdfs[0]

    return pdf_path


def compute_template_hash(pdf_path: Path) -> str:
    return _md5(pdf_path)


# ── Health check ──────────────────────────────────────────────────────────────

def health_check(settings: dict, forms: dict) -> list[dict]:
    """
    Verify forms_folder + each mapped form subfolder.
    Statuses: ok | warn | error | empty | unmapped
    Also reports discovered subfolders not yet in forms.json (unmapped).
    """
    results = []
    forms_folder = forms_folder_path(settings)

    if not forms_folder.exists():
        return [{"name": "forms_folder", "status": "error",
                 "message": f"Folder not found: {forms_folder}"}]

    # Empty forms folder — friendly first-run message.
    # Only subdirectories count as forms; stray loose files are ignored
    # (they're not pickable as forms anyway).
    if not any(p.is_dir() for p in forms_folder.iterdir()):
        return [{"name": "forms_folder", "status": "empty",
                 "message": (f"{forms_folder} has no form subfolders. "
                             "Create a subfolder per form and drop the blank PDF inside, "
                             "then click ↻ Reload.")}]

    # Check each registered form
    registered_subfolders = set()
    for form_id, form_cfg in forms.items():
        if form_id.startswith("_"):
            continue
        subfolder = form_cfg.get("template_subfolder", "")
        if not subfolder:
            continue
        registered_subfolders.add(subfolder)
        path = forms_folder / subfolder

        if not path.exists():
            results.append({"name": form_id, "status": "error",
                            "message": f"Subfolder missing: {subfolder}"})
            continue

        pdfs = [p for p in path.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"]

        if not pdfs:
            results.append({"name": form_id, "status": "warn",
                            "message": f"No PDF in {subfolder}\\"})
        elif len(pdfs) > 1:
            names = ", ".join(p.name for p in sorted(pdfs))
            results.append({"name": form_id, "status": "warn",
                            "message": f"Multiple PDFs: {names}"})
        else:
            results.append({"name": form_id, "status": "ok",
                            "message": pdfs[0].name})

    # Report unregistered subfolders (discovered but not in forms.json)
    for sub in scan_forms_folder(settings):
        if sub not in registered_subfolders:
            results.append({"name": f"[{sub}]", "status": "unmapped",
                            "message": f"Found in {forms_folder.name}/ but not in forms.json"})

    return results


# ── Data / state paths ────────────────────────────────────────────────────────

def data_path(filename: str) -> Path:
    return _base_dir() / "data" / filename


def state_path() -> Path:
    return _base_dir() / "data" / "state.json"


def load_state() -> dict:
    p = state_path()
    if p.exists():
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_state(data: dict):
    p = state_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f)
