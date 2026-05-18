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


def output_folder_path(settings: dict) -> Path:
    return resolve_path(settings.get("output_folder", "filled"))


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
    info = {"forms_created": False, "output_created": False, "forms_empty": False}

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

    return info


def ensure_clients_xlsx() -> dict:
    """If data/clients.xlsx missing, copy data/clients_template.xlsx to it.
    Returns dict with keys: copied, source_missing."""
    info = {"copied": False, "source_missing": False}
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

def load_forms() -> dict:
    with open(_config_path("forms.json"), encoding="utf-8") as f:
        return json.load(f)


def save_forms(data: dict):
    """P3-11: Back up before overwriting."""
    forms_path = _config_path("forms.json")
    bak_path   = _config_path("forms.json.bak")
    if forms_path.exists():
        shutil.copy(forms_path, bak_path)
    with open(forms_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── Applications ──────────────────────────────────────────────────────────────

def load_applications() -> list:
    with open(_config_path("applications.json"), encoding="utf-8") as f:
        return json.load(f)


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


def log_path() -> Path:
    return _base_dir() / "data" / "fill_log.csv"


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
