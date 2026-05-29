import json
from pathlib import Path
from typing import Optional

APP_ROOT = Path(__file__).resolve().parent.parent

TEMPLATE_DIR = APP_ROOT / "templates"
UPLOAD_DIR = APP_ROOT / "uploads"
OUTPUT_DIR = APP_ROOT / "outputs"
LOG_DIR = APP_ROOT / "logs"
ASSETS_DIR = APP_ROOT / "assets"
AUDIT_DIR = APP_ROOT / "audits"

USERS_FILE = APP_ROOT / "users.json"
CONFIG_FILE = APP_ROOT / "app_config.json"

SHEET_NAME = "Sheet1"
PREFERRED_TEMPLATE_FILENAME = "Buckler Excel Credit WS Template.xlsx"
SESSION_COOKIE = "pc_session"

DEFAULT_CONFIG = {
    "max_pdf_size_mb": 50,
    "max_batch_size": 20,
    "session_timeout_minutes": 60,
    "cleanup_uploads_days": 7,
    "cleanup_outputs_days": 30,
    "cleanup_logs_days": 30,
}


def ensure_directories():
    for folder in [TEMPLATE_DIR, UPLOAD_DIR, OUTPUT_DIR, LOG_DIR, ASSETS_DIR, AUDIT_DIR]:
        folder.mkdir(parents=True, exist_ok=True)


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            merged = dict(DEFAULT_CONFIG)
            merged.update({k: data[k] for k in data if k in DEFAULT_CONFIG})
            return merged
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(config: dict):
    clean = dict(DEFAULT_CONFIG)
    for key in DEFAULT_CONFIG:
        if key in config:
            clean[key] = config[key]
    CONFIG_FILE.write_text(json.dumps(clean, indent=2), encoding="utf-8")


def detect_backend_template() -> Path:
    preferred = TEMPLATE_DIR / PREFERRED_TEMPLATE_FILENAME
    if preferred.exists() and not preferred.name.startswith("~$"):
        return preferred
    candidates = [
        path for path in TEMPLATE_DIR.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No Excel template found in: {TEMPLATE_DIR}. "
            "Place your blank .xlsx template inside the templates folder."
        )
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def get_template_status() -> dict:
    try:
        template_path = detect_backend_template()
        return {
            "exists": True,
            "filename": template_path.name,
            "folder": str(TEMPLATE_DIR),
            "path": str(template_path),
            "error": "",
        }
    except Exception as e:
        return {
            "exists": False,
            "filename": "",
            "folder": str(TEMPLATE_DIR),
            "path": "",
            "error": str(e),
        }


def get_logo_path() -> Optional[Path]:
    preferred = ASSETS_DIR / "phillipcapital_logo.png"
    if preferred.exists():
        return preferred
    candidates = [
        path for path in ASSETS_DIR.glob("*.png")
        if not path.name.startswith("~$")
    ]
    if candidates:
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates[0]
    return None


def get_logo_url() -> str:
    logo_path = get_logo_path()
    if logo_path:
        return f"/assets/{logo_path.name}"
    return ""
