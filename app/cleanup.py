import time
from pathlib import Path
from typing import Optional

from app.config import UPLOAD_DIR, OUTPUT_DIR, LOG_DIR, AUDIT_DIR, load_config


def cleanup_directory(directory: Path, older_than_days: int, extensions: Optional[list[str]] = None) -> int:
    if older_than_days <= 0 or not directory.exists():
        return 0
    cutoff = time.time() - (older_than_days * 86400)
    deleted = 0
    for path in directory.glob("*"):
        if not path.is_file():
            continue
        if extensions and path.suffix.lower() not in extensions:
            continue
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink()
                deleted += 1
        except Exception:
            pass
    return deleted


def run_age_based_cleanup() -> dict:
    config = load_config()
    return {
        "uploads": cleanup_directory(UPLOAD_DIR, config["cleanup_uploads_days"]),
        "outputs": cleanup_directory(OUTPUT_DIR, config["cleanup_outputs_days"]),
        "audits": cleanup_directory(AUDIT_DIR, config["cleanup_outputs_days"]),
        "logs": cleanup_directory(LOG_DIR, config["cleanup_logs_days"]),
    }


def force_clear_directory(directory: Path, extensions: Optional[list[str]] = None) -> int:
    deleted = 0
    if not directory.exists():
        return 0
    for path in directory.glob("*"):
        if not path.is_file():
            continue
        if extensions and path.suffix.lower() not in extensions:
            continue
        try:
            path.unlink()
            deleted += 1
        except Exception:
            pass
    return deleted
