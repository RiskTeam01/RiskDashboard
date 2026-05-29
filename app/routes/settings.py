from urllib.parse import quote

from fastapi import APIRouter, Form, Request
from fastapi.responses import RedirectResponse

from app.auth import current_user_or_none, add_user, delete_user, update_password, authenticate
from app.config import load_config, save_config, UPLOAD_DIR, OUTPUT_DIR, LOG_DIR, AUDIT_DIR
from app.cleanup import run_age_based_cleanup, force_clear_directory
from app.customers import delete_all_customers

router = APIRouter()


def settings_redirect(message: str, flash_type: str = "ok") -> RedirectResponse:
    return RedirectResponse(f"/settings?msg={quote(message)}&t={flash_type}", status_code=303)


@router.post("/settings/add-user")
def settings_add_user(request: Request, new_username: str = Form(...), new_password: str = Form(...)):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    ok, message = add_user(new_username, new_password, created_by=user)
    return settings_redirect(message, "ok" if ok else "err")


@router.post("/settings/delete-user")
def settings_delete_user(request: Request, username: str = Form(...)):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    ok, message = delete_user(username, current_user=user)
    return settings_redirect(message, "ok" if ok else "err")


@router.post("/settings/update-password")
def settings_update_password(
    request: Request,
    username: str = Form(...),
    new_password: str = Form(...),
):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    ok, message = update_password(username, new_password, actor=user)
    return settings_redirect(message, "ok" if ok else "err")


@router.post("/settings/limits")
def settings_limits(
    request: Request,
    max_pdf_size_mb: int = Form(...),
    max_batch_size: int = Form(...),
    session_timeout_minutes: int = Form(...),
):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    config = load_config()
    config["max_pdf_size_mb"] = max(1, min(int(max_pdf_size_mb), 500))
    config["max_batch_size"] = max(1, min(int(max_batch_size), 200))
    config["session_timeout_minutes"] = max(5, min(int(session_timeout_minutes), 1440))
    save_config(config)
    return settings_redirect("Upload limits saved.")


@router.post("/settings/retention")
def settings_retention(
    request: Request,
    cleanup_uploads_days: int = Form(...),
    cleanup_outputs_days: int = Form(...),
    cleanup_logs_days: int = Form(...),
):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    config = load_config()
    config["cleanup_uploads_days"] = max(0, min(int(cleanup_uploads_days), 365))
    config["cleanup_outputs_days"] = max(0, min(int(cleanup_outputs_days), 365))
    config["cleanup_logs_days"] = max(0, min(int(cleanup_logs_days), 365))
    save_config(config)
    return settings_redirect("Retention settings saved.")


@router.post("/settings/cleanup")
def settings_cleanup(request: Request, target: str = Form(...)):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    if target == "age":
        result = run_age_based_cleanup()
        total = sum(result.values())
        return settings_redirect(f"Age-based cleanup removed {total} file(s).")

    if target == "uploads":
        deleted = force_clear_directory(UPLOAD_DIR)
        return settings_redirect(f"Cleared {deleted} uploaded file(s).")

    if target == "logs":
        deleted = force_clear_directory(LOG_DIR)
        return settings_redirect(f"Cleared {deleted} log file(s).")

    if target == "outputs":
        deleted = force_clear_directory(OUTPUT_DIR) + force_clear_directory(AUDIT_DIR)
        return settings_redirect(f"Cleared {deleted} output/audit file(s).")

    return settings_redirect("Unknown cleanup target.", "err")


@router.post("/settings/delete-all-customers")
def settings_delete_all_customers(request: Request, password: str = Form(...)):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    # Require the signed-in user's own password to authorize this destructive action.
    if not authenticate(user, password):
        return settings_redirect("Incorrect password. No customer data was deleted.", "err")

    result = delete_all_customers()
    return settings_redirect(
        f"Deleted {result['customers']} customer account(s), "
        f"{result['reports']} report record(s), and "
        f"{result['net_capital_files']} Net Capital workbook(s)."
    )
