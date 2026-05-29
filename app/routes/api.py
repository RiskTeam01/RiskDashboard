from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse

from app.auth import require_api_user
from app.config import get_template_status, get_logo_path, get_logo_url
from app.utils import get_job_payload
from app.automations.credit_worksheet import get_status_metadata

router = APIRouter()


@router.get("/favicon.ico")
def favicon():
    logo_path = get_logo_path()
    if not logo_path:
        raise HTTPException(status_code=404, detail="Favicon not found.")
    return FileResponse(path=str(logo_path), media_type="image/png", filename=logo_path.name)


@router.get("/health")
def health(request: Request):
    require_api_user(request)
    meta = get_status_metadata()
    return {
        "status": "running",
        "template": get_template_status(),
        "logo_url": get_logo_url(),
        "field_count": meta["field_count"],
    }


@router.get("/api/status")
def api_status(request: Request):
    require_api_user(request)
    meta = get_status_metadata()
    return {
        "status": "running",
        "template": get_template_status(),
        "field_count": meta["field_count"],
        "expected_sheet": meta["expected_sheet"],
    }


@router.get("/api/latest-result")
def api_latest_result(request: Request):
    require_api_user(request)
    return get_job_payload()


@router.get("/api/result/{job_id}")
def api_result(request: Request, job_id: str):
    require_api_user(request)
    return get_job_payload(job_id)
