from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse

from app.auth import require_api_user
from app.config import OUTPUT_DIR, AUDIT_DIR, LOG_DIR
from app.utils import safe_filename

router = APIRouter()


@router.get("/download-output/{filename}")
def download_output(request: Request, filename: str):
    require_api_user(request)
    safe_name = safe_filename(filename)
    path = OUTPUT_DIR / safe_name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Output file not found.")
    if path.suffix.lower() == ".zip":
        media_type = "application/zip"
    elif path.suffix.lower() == ".xlsx":
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        media_type = "application/octet-stream"
    return FileResponse(path=str(path), filename=path.name, media_type=media_type)


@router.get("/download-audit/{filename}")
def download_audit(request: Request, filename: str):
    require_api_user(request)
    safe_name = safe_filename(filename)
    path = AUDIT_DIR / safe_name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Audit file not found.")
    return FileResponse(path=str(path), filename=path.name, media_type="text/plain")


@router.get("/download-log/{filename}")
def download_log(request: Request, filename: str):
    require_api_user(request)
    safe_name = safe_filename(filename)
    path = LOG_DIR / safe_name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Debug log not found.")
    return FileResponse(path=str(path), filename=path.name, media_type="text/plain")
