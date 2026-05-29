from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi import HTTPException

from app.auth import current_user_or_none
from app.config import NET_CAPITAL_DIR
from app.ui.pages.customers import customers_list_page_html, customer_detail_page_html

router = APIRouter()


@router.get("/customers", response_class=HTMLResponse)
def customers_list(request: Request):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return HTMLResponse(customers_list_page_html(user))


@router.get("/customers/{customer_id}", response_class=HTMLResponse)
def customer_detail(customer_id: str, request: Request):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return HTMLResponse(customer_detail_page_html(user, customer_id))


@router.get("/download-net-capital/{filename}")
def download_net_capital(filename: str, request: Request):
    user = current_user_or_none(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    safe = filename.replace("/", "").replace("..", "")
    path = NET_CAPITAL_DIR / safe
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path=str(path), filename=safe, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
