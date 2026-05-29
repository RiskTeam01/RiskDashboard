from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

from app.auth import authenticate, create_session, destroy_session, current_user_or_none
from app.config import SESSION_COOKIE, load_config
from app.ui.pages.login import login_page_html

router = APIRouter()


@router.get("/login", response_class=HTMLResponse)
def login_get(request: Request):
    if current_user_or_none(request):
        return RedirectResponse("/", status_code=303)
    return HTMLResponse(login_page_html())


@router.post("/login")
def login_post(request: Request, username: str = Form(...), password: str = Form(...)):
    resolved = authenticate(username, password)
    if not resolved:
        return HTMLResponse(login_page_html("Invalid username or password."), status_code=401)

    token = create_session(resolved)
    timeout_seconds = load_config()["session_timeout_minutes"] * 60

    response = RedirectResponse("/", status_code=303)
    response.set_cookie(
        SESSION_COOKIE,
        token,
        httponly=True,
        samesite="lax",
        max_age=timeout_seconds,
        path="/",
    )
    return response


@router.get("/logout")
def logout(request: Request):
    destroy_session(request.cookies.get(SESSION_COOKIE))
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(SESSION_COOKIE, path="/")
    return response
