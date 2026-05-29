import html

from app.config import get_logo_url
from app.ui.components import head_html


def login_page_html(error: str = "") -> str:
    logo_url = get_logo_url()
    logo_block = (
        f'<div class="login-logo"><img src="{html.escape(logo_url)}" alt="PhillipCapital" /></div>'
        if logo_url else ""
    )
    error_block = (
        f'<div class="login-error">{html.escape(error)}</div>'
        if error else ""
    )

    return f"""
<!doctype html>
<html>
{head_html("Sign In | Phillip Capital Risk Management")}
<body>
<div class="shell">
    <div class="login-shell">
        <div class="login-card">
            <div class="login-topline"></div>
            <div class="login-body">
                {logo_block}
                <div class="login-title">Risk Management Portal</div>
                <div class="login-sub">Credit Worksheet Processor &middot; Sign in to continue</div>
                {error_block}
                <form method="post" action="/login">
                    <div class="field-group">
                        <label for="username">Username</label>
                        <input type="text" id="username" name="username" autocomplete="username" autofocus required />
                    </div>
                    <div class="field-group">
                        <label for="password">Password</label>
                        <input type="password" id="password" name="password" autocomplete="current-password" required />
                    </div>
                    <button type="submit" class="orange full-btn">Sign In</button>
                </form>
            </div>
        </div>
    </div>
</div>
</body>
</html>
    """
