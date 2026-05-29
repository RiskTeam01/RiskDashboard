import html
from typing import Optional

from app.config import get_logo_url
from app.ui.styles import base_css


def head_html(title: str) -> str:
    return f"""
    <head>
        <title>{html.escape(title)}</title>
        <link rel="icon" type="image/png" href="/favicon.ico">
        {base_css()}
    </head>
    """


def nav_html(active: str) -> str:
    home_active = "active" if active == "home" else ""
    console_active = "active" if active == "console" else ""
    outputs_active = "active" if active == "outputs" else ""
    settings_active = "active" if active == "settings" else ""

    return f"""
    <div class="menu-wrap">
        <button class="menu-button" onclick="toggleMenu(event)" title="Menu">&#9776;</button>
        <div id="pageMenu" class="menu-dropdown">
            <a class="{home_active}" href="/">Home</a>
            <a class="{console_active}" href="/console">Console</a>
            <a class="{outputs_active}" href="/outputs">Outputs</a>
            <a class="{settings_active}" href="/settings">Settings</a>
            <a class="logout" href="/logout">Sign out</a>
        </div>
    </div>

    <script>
        function toggleMenu(event) {{
            event.stopPropagation();
            const menu = document.getElementById("pageMenu");
            menu.classList.toggle("show");
        }}

        document.addEventListener("click", function() {{
            const menu = document.getElementById("pageMenu");
            if (menu) {{
                menu.classList.remove("show");
            }}
        }});
    </script>
    """


def topbar_html(active: str, user: Optional[str] = None) -> str:
    logo_url = get_logo_url()
    logo_block = (
        f'<img src="{html.escape(logo_url)}" alt="PhillipCapital" class="brand-logo" />'
        if logo_url else ""
    )
    user_pill = (
        f'<div class="user-pill"><span class="dot"></span>{html.escape(user)}</div>'
        if user else ""
    )
    return f"""
    <div class="topbar">
        <div class="brand">
            {logo_block}
            <div class="brand-text-wrap">
                <div class="brand-title">Phillip Capital</div>
                <div class="brand-subtitle">Risk Management</div>
            </div>
        </div>
        <div class="topbar-right">
            {user_pill}
            {nav_html(active)}
        </div>
    </div>
    """


def hero_html(title: str, subtitle: str) -> str:
    return f"""
    <div class="hero">
        <div class="hero-topline"></div>
        <div class="hero-inner">
            <div class="hero-left">
                <div class="hero-kicker">Credit Worksheet Automation</div>
                <h1>{html.escape(title)}</h1>
                <p>{html.escape(subtitle)}</p>
            </div>
            <div class="hero-right">
                <div class="hero-right-brand">
                    <div class="hero-right-title">PhillipCapital</div>
                    <div class="hero-right-subtitle">
                        Securities, Fixed Income,<br />
                        Futures, Options.
                    </div>
                </div>
            </div>
        </div>
    </div>
    """
