import html

from app.config import load_config
from app.auth import load_users
from app.ui.components import head_html, topbar_html, hero_html


def settings_page_html(user: str, flash: str = "", flash_type: str = "ok") -> str:
    config = load_config()
    users = load_users()

    flash_block = (
        f'<div class="flash {html.escape(flash_type)}">{html.escape(flash)}</div>'
        if flash else ""
    )

    user_rows = []
    for record in sorted(users.values(), key=lambda r: r.get("username", "").lower()):
        uname = record.get("username", "")
        created_at = record.get("created_at", "")
        created_by = record.get("created_by", "")
        is_self = uname.strip().lower() == user.strip().lower()

        delete_control = (
            '<span class="muted" style="font-size:12px;">Current session</span>'
            if is_self else
            f"""
            <form method="post" action="/settings/delete-user" onsubmit="return confirm('Remove {html.escape(uname)}?');" style="margin:0;">
                <input type="hidden" name="username" value="{html.escape(uname)}" />
                <button type="submit" class="danger" style="padding:8px 12px; font-size:12px;">Remove</button>
            </form>
            """
        )

        pw_updated_at = record.get("password_updated_at", "")
        pw_updated_by = record.get("password_updated_by", "")
        meta_pw_line = (
            f'<div class="ur-meta">Password updated {html.escape(pw_updated_at)}'
            f' by {html.escape(pw_updated_by)}</div>'
            if pw_updated_at else ""
        )

        self_tag = '<span class="ur-self-tag">You</span>' if is_self else ""
        confirm_msg = (
            "Change your own password?"
            if is_self
            else f"Update password for {uname}? Any existing sessions for this user will be signed out."
        )

        user_rows.append(f"""
            <div class="user-row">
                <div class="ur-top">
                    <div style="min-width:0;">
                        <div class="ur-name">{html.escape(uname)}{self_tag}</div>
                        <div class="ur-meta">Added {html.escape(created_at)} by {html.escape(created_by)}</div>
                        {meta_pw_line}
                    </div>
                    <div>{delete_control}</div>
                </div>
                <form method="post" action="/settings/update-password" class="ur-pwform"
                      onsubmit="return confirm('{html.escape(confirm_msg)}');">
                    <input type="hidden" name="username" value="{html.escape(uname)}" />
                    <input type="password" name="new_password" placeholder="New password for {html.escape(uname)}"
                           required minlength="1" autocomplete="new-password" />
                    <button type="submit" class="primary">Update password</button>
                </form>
            </div>
        """)

    users_block = "".join(user_rows) if user_rows else '<p class="muted">No users found.</p>'

    return f"""
<!doctype html>
<html>
{head_html("Settings | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("settings", user)}
    {hero_html("Settings", "Manage portal logins, upload limits, session timeout, and folder cleanup.")}

    {flash_block}

    <div class="settings-grid">
        <div class="card">
            <h2>Login Credentials</h2>
            <p class="muted">Existing users who can sign in to the portal.</p>
            {users_block}
        </div>

        <div class="card">
            <h2>Add New Login</h2>
            <p class="muted">Create a username and password. Passwords are stored hashed.</p>
            <form method="post" action="/settings/add-user">
                <div class="field-group">
                    <label for="new_username">Username</label>
                    <input type="text" id="new_username" name="new_username" placeholder="name@phillipcapital.com" required />
                </div>
                <div class="field-group">
                    <label for="new_password">Password</label>
                    <input type="text" id="new_password" name="new_password" placeholder="Set a password" required />
                </div>
                <button type="submit" class="orange">Add User</button>
            </form>
        </div>
    </div>

    <div class="card">
        <h2>Upload Limits</h2>
        <p class="muted">Controls applied to every batch. Session timeout also signs users out after inactivity.</p>
        <form method="post" action="/settings/limits">
            <div class="limit-grid">
                <div class="field-group">
                    <label for="max_pdf_size_mb">Max PDF size (MB)</label>
                    <input type="number" min="1" max="500" id="max_pdf_size_mb" name="max_pdf_size_mb" value="{config['max_pdf_size_mb']}" />
                </div>
                <div class="field-group">
                    <label for="max_batch_size">Max PDFs per batch</label>
                    <input type="number" min="1" max="200" id="max_batch_size" name="max_batch_size" value="{config['max_batch_size']}" />
                </div>
                <div class="field-group">
                    <label for="session_timeout_minutes">Session timeout (minutes)</label>
                    <input type="number" min="5" max="1440" id="session_timeout_minutes" name="session_timeout_minutes" value="{config['session_timeout_minutes']}" />
                </div>
            </div>
            <button type="submit">Save Limits</button>
        </form>
    </div>

    <div class="card">
        <h2>Cleanup</h2>
        <p class="muted">
            Automatic cleanup runs on startup. Uploads older than <strong>{config['cleanup_uploads_days']}</strong> days,
            outputs/audits older than <strong>{config['cleanup_outputs_days']}</strong> days, and logs older than <strong>{config['cleanup_logs_days']}</strong> days
            are removed. You can also change retention or clear folders now.
        </p>

        <form method="post" action="/settings/retention">
            <div class="limit-grid">
                <div class="field-group">
                    <label for="cleanup_uploads_days">Keep uploads (days)</label>
                    <input type="number" min="0" max="365" id="cleanup_uploads_days" name="cleanup_uploads_days" value="{config['cleanup_uploads_days']}" />
                </div>
                <div class="field-group">
                    <label for="cleanup_outputs_days">Keep outputs/audits (days)</label>
                    <input type="number" min="0" max="365" id="cleanup_outputs_days" name="cleanup_outputs_days" value="{config['cleanup_outputs_days']}" />
                </div>
                <div class="field-group">
                    <label for="cleanup_logs_days">Keep logs (days)</label>
                    <input type="number" min="0" max="365" id="cleanup_logs_days" name="cleanup_logs_days" value="{config['cleanup_logs_days']}" />
                </div>
            </div>
            <button type="submit">Save Retention</button>
        </form>

        <hr style="border:none; border-top:1px solid var(--border); margin:18px 0;" />

        <p class="muted">Run cleanup immediately:</p>
        <form method="post" action="/settings/cleanup" style="display:inline;">
            <input type="hidden" name="target" value="age" />
            <button type="submit" class="secondary">Run Age-Based Cleanup</button>
        </form>
        <form method="post" action="/settings/cleanup" style="display:inline;" onsubmit="return confirm('Clear ALL uploaded PDFs now?');">
            <input type="hidden" name="target" value="uploads" />
            <button type="submit" class="danger">Clear All Uploads</button>
        </form>
        <form method="post" action="/settings/cleanup" style="display:inline;" onsubmit="return confirm('Clear ALL logs now?');">
            <input type="hidden" name="target" value="logs" />
            <button type="submit" class="danger">Clear All Logs</button>
        </form>
        <form method="post" action="/settings/cleanup" style="display:inline;" onsubmit="return confirm('Clear ALL generated outputs and audits now?');">
            <input type="hidden" name="target" value="outputs" />
            <button type="submit" class="danger">Clear All Outputs</button>
        </form>
    </div>

    <div class="card" style="border:1.5px solid var(--red);">
        <h2 style="color:var(--red);">Danger Zone — Customer Data</h2>
        <p class="muted">
            Permanently delete <strong>all customer accounts</strong>, their report
            history, and every accumulated Net Capital workbook. This cannot be undone.
            You must confirm and re-enter your password to proceed.
        </p>
        <form method="post" action="/settings/delete-all-customers"
              onsubmit="return confirm('Delete ALL customer accounts and Net Capital data? This cannot be undone.');">
            <div class="field-group" style="max-width:360px;">
                <label for="delete_customers_password">Your password</label>
                <input type="password" id="delete_customers_password" name="password"
                       placeholder="Enter your password to confirm" required autocomplete="current-password" />
            </div>
            <button type="submit" class="danger">Delete All Customer Accounts &amp; Data</button>
        </form>
    </div>
</div>
</body>
</html>
    """
