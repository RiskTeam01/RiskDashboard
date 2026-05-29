import html
from datetime import datetime

from app.customers import load_customers, get_customer
from app.config import NET_CAPITAL_DIR, OUTPUT_DIR, AUDIT_DIR
from app.ui.components import head_html, topbar_html, hero_html


def customers_list_page_html(user: str) -> str:
    data = load_customers()
    customers = sorted(
        [{"id": cid, **cust} for cid, cust in data.items()],
        key=lambda c: c.get("name", "").lower(),
    )

    if not customers:
        body = """
        <div class="empty-state">
            <h3 style="margin-bottom:8px;">No customers yet</h3>
            <p style="margin:0;">Customer accounts are created automatically when you process a PDF.
            The company name is detected from helper code 13.</p>
        </div>
        """
    else:
        rows = []
        for c in customers:
            reports = c.get("reports", [])
            last_run = ""
            if reports:
                latest = max(reports, key=lambda r: r.get("created_at", ""))
                try:
                    dt = datetime.fromisoformat(latest["created_at"])
                    last_run = dt.strftime("%b %d, %Y %I:%M %p").lstrip("0")
                except Exception:
                    last_run = latest.get("created_at", "")
            report_count = len(reports)
            word = "report" if report_count == 1 else "reports"
            rows.append(f"""
            <a class="customer-card" href="/customers/{html.escape(c['id'])}">
                <div class="customer-name">{html.escape(c['name'])}</div>
                <div class="customer-meta">
                    <span>{report_count} {word}</span>
                    {f'<span class="dot">·</span><span>Last run: {html.escape(last_run)}</span>' if last_run else ''}
                </div>
            </a>
            """)
        body = '<div class="customer-list">' + "".join(rows) + "</div>"

    return f"""
<!doctype html>
<html>
{head_html("Customers | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("customers", user)}
    {hero_html("Customer Accounts", "Each customer account collects all reports processed for that firm.")}
    <div class="card">
        <h2>All Customers</h2>
        <p class="muted">Accounts are created automatically from helper code 13 (company name) on each PDF run.</p>
        {body}
    </div>
</div>
</body>
</html>
    """


def customer_detail_page_html(user: str, customer_id: str) -> str:
    customer = get_customer(customer_id)
    if not customer:
        return f"""
<!doctype html><html>
{head_html("Not Found | Phillip Capital Risk Management")}
<body><div class="shell">
{topbar_html("customers", user)}
<div class="card"><h2>Customer not found</h2>
<p><a href="/customers">Back to Customers</a></p></div>
</div></body></html>
        """

    reports = sorted(
        customer.get("reports", []),
        key=lambda r: r.get("created_at", ""),
        reverse=True,
    )

    if not reports:
        reports_html = """
        <div class="empty-state">
            <p style="margin:0;">No reports on file for this customer yet.</p>
        </div>
        """
    else:
        rows = []
        for r in reports:
            try:
                dt = datetime.fromisoformat(r["created_at"])
                date_str = dt.strftime("%b %d, %Y %I:%M %p").lstrip("0")
            except Exception:
                date_str = r.get("created_at", "")

            period = r.get("period_label", "")
            output_name = r.get("output_filename", "")
            audit_name = r.get("audit_filename", "")
            nc_name = r.get("net_capital_filename", "")
            orig = r.get("original_filename", "")

            dl_btn = (
                f'<a class="button-link" href="/download-output/{html.escape(output_name)}">Credit WS</a>'
                if output_name and (OUTPUT_DIR / output_name).exists() else ""
            )
            audit_btn = (
                f'<a class="button-link secondary" href="/download-audit/{html.escape(audit_name)}">Audit</a>'
                if audit_name and (AUDIT_DIR / audit_name).exists() else ""
            )
            nc_btn = (
                f'<a class="button-link secondary" href="/download-net-capital/{html.escape(nc_name)}">Net Capital</a>'
                if nc_name and (NET_CAPITAL_DIR / nc_name).exists() else ""
            )

            rows.append(f"""
            <div class="file-row">
                <div class="file-info">
                    <div class="file-name">{html.escape(orig)}</div>
                    <div class="file-meta">
                        <span>{html.escape(date_str)}</span>
                        {f'<span class="dot">·</span><span>{html.escape(period)}</span>' if period else ''}
                    </div>
                </div>
                <div class="file-actions">
                    {dl_btn}
                    {audit_btn}
                    {nc_btn}
                </div>
            </div>
            """)
        reports_html = '<div class="outputs-list">' + "".join(rows) + "</div>"

    return f"""
<!doctype html>
<html>
{head_html(f"{html.escape(customer['name'])} | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("customers", user)}
    {hero_html(customer['name'], f"Customer account — all processed reports for this firm.")}
    <div class="card">
        <div style="margin-bottom:12px;">
            <a href="/customers" style="color:var(--pc-blue);text-decoration:none;font-size:0.875rem;">
                &larr; All Customers
            </a>
        </div>
        <h2>Reports</h2>
        <p class="muted">All credit worksheet runs and net capital workbooks processed for this customer.</p>
        {reports_html}
    </div>
</div>
</body>
</html>
    """
