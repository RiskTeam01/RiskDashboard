import html
from datetime import datetime
from pathlib import Path

from app.customers import load_customers, get_customer
from app.config import NET_CAPITAL_DIR, AUDIT_DIR
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
            The company name is detected from helper code 13 on the cover page.</p>
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

    # Group reports by workbook file (customer_name_year.xlsx)
    workbooks: dict[str, dict] = {}
    for r in reports:
        wb_name = r.get("output_filename", "")
        if wb_name not in workbooks:
            workbooks[wb_name] = {"filename": wb_name, "runs": []}
        workbooks[wb_name]["runs"].append(r)

    if not reports:
        reports_html = """
        <div class="empty-state">
            <p style="margin:0;">No reports on file for this customer yet.</p>
        </div>
        """
    else:
        wb_cards = []
        for wb_name, wb_info in workbooks.items():
            wb_path = NET_CAPITAL_DIR / wb_name if wb_name else None
            wb_exists = wb_path and wb_path.exists()

            dl_btn = (
                f'<a class="button-link orange" href="/download-net-capital/{html.escape(wb_name)}">Download Workbook</a>'
                if wb_exists else
                '<span class="muted" style="font-size:12px;">File not on disk</span>'
            )

            run_rows = []
            for r in wb_info["runs"]:
                try:
                    dt = datetime.fromisoformat(r["created_at"])
                    date_str = dt.strftime("%b %d, %Y %I:%M %p").lstrip("0")
                except Exception:
                    date_str = r.get("created_at", "")

                orig = r.get("original_filename", "")
                period = r.get("period_label", "")
                credit_sheet = r.get("credit_sheet", "")
                nc_sheet = r.get("net_capital_sheet", "")
                audit_name = r.get("audit_filename", "")

                audit_btn = (
                    f'<a class="button-link secondary" href="/download-audit/{html.escape(audit_name)}">Audit</a>'
                    if audit_name and (AUDIT_DIR / audit_name).exists() else ""
                )
                sheets_label = ""
                if credit_sheet or nc_sheet:
                    parts = []
                    if credit_sheet:
                        parts.append(f"Sheet: <em>{html.escape(credit_sheet)}</em>")
                    if nc_sheet:
                        parts.append(f"Net Capital: <em>{html.escape(nc_sheet)}</em>")
                    sheets_label = f'<span class="dot">·</span><span>{" &amp; ".join(parts)}</span>'

                run_rows.append(f"""
                <div class="file-row">
                    <div class="file-info">
                        <div class="file-name">{html.escape(orig)}</div>
                        <div class="file-meta">
                            <span>{html.escape(date_str)}</span>
                            {f'<span class="dot">·</span><span>{html.escape(period)}</span>' if period else ''}
                            {sheets_label}
                        </div>
                    </div>
                    <div class="file-actions">{audit_btn}</div>
                </div>
                """)

            wb_cards.append(f"""
            <div class="single-card" style="flex-direction:column;align-items:stretch;gap:0;padding:0;overflow:hidden;">
                <div style="display:flex;align-items:center;justify-content:space-between;padding:14px 18px;background:var(--pc-blue-soft);border-bottom:1px solid var(--border);">
                    <div>
                        <div class="folder-name" style="font-weight:700;color:var(--pc-blue-dark);">{html.escape(wb_name or "Unknown workbook")}</div>
                        <div class="folder-meta" style="font-size:12px;color:var(--muted);">{len(wb_info['runs'])} run(s) in this workbook</div>
                    </div>
                    <div>{dl_btn}</div>
                </div>
                <div style="padding:0 18px;">{''.join(run_rows)}</div>
            </div>
            """)

        reports_html = '<div class="outputs-list" style="gap:16px;">' + "".join(wb_cards) + "</div>"

    return f"""
<!doctype html>
<html>
{head_html(f"{html.escape(customer['name'])} | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("customers", user)}
    {hero_html(customer['name'], "Customer account — all processed reports for this firm.")}
    <div class="card">
        <div style="margin-bottom:12px;">
            <a href="/customers" style="color:var(--pc-blue);text-decoration:none;font-size:0.875rem;">
                &larr; All Customers
            </a>
        </div>
        <h2>Workbooks &amp; Runs</h2>
        <p class="muted">Each workbook groups monthly credit sheets and an accumulating Net Capital sheet for one year.
        Clicking Download gives you the complete combined Excel file.</p>
        {reports_html}
    </div>
</div>
</body>
</html>
    """
