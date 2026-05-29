import html
import re
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


def _year_from_wb_name(wb_name: str) -> int:
    m = re.search(r"_(\d{4})\.xlsx$", wb_name or "")
    return int(m.group(1)) if m else 0


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

    # Group by year (derived from workbook filename)
    years: dict[int, dict] = {}
    for r in reports:
        yr = _year_from_wb_name(r.get("output_filename", ""))
        if yr not in years:
            years[yr] = {"workbook_filename": r.get("output_filename", ""), "runs": []}
        years[yr]["runs"].append(r)

    all_years = sorted(years.keys(), reverse=True)  # newest first

    year_cards_html = ""
    for yr in all_years:
        info = years[yr]
        wb_name = info["workbook_filename"]
        wb_path = NET_CAPITAL_DIR / wb_name if wb_name else None
        wb_exists = wb_path and wb_path.exists()
        dl_btn = (
            f'<a class="button-link orange" href="/download-net-capital/{html.escape(wb_name)}">Download</a>'
            if wb_exists else
            '<span class="muted" style="font-size:12px;">File missing</span>'
        )

        run_rows = []
        for r in info["runs"]:
            try:
                dt = datetime.fromisoformat(r["created_at"])
                date_str = dt.strftime("%b %d, %Y %I:%M %p").lstrip("0")
            except Exception:
                date_str = r.get("created_at", "")
            orig = r.get("original_filename", "")
            period = r.get("period_label", "")
            credit_sheet = r.get("credit_sheet", "")
            run_rows.append(f"""
            <div class="file-row run-row" data-period="{html.escape(period.lower())}" data-file="{html.escape(orig.lower())}">
                <div class="file-info">
                    <div class="file-name" style="font-size:0.875rem;">{html.escape(orig)}</div>
                    <div class="file-meta">
                        <span>{html.escape(date_str)}</span>
                        {f'<span class="dot">·</span><span>{html.escape(period)}</span>' if period else ''}
                        {f'<span class="dot">·</span><span style="color:var(--pc-blue);">{html.escape(credit_sheet)}</span>' if credit_sheet else ''}
                    </div>
                </div>
            </div>
            """)

        year_cards_html += f"""
        <div class="year-card" data-year="{yr}">
            <div style="display:flex;align-items:center;justify-content:space-between;
                        padding:14px 20px;background:var(--pc-blue-soft);
                        border-bottom:1px solid var(--border);border-radius:10px 10px 0 0;">
                <div>
                    <span style="font-size:1.35rem;font-weight:800;color:var(--pc-blue-dark);">{yr}</span>
                    <span class="muted" style="margin-left:10px;font-size:0.8rem;">{len(info['runs'])} run{"s" if len(info["runs"]) != 1 else ""}</span>
                </div>
                {dl_btn}
            </div>
            <div class="year-runs">{"".join(run_rows)}</div>
        </div>
        """

    if not year_cards_html:
        year_cards_html = '<div class="empty-state"><p style="margin:0;">No reports on file yet.</p></div>'

    return f"""
<!doctype html>
<html>
{head_html(f"{html.escape(customer['name'])} | Phillip Capital Risk Management")}
<style>
  .year-card {{
    border: 1px solid var(--border);
    border-radius: 10px;
    margin-bottom: 16px;
    overflow: hidden;
  }}
  .year-runs {{ padding: 0 20px; }}
  .year-card .file-row:last-child {{ border-bottom: none; }}

  .search-bar {{
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    flex-wrap: wrap;
  }}
  .search-bar input {{
    flex: 1;
    min-width: 200px;
    padding: 8px 14px;
    border: 1px solid var(--border);
    border-radius: 8px;
    font-size: 0.9rem;
    outline: none;
    background: var(--bg);
    color: var(--text);
  }}
  .search-bar input:focus {{ border-color: var(--pc-blue); }}
  .filter-pills {{ display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
  .pill {{
    padding: 5px 14px;
    border-radius: 20px;
    border: 1px solid var(--border);
    font-size: 0.8rem;
    cursor: pointer;
    background: var(--bg);
    color: var(--text);
    transition: background 0.15s, color 0.15s;
    user-select: none;
  }}
  .pill.active {{
    background: var(--pc-blue);
    color: #fff;
    border-color: var(--pc-blue);
  }}
  .year-card.hidden {{ display: none; }}
  .run-row.hidden {{ display: none; }}
</style>
<body>
<div class="shell">
    {topbar_html("customers", user)}
    {hero_html(customer['name'], "All processed workbooks for this firm.")}
    <div class="card">
        <div style="margin-bottom:14px;">
            <a href="/customers" style="color:var(--pc-blue);text-decoration:none;font-size:0.875rem;">&larr; All Customers</a>
        </div>

        <div class="search-bar">
            <input type="text" id="searchInput" placeholder="Search by filename or period (e.g. March, 2025)&hellip;" oninput="applyFilters()">
            <div class="filter-pills" id="yearPills">
                <span class="pill active" data-year="all" onclick="setPill(this)">All Years</span>
                {"".join(f'<span class="pill" data-year="{yr}" onclick="setPill(this)">{yr}</span>' for yr in all_years)}
            </div>
        </div>

        <div id="yearList">
            {year_cards_html}
        </div>
    </div>
</div>

<script>
let activeYear = "all";

function setPill(el) {{
    document.querySelectorAll("#yearPills .pill").forEach(p => p.classList.remove("active"));
    el.classList.add("active");
    activeYear = el.dataset.year;
    applyFilters();
}}

function applyFilters() {{
    const q = document.getElementById("searchInput").value.trim().toLowerCase();
    document.querySelectorAll("#yearList .year-card").forEach(card => {{
        const yr = card.dataset.year;
        const yearMatch = activeYear === "all" || activeYear === yr;
        if (!yearMatch) {{ card.classList.add("hidden"); return; }}
        if (!q) {{
            card.classList.remove("hidden");
            card.querySelectorAll(".run-row").forEach(r => r.classList.remove("hidden"));
            return;
        }}
        let anyVisible = false;
        card.querySelectorAll(".run-row").forEach(row => {{
            const text = (row.dataset.period + " " + row.dataset.file + " " + yr);
            if (text.includes(q)) {{ row.classList.remove("hidden"); anyVisible = true; }}
            else {{ row.classList.add("hidden"); }}
        }});
        if (anyVisible) card.classList.remove("hidden");
        else card.classList.add("hidden");
    }});
}}
</script>
</body>
</html>
    """
