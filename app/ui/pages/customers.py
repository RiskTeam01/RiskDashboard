import html
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from app.customers import load_customers, get_customer
from app.config import NET_CAPITAL_DIR, AUDIT_DIR
from app.ui.components import head_html, topbar_html, hero_html


# Net Capital rows we surface in the per-year financial summary.
SUMMARY_METRICS = [
    {"row": 7,  "label": "Total Equity",        "key": "total_equity"},
    {"row": 47, "label": "Net Capital",         "key": "net_capital"},
    {"row": 53, "label": "Excess Net Capital",  "key": "excess_net_capital"},
]
_MONTH_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _abbrev(val: float) -> str:
    """Compact money label: 1,250,000 -> 1.25M."""
    if val is None:
        return "—"
    sign = "-" if val < 0 else ""
    a = abs(val)
    if a >= 1_000_000_000:
        return f"{sign}{a/1_000_000_000:.2f}B"
    if a >= 1_000_000:
        return f"{sign}{a/1_000_000:.2f}M"
    if a >= 1_000:
        return f"{sign}{a/1_000:.1f}K"
    return f"{sign}{a:,.0f}"


def _full_money(val: float) -> str:
    if val is None:
        return "—"
    return f"${val:,.0f}"


def read_year_metrics(wb_path: Path, year: int) -> dict:
    """Read the three summary rows from a workbook's Net Capital sheet.
    Returns {metric_key: [(month_num, value), ...]} for months that have data."""
    result = {m["key"]: [] for m in SUMMARY_METRICS}
    if not wb_path or not wb_path.exists():
        return result
    try:
        wb = load_workbook(wb_path, data_only=True)
    except Exception:
        return result
    nc_name = f"Net Capital {year}"
    if nc_name not in wb.sheetnames:
        wb.close()
        return result
    ws = wb[nc_name]
    for metric in SUMMARY_METRICS:
        for idx, col in enumerate(_MONTH_COLUMNS):
            cell = ws[f"{col}{metric['row']}"]
            v = cell.value
            if isinstance(v, (int, float)):
                result[metric["key"]].append((idx + 1, float(v)))
    wb.close()
    return result


def _build_metrics_summary_html(metrics: dict) -> str:
    """Build the visual financial summary for one year."""
    has_any = any(len(v) >= 1 for v in metrics.values())
    if not has_any:
        return ""

    cards = []
    for metric in SUMMARY_METRICS:
        series = metrics[metric["key"]]  # list of (month_num, value), in column order
        if not series:
            cards.append(f"""
            <div class="metric-card">
                <div class="metric-label">{metric['label']}</div>
                <div class="metric-current muted">No data</div>
            </div>
            """)
            continue

        first_month, first_val = series[0]
        last_month, last_val = series[-1]
        delta = last_val - first_val
        pct = (delta / abs(first_val) * 100) if first_val else 0.0
        up = delta >= 0
        arrow = "&#9650;" if up else "&#9660;"
        delta_cls = "delta-up" if up else "delta-down"

        # Bar chart scaling
        vals = [v for _, v in series]
        vmax = max(vals)
        vmin = min(vals)
        span = (vmax - vmin) or abs(vmax) or 1
        bars = []
        for m_num, v in series:
            # Height 18%..100% so even the smallest bar is visible
            h = 18 + 82 * ((v - vmin) / span) if span else 60
            bar_cls = "bar-pos" if v >= 0 else "bar-neg"
            bars.append(f"""
            <div class="bar-col" title="{_MONTH_ABBR[m_num-1]}: {_full_money(v)}">
                <div class="bar-val">{_abbrev(v)}</div>
                <div class="bar {bar_cls}" style="height:{h:.0f}%;"></div>
                <div class="bar-month">{_MONTH_ABBR[m_num-1]}</div>
            </div>
            """)

        change_label = (
            f'<span class="metric-delta {delta_cls}">{arrow} {abs(pct):.1f}%</span>'
            if first_month != last_month else
            '<span class="metric-delta delta-flat">first reading</span>'
        )
        since_note = (
            f'<div class="metric-since">{_abbrev(delta) if delta < 0 else "+" + _abbrev(delta)} since {_MONTH_ABBR[first_month-1]}</div>'
            if first_month != last_month else
            f'<div class="metric-since">as of {_MONTH_ABBR[last_month-1]}</div>'
        )

        cards.append(f"""
        <div class="metric-card">
            <div class="metric-label">{metric['label']}</div>
            <div class="metric-current">{_full_money(last_val)} {change_label}</div>
            {since_note}
            <div class="bar-chart">{"".join(bars)}</div>
        </div>
        """)

    return f'<div class="metrics-summary">{"".join(cards)}</div>'


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

    QUARTERLY_MONTHS = [("March", 3), ("June", 6), ("September", 9), ("December", 12)]

    # Group by year (derived from workbook filename)
    years: dict[int, dict] = {}
    for r in reports:
        yr = _year_from_wb_name(r.get("output_filename", ""))
        if yr not in years:
            years[yr] = {"workbook_filename": r.get("output_filename", ""), "runs": []}
        years[yr]["runs"].append(r)

    all_years = sorted(years.keys(), reverse=True)  # newest first
    current_year = datetime.now().year
    current_month = datetime.now().month

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

        # Collect which credit sheets exist for this year
        sheets_received = {r.get("credit_sheet", "") for r in info["runs"]}

        quarterly_badges = []
        any_missing = False
        for month_name, month_num in QUARTERLY_MONTHS:
            sheet_name = f"{month_name} {yr}"
            received = sheet_name in sheets_received
            # Don't flag future quarters as missing
            is_future = (yr > current_year) or (yr == current_year and month_num > current_month)
            if received:
                quarterly_badges.append(
                    f'<span class="q-badge q-received" title="{sheet_name} received">&#10003; {month_name}</span>'
                )
            elif is_future:
                quarterly_badges.append(
                    f'<span class="q-badge q-future" title="{sheet_name} not yet due">&ndash; {month_name}</span>'
                )
            else:
                quarterly_badges.append(
                    f'<span class="q-badge q-missing" title="{sheet_name} missing">&#33; {month_name}</span>'
                )
                any_missing = True

        quarterly_html = f"""
        <div class="quarterly-summary">
            <span class="q-label">Quarterlies:</span>
            {"".join(quarterly_badges)}
            {f'<span class="q-alert">Missing data</span>' if any_missing else '<span class="q-ok">All received</span>'}
        </div>
        """

        # Financial summary read from the saved workbook's Net Capital sheet
        year_metrics = read_year_metrics(wb_path, yr)
        metrics_html = _build_metrics_summary_html(year_metrics)

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
        <div class="year-card collapsed" data-year="{yr}">
            <div class="year-header" onclick="toggleYear(this)">
                <div style="display:flex;align-items:center;gap:12px;">
                    <span class="chevron">&#9656;</span>
                    <span style="font-size:1.35rem;font-weight:800;color:var(--pc-blue-dark);">{yr}</span>
                    <span class="muted" style="font-size:0.8rem;">{len(info['runs'])} run{"s" if len(info["runs"]) != 1 else ""}</span>
                    {f'<span class="header-alert">! Missing</span>' if any_missing else ''}
                </div>
                <span onclick="event.stopPropagation();">{dl_btn}</span>
            </div>
            <div class="year-body">
                {quarterly_html}
                {metrics_html}
                <div class="year-runs">{"".join(run_rows)}</div>
            </div>
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

  .year-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 20px;
    background: var(--pc-blue-soft);
    border-radius: 10px 10px 0 0;
    cursor: pointer;
    user-select: none;
    transition: background 0.15s;
  }}
  .year-header:hover {{ filter: brightness(0.98); }}
  .chevron {{
    display: inline-block;
    transition: transform 0.18s ease;
    color: var(--pc-blue);
    font-size: 0.9rem;
  }}
  .year-card.collapsed .year-header {{ border-radius: 10px; }}
  .year-card:not(.collapsed) .chevron {{ transform: rotate(90deg); }}
  .year-body {{
    max-height: 4000px;
    overflow: hidden;
    transition: max-height 0.3s ease;
  }}
  .year-card.collapsed .year-body {{ max-height: 0; }}
  .header-alert {{
    font-size: 0.7rem;
    font-weight: 700;
    color: #b91c1c;
    background: #fee2e2;
    padding: 2px 9px;
    border-radius: 10px;
  }}
  .expand-toggle {{ font-weight: 600; }}

  /* Financial summary */
  .metrics-summary {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 14px;
    padding: 18px 20px;
    background: linear-gradient(180deg, var(--pc-blue-soft) 0%, transparent 100%);
    border-bottom: 1px solid var(--border);
  }}
  .metric-card {{
    background: var(--bg, #fff);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 14px 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  }}
  .metric-label {{
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--muted);
    margin-bottom: 6px;
  }}
  .metric-current {{
    font-size: 1.45rem;
    font-weight: 800;
    color: var(--text);
    display: flex;
    align-items: baseline;
    gap: 8px;
    flex-wrap: wrap;
  }}
  .metric-delta {{
    font-size: 0.78rem;
    font-weight: 700;
    padding: 1px 8px;
    border-radius: 10px;
  }}
  .delta-up {{ color: #065f46; background: #d1fae5; }}
  .delta-down {{ color: #991b1b; background: #fee2e2; }}
  .delta-flat {{ color: #6b7280; background: #f3f4f6; }}
  .metric-since {{
    font-size: 0.74rem;
    color: var(--muted);
    margin: 4px 0 12px;
  }}
  .bar-chart {{
    display: flex;
    align-items: flex-end;
    gap: 6px;
    height: 90px;
  }}
  .bar-col {{
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-end;
    height: 100%;
    min-width: 0;
  }}
  .bar-val {{
    font-size: 0.6rem;
    font-weight: 700;
    color: var(--muted);
    margin-bottom: 3px;
    white-space: nowrap;
  }}
  .bar {{
    width: 100%;
    max-width: 26px;
    border-radius: 4px 4px 0 0;
    transition: opacity 0.15s;
  }}
  .bar:hover {{ opacity: 0.8; }}
  .bar-pos {{ background: linear-gradient(180deg, var(--pc-blue) 0%, var(--pc-blue-dark) 100%); }}
  .bar-neg {{ background: linear-gradient(180deg, #f87171 0%, #b91c1c 100%); }}
  .bar-month {{
    font-size: 0.62rem;
    color: var(--muted);
    margin-top: 4px;
  }}

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

  .quarterly-summary {{
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
    padding: 10px 20px;
    background: var(--surface, #fafafa);
    border-bottom: 1px solid var(--border);
    font-size: 0.8rem;
  }}
  .q-label {{
    font-weight: 600;
    color: var(--muted);
    margin-right: 4px;
  }}
  .q-badge {{
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 0.78rem;
    font-weight: 600;
  }}
  .q-received {{
    background: #d1fae5;
    color: #065f46;
  }}
  .q-missing {{
    background: #fee2e2;
    color: #991b1b;
  }}
  .q-future {{
    background: #f3f4f6;
    color: #9ca3af;
  }}
  .q-alert {{
    margin-left: auto;
    font-size: 0.75rem;
    font-weight: 700;
    color: #b91c1c;
    background: #fee2e2;
    padding: 2px 10px;
    border-radius: 10px;
  }}
  .q-ok {{
    margin-left: auto;
    font-size: 0.75rem;
    font-weight: 700;
    color: #065f46;
    background: #d1fae5;
    padding: 2px 10px;
    border-radius: 10px;
  }}
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
            <span class="pill expand-toggle" onclick="toggleAll(this)" data-state="collapsed">Expand all</span>
        </div>

        <div id="yearList">
            {year_cards_html}
        </div>
    </div>
</div>

<script>
let activeYear = "all";

function toggleYear(headerEl) {{
    headerEl.closest(".year-card").classList.toggle("collapsed");
}}

function toggleAll(el) {{
    const expanding = el.dataset.state === "collapsed";
    document.querySelectorAll("#yearList .year-card").forEach(card => {{
        card.classList.toggle("collapsed", !expanding);
    }});
    el.dataset.state = expanding ? "expanded" : "collapsed";
    el.textContent = expanding ? "Collapse all" : "Expand all";
}}

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
        if (anyVisible) {{
            card.classList.remove("hidden");
            card.classList.remove("collapsed");  // auto-expand matches while searching
        }} else {{
            card.classList.add("hidden");
        }}
    }});
}}
</script>
</body>
</html>
    """
