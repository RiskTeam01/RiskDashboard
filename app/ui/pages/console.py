import html
from typing import Optional

from app.utils import get_job_payload, build_full_console_text, safe_js_json
from app.ui.components import head_html, topbar_html, hero_html


def console_page_html(user: str, job_id: Optional[str] = None) -> str:
    try:
        payload = get_job_payload(job_id)
    except Exception:
        payload = {}

    console_text = build_full_console_text(payload) if payload else "No console data yet."
    payload_json = safe_js_json(payload)

    return f"""
<!doctype html>
<html>
{head_html("Console | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("console", user)}
    {hero_html(
        "Processing Console",
        "Review field results, raw helper-code occurrences, and debug logs from the latest workbook generation.",
    )}

    <div class="card">
        <h2>Run Summary</h2>
        <div id="summaryBox" class="summary-list">No run found yet.</div>
        <br />
        <button onclick="copyEverything()">Copy Full Console</button>
        <button onclick="copyFieldResults()" class="secondary">Copy Field Results</button>
        <button onclick="copyRawOccurrences()" class="secondary">Copy Raw Occurrences</button>
        <button onclick="copyDebugLog()" class="secondary">Copy Debug Log</button>
    </div>

    <div class="card">
        <h2>Combined Console</h2>
        <div id="consoleBox" class="console">{html.escape(console_text)}</div>
    </div>
</div>

<script>
const payload = {payload_json};

let latestFieldResults = payload.field_results || [];
let latestRawOccurrences = payload.raw_occurrences || [];
let latestDebugLog = payload.debug_log || [];
let latestSummary = payload.summary || [];

if (payload.type === "batch" && payload.batch_results && payload.batch_results.length > 0) {{
    const firstSuccess = payload.batch_results.find(item => item.status === "SUCCESS");
    if (firstSuccess) {{
        latestFieldResults = firstSuccess.field_results || [];
        latestRawOccurrences = firstSuccess.raw_occurrences || [];
        latestDebugLog = firstSuccess.debug_log || [];
        latestSummary = payload.batch_summary || [];
    }}
}}

function rowsToTsv(rows) {{
    return rows.map(row => row.map(cell => {{
        if (cell === null || cell === undefined) return "";
        return String(cell).replaceAll("\\t", " ").replaceAll("\\n", " ").replaceAll("\\r", " ").trim();
    }}).join("\\t")).join("\\n");
}}

function buildFieldResultsTsv() {{
    const rows = [[
        "PDF Code(s)", "Field", "Excel Cell", "Preview Value", "Status", "Difficulty", "Notes", "Component Details"
    ]];
    latestFieldResults.forEach(r => rows.push([
        r.expression, r.label, r.excel_cell, r.display_value, r.status, r.difficulty, r.notes, r.component_details
    ]));
    return rowsToTsv(rows);
}}

function buildRawOccurrencesTsv() {{
    const rows = [[
        "Code", "Selected", "Page", "X0", "Y0", "X1", "Y1", "Nearby Amount", "Confidence Score", "Note", "Nearby Context"
    ]];
    latestRawOccurrences.forEach(r => rows.push([
        r.code, r.selected ? "YES" : "NO", r.page_number,
        Number(r.x0 || 0).toFixed(2), Number(r.y0 || 0).toFixed(2),
        Number(r.x1 || 0).toFixed(2), Number(r.y1 || 0).toFixed(2),
        r.nearby_amount_text, r.confidence_score, r.note, r.nearby_context
    ]));
    return rowsToTsv(rows);
}}

function buildFullConsoleText() {{
    return document.getElementById("consoleBox").textContent || "";
}}

async function copyText(text, label) {{
    if (!text || !text.trim()) {{ alert("No " + label + " available to copy."); return; }}
    await navigator.clipboard.writeText(text);
    alert(label + " copied to clipboard.");
}}

function copyEverything() {{ copyText(buildFullConsoleText(), "Full Console"); }}
function copyFieldResults() {{ copyText(buildFieldResultsTsv(), "Field Results"); }}
function copyRawOccurrences() {{ copyText(buildRawOccurrencesTsv(), "Raw Occurrences"); }}
function copyDebugLog() {{ copyText(latestDebugLog.join("\\n"), "Debug Log"); }}

function render() {{
    if (payload.type === "batch") {{
        document.getElementById("summaryBox").textContent = (payload.batch_summary || []).join("\\n") || "No run found yet.";
    }} else {{
        document.getElementById("summaryBox").textContent = latestSummary.join("\\n") || "No run found yet.";
    }}
}}

render();
</script>
</body>
</html>
    """
