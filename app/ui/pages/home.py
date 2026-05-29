import html

from app.config import load_config, get_template_status
from app.ui.components import head_html, topbar_html, hero_html


def home_page_html(user: str) -> str:
    config = load_config()
    template = get_template_status()
    template_class = "good" if template["exists"] else "bad"
    template_label = "Template detected" if template["exists"] else "Template missing"
    template_note = (
        f"Using backend Excel template: {template['filename']}"
        if template["exists"]
        else f"Place a blank .xlsx template inside: {template['folder']}"
    )

    max_batch = config["max_batch_size"]
    max_size = config["max_pdf_size_mb"]

    return f"""
<!doctype html>
<html>
{head_html("Phillip Capital Risk Management | Credit Worksheet Processor")}
<body>
<div class="shell">
    {topbar_html("home", user)}
    {hero_html(
        "Generate Credit Worksheet",
        "Upload customer PDF files and automatically produce completed Excel credit worksheets using the backend template.",
    )}

    <div class="card center-card">
        <h2>Upload Customer PDF(s)</h2>
        <p>Status: <span class="status-pill {html.escape(template_class)}">{html.escape(template_label)}</span></p>
        <p class="muted">{html.escape(template_note)}</p>

        <form id="generateForm">
            <div class="upload-zone">
                <input type="file" id="generatePdf" name="pdfs" accept="application/pdf" multiple required />
            </div>

            <div class="batch-note">
                Upload one PDF to generate one Excel file, or upload multiple PDFs to generate a ZIP containing one Excel workbook per PDF.
                Limits: up to <strong>{max_batch}</strong> PDFs per batch, <strong>{max_size} MB</strong> per file.
            </div>

            <label class="muted" style="display:block; margin-top:14px;">
                <input type="checkbox" id="recalculateExcel" name="recalculate_with_excel" value="true" />
                Recalculate formulas with Excel COM if available
            </label>

            <br />
            <button type="submit" class="orange">Generate Completed Excel</button>
        </form>

        <div id="errorBox" class="error-box"></div>

        <div id="progressWrap" class="progress-wrap">
            <div class="progress-head">
                <h3 style="margin:0;">Batch Progress</h3>
                <div id="progressCounts" class="progress-counts"></div>
            </div>
            <div id="progressList" class="progress-list"></div>
            <div id="downloadBar" class="download-bar"></div>
        </div>

        <div class="metric-grid">
            <div class="metric"><b>Batch</b>Process one or many PDFs</div>
            <div class="metric"><b>Live</b>Per-PDF status updates</div>
            <div class="metric"><b>Audit</b>Report saved per workbook</div>
        </div>
    </div>
</div>

<script>
const MAX_BATCH = {max_batch};
const MAX_SIZE_BYTES = {max_size} * 1024 * 1024;
let pollTimer = null;

function escapeHtml(value) {{
    if (value === null || value === undefined) return "";
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}}

function showError(message) {{
    const box = document.getElementById("errorBox");
    box.style.display = "block";
    box.textContent = message;
}}

function clearError() {{
    const box = document.getElementById("errorBox");
    box.style.display = "none";
    box.textContent = "";
}}

function statusBadge(status) {{
    const labels = {{ queued: "Queued", processing: "Processing", complete: "Complete", failed: "Failed" }};
    const label = labels[status] || status;
    return `<span class="pstat ${{status}}"><span class="pdot"></span>${{escapeHtml(label)}}</span>`;
}}

function renderProgress(job) {{
    const wrap = document.getElementById("progressWrap");
    wrap.style.display = "block";

    const counts = document.getElementById("progressCounts");
    counts.innerHTML =
        `<span class="count-chip">Total: ${{job.total}}</span>` +
        `<span class="count-chip ok">Complete: ${{job.success_count}}</span>` +
        `<span class="count-chip fail">Failed: ${{job.failure_count}}</span>`;

    const list = document.getElementById("progressList");
    list.innerHTML = (job.files || []).map(function(f) {{
        let detail = "";
        if (f.status === "complete") {{
            detail = `<div class="pr-detail">Output: ${{escapeHtml(f.output_filename || "")}} &middot; Fields ${{escapeHtml(f.fields_found || "")}} &middot; Needs review: ${{escapeHtml(String(f.needs_review))}}</div>`;
        }} else if (f.status === "failed") {{
            detail = `<div class="pr-detail err">${{escapeHtml(f.error || "Failed")}}</div>`;
        }} else if (f.status === "processing") {{
            detail = `<div class="pr-detail">Extracting values and writing workbook...</div>`;
        }} else {{
            detail = `<div class="pr-detail">Waiting in queue</div>`;
        }}

        let actions = "";
        if (f.status === "complete") {{
            if (f.download_url) actions += `<a class="button-link" href="${{f.download_url}}">Excel</a>`;
            if (f.audit_url) actions += `<a class="button-link secondary" href="${{f.audit_url}}">Audit</a>`;
        }}

        return `<div class="progress-row">
            <div class="pr-left">
                <div class="pr-name">${{escapeHtml(f.filename)}}</div>
                ${{detail}}
            </div>
            <div class="pr-actions">
                ${{statusBadge(f.status)}}
                ${{actions}}
            </div>
        </div>`;
    }}).join("");

    const bar = document.getElementById("downloadBar");
    if (job.status === "complete") {{
        bar.style.display = "block";
        bar.className = job.failure_count > 0 ? "download-bar has-fail" : "download-bar";

        let summary = `<strong>${{job.success_count}}</strong> successful`;
        if (job.failure_count > 0) summary += ` &middot; <strong>${{job.failure_count}}</strong> failed`;

        let links = "";
        if (job.zip_url) {{
            links += `<a class="button-link orange" href="${{job.zip_url}}">Download All (ZIP)</a>`;
        }} else if (job.single_url) {{
            links += `<a class="button-link orange" href="${{job.single_url}}">Download Excel</a>`;
        }}
        if (job.console_url) {{
            links += `<a class="button-link secondary" href="${{job.console_url}}">Open Console</a>`;
        }}

        bar.innerHTML = `<div style="margin-bottom:10px;font-weight:800;">${{summary}}</div>${{links}}`;
    }} else {{
        bar.style.display = "none";
    }}
}}

async function poll(jobId) {{
    try {{
        const res = await fetch(`/batch-status/${{jobId}}`);
        if (res.status === 401) {{ window.location = "/login"; return; }}
        if (!res.ok) {{ showError("Lost connection to batch job."); return; }}

        const job = await res.json();
        renderProgress(job);

        if (job.status === "complete") {{
            if (pollTimer) clearTimeout(pollTimer);
            return;
        }}
        pollTimer = setTimeout(() => poll(jobId), 900);
    }} catch (err) {{
        showError("Status polling failed: " + err.message);
    }}
}}

document.getElementById("generateForm").addEventListener("submit", async function(e) {{
    e.preventDefault();
    clearError();

    const files = document.getElementById("generatePdf").files;

    if (!files || files.length === 0) {{
        showError("Please select at least one PDF.");
        return;
    }}

    if (files.length > MAX_BATCH) {{
        showError(`Too many files. The batch limit is ${{MAX_BATCH}} PDFs.`);
        return;
    }}

    for (let i = 0; i < files.length; i++) {{
        if (files[i].size > MAX_SIZE_BYTES) {{
            showError(`"${{files[i].name}}" is larger than the ${{MAX_SIZE_BYTES / (1024*1024)}} MB limit.`);
            return;
        }}
    }}

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {{
        formData.append("pdfs", files[i]);
    }}
    if (document.getElementById("recalculateExcel").checked) {{
        formData.append("recalculate_with_excel", "true");
    }}

    try {{
        const response = await fetch("/start-batch", {{ method: "POST", body: formData }});
        if (response.status === 401) {{ window.location = "/login"; return; }}

        if (!response.ok) {{
            const errText = await response.text();
            throw new Error(errText);
        }}

        const data = await response.json();
        if (!data.job_id) throw new Error("Server did not return a job id.");

        document.getElementById("downloadBar").style.display = "none";
        poll(data.job_id);
    }} catch (err) {{
        showError("Could not start batch:\\n" + err.message);
    }}
}});
</script>
</body>
</html>
    """
