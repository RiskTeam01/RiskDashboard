import html
import re
import zipfile
from datetime import date, datetime, timedelta

from app.config import OUTPUT_DIR, AUDIT_DIR
from app.utils import get_file_size_label
from app.ui.components import head_html, topbar_html, hero_html

_FOLDER_SVG = (
    '<svg viewBox="0 0 20 16" fill="currentColor" aria-hidden="true">'
    '<path d="M0 2.5C0 1.12 1.12 0 2.5 0H7l2 2h8.5C18.88 2 20 3.12 20 4.5v9c0 1.38-1.12 2.5-2.5 2.5h-15C1.12 16 0 14.88 0 13.5v-11z"/>'
    '</svg>'
)
_DOC_SVG = (
    '<svg viewBox="0 0 16 20" fill="currentColor" aria-hidden="true">'
    '<path d="M2 0C0.9 0 0 0.9 0 2v16c0 1.1 0.9 2 2 2h12c1.1 0 2-0.9 2-2V6L10 0H2zm8 1.5L14.5 6H10V1.5z"/>'
    '</svg>'
)
_AUDIT_SVG = (
    '<svg viewBox="0 0 16 20" fill="currentColor" aria-hidden="true">'
    '<path d="M2 0C0.9 0 0 0.9 0 2v16c0 1.1 0.9 2 2 2h12c1.1 0 2-0.9 2-2V2c0-1.1-0.9-2-2-2H2zm2 6h8v1.5H4V6zm0 3h8v1.5H4V9zm0 3h5v1.5H4V12z"/>'
    '</svg>'
)


def outputs_page_html(user: str) -> str:
    def fmt_time(mtime: float) -> str:
        return datetime.fromtimestamp(mtime).strftime("%I:%M %p").lstrip("0")

    def fmt_size(n: int) -> str:
        return get_file_size_label(n)

    batches = []
    for zip_path in OUTPUT_DIR.glob("*.zip"):
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                members = {n for n in zf.namelist() if not n.endswith("/")}
        except Exception:
            members = set()
        batches.append({
            "kind": "batch",
            "zip_path": zip_path,
            "members": members,
            "mtime": zip_path.stat().st_mtime,
            "files": [],
        })

    xlsx_to_batch = {}
    for idx, b in enumerate(batches):
        for name in b["members"]:
            xlsx_to_batch[name] = idx

    audit_pattern = re.compile(r"^(.+?)_audit(?:_\d+)?$")
    audits_by_xlsx_stem: dict[str, list] = {}
    for audit_path in AUDIT_DIR.glob("*.txt"):
        m = audit_pattern.match(audit_path.stem)
        if m:
            audits_by_xlsx_stem.setdefault(m.group(1), []).append(audit_path)

    matched_audits: set = set()
    singles = []

    for xlsx_path in OUTPUT_DIR.glob("*.xlsx"):
        these_audits = audits_by_xlsx_stem.get(xlsx_path.stem, [])
        for a in these_audits:
            matched_audits.add(a)
        entry = {"xlsx_path": xlsx_path, "audit_paths": these_audits}
        if xlsx_path.name in xlsx_to_batch:
            batches[xlsx_to_batch[xlsx_path.name]]["files"].append(entry)
        else:
            singles.append({
                "kind": "single",
                "xlsx_path": xlsx_path,
                "audit_paths": these_audits,
                "mtime": xlsx_path.stat().st_mtime,
            })

    orphan_audits = []
    for audits in audits_by_xlsx_stem.values():
        for a in audits:
            if a not in matched_audits:
                orphan_audits.append({
                    "kind": "audit_only",
                    "audit_path": a,
                    "mtime": a.stat().st_mtime,
                })

    items = batches + singles + orphan_audits
    items.sort(key=lambda it: it["mtime"], reverse=True)

    today = date.today()
    yesterday = today - timedelta(days=1)

    def date_labels(d):
        if d == today:
            return "Today", d.strftime("%A, %B %d")
        if d == yesterday:
            return "Yesterday", d.strftime("%A, %B %d")
        if d.year == today.year:
            return d.strftime("%A, %B %d"), ""
        return d.strftime("%A, %B %d, %Y"), ""

    date_groups: list[tuple] = []
    for it in items:
        d = datetime.fromtimestamp(it["mtime"]).date()
        if not date_groups or date_groups[-1][0] != d:
            date_groups.append((d, []))
        date_groups[-1][1].append(it)

    def render_file_row(entry):
        xlsx_path = entry["xlsx_path"]
        audit_paths = entry["audit_paths"]
        if not xlsx_path.exists():
            return ""
        stat = xlsx_path.stat()
        audit_btns = "".join(
            f'<a class="button-link secondary" href="/download-audit/{html.escape(a.name)}">Audit</a>'
            for a in audit_paths
        )
        return f"""
        <div class="file-row">
            <span class="file-icon xlsx">{_DOC_SVG}</span>
            <div class="file-info">
                <div class="file-name">{html.escape(xlsx_path.name)}</div>
                <div class="file-meta">
                    <span>{html.escape(fmt_size(stat.st_size))}</span>
                    <span class="dot">·</span>
                    <span>{html.escape(fmt_time(stat.st_mtime))}</span>
                </div>
            </div>
            <div class="file-actions">
                <a class="button-link" href="/download-output/{html.escape(xlsx_path.name)}">Download</a>
                {audit_btns}
            </div>
        </div>
        """

    def render_batch(batch):
        zip_path = batch["zip_path"]
        files = batch["files"]
        n_files = len(files)
        n_audits = sum(len(e["audit_paths"]) for e in files)
        total_size = zip_path.stat().st_size
        for e in files:
            if e["xlsx_path"].exists():
                total_size += e["xlsx_path"].stat().st_size
            for a in e["audit_paths"]:
                if a.exists():
                    total_size += a.stat().st_size

        if not files:
            inner = (
                '<p class="muted" style="margin:8px 0;">'
                'The individual workbooks for this batch are no longer on disk. '
                'You can still download the ZIP below.'
                '</p>'
            )
        else:
            inner = "".join(
                render_file_row(e)
                for e in sorted(files, key=lambda e: e["xlsx_path"].name.lower())
            )

        files_word = "workbook" if n_files == 1 else "workbooks"
        audits_word = "audit" if n_audits == 1 else "audits"
        time_label = fmt_time(batch["mtime"])

        return f"""
        <details class="folder-card" open>
            <summary>
                <span class="file-icon folder">{_FOLDER_SVG}</span>
                <div class="folder-info">
                    <div class="folder-name">Batch &middot; {n_files} {files_word} &middot; {html.escape(time_label)}</div>
                    <div class="folder-meta">
                        <span>{n_audits} {audits_word}</span>
                        <span class="dot">·</span>
                        <span>{html.escape(fmt_size(total_size))}</span>
                        <span class="dot">·</span>
                        <code>{html.escape(zip_path.name)}</code>
                    </div>
                </div>
                <div class="folder-actions">
                    <a class="button-link orange" href="/download-output/{html.escape(zip_path.name)}">Download all (ZIP)</a>
                </div>
            </summary>
            <div class="folder-contents">{inner}</div>
        </details>
        """

    def render_single(item):
        xlsx_path = item["xlsx_path"]
        audit_paths = item["audit_paths"]
        if not xlsx_path.exists():
            return ""
        stat = xlsx_path.stat()
        audit_btns = "".join(
            f'<a class="button-link secondary" href="/download-audit/{html.escape(a.name)}">Audit</a>'
            for a in audit_paths
        )
        audit_word = "audit attached" if audit_paths else "no audit on file"
        return f"""
        <div class="single-card">
            <span class="file-icon xlsx">{_DOC_SVG}</span>
            <div class="folder-info">
                <div class="folder-name">{html.escape(xlsx_path.name)}</div>
                <div class="folder-meta">
                    <span>Single workbook</span>
                    <span class="dot">·</span>
                    <span>{html.escape(fmt_size(stat.st_size))}</span>
                    <span class="dot">·</span>
                    <span>{html.escape(fmt_time(stat.st_mtime))}</span>
                    <span class="dot">·</span>
                    <span>{audit_word}</span>
                </div>
            </div>
            <div class="folder-actions">
                <a class="button-link orange" href="/download-output/{html.escape(xlsx_path.name)}">Download</a>
                {audit_btns}
            </div>
        </div>
        """

    def render_audit_only(item):
        audit_path = item["audit_path"]
        return f"""
        <div class="single-card audit-only">
            <span class="file-icon audit">{_AUDIT_SVG}</span>
            <div class="folder-info">
                <div class="folder-name">{html.escape(audit_path.name)}</div>
                <div class="folder-meta">
                    <span>Audit report only</span>
                    <span class="dot">·</span>
                    <span>workbook no longer on disk</span>
                    <span class="dot">·</span>
                    <span>{html.escape(fmt_size(audit_path.stat().st_size))}</span>
                    <span class="dot">·</span>
                    <span>{html.escape(fmt_time(item["mtime"]))}</span>
                </div>
            </div>
            <div class="folder-actions">
                <a class="button-link secondary" href="/download-audit/{html.escape(audit_path.name)}">Download audit</a>
            </div>
        </div>
        """

    if not date_groups:
        outputs_html = """
        <div class="empty-state">
            <h3 style="margin-bottom:8px;">No outputs yet</h3>
            <p style="margin:0;">Generate a workbook from the Home page and it will appear here.</p>
        </div>
        """
    else:
        sections = []
        for grp_date, grp_items in date_groups:
            primary, secondary = date_labels(grp_date)
            secondary_html = (
                f'<span class="date-meta">{html.escape(secondary)}</span>'
                if secondary else ""
            )
            item_word = "item" if len(grp_items) == 1 else "items"
            divider = f"""
            <div class="date-divider">
                <span class="date-label">{html.escape(primary)}</span>
                {secondary_html}
                <span class="date-count">{len(grp_items)} {item_word}</span>
            </div>
            """
            cards = []
            for it in grp_items:
                if it["kind"] == "batch":
                    cards.append(render_batch(it))
                elif it["kind"] == "single":
                    cards.append(render_single(it))
                else:
                    cards.append(render_audit_only(it))
            sections.append(divider + "".join(cards))
        outputs_html = '<div class="outputs-list">' + "".join(sections) + "</div>"

    return f"""
<!doctype html>
<html>
{head_html("Outputs | Phillip Capital Risk Management")}
<body>
<div class="shell">
    {topbar_html("outputs", user)}
    {hero_html(
        "Completed Outputs",
        "View and download generated Excel workbooks, batch ZIP files, and audit reports.",
    )}

    <div class="card">
        <h2>Completed Workbooks &amp; Audits</h2>
        <p class="muted">Batches are shown as folders containing each generated workbook. Single runs appear as individual cards. Everything is grouped by the day it was produced.</p>
        {outputs_html}
    </div>
</div>
</body>
</html>
    """
