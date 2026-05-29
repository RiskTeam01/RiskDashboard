import json
import re
import threading
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import HTTPException

from app.config import LOG_DIR, OUTPUT_DIR

JOB_RESULTS: dict[str, dict] = {}
LATEST_JOB_ID: Optional[str] = None

BATCH_JOBS: dict[str, dict] = {}
BATCH_JOBS_LOCK = threading.Lock()


def safe_filename(filename: str) -> str:
    filename = Path(filename).name
    filename = re.sub(r"[^A-Za-z0-9_.\- ]+", "_", filename).strip()
    return filename or "uploaded_file"


def save_json_debug(payload: dict, job_id: str) -> Path:
    path = LOG_DIR / f"credit_ws_payload_{job_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return path


def store_job_result(job_id: str, payload: dict):
    global LATEST_JOB_ID
    JOB_RESULTS[job_id] = payload
    LATEST_JOB_ID = job_id
    if len(JOB_RESULTS) > 25:
        oldest_keys = list(JOB_RESULTS.keys())[:-25]
        for key in oldest_keys:
            JOB_RESULTS.pop(key, None)


def get_job_payload(job_id: Optional[str] = None) -> dict:
    if job_id:
        if job_id not in JOB_RESULTS:
            raise HTTPException(status_code=404, detail="Job result not found in current server memory.")
        return JOB_RESULTS[job_id]
    if LATEST_JOB_ID and LATEST_JOB_ID in JOB_RESULTS:
        return JOB_RESULTS[LATEST_JOB_ID]
    json_files = sorted(LOG_DIR.glob("credit_ws_payload_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not json_files:
        return {}
    return json.loads(json_files[0].read_text(encoding="utf-8"))


def safe_js_json(payload: dict) -> str:
    return json.dumps(payload, default=str).replace("</", "<\\/")


def build_single_run_console_text(payload: dict) -> str:
    field_rows = [[
        "PDF Code(s)", "Field", "Excel Cell", "Preview Value",
        "Status", "Difficulty", "Notes", "Component Details",
    ]]
    for r in payload.get("field_results", []):
        field_rows.append([
            r.get("expression", ""), r.get("label", ""), r.get("excel_cell", ""),
            r.get("display_value", ""), r.get("status", ""), r.get("difficulty", ""),
            r.get("notes", ""), r.get("component_details", ""),
        ])

    raw_rows = [[
        "Code", "Selected", "Page", "X0", "Y0", "X1", "Y1",
        "Nearby Amount", "Confidence Score", "Note", "Nearby Context",
    ]]
    for r in payload.get("raw_occurrences", []):
        raw_rows.append([
            r.get("code", ""), "YES" if r.get("selected") else "NO",
            r.get("page_number", ""),
            round(float(r.get("x0", 0)), 2), round(float(r.get("y0", 0)), 2),
            round(float(r.get("x1", 0)), 2), round(float(r.get("y1", 0)), 2),
            r.get("nearby_amount_text", ""), r.get("confidence_score", ""),
            r.get("note", ""), r.get("nearby_context", ""),
        ])

    def clean(value):
        if value is None:
            return ""
        return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ").strip()

    def to_tsv(rows):
        return "\n".join("\t".join(clean(cell) for cell in row) for row in rows)

    return (
        "SUMMARY\n"
        + "\n".join(payload.get("summary", []))
        + "\n\nFIELD RESULTS\n"
        + to_tsv(field_rows)
        + "\n\nRAW HELPER CODE OCCURRENCES\n"
        + to_tsv(raw_rows)
        + "\n\nDEBUG LOG\n"
        + "\n".join(payload.get("debug_log", []))
    )


def build_full_console_text(payload: dict) -> str:
    if not payload:
        return "No console data yet."

    if payload.get("type") == "batch":
        lines = []
        lines.append("BATCH SUMMARY")
        lines.append(f"Batch ID: {payload.get('job_id', '')}")
        lines.append(f"Created: {payload.get('created_at', '')}")
        lines.append(f"PDFs submitted: {payload.get('submitted_count', 0)}")
        lines.append(f"Successful outputs: {payload.get('success_count', 0)}")
        lines.append(f"Failed outputs: {payload.get('failure_count', 0)}")
        lines.append("")
        for index, item in enumerate(payload.get("batch_results", []), start=1):
            lines.append("=" * 100)
            lines.append(f"PDF {index}: {item.get('original_filename', '')}")
            lines.append(f"Status: {item.get('status', '')}")
            if item.get("status") == "SUCCESS":
                lines.append(f"Output: {item.get('output_filename', '')}")
                lines.append("")
                lines.append(build_single_run_console_text(item))
            else:
                lines.append(f"Error: {item.get('error', '')}")
                lines.append("")
                lines.append("DEBUG LOG")
                lines.extend(item.get("debug_log", []))
            lines.append("")
        return "\n".join(lines)

    return build_single_run_console_text(payload)


def create_batch_zip(output_paths: list[Path], batch_id: str) -> Path:
    zip_path = OUTPUT_DIR / f"Credit_Worksheet_Batch_{batch_id[:8]}.zip"
    counter = 2
    while zip_path.exists():
        zip_path = OUTPUT_DIR / f"Credit_Worksheet_Batch_{batch_id[:8]}_{counter}.zip"
        counter += 1
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for output_path in output_paths:
            zip_file.write(output_path, arcname=output_path.name)
    return zip_path


def get_file_size_label(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"
