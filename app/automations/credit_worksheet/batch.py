import re
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import UploadFile

from app.config import UPLOAD_DIR, LOG_DIR, get_template_status, SHEET_NAME
from app.utils import (
    BATCH_JOBS, BATCH_JOBS_LOCK,
    safe_filename, save_json_debug, store_job_result,
    build_full_console_text, create_batch_zip,
)
from app.automations.credit_worksheet.engine import CreditWorksheetEngine
from app.automations.credit_worksheet.audit import write_audit_report
from app.automations.net_capital.engine import NetCapitalEngine
from app.automations.net_capital.fields import NET_CAPITAL_ROW_MAP, MONTH_COLUMNS, MONTH_NAMES
from app.automations.combined_workbook import (
    customer_workbook_path, open_or_create,
    add_credit_month_sheet, add_or_update_net_capital_sheet,
    set_recalc_on_open,
)
from app.customers import find_or_create_customer, add_report_to_customer
from app.config import NET_CAPITAL_DIR, detect_backend_template


async def stream_save_pdf(
    pdf: UploadFile,
    batch_id: str,
    index: int,
    max_bytes: int,
) -> tuple[Optional[Path], Optional[str]]:
    if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
        return None, "Not a PDF file (only .pdf is accepted)."

    upload_name = safe_filename(pdf.filename)
    upload_path = UPLOAD_DIR / f"{batch_id}_{index}_{upload_name}"

    size = 0
    try:
        with open(upload_path, "wb") as f:
            while True:
                chunk = await pdf.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                if size > max_bytes:
                    f.close()
                    upload_path.unlink(missing_ok=True)
                    return None, f"Exceeds the {max_bytes // (1024 * 1024)} MB size limit."
                f.write(chunk)
    except Exception as e:
        upload_path.unlink(missing_ok=True)
        return None, f"Could not save upload: {e}"

    if size == 0:
        upload_path.unlink(missing_ok=True)
        return None, "File is empty."

    return upload_path, None


def update_file_entry(job_id: str, index: int, **changes):
    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        if not job:
            return
        for entry in job["files"]:
            if entry["index"] == index:
                entry.update(changes)
                break


def run_batch_job(job_id: str):
    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        if not job:
            return
        queue = list(job.get("_queue", []))
        recalc = job.get("_recalc", False)
        user = job.get("user", "unknown")

    batch_results = []
    successful_workbook_paths: list[Path] = []
    combined_debug_lines = []
    template_status = get_template_status()
    template_name = template_status.get("filename", "")

    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        for entry in job["files"]:
            if entry["status"] == "failed":
                batch_results.append({
                    "status": "FAILED",
                    "job_id": f"{job_id}_{entry['index']}",
                    "original_filename": entry["filename"],
                    "error": entry.get("error", "Validation failed"),
                    "debug_log": [f"Validation failed: {entry.get('error', '')}"],
                })

    for index, original_filename, upload_path in queue:
        update_file_entry(job_id, index, status="processing")
        engine = CreditWorksheetEngine()

        try:
            # ── 1. Extract PDF ──────────────────────────────────────────────
            credit_template_path = detect_backend_template()
            engine.log("[TEMPLATE] Auto-detected backend template:")
            engine.log(str(credit_template_path))
            engine.validate_template(credit_template_path)

            field_results, occurrences_by_code, readability_info, summary = (
                engine.extract_pdf_field_results(upload_path)
            )
            engine.validate_no_missing(field_results)

            # ── 2. Customer detection ───────────────────────────────────────
            engine.log("")
            engine.log("=" * 100)
            engine.log("[CUSTOMER & NET CAPITAL]")

            nc_engine = NetCapitalEngine()
            company_name = nc_engine.extract_company_name(engine.all_words)
            customer = find_or_create_customer(company_name or "Unknown")
            customer_id = customer["id"]
            engine.log(f"[CUSTOMER] '{customer['name']}' (id={customer_id})")

            # ── 3. Determine period month/year ──────────────────────────────
            date_text = nc_engine.extract_period_end_date(engine.all_words)
            month: Optional[int] = None
            year: int = datetime.now().year

            if date_text:
                month = nc_engine.determine_month(date_text)
                year_match = re.search(r"\d{4}", date_text)
                if year_match:
                    year = int(year_match.group(0))
                elif re.search(r"\d{2}$", date_text.split("/")[-1] if "/" in date_text else ""):
                    short_year = int(date_text.split("/")[-1])
                    year = 2000 + short_year if short_year < 100 else short_year

            if not month:
                engine.log(f"[WARN] Could not determine month from '{date_text}'. Defaulting to current month.")
                month = datetime.now().month

            engine.log(f"[PERIOD] month={month} ({MONTH_NAMES[month-1]}), year={year}, date='{date_text}'")

            # ── 4. Build Net Capital occurrence map ─────────────────────────
            nc_occurrences = nc_engine.build_occurrences(engine.all_words)

            # ── 5. Open/create combined customer workbook ───────────────────
            NET_CAPITAL_DIR.mkdir(parents=True, exist_ok=True)
            wb_path = customer_workbook_path(customer["name"], year)
            wb = open_or_create(wb_path)
            engine.log(f"[WORKBOOK] {'Opened' if wb_path.exists() else 'Creating'}: {wb_path.name}")

            # ── 6. Add credit month sheet ───────────────────────────────────
            credit_sheet = add_credit_month_sheet(
                wb=wb,
                month=month,
                year=year,
                credit_template_path=credit_template_path,
                credit_template_sheet=SHEET_NAME,
                field_results=field_results,
                log=engine.log,
            )

            # ── 7. Add/update Net Capital sheet ────────────────────────────
            from app.automations.net_capital.engine import detect_net_capital_template
            nc_template = detect_net_capital_template()
            nc_sheet = add_or_update_net_capital_sheet(
                wb=wb,
                month=month,
                year=year,
                company_name=customer["name"],
                date_text=date_text,
                net_capital_template_path=nc_template,
                occurrences_by_code=nc_occurrences,
                row_map=NET_CAPITAL_ROW_MAP,
                log=engine.log,
            )

            set_recalc_on_open(wb)
            wb.save(wb_path)
            engine.log(f"[WORKBOOK] Saved: {wb_path}")
            engine.debug_lines.extend(nc_engine.debug_lines)
            engine.log("=" * 100)

            # ── 8. Audit report ─────────────────────────────────────────────
            audit_path, metrics = write_audit_report(
                user=user,
                original_filename=original_filename,
                output_path=wb_path,
                template_name=template_name,
                field_results=field_results,
                readability_info=readability_info,
            )

            log_path = engine.save_debug_log(job_id, suffix=f"pdf_{index}")

            # ── 9. File report under customer ───────────────────────────────
            add_report_to_customer(
                customer_id=customer_id,
                report_type="credit_worksheet",
                original_filename=original_filename,
                output_filename=wb_path.name,
                period_label=f"{MONTH_NAMES[month-1]} {year}",
                audit_filename=audit_path.name,
                net_capital_filename="",  # same workbook now
                credit_sheet=credit_sheet,
                net_capital_sheet=nc_sheet,
            )

            successful_workbook_paths.append(wb_path)

            item_payload = {
                "status": "SUCCESS",
                "job_id": f"{job_id}_{index}",
                "original_filename": original_filename,
                "uploaded_pdf": str(upload_path),
                "workbook_path": str(wb_path),
                "workbook_filename": wb_path.name,
                "credit_sheet": credit_sheet,
                "net_capital_sheet": nc_sheet,
                "audit_filename": audit_path.name,
                "customer_id": customer_id,
                "customer_name": customer["name"],
                "period_label": f"{MONTH_NAMES[month-1]} {year}",
                "template": template_status,
                "readability": readability_info,
                "summary": summary,
                "field_results": engine.serialize_field_results(field_results),
                "raw_occurrences": engine.serialize_occurrences(occurrences_by_code),
                "metrics": metrics,
                "debug_log_filename": log_path.name,
                "debug_log_download_url": f"/download-log/{log_path.name}",
                "debug_log": engine.debug_lines,
            }

            batch_results.append(item_payload)

            update_file_entry(
                job_id, index,
                status="complete",
                output_filename=wb_path.name,
                download_url=f"/download-net-capital/{wb_path.name}",
                audit_url=f"/download-audit/{audit_path.name}",
                fields_found=metrics["fields_found_label"],
                needs_review=metrics["needs_review"],
                valid_blanks=metrics["valid_blanks"],
                customer_id=customer_id,
                customer_name=customer["name"],
                period_label=f"{MONTH_NAMES[month-1]} {year}",
                credit_sheet=credit_sheet,
            )

        except Exception as e:
            engine.log("[FATAL ERROR]")
            engine.log(traceback.format_exc())
            engine.save_debug_log(job_id, suffix=f"pdf_{index}_failed")

            batch_results.append({
                "status": "FAILED",
                "job_id": f"{job_id}_{index}",
                "original_filename": original_filename,
                "error": str(e),
                "debug_log": engine.debug_lines,
            })
            update_file_entry(job_id, index, status="failed", error=str(e))

        combined_debug_lines.extend([f"===== {original_filename} ====="])
        combined_debug_lines.extend(engine.debug_lines)
        combined_debug_lines.append("")

        with BATCH_JOBS_LOCK:
            job = BATCH_JOBS.get(job_id)
            if job:
                job["success_count"] = sum(1 for f in job["files"] if f["status"] == "complete")
                job["failure_count"] = sum(1 for f in job["files"] if f["status"] == "failed")

    success_count = sum(1 for item in batch_results if item.get("status") == "SUCCESS")
    failure_count = sum(1 for item in batch_results if item.get("status") == "FAILED")
    submitted_count = len(batch_results)

    batch_summary = [
        f"Batch ID: {job_id}",
        f"Created by: {user}",
        f"PDFs submitted: {submitted_count}",
        f"Successful outputs: {success_count}",
        f"Failed outputs: {failure_count}",
        "",
        "All outputs are written into per-customer combined workbooks.",
        "Each workbook contains a credit sheet per month and one accumulating Net Capital sheet.",
    ]

    payload = {
        "type": "batch" if submitted_count > 1 else "single",
        "job_id": job_id,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "submitted_count": submitted_count,
        "success_count": success_count,
        "failure_count": failure_count,
        "batch_summary": batch_summary,
        "batch_results": batch_results,
        "debug_log": combined_debug_lines,
    }

    if submitted_count == 1 and batch_results and batch_results[0].get("status") == "SUCCESS":
        payload.update(batch_results[0])
        payload["type"] = "single"

    try:
        save_json_debug(payload, job_id)
        store_job_result(job_id, payload)
        combined_log_path = (
            LOG_DIR / f"credit_ws_batch_debug_{job_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        combined_log_path.write_text(build_full_console_text(payload), encoding="utf-8")
    except Exception:
        pass

    # Deduplicate workbook paths for zip (multiple PDFs may write to the same file)
    unique_paths = list({str(p): p for p in successful_workbook_paths}.values())
    zip_url = None
    single_url = None
    if len(unique_paths) > 1:
        zip_path = create_batch_zip(unique_paths, job_id)
        zip_url = f"/download-net-capital/{zip_path.name}"
    elif len(unique_paths) == 1:
        single_url = f"/download-net-capital/{unique_paths[0].name}"

    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        if job:
            job["status"] = "complete"
            job["success_count"] = success_count
            job["failure_count"] = failure_count
            job["zip_url"] = zip_url
            job["single_url"] = single_url
            job["console_url"] = f"/console?job_id={job_id}"
            job.pop("_queue", None)
            job.pop("_recalc", None)
