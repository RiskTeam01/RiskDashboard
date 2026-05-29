import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import UploadFile

from app.config import UPLOAD_DIR, LOG_DIR, get_template_status
from app.utils import (
    BATCH_JOBS, BATCH_JOBS_LOCK,
    safe_filename, save_json_debug, store_job_result,
    build_full_console_text, create_batch_zip,
)
from app.automations.credit_worksheet.engine import CreditWorksheetEngine
from app.automations.credit_worksheet.audit import write_audit_report
from app.automations.net_capital.engine import NetCapitalEngine
from app.customers import find_or_create_customer, add_report_to_customer


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
    successful_output_paths: list[Path] = []
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
            output_path, field_results, occurrences_by_code, readability_info, summary = (
                engine.generate_excel_from_pdf(
                    pdf_path=upload_path,
                    recalculate_with_excel=recalc,
                )
            )

            audit_path, metrics = write_audit_report(
                user=user,
                original_filename=original_filename,
                output_path=output_path,
                template_name=template_name,
                field_results=field_results,
                readability_info=readability_info,
            )

            # Detect customer from code 13 (company name) using all_words from engine
            engine.log("")
            engine.log("=" * 100)
            engine.log("[NET CAPITAL AUTOMATION]")
            nc_engine = NetCapitalEngine()
            company_name = nc_engine.extract_company_name(engine.all_words)
            customer = find_or_create_customer(company_name or "Unknown")
            customer_id = customer["id"]
            engine.log(f"[CUSTOMER] Matched/created account: '{customer['name']}' (id={customer_id})")

            # Run Net Capital automation
            nc_workbook_path = None
            try:
                nc_workbook_path = nc_engine.run(
                    all_words=engine.all_words,
                    occurrences_by_code=occurrences_by_code,
                    customer_id=customer_id,
                    customer_name=customer["name"],
                )
            except Exception as nc_err:
                nc_engine.log(f"[NET CAPITAL] Error: {nc_err}")
            # Surface Net Capital log inside the main console output
            engine.debug_lines.extend(nc_engine.debug_lines)
            if nc_workbook_path:
                engine.log(f"[NET CAPITAL] Workbook ready: {nc_workbook_path.name}")
            else:
                engine.log("[NET CAPITAL] No workbook produced (see messages above).")
            engine.log("=" * 100)

            # File report under customer account
            add_report_to_customer(
                customer_id=customer_id,
                report_type="credit_worksheet",
                original_filename=original_filename,
                output_filename=output_path.name,
                audit_filename=audit_path.name,
                net_capital_filename=nc_workbook_path.name if nc_workbook_path else "",
            )

            log_path = engine.save_debug_log(job_id, suffix=f"pdf_{index}")

            item_payload = {
                "status": "SUCCESS",
                "job_id": f"{job_id}_{index}",
                "original_filename": original_filename,
                "uploaded_pdf": str(upload_path),
                "output_path": str(output_path),
                "output_filename": output_path.name,
                "audit_filename": audit_path.name,
                "net_capital_filename": nc_workbook_path.name if nc_workbook_path else "",
                "customer_id": customer_id,
                "customer_name": customer["name"],
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

            successful_output_paths.append(output_path)
            batch_results.append(item_payload)

            update_file_entry(
                job_id, index,
                status="complete",
                output_filename=output_path.name,
                download_url=f"/download-output/{output_path.name}",
                audit_url=f"/download-audit/{audit_path.name}",
                fields_found=metrics["fields_found_label"],
                needs_review=metrics["needs_review"],
                valid_blanks=metrics["valid_blanks"],
                customer_id=customer_id,
                customer_name=customer["name"],
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
        "Output naming rule: each Excel file is named after its source PDF.",
        "An audit report (.txt) is written alongside each successful workbook.",
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

    zip_url = None
    single_url = None
    if len(successful_output_paths) > 1:
        zip_path = create_batch_zip(successful_output_paths, job_id)
        zip_url = f"/download-output/{zip_path.name}"
    elif len(successful_output_paths) == 1:
        single_url = f"/download-output/{successful_output_paths[0].name}"

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
