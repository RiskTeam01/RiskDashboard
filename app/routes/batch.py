import threading
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from app.auth import require_api_user
from app.config import load_config, get_template_status
from app.utils import (
    BATCH_JOBS, BATCH_JOBS_LOCK,
    save_json_debug, store_job_result,
    create_batch_zip,
)
from app.automations.credit_worksheet.batch import (
    stream_save_pdf, run_batch_job,
)
from app.automations.credit_worksheet.engine import CreditWorksheetEngine
from app.automations.credit_worksheet.audit import write_audit_report

router = APIRouter()


@router.post("/start-batch")
async def start_batch(
    request: Request,
    pdfs: list[UploadFile] = File(...),
    recalculate_with_excel: bool = Form(False),
):
    user = require_api_user(request)
    config = load_config()

    if not pdfs:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF.")

    if len(pdfs) > config["max_batch_size"]:
        raise HTTPException(
            status_code=400,
            detail=f"Too many files. The batch limit is {config['max_batch_size']} PDFs.",
        )

    batch_id = str(uuid.uuid4())
    max_bytes = config["max_pdf_size_mb"] * 1024 * 1024

    file_entries = []
    queue = []

    for index, pdf in enumerate(pdfs, start=1):
        original_filename = pdf.filename or f"file_{index}.pdf"
        upload_path, error = await stream_save_pdf(pdf, batch_id, index, max_bytes)

        if error:
            file_entries.append({
                "index": index,
                "filename": original_filename,
                "status": "failed",
                "error": error,
            })
        else:
            file_entries.append({
                "index": index,
                "filename": original_filename,
                "status": "queued",
            })
            queue.append((index, original_filename, upload_path))

    job = {
        "job_id": batch_id,
        "user": user,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "status": "running",
        "total": len(pdfs),
        "success_count": 0,
        "failure_count": sum(1 for f in file_entries if f["status"] == "failed"),
        "files": file_entries,
        "zip_url": None,
        "single_url": None,
        "console_url": f"/console?job_id={batch_id}",
        "_queue": queue,
        "_recalc": bool(recalculate_with_excel),
    }

    with BATCH_JOBS_LOCK:
        BATCH_JOBS[batch_id] = job
        if len(BATCH_JOBS) > 40:
            for key in list(BATCH_JOBS.keys())[:-40]:
                BATCH_JOBS.pop(key, None)

    if queue:
        thread = threading.Thread(target=run_batch_job, args=(batch_id,), daemon=True)
        thread.start()
    else:
        with BATCH_JOBS_LOCK:
            job["status"] = "complete"
            job.pop("_queue", None)
            job.pop("_recalc", None)

    return JSONResponse({"job_id": batch_id})


@router.get("/batch-status/{job_id}")
def batch_status(request: Request, job_id: str):
    require_api_user(request)
    with BATCH_JOBS_LOCK:
        job = BATCH_JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Batch job not found.")
        public = {k: v for k, v in job.items() if not k.startswith("_")}
    return JSONResponse(public)


@router.post("/generate-excel")
async def generate_excel(
    request: Request,
    pdfs: list[UploadFile] = File(...),
    recalculate_with_excel: bool = Form(False),
):
    """Backward-compatible synchronous endpoint."""
    user = require_api_user(request)
    config = load_config()

    batch_id = str(uuid.uuid4())
    batch_created_at = datetime.now().isoformat(timespec="seconds")

    if not pdfs:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF.")

    if len(pdfs) > config["max_batch_size"]:
        raise HTTPException(status_code=400, detail=f"Batch limit is {config['max_batch_size']} PDFs.")

    max_bytes = config["max_pdf_size_mb"] * 1024 * 1024

    batch_results = []
    successful_output_paths: list[Path] = []
    combined_debug_lines = []
    template_status = get_template_status()
    template_name = template_status.get("filename", "")

    for index, pdf in enumerate(pdfs, start=1):
        engine = CreditWorksheetEngine()
        original_filename = pdf.filename or f"file_{index}.pdf"

        upload_path, error = await stream_save_pdf(pdf, batch_id, index, max_bytes)

        if error:
            batch_results.append({
                "status": "FAILED",
                "original_filename": original_filename,
                "error": error,
                "debug_log": [f"Validation failed: {error}"],
            })
            combined_debug_lines.extend([f"===== {original_filename} =====", f"Validation failed: {error}", ""])
            continue

        try:
            output_path, field_results, occurrences_by_code, readability_info, summary = (
                engine.generate_excel_from_pdf(
                    pdf_path=upload_path,
                    recalculate_with_excel=recalculate_with_excel,
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

            log_path = engine.save_debug_log(batch_id, suffix=f"pdf_{index}")

            batch_results.append({
                "status": "SUCCESS",
                "original_filename": original_filename,
                "output_filename": output_path.name,
                "audit_filename": audit_path.name,
                "template": template_status,
                "readability": readability_info,
                "summary": summary,
                "field_results": engine.serialize_field_results(field_results),
                "raw_occurrences": engine.serialize_occurrences(occurrences_by_code),
                "metrics": metrics,
                "debug_log": engine.debug_lines,
            })
            successful_output_paths.append(output_path)

        except Exception as e:
            engine.log("[FATAL ERROR]")
            engine.log(traceback.format_exc())
            engine.save_debug_log(batch_id, suffix=f"pdf_{index}_failed")
            batch_results.append({
                "status": "FAILED",
                "original_filename": original_filename,
                "error": str(e),
                "debug_log": engine.debug_lines,
            })

        combined_debug_lines.extend([f"===== {original_filename} ====="])
        combined_debug_lines.extend(engine.debug_lines)
        combined_debug_lines.append("")

    success_count = sum(1 for item in batch_results if item.get("status") == "SUCCESS")
    failure_count = sum(1 for item in batch_results if item.get("status") == "FAILED")

    payload = {
        "type": "batch" if len(pdfs) > 1 else "single",
        "job_id": batch_id,
        "created_at": batch_created_at,
        "submitted_count": len(pdfs),
        "success_count": success_count,
        "failure_count": failure_count,
        "batch_summary": [
            f"Batch ID: {batch_id}",
            f"Created by: {user}",
            f"PDFs submitted: {len(pdfs)}",
            f"Successful outputs: {success_count}",
            f"Failed outputs: {failure_count}",
        ],
        "batch_results": batch_results,
        "debug_log": combined_debug_lines,
    }

    if len(pdfs) == 1 and batch_results and batch_results[0].get("status") == "SUCCESS":
        payload.update(batch_results[0])
        payload["type"] = "single"

    save_json_debug(payload, batch_id)
    store_job_result(batch_id, payload)

    if success_count == 0:
        raise HTTPException(status_code=500, detail="No PDFs could be processed.")

    if len(successful_output_paths) == 1:
        output_path = successful_output_paths[0]
        return FileResponse(
            path=str(output_path),
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "X-Job-ID": batch_id,
                "X-Success-Count": str(success_count),
                "X-Failure-Count": str(failure_count),
            },
        )

    zip_path = create_batch_zip(successful_output_paths, batch_id)
    return FileResponse(
        path=str(zip_path),
        filename=zip_path.name,
        media_type="application/zip",
        headers={
            "X-Job-ID": batch_id,
            "X-Success-Count": str(success_count),
            "X-Failure-Count": str(failure_count),
        },
    )
