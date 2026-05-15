"""文件处理路由。"""

import os
import threading
import time
import uuid
from pathlib import Path

from flask import jsonify, request

from docflow.settings import DEFAULT_PDF_MODE
from docflow.settings import normalize_pdf_mode as _normalize_pdf_mode
from docflow.support import augment_result_payload

from ..core import UPLOAD_FOLDER, app
from ..responses import error_response
from ..services.ocr import IMAGE_EXTS, process_image_ocr
from ..services.process_jobs import (
    PROCESS_JOBS,
    PROCESS_JOB_LOCK,
    _build_cancelled_process_result,
    _maybe_attach_invoice_fields,
    _maybe_save_invoice_record,
    _run_process_job,
    _serialize_process_job,
    processor,
)


def _get_uploaded_file():
    if "file" not in request.files:
        return None, error_response("没有收到文件")

    file = request.files["file"]
    if file.filename == "":
        return None, error_response("文件名为空")
    return file, None


@app.route("/process/start", methods=["POST"])
def start_process_file():
    file, err = _get_uploaded_file()
    if err:
        return err

    safe_name = Path(file.filename).name
    job_id = uuid.uuid4().hex[:12]
    temp_name = f"{job_id}_{safe_name}"
    save_path = os.path.join(UPLOAD_FOLDER, temp_name)
    file.save(save_path)

    output_format = request.form.get("format", "txt")
    pdf_mode = _normalize_pdf_mode(request.form.get("pdf_mode", DEFAULT_PDF_MODE))
    invoice_extract = request.form.get("invoice_extract", "0") == "1"
    force_reprocess = request.form.get("force_reprocess", "0") == "1"
    try:
        confidence_threshold = max(0.0, min(1.0, float(request.form.get("confidence_threshold", "0.0"))))
    except (ValueError, TypeError):
        confidence_threshold = 0.0
    file_ext = Path(safe_name).suffix.lower()
    now = time.time()

    with PROCESS_JOB_LOCK:
        PROCESS_JOBS[job_id] = {
            "job_id": job_id,
            "state": "queued",
            "progress_pct": 0.0,
            "stage": "queued",
            "message": "文件已入队，等待处理",
            "file_name": safe_name,
            "file_ext": file_ext,
            "save_path": save_path,
            "output_format": output_format,
            "pdf_mode": pdf_mode,
            "invoice_extract": invoice_extract,
            "confidence_threshold": confidence_threshold,
            "force_reprocess": force_reprocess,
            "result": None,
            "error": "",
            "cancel_requested": False,
            "created_at": now,
            "started_at": None,
            "updated_at": now,
            "finished_at": None,
        }

    thread = threading.Thread(target=_run_process_job, args=(job_id,), daemon=True)
    thread.start()

    return jsonify(
        {
            "success": True,
            "job_id": job_id,
            "state": "queued",
            "progress_pct": 0.0,
            "file_name": safe_name,
            "poll_url": f"/process/{job_id}",
        }
    )


@app.route("/process/<job_id>", methods=["GET"])
def get_process_job(job_id: str):
    data = _serialize_process_job(job_id)
    if not data:
        return error_response("处理任务不存在", status_code=404)
    return jsonify({"success": True, **data})


@app.route("/process/<job_id>/cancel", methods=["POST"])
def cancel_process_job(job_id: str):
    now = time.time()
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return error_response("处理任务不存在", status_code=404)

        state = job.get("state")
        if state in ("completed", "failed", "cancelled", "cancelling"):
            data = _serialize_process_job(job_id)
            if not data:
                return error_response("处理任务不存在", status_code=404)
            return jsonify({"success": True, **data})

        job["cancel_requested"] = True
        job["updated_at"] = now

        if state == "queued":
            job["state"] = "cancelled"
            job["progress_pct"] = 100.0
            job["stage"] = "cancelled"
            job["message"] = "任务已取消"
            job["error"] = "任务已取消"
            job["result"] = _build_cancelled_process_result(
                job.get("file_name", ""),
                job.get("file_ext", ""),
            )
            job["finished_at"] = now
        else:
            job["state"] = "cancelling"
            job["stage"] = "cancelling"
            job["message"] = "正在取消任务..."

    data = _serialize_process_job(job_id)
    if not data:
        return error_response("处理任务不存在", status_code=404)
    return jsonify({"success": True, **data})


@app.route("/process", methods=["POST"])
def process_file():
    file, err = _get_uploaded_file()
    if err:
        return err

    safe_name = Path(file.filename).name
    temp_name = f"sync_{uuid.uuid4().hex[:12]}_{safe_name}"
    save_path = os.path.join(UPLOAD_FOLDER, temp_name)
    file.save(save_path)

    output_format = request.form.get("format", "txt")
    pdf_mode = _normalize_pdf_mode(request.form.get("pdf_mode", DEFAULT_PDF_MODE))
    invoice_extract = request.form.get("invoice_extract", "0") == "1"
    force_reprocess = request.form.get("force_reprocess", "0") == "1"
    ext = Path(safe_name).suffix.lower()

    try:
        if ext in IMAGE_EXTS:
            result = process_image_ocr(save_path, safe_name, force_reprocess=force_reprocess)
            result = _maybe_attach_invoice_fields(result, invoice_extract)
        else:
            result = processor.process(
                save_path,
                extract_keywords=True,
                extract_invoice=invoice_extract,
                output_format=output_format,
                pdf_mode=pdf_mode,
            )
        result = augment_result_payload(
            result,
            file_name=safe_name,
            file_ext=ext,
            source="process",
        )
        result = _maybe_save_invoice_record(
            result,
            invoice_extract,
            save_path,
            safe_name,
            ext,
        )
    finally:
        try:
            os.remove(save_path)
        except Exception:
            pass

    return jsonify(result)
