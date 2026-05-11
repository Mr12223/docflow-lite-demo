"""批量测试路由。"""

import threading
import time
import uuid

from flask import jsonify, request

from docflow.settings import DEFAULT_PDF_MODE
from docflow.settings import normalize_pdf_mode as _normalize_pdf_mode

from ..core import app
from ..responses import error_response
from ..services.batch_jobs import (
    BATCH_TEST_JOBS,
    BATCH_TEST_LOCK,
    _append_job_log,
    _count_suite_cases,
    _resolve_batch_suites,
    _run_batch_test_job,
    _serialize_batch_job,
    _terminate_batch_process,
)


@app.route("/run-batch-tests", methods=["POST"])
def run_batch_tests():
    payload = request.get_json(silent=True) or {}
    suites = _resolve_batch_suites(payload.get("suites") or ["test_documents", "test_documents_edge_cases"])
    keywords = bool(payload.get("keywords"))
    strict = bool(payload.get("strict"))
    pdf_mode = _normalize_pdf_mode(payload.get("pdf_mode", DEFAULT_PDF_MODE))

    if not suites:
        return error_response("未选择有效的测试目录")

    job_id = uuid.uuid4().hex[:12]
    total = _count_suite_cases(suites)
    now = time.time()
    with BATCH_TEST_LOCK:
        BATCH_TEST_JOBS[job_id] = {
            "job_id": job_id,
            "state": "queued",
            "cancel_requested": False,
            "success": False,
            "command_ok": False,
            "return_code": None,
            "suites": suites,
            "keywords": keywords,
            "strict": strict,
            "pdf_mode": pdf_mode,
            "total": total,
            "current_index": 0,
            "completed_count": 0,
            "current_file": "",
            "current_suite": "",
            "summary": {},
            "records_count": 0,
            "failed_cases": [],
            "unexpected_cases": [],
            "report_dir": "",
            "report_urls": {},
            "logs": [],
            "log_seq": 0,
            "error": "",
            "created_at": now,
            "started_at": None,
            "updated_at": now,
            "finished_at": None,
        }

    thread = threading.Thread(target=_run_batch_test_job, args=(job_id,), daemon=True)
    thread.start()

    return jsonify(
        {
            "success": True,
            "job_id": job_id,
            "state": "queued",
            "total": total,
            "suites": suites,
            "pdf_mode": pdf_mode,
            "poll_url": f"/run-batch-tests/{job_id}",
        }
    )


@app.route("/run-batch-tests/<job_id>", methods=["GET"])
def get_batch_test_status(job_id: str):
    data = _serialize_batch_job(job_id)
    if data is None:
        return error_response("批量测试任务不存在", status_code=404)
    data["job_success"] = data.get("success", False)
    data["success"] = True
    return jsonify(data)


@app.route("/run-batch-tests/<job_id>/cancel", methods=["POST"])
def cancel_batch_test(job_id: str):
    now = time.time()
    with BATCH_TEST_LOCK:
        job = BATCH_TEST_JOBS.get(job_id)
        if not job:
            return error_response("批量测试任务不存在", status_code=404)

        state = job.get("state")
        if state in ("completed", "failed", "cancelled", "cancelling"):
            data = _serialize_batch_job(job_id)
            if data is None:
                return error_response("批量测试任务不存在", status_code=404)
            data["job_success"] = data.get("success", False)
            data["success"] = True
            return jsonify(data)

        job["cancel_requested"] = True
        job["updated_at"] = now

        if state == "queued":
            job["state"] = "cancelled"
            job["success"] = False
            job["error"] = "任务已取消"
            job["finished_at"] = now
        else:
            job["state"] = "cancelling"

    if state == "queued":
        _append_job_log(job_id, "WARN", "批量测试任务已取消")
    else:
        _append_job_log(job_id, "WARN", "正在取消批量测试任务...")
        _terminate_batch_process(job_id)

    data = _serialize_batch_job(job_id)
    if data is None:
        return error_response("批量测试任务不存在", status_code=404)
    data["job_success"] = data.get("success", False)
    data["success"] = True
    return jsonify(data)
