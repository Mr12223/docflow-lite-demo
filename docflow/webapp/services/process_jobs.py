"""文件处理任务服务。"""

import os
import threading
import time
from typing import Optional

from docflow.core import DocFlowCancelledError, DocFlowProcessor
from docflow.support import augment_result_payload, build_error_info

from ..core import invoice_db, logger
from .ocr import IMAGE_EXTS, extract_invoice_fields, process_image_ocr

PROCESS_JOBS = {}
PROCESS_JOB_LOCK = threading.Lock()

# 字段置信度字符串到数值的映射（内部合并逻辑使用，不对外暴露）
_CONF_SCORE = {"high": 0.9, "medium": 0.6, "low": 0.3, "none": 0.0}


processor = DocFlowProcessor()


def _update_process_job(job_id: str, **fields) -> None:
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return
        if "progress_pct" in fields:
            try:
                fields["progress_pct"] = round(max(0.0, min(float(fields["progress_pct"]), 100.0)), 2)
            except Exception:
                fields["progress_pct"] = job.get("progress_pct", 0.0)
        job.update(fields)
        job["updated_at"] = time.time()


def _serialize_process_job(job_id: str) -> Optional[dict]:
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return None
        data = dict(job)
    if isinstance(data.get("result"), dict):
        data["result"] = dict(data["result"])
    return data


def _is_process_job_cancel_requested(job_id: str) -> bool:
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return False
        return bool(job.get("cancel_requested"))


def _build_cancelled_process_result(file_name: str, file_ext: str) -> dict:
    message = "任务已取消"
    return {
        "success": False,
        "cancelled": True,
        "error": message,
        "error_info": build_error_info(
            message,
            file_name=file_name,
            file_ext=file_ext,
            metadata_dict={},
            source="process",
        ),
    }


def _maybe_attach_invoice_fields(result: dict, enabled: bool) -> dict:
    if not enabled or not isinstance(result, dict) or not result.get("success"):
        return result

    text = str(result.get("text") or "").strip()
    if not text:
        return result

    stats = result.setdefault("statistics", {})
    if isinstance(stats.get("invoice_fields"), dict):
        return result

    try:
        inv_data = extract_invoice_fields(text)
        stats["invoice_fields"] = inv_data
    except Exception as exc:
        logger.warning("发票提取失败: %s", exc)
        stats["invoice_fields"] = {
            "is_invoice": False,
            "confidence": "none",
            "field_count": 0,
            "fields": {},
        }
    return result


def _maybe_save_invoice_record(
    result: dict,
    enabled: bool,
    save_path: str,
    file_name: str,
    file_ext: str,
) -> dict:
    if not enabled or not isinstance(result, dict) or not result.get("success"):
        return result

    inv_data = (result.get("statistics") or {}).get("invoice_fields") or {}
    if not isinstance(inv_data, dict) or not inv_data.get("is_invoice"):
        return result

    try:
        file_size = 0
        try:
            file_size = os.path.getsize(save_path)
        except Exception:
            pass
        record_id = invoice_db.save_record(
            fields=inv_data.get("fields", {}),
            file_name=file_name,
            file_type=file_ext,
            file_size=file_size,
            confidence=inv_data.get("confidence", "low"),
            field_count=inv_data.get("field_count", 0),
            raw_text=result.get("text", ""),
        )
        inv_data["saved_record_id"] = record_id
    except Exception as exc:
        logger.warning("发票记录保存失败: %s", exc)

    return result


def _run_process_job(job_id: str) -> None:
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return
        file_name = job["file_name"]
        save_path = job["save_path"]
        output_format = job["output_format"]
        pdf_mode = job["pdf_mode"]
        file_ext = job["file_ext"]
        invoice_extract = job.get("invoice_extract", False)
        force_reprocess = job.get("force_reprocess", False)
        if job.get("state") == "cancelled" or job.get("cancel_requested"):
            PROCESS_JOBS[job_id]["state"] = "cancelled"
            PROCESS_JOBS[job_id]["progress_pct"] = 100.0
            PROCESS_JOBS[job_id]["stage"] = "cancelled"
            PROCESS_JOBS[job_id]["message"] = "任务已取消"
            PROCESS_JOBS[job_id]["error"] = "任务已取消"
            PROCESS_JOBS[job_id]["result"] = _build_cancelled_process_result(file_name, file_ext)
            PROCESS_JOBS[job_id]["finished_at"] = time.time()
            PROCESS_JOBS[job_id]["updated_at"] = time.time()
            return
        PROCESS_JOBS[job_id]["state"] = "running"
        PROCESS_JOBS[job_id]["started_at"] = time.time()
        PROCESS_JOBS[job_id]["message"] = "任务已启动"
        PROCESS_JOBS[job_id]["stage"] = "running"
        PROCESS_JOBS[job_id]["updated_at"] = time.time()

    job_processor = DocFlowProcessor()
    cancelled_message = "任务已取消"

    def cancel_requested() -> bool:
        return _is_process_job_cancel_requested(job_id)

    def report(progress_pct: float, stage: str, message: str = "", **extra) -> None:
        payload = {
            "progress_pct": progress_pct,
            "stage": stage,
            "message": message,
        }
        payload.update(extra)
        _update_process_job(job_id, **payload)
        if cancel_requested():
            raise DocFlowCancelledError(cancelled_message)

    try:
        if cancel_requested():
            raise DocFlowCancelledError(cancelled_message)

        if file_ext in IMAGE_EXTS:
            result = process_image_ocr(
                save_path,
                file_name,
                progress_callback=report,
                cancel_callback=cancel_requested,
                force_reprocess=force_reprocess,
                extract_invoice=invoice_extract,
            )
            result = _maybe_attach_invoice_fields(result, invoice_extract)
        else:
            result = job_processor.process(
                save_path,
                extract_keywords=True,
                extract_invoice=invoice_extract,
                output_format=output_format,
                pdf_mode=pdf_mode,
                progress_callback=report,
                cancel_callback=cancel_requested,
            )

        if cancel_requested():
            raise DocFlowCancelledError(cancelled_message)

        result = augment_result_payload(result, file_name=file_name, file_ext=file_ext, source="process")
        result = _maybe_save_invoice_record(result, invoice_extract, save_path, file_name, file_ext)

        final_state = "completed" if result.get("success") else "failed"
        _update_process_job(
            job_id,
            state=final_state,
            progress_pct=100.0,
            stage="done" if final_state == "completed" else "error",
            message="处理完成" if final_state == "completed" else (result.get("error") or "处理失败"),
            result=result,
            error="" if final_state == "completed" else (result.get("error") or "处理失败"),
            finished_at=time.time(),
        )
    except DocFlowCancelledError:
        _update_process_job(
            job_id,
            state="cancelled",
            progress_pct=100.0,
            stage="cancelled",
            message=cancelled_message,
            error=cancelled_message,
            result=_build_cancelled_process_result(file_name, file_ext),
            finished_at=time.time(),
        )
    except Exception as exc:
        _update_process_job(
            job_id,
            state="failed",
            progress_pct=100.0,
            stage="error",
            message=str(exc),
            error=str(exc),
            result={
                "success": False,
                "error": str(exc),
                "error_info": build_error_info(
                    str(exc),
                    file_name=file_name,
                    file_ext=file_ext,
                    metadata_dict={},
                    source="process",
                ),
            },
            finished_at=time.time(),
        )
    finally:
        try:
            os.remove(save_path)
        except Exception:
            pass
