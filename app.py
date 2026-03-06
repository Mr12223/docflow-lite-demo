"""
app.py — DocFlow 后端服务
作用：充当网页和 docflow_core.py 之间的"中间人"
运行方式：python app.py
"""

import os
import sys
import base64
import hashlib
import json
import time
import subprocess
import threading
import uuid
import re
from pathlib import Path
from typing import Optional
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

# 修复 Windows 编码问题
if sys.platform == 'win32':
    import io
    if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if hasattr(sys.stderr, 'buffer') and not isinstance(sys.stderr, io.TextIOWrapper):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

from docflow_core import DocFlowCancelledError, DocFlowProcessor, import_with_base_fallback
from docflow_support import (
    augment_result_payload,
    build_error_info,
    collect_dependency_status,
    install_missing_dependencies,
    prepare_pytesseract,
)

PROJECT_ROOT = Path(__file__).resolve().parent
FRONTEND_DIR = PROJECT_ROOT / "frontend"
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
SAMPLE_DATA_DIR = PROJECT_ROOT / "sample_data"
BATCH_SUITE_ALIASES = {
    "test_documents": SAMPLE_DATA_DIR / "test_documents",
    "test_documents_edge_cases": SAMPLE_DATA_DIR / "test_documents_edge_cases",
}
_EASYOCR_READER_CACHE = None
_EASYOCR_READER_ERROR = None
IMAGE_OCR_CACHE_VERSION = "v1"
IMAGE_OCR_CACHE_DIR = PROJECT_ROOT / "uploads_temp" / "ocr_cache"
IMAGE_OCR_MEMORY_CACHE = {}
IMAGE_OCR_CACHE_LOCK = threading.Lock()


def _get_default_pdf_mode() -> str:
    value = str(os.getenv("DOCFLOW_DEFAULT_PDF_MODE", "balanced")).strip().lower()
    return value if value in {"accurate", "balanced", "fast"} else "balanced"


DEFAULT_PDF_MODE = _get_default_pdf_mode()


def _normalize_pdf_mode(value: str) -> str:
    value = str(value or "").strip().lower()
    return value if value in {"accurate", "balanced", "fast"} else DEFAULT_PDF_MODE


def _env_flag(name: str, default: bool = False) -> bool:
    value = str(os.getenv(name, "")).strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    try:
        value = int(str(os.getenv(name, "")).strip())
        return value if value > 0 else default
    except Exception:
        return default


def _get_image_ocr_order() -> list[str]:
    raw = str(os.getenv("DOCFLOW_IMAGE_OCR_ORDER", "tesseract,easyocr")).strip().lower()
    engines = [item.strip() for item in raw.split(",") if item.strip()]
    valid = [item for item in engines if item in {"tesseract", "easyocr"}]
    return valid or ["tesseract", "easyocr"]


def _get_easyocr_reader():
    global _EASYOCR_READER_CACHE, _EASYOCR_READER_ERROR
    if _EASYOCR_READER_CACHE is not None:
        return _EASYOCR_READER_CACHE
    if _EASYOCR_READER_ERROR is not None:
        raise _EASYOCR_READER_ERROR

    try:
        easyocr = import_with_base_fallback("easyocr")
        _EASYOCR_READER_CACHE = easyocr.Reader(["ch_sim", "en"], verbose=False)
        return _EASYOCR_READER_CACHE
    except Exception as exc:
        _EASYOCR_READER_ERROR = exc
        raise


def _prepare_image_for_tesseract(image_path: str):
    from PIL import Image, ImageOps

    max_long_edge = _env_int("DOCFLOW_IMAGE_OCR_MAX_LONG_EDGE", 1600)
    huge_trigger = _env_int("DOCFLOW_IMAGE_OCR_FAST_EDGE_TRIGGER", 2400)
    huge_long_edge = _env_int("DOCFLOW_IMAGE_OCR_HUGE_LONG_EDGE", 1280)
    apply_gray = _env_flag("DOCFLOW_IMAGE_OCR_GRAYSCALE", True)
    apply_autocontrast = _env_flag("DOCFLOW_IMAGE_OCR_AUTOCONTRAST", True)

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")

        original_size = image.size
        width, height = image.size
        long_edge = max(width, height)
        target_long_edge = huge_long_edge if long_edge >= huge_trigger else max_long_edge

        if long_edge > target_long_edge:
            scale = target_long_edge / max(long_edge, 1)
            image = image.resize(
                (max(1, int(width * scale)), max(1, int(height * scale))),
                Image.LANCZOS,
            )

        if apply_gray and image.mode != "L":
            image = ImageOps.grayscale(image)
        if apply_autocontrast:
            image = ImageOps.autocontrast(image)

        prepared = image.copy()

    return prepared, {
        "original_size": original_size,
        "prepared_size": prepared.size,
        "long_edge_target": target_long_edge,
        "grayscale": prepared.mode == "L",
    }


def _build_image_tesseract_config(base_config: str = "") -> str:
    parts = [str(base_config or "").strip()]
    psm = _env_int("DOCFLOW_IMAGE_OCR_PSM", 6)
    oem = os.getenv("DOCFLOW_IMAGE_OCR_OEM", "").strip()
    if psm:
        parts.append(f"--psm {psm}")
    if oem:
        parts.append(f"--oem {oem}")
    return " ".join(part for part in parts if part).strip()


def _clone_json_payload(payload):
    return json.loads(json.dumps(payload, ensure_ascii=False))


def _is_image_ocr_cache_enabled() -> bool:
    return _env_flag("DOCFLOW_ENABLE_IMAGE_OCR_CACHE", True)


def _compute_file_sha256(file_path: str) -> str:
    digest = hashlib.sha256()
    with open(file_path, "rb") as file_obj:
        while True:
            chunk = file_obj.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _get_image_ocr_profile() -> dict:
    result = {
        "version": IMAGE_OCR_CACHE_VERSION,
        "order": _get_image_ocr_order(),
        "max_long_edge": _env_int("DOCFLOW_IMAGE_OCR_MAX_LONG_EDGE", 1600),
        "fast_edge_trigger": _env_int("DOCFLOW_IMAGE_OCR_FAST_EDGE_TRIGGER", 2400),
        "huge_long_edge": _env_int("DOCFLOW_IMAGE_OCR_HUGE_LONG_EDGE", 1280),
        "grayscale": _env_flag("DOCFLOW_IMAGE_OCR_GRAYSCALE", True),
        "autocontrast": _env_flag("DOCFLOW_IMAGE_OCR_AUTOCONTRAST", True),
        "psm": _env_int("DOCFLOW_IMAGE_OCR_PSM", 6),
        "oem": os.getenv("DOCFLOW_IMAGE_OCR_OEM", "").strip(),
    }
    return result


def _build_image_ocr_cache_key(image_path: str) -> tuple[str, str, dict]:
    file_sha256 = _compute_file_sha256(image_path)
    profile = _get_image_ocr_profile()
    profile_json = json.dumps(profile, ensure_ascii=False, sort_keys=True)
    cache_key = hashlib.sha256(f"{file_sha256}|{profile_json}".encode("utf-8")).hexdigest()
    return cache_key, file_sha256, profile


def _get_image_ocr_cache_file(cache_key: str) -> Path:
    return IMAGE_OCR_CACHE_DIR / f"{cache_key}.json"


def _load_image_ocr_cache(cache_key: str) -> Optional[dict]:
    if not _is_image_ocr_cache_enabled():
        return None

    with IMAGE_OCR_CACHE_LOCK:
        cached_payload = IMAGE_OCR_MEMORY_CACHE.get(cache_key)
    if cached_payload:
        return _clone_json_payload(cached_payload)

    cache_file = _get_image_ocr_cache_file(cache_key)
    if not cache_file.exists():
        return None

    try:
        payload = json.loads(cache_file.read_text(encoding="utf-8"))
        if payload.get("version") != IMAGE_OCR_CACHE_VERSION:
            return None
        result = payload.get("result")
        if not isinstance(result, dict):
            return None
    except Exception:
        return None

    with IMAGE_OCR_CACHE_LOCK:
        IMAGE_OCR_MEMORY_CACHE[cache_key] = payload
    return _clone_json_payload(payload)


def _save_image_ocr_cache(cache_key: str, file_sha256: str, profile: dict, result: dict) -> None:
    if not _is_image_ocr_cache_enabled():
        return
    if not isinstance(result, dict):
        return
    if result.get("metadata", {}).get("engine") not in {"Tesseract", "EasyOCR"}:
        return

    payload = {
        "version": IMAGE_OCR_CACHE_VERSION,
        "saved_at": time.time(),
        "file_sha256": file_sha256,
        "profile": profile,
        "result": _clone_json_payload(result),
    }
    cache_file = _get_image_ocr_cache_file(cache_key)
    try:
        cache_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        with IMAGE_OCR_CACHE_LOCK:
            IMAGE_OCR_MEMORY_CACHE[cache_key] = payload
    except Exception:
        return


def _format_image_ocr_output(
    filename: str,
    engine_used: str,
    char_count: int,
    elapsed_ms: float,
    text: str,
    *,
    cache_hit: bool = False,
    cache_original_processing_ms: Optional[float] = None,
) -> str:
    cache_lines = [f"缓存命中: {'是' if cache_hit else '否'}"]
    if cache_hit and cache_original_processing_ms is not None:
        cache_lines.append(f"原始OCR耗时: {cache_original_processing_ms:.0f}ms")

    return f"""[图片 OCR 结果] {filename}
{'━' * 40}
OCR 引擎: {engine_used}
识别字符数: {char_count}
处理耗时: {elapsed_ms:.0f}ms
{chr(10).join(cache_lines)}

识别内容:
{'━' * 20}
{text}
"""

# ── 创建 Flask 应用
def _build_image_ocr_cache_meta(
    *,
    cache_hit: bool,
    file_sha256: str = "",
    profile: Optional[dict] = None,
    saved_at: Optional[float] = None,
    original_processing_ms: Optional[float] = None,
) -> dict:
    payload = {
        "enabled": _is_image_ocr_cache_enabled(),
        "hit": bool(cache_hit),
        "file_sha256": file_sha256,
        "profile": profile or {},
    }
    if saved_at is not None:
        payload["saved_at"] = saved_at
    if original_processing_ms is not None:
        payload["original_processing_ms"] = round(float(original_processing_ms), 2)
    return payload


def _restore_cached_image_ocr_result(filename: str, cached_payload: dict, elapsed_ms: float) -> Optional[dict]:
    if not isinstance(cached_payload, dict):
        return None

    cached_result = cached_payload.get("result")
    if not isinstance(cached_result, dict):
        return None

    result = _clone_json_payload(cached_result)
    metadata = result.setdefault("metadata", {})
    statistics = result.setdefault("statistics", {})
    engine_used = str(metadata.get("engine") or "")
    text = str(result.get("text") or "")
    char_count = int(statistics.get("char_count") or len(text))
    original_processing_ms = result.get("processing_ms")
    if original_processing_ms is None:
        original_processing_ms = statistics.get("processing_ms")

    result["success"] = True
    result["file"] = filename
    result["error"] = ""
    result["processing_ms"] = elapsed_ms
    statistics["char_count"] = char_count
    statistics["paragraph_count"] = int(statistics.get("paragraph_count") or len(text.splitlines()))
    statistics["table_count"] = int(statistics.get("table_count") or 0)
    statistics["processing_ms"] = elapsed_ms
    metadata["file"] = filename
    metadata["image_ocr_cache"] = _build_image_ocr_cache_meta(
        cache_hit=True,
        file_sha256=str(cached_payload.get("file_sha256") or ""),
        profile=cached_payload.get("profile") if isinstance(cached_payload.get("profile"), dict) else {},
        saved_at=cached_payload.get("saved_at"),
        original_processing_ms=original_processing_ms,
    )
    result["formatted_output"] = _format_image_ocr_output(
        filename,
        engine_used,
        char_count,
        elapsed_ms,
        text,
        cache_hit=True,
        cache_original_processing_ms=original_processing_ms,
    )
    return result


app = Flask(__name__)
CORS(app)  # 允许网页跨域访问

# ── 上传文件临时存放的文件夹
UPLOAD_FOLDER = str(PROJECT_ROOT / "uploads_temp")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGE_OCR_CACHE_DIR, exist_ok=True)
REPORTS_FOLDER = str(PROJECT_ROOT / "reports")
os.makedirs(REPORTS_FOLDER, exist_ok=True)
BATCH_TEST_JOBS = {}
BATCH_TEST_PROCESSES = {}
BATCH_TEST_LOCK = threading.Lock()
PROCESS_JOBS = {}
PROCESS_JOB_LOCK = threading.Lock()

# ── 创建处理器（全局复用）
processor = DocFlowProcessor()

# ── 支持 OCR 的图片格式
IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp'}


def _resolve_batch_suites(suites) -> list[str]:
    resolved = []
    seen = set()
    for suite in suites or []:
        suite_name = str(suite or "").strip()
        if not suite_name:
            continue

        candidates = []
        if suite_name in BATCH_SUITE_ALIASES:
            candidates.append(BATCH_SUITE_ALIASES[suite_name])
        suite_path = Path(suite_name)
        candidates.extend([suite_path, PROJECT_ROOT / suite_path])

        for candidate in candidates:
            try:
                resolved_path = candidate.resolve()
            except Exception:
                resolved_path = candidate
            key = os.path.normcase(str(resolved_path))
            if key in seen:
                continue
            if resolved_path.exists() and resolved_path.is_dir():
                resolved.append(str(resolved_path))
                seen.add(key)
                break
    return resolved


def _count_suite_cases(suites: list[str]) -> int:
    total = 0
    for suite in suites:
        suite_path = Path(suite)
        if not suite_path.exists() or not suite_path.is_dir():
            continue
        total += len([p for p in suite_path.iterdir() if p.is_file() and p.name.lower() != "readme.md"])
    return total


def _error_response(message: str, status_code: int = 400, file_name: str = "", file_ext: str = ""):
    payload = {
        "success": False,
        "error": message,
        "error_info": build_error_info(
            message,
            file_name=file_name,
            file_ext=file_ext,
            metadata_dict={},
            source="api",
        ),
    }
    return jsonify(payload), status_code


def _append_job_log(job_id: str, level: str, message: str) -> None:
    with BATCH_TEST_LOCK:
        job = BATCH_TEST_JOBS.get(job_id)
        if not job:
            return
        job["log_seq"] += 1
        job["logs"].append(
            {
                "id": job["log_seq"],
                "time": time.strftime("%H:%M:%S"),
                "level": level,
                "message": message,
            }
        )
        if len(job["logs"]) > 400:
            job["logs"] = job["logs"][-400:]
        job["updated_at"] = time.time()


def _parse_log_level(line: str) -> str:
    m = re.search(r"\[(INFO|WARNING|ERROR|DEBUG)\]", line)
    if not m:
        return "INFO"
    return {"WARNING": "WARN"}.get(m.group(1), m.group(1))


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
    result = {
        "success": False,
        "cancelled": True,
        "error": message,
        "error_info": build_error_info(message, file_name=file_name, file_ext=file_ext, metadata_dict={}, source="process"),
    }


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
            result = process_image_ocr(save_path, file_name, progress_callback=report, cancel_callback=cancel_requested)
        else:
            result = job_processor.process(
                save_path,
                extract_keywords=True,
                output_format=output_format,
                pdf_mode=pdf_mode,
                progress_callback=report,
                cancel_callback=cancel_requested,
            )
        if cancel_requested():
            raise DocFlowCancelledError(cancelled_message)
        result = augment_result_payload(result, file_name=file_name, file_ext=file_ext, source="process")
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
                "error_info": build_error_info(str(exc), file_name=file_name, file_ext=file_ext, metadata_dict={}, source="process"),
            },
            finished_at=time.time(),
        )
    finally:
        try:
            os.remove(save_path)
        except Exception:
            pass


def _build_report_payload(report_dir: Optional[Path]) -> tuple[dict, list, list, list, dict]:
    summary = {}
    records = []
    failed_cases = []
    unexpected_cases = []
    report_urls = {}

    if not report_dir:
        return summary, records, failed_cases, unexpected_cases, report_urls

    results_json = report_dir / "results.json"
    if results_json.exists():
        payload = json.loads(results_json.read_text(encoding="utf-8"))
        summary = payload.get("summary", {})
        records = payload.get("records", [])

    failed_cases = [
        {
            "suite": item.get("suite"),
            "filename": item.get("filename"),
            "error": item.get("error", ""),
            "expected_success": item.get("expected_success"),
            "success": item.get("success"),
        }
        for item in records
        if not item.get("success")
    ][:20]

    unexpected_cases = [
        {
            "suite": item.get("suite"),
            "filename": item.get("filename"),
            "expected_success": item.get("expected_success"),
            "success": item.get("success"),
        }
        for item in records
        if item.get("matches_expectation") is False
    ][:20]

    report_urls = {
        "html": f"/reports/{report_dir.name}/summary.html",
        "markdown": f"/reports/{report_dir.name}/report.md",
        "json": f"/reports/{report_dir.name}/results.json",
        "csv": f"/reports/{report_dir.name}/results.csv",
    }
    return summary, records, failed_cases, unexpected_cases, report_urls


def _terminate_batch_process(job_id: str) -> None:
    with BATCH_TEST_LOCK:
        proc = BATCH_TEST_PROCESSES.get(job_id)

    if not proc or proc.poll() is not None:
        return

    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                capture_output=True,
                text=True,
                timeout=10,
            )
        else:
            proc.terminate()
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def _run_batch_test_job(job_id: str) -> None:
    with BATCH_TEST_LOCK:
        job = BATCH_TEST_JOBS.get(job_id)
        if not job:
            return
        if job.get("state") == "cancelled" or job.get("cancel_requested"):
            job["finished_at"] = time.time()
            job["updated_at"] = time.time()
            return
        job["state"] = "running"
        job["started_at"] = time.time()
        job["updated_at"] = time.time()
        suites = list(job["suites"])
        keywords = bool(job["keywords"])
        strict = bool(job["strict"])
        pdf_mode = _normalize_pdf_mode(job.get("pdf_mode", DEFAULT_PDF_MODE))

    reports_dir = Path(REPORTS_FOLDER)
    before = {p.name for p in reports_dir.glob("batch_test_*") if p.is_dir()}
    cmd = [sys.executable, "-u", str(SCRIPTS_DIR / "run_batch_tests.py"), *suites]
    if keywords:
        cmd.append("--keywords")
    if strict:
        cmd.append("--strict")
    if pdf_mode in ("accurate", "balanced", "fast"):
        cmd.extend(["--pdf-mode", pdf_mode])

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    _append_job_log(job_id, "INFO", f"开始批测：{', '.join(Path(s).name for s in suites)} ｜ PDF模式: {pdf_mode}")

    try:
        proc = subprocess.Popen(
            cmd,
            cwd=Path(__file__).resolve().parent,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )
        with BATCH_TEST_LOCK:
            if job_id in BATCH_TEST_JOBS:
                BATCH_TEST_PROCESSES[job_id] = proc
    except Exception as exc:
        with BATCH_TEST_LOCK:
            job = BATCH_TEST_JOBS.get(job_id)
            if job:
                job["state"] = "failed"
                job["error"] = f"批量测试启动失败: {exc}"
                job["finished_at"] = time.time()
                job["updated_at"] = time.time()
        _append_job_log(job_id, "ERROR", f"批量测试启动失败: {exc}")
        return

    current_total = 0
    cancelled = False
    try:
        for raw_line in proc.stdout or []:
            with BATCH_TEST_LOCK:
                job = BATCH_TEST_JOBS.get(job_id)
                cancel_requested = bool(job and job.get("cancel_requested"))
            if cancel_requested:
                cancelled = True
                _terminate_batch_process(job_id)
                break

            line = raw_line.strip()
            if not line:
                continue

            _append_job_log(job_id, _parse_log_level(line), line)
            m = re.search(r"\[(\d+)/(\d+)\]\s+(.*?)\s+->\s+(.+)$", line)
            if m:
                index = int(m.group(1))
                current_total = int(m.group(2))
                suite_name = m.group(3).strip()
                current_file = m.group(4).strip()
                with BATCH_TEST_LOCK:
                    job = BATCH_TEST_JOBS.get(job_id)
                    if job:
                        job["current_index"] = index
                        job["total"] = current_total
                        job["current_suite"] = suite_name
                        job["current_file"] = current_file
                        job["completed_count"] = max(0, index - 1)
                        job["updated_at"] = time.time()

        return_code = proc.wait(timeout=30)
        with BATCH_TEST_LOCK:
            job = BATCH_TEST_JOBS.get(job_id)
            if job and job.get("cancel_requested"):
                cancelled = True
    except Exception as exc:
        try:
            proc.kill()
        except Exception:
            pass
        with BATCH_TEST_LOCK:
            BATCH_TEST_PROCESSES.pop(job_id, None)
        with BATCH_TEST_LOCK:
            job = BATCH_TEST_JOBS.get(job_id)
            if job:
                if job.get("cancel_requested"):
                    job["state"] = "cancelled"
                    job["error"] = "任务已取消"
                else:
                    job["state"] = "failed"
                    job["error"] = f"批量测试执行异常: {exc}"
                job["finished_at"] = time.time()
                job["updated_at"] = time.time()
        if cancelled:
            _append_job_log(job_id, "WARN", "批量测试任务已取消")
        else:
            _append_job_log(job_id, "ERROR", f"批量测试执行异常: {exc}")
        return
    finally:
        with BATCH_TEST_LOCK:
            BATCH_TEST_PROCESSES.pop(job_id, None)

    after_dirs = sorted(
        [p for p in reports_dir.glob("batch_test_*") if p.is_dir()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    report_dir = next((p for p in after_dirs if p.name not in before), after_dirs[0] if after_dirs else None)

    try:
        summary, records, failed_cases, unexpected_cases, report_urls = _build_report_payload(report_dir)
        total = summary.get("total", current_total)
        completed = total or current_total
        if cancelled:
            success = False
            final_state = "cancelled"
            failed_cases = failed_cases or []
            unexpected_cases = unexpected_cases or []
        else:
            success = bool(summary) or return_code == 0
            final_state = "completed" if success or report_dir else "failed"

        with BATCH_TEST_LOCK:
            job = BATCH_TEST_JOBS.get(job_id)
            if job:
                job.update(
                    {
                        "state": final_state,
                        "success": success,
                        "command_ok": return_code == 0,
                        "return_code": return_code,
                        "summary": summary,
                        "records_count": len(records),
                        "failed_cases": failed_cases,
                        "unexpected_cases": unexpected_cases,
                        "report_dir": str(report_dir) if report_dir else "",
                        "report_urls": report_urls,
                        "finished_at": time.time(),
                        "updated_at": time.time(),
                        "completed_count": completed if not cancelled else job.get("completed_count", completed),
                        "current_index": total if not cancelled else job.get("current_index", completed),
                        "total": total,
                        "current_file": "",
                        "current_suite": "",
                        "error": "任务已取消" if cancelled else ("" if success else ((failed_cases[0]["error"] if failed_cases else "批量测试失败"))),
                    }
                )

        if cancelled:
            _append_job_log(job_id, "WARN", "批量测试任务已取消")
        elif success:
            matched = summary.get("expected_matched", 0)
            checked = summary.get("expected_checked", 0)
            _append_job_log(
                job_id,
                "INFO",
                f"批量测试完成 ✓ 成功 {summary.get('success', 0)}/{summary.get('total', 0)}，预期匹配 {matched}/{checked}",
            )
        else:
            _append_job_log(job_id, "ERROR", "批量测试失败")
    except Exception as exc:
        with BATCH_TEST_LOCK:
            job = BATCH_TEST_JOBS.get(job_id)
            if job:
                job["state"] = "failed"
                job["error"] = f"批量测试结果整理失败: {exc}"
                job["return_code"] = return_code
                job["finished_at"] = time.time()
                job["updated_at"] = time.time()
        _append_job_log(job_id, "ERROR", f"批量测试结果整理失败: {exc}")


def _serialize_batch_job(job_id: str) -> Optional[dict]:
    with BATCH_TEST_LOCK:
        job = BATCH_TEST_JOBS.get(job_id)
        if not job:
            return None
        data = dict(job)

    total = data.get("total", 0) or 0
    completed = data.get("completed_count", 0) or 0
    if data.get("state") == "completed" and total:
        completed = total
    data["completed_count"] = completed
    data["progress_pct"] = round((completed / total) * 100, 2) if total else 0.0
    data["logs"] = list(data.get("logs", []))
    data["error_info"] = build_error_info(
        data.get("error", ""),
        file_name=data.get("current_file", ""),
        file_ext="",
        metadata_dict={},
        source="batch",
    )
    return data


# ────────────────────────────────────────
#  路由 1：打开网页
# ────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(FRONTEND_DIR, "doc_tool.html")


@app.route("/reports/<path:report_path>")
def serve_report(report_path):
    return send_from_directory(REPORTS_FOLDER, report_path)


@app.route("/system/dependencies", methods=["GET"])
def get_system_dependencies():
    return jsonify({"success": True, **collect_dependency_status()})


@app.route("/system/dependencies/install", methods=["POST"])
def install_system_dependencies():
    payload = request.get_json(silent=True) or {}
    include_optional = bool(payload.get("include_optional", True))
    result = install_missing_dependencies(include_optional=include_optional)
    return jsonify({"success": True, **result})


@app.route("/process/start", methods=["POST"])
def start_process_file():
    if "file" not in request.files:
        return _error_response("没有收到文件")

    file = request.files["file"]
    if file.filename == "":
        return _error_response("文件名为空")

    safe_name = Path(file.filename).name
    job_id = uuid.uuid4().hex[:12]
    temp_name = f"{job_id}_{safe_name}"
    save_path = os.path.join(UPLOAD_FOLDER, temp_name)
    file.save(save_path)

    output_format = request.form.get("format", "txt")
    pdf_mode = _normalize_pdf_mode(request.form.get("pdf_mode", DEFAULT_PDF_MODE))
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
        return _error_response("处理任务不存在", status_code=404)
    return jsonify({"success": True, **data})


@app.route("/process/<job_id>/cancel", methods=["POST"])
def cancel_process_job(job_id: str):
    now = time.time()
    with PROCESS_JOB_LOCK:
        job = PROCESS_JOBS.get(job_id)
        if not job:
            return _error_response("处理任务不存在", status_code=404)

        state = job.get("state")
        if state in ("completed", "failed", "cancelled", "cancelling"):
            data = _serialize_process_job(job_id)
            if not data:
                return _error_response("处理任务不存在", status_code=404)
            return jsonify({"success": True, **data})

        job["cancel_requested"] = True
        job["updated_at"] = now

        if state == "queued":
            job["state"] = "cancelled"
            job["progress_pct"] = 100.0
            job["stage"] = "cancelled"
            job["message"] = "任务已取消"
            job["error"] = "任务已取消"
            job["result"] = _build_cancelled_process_result(job.get("file_name", ""), job.get("file_ext", ""))
            job["finished_at"] = now
        else:
            job["state"] = "cancelling"
            job["stage"] = "cancelling"
            job["message"] = "正在取消任务..."

    data = _serialize_process_job(job_id)
    if not data:
        return _error_response("处理任务不存在", status_code=404)
    return jsonify({"success": True, **data})


# ────────────────────────────────────────
#  路由 2：上传并处理文件
# ────────────────────────────────────────
@app.route("/process", methods=["POST"])
def process_file():
    if "file" not in request.files:
        return _error_response("没有收到文件")

    file = request.files["file"]
    if file.filename == "":
        return _error_response("文件名为空")

    # 保存临时文件
    save_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(save_path)

    output_format = request.form.get("format", "txt")
    pdf_mode = _normalize_pdf_mode(request.form.get("pdf_mode", DEFAULT_PDF_MODE))
    ext = os.path.splitext(file.filename)[1].lower()

    # 图片走 OCR 流程
    if ext in IMAGE_EXTS:
        result = process_image_ocr(save_path, file.filename)
    else:
        result = processor.process(
            save_path,
            extract_keywords=True,
            output_format=output_format,
            pdf_mode=pdf_mode,
        )
    result = augment_result_payload(result, file_name=file.filename, file_ext=ext, source="process")

    # 删除临时文件
    try:
        os.remove(save_path)
    except Exception:
        pass

    return jsonify(result)


@app.route("/run-batch-tests", methods=["POST"])
def run_batch_tests():
    payload = request.get_json(silent=True) or {}
    suites = _resolve_batch_suites(payload.get("suites") or ["test_documents", "test_documents_edge_cases"])
    keywords = bool(payload.get("keywords"))
    strict = bool(payload.get("strict"))
    pdf_mode = _normalize_pdf_mode(payload.get("pdf_mode", DEFAULT_PDF_MODE))

    if not suites:
        return _error_response("未选择有效的测试目录")

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
        return _error_response("批量测试任务不存在", status_code=404)
    data["job_success"] = data.get("success", False)
    data["success"] = True
    return jsonify(data)


@app.route("/run-batch-tests/<job_id>/cancel", methods=["POST"])
def cancel_batch_test(job_id: str):
    now = time.time()
    with BATCH_TEST_LOCK:
        job = BATCH_TEST_JOBS.get(job_id)
        if not job:
            return _error_response("批量测试任务不存在", status_code=404)

        state = job.get("state")
        if state in ("completed", "failed", "cancelled", "cancelling"):
            data = _serialize_batch_job(job_id)
            if data is None:
                return _error_response("批量测试任务不存在", status_code=404)
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
        return _error_response("批量测试任务不存在", status_code=404)
    data["job_success"] = data.get("success", False)
    data["success"] = True
    return jsonify(data)


# ────────────────────────────────────────
#  OCR 处理（优先用 easyocr，备用 pytesseract）
# ────────────────────────────────────────
def process_image_ocr(image_path: str, filename: str, progress_callback=None, cancel_callback=None) -> dict:
    import time
    start = time.time()

    def ensure_not_cancelled() -> None:
        if callable(cancel_callback) and cancel_callback():
            raise DocFlowCancelledError("任务已取消")

    def emit(progress_pct: float, stage: str, message: str = "", **extra) -> None:
        ensure_not_cancelled()
        if callable(progress_callback):
            progress_callback(
                progress_pct=max(0.0, min(float(progress_pct), 100.0)),
                stage=stage,
                message=message,
                **extra,
            )

    emit(5, "image_prepare", f"正在读取图片：{filename}")
    text = ""
    engine_used = ""
    easyocr_error = ""
    tesseract_error = ""
    prepared_meta = {}
    engine_order = _get_image_ocr_order()
    cache_key = ""
    file_sha256 = ""
    cache_profile = {}

    if _is_image_ocr_cache_enabled():
        try:
            emit(12, "image_cache_lookup", "正在检查图片 OCR 缓存")
            cache_key, file_sha256, cache_profile = _build_image_ocr_cache_key(image_path)
            cached_payload = _load_image_ocr_cache(cache_key)
            if cached_payload:
                elapsed = (time.time() - start) * 1000
                ensure_not_cancelled()
                emit(90, "image_cache_hit", "命中图片 OCR 缓存，正在返回结果")
                cached_result = _restore_cached_image_ocr_result(filename, cached_payload, elapsed)
                if cached_result:
                    return cached_result
        except Exception:
            cache_key = ""
            file_sha256 = ""
            cache_profile = {}

    for engine_name in engine_order:
        if text or engine_used:
            break

        if engine_name == "tesseract":
            try:
                ensure_not_cancelled()
                emit(24, "image_preprocess", "正在优化图片尺寸与对比度")
                import pytesseract

                prepared = prepare_pytesseract()
                img, prepared_meta = _prepare_image_for_tesseract(image_path)
                try:
                    ensure_not_cancelled()
                    emit(58, "image_tesseract", "正在尝试 Tesseract 识别")
                    text = pytesseract.image_to_string(
                        img,
                        lang=prepared.get("lang", "chi_sim+eng"),
                        config=_build_image_tesseract_config(prepared.get("config", "")),
                    )
                finally:
                    img.close()
                engine_used = "Tesseract"
            except ImportError:
                pass
            except Exception as e:
                tesseract_error = str(e)

        elif engine_name == "easyocr":
            try:
                ensure_not_cancelled()
                emit(72, "image_easyocr", "正在尝试 EasyOCR 识别")
                reader = _get_easyocr_reader()
                results = reader.readtext(image_path, detail=0)
                text = "\n".join(results)
                engine_used = "EasyOCR"
            except ImportError:
                pass
            except Exception as e:
                easyocr_error = str(e)

    # 两种都没装：返回提示
    if not engine_used:
        error_lines = []
        if easyocr_error:
            error_lines.append(f"EasyOCR: {easyocr_error}")
        if tesseract_error:
            error_lines.append(f"Tesseract: {tesseract_error}")
        detail = "\n".join(error_lines)
        text = (
            "⚠ OCR 不可用\n\n"
            "请安装 `easyocr`，或补齐 `pytesseract + Tesseract OCR` 运行环境。"
        )
        if detail:
            text += f"\n\n{detail}"
            engine_used = "未安装"

    elapsed = (time.time() - start) * 1000
    ensure_not_cancelled()
    emit(92, "image_finalize", "正在整理图片 OCR 结果")
    char_count = len(text)

    formatted = f"""[图片 OCR 结果] {filename}
{'─' * 40}
OCR 引擎: {engine_used}
识别字符数: {char_count}
处理耗时: {elapsed:.0f}ms

识别内容:
{'─' * 20}
{text}
"""
    result = {
        "success": True,
        "file": filename,
        "format": "image",
        "text": text,
        "tables": [],
        "metadata": {
            "engine": engine_used,
            "file": filename,
            "ocr_preprocess": prepared_meta,
            "image_ocr_cache": _build_image_ocr_cache_meta(
                cache_hit=False,
                file_sha256=file_sha256,
                profile=cache_profile,
            ),
        },
        "statistics": {
            "char_count": char_count,
            "paragraph_count": len(text.splitlines()),
            "table_count": 0,
            "keywords": [],
            "processing_ms": elapsed,
        },
        "processing_ms": elapsed,
        "formatted_output": _format_image_ocr_output(
            filename,
            engine_used,
            char_count,
            elapsed,
            text,
            cache_hit=False,
        ),
        "error": "",
    }
    if cache_key and file_sha256 and cache_profile and engine_used in {"Tesseract", "EasyOCR"}:
        _save_image_ocr_cache(cache_key, file_sha256, cache_profile, result)
    return result


# ────────────────────────────────────────
#  启动服务
# ────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 45)
    print("  DocFlow 服务已启动！")
    print("  请用浏览器打开：http://127.0.0.1:5000")
    print("=" * 45)
    app.run(debug=True, port=5000)


# ────────────────────────────────────────
#  临时诊断路由 — 测完可删
# ────────────────────────────────────────
@app.route("/debug-doc", methods=["GET"])
def debug_doc():
    import platform, shutil, subprocess, tempfile, os
    info = {}
    info["platform"] = platform.system()
    info["python"] = platform.python_version()
    
    candidates = [
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
        '/usr/lib/libreoffice/program/soffice',
        '/opt/libreoffice/program/soffice',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    ]
    found = []
    for c in candidates:
        if os.path.isfile(c) and os.access(c, os.X_OK):
            found.append(c)
    info["soffice_found"] = found
    info["which_soffice"] = shutil.which('soffice')
    info["PATH"] = os.environ.get('PATH', '')
    
    # 尝试实际运行 soffice --version
    if found:
        try:
            r = subprocess.run([found[0], '--version'], capture_output=True, text=True, timeout=10)
            info["soffice_version"] = r.stdout.strip() or r.stderr.strip()
            info["soffice_version_code"] = r.returncode
        except Exception as e:
            info["soffice_version_err"] = str(e)
    
    # 检查 lock 文件
    import glob
    locks = glob.glob(os.path.expanduser('~/.config/libreoffice/**/.~lock*'), recursive=True)
    info["lock_files"] = locks
    
    # 尝试实际转换一个测试文件
    test_files = []
    for ext in ['docx', 'doc']:
        for root, dirs, files in os.walk(UPLOAD_FOLDER):
            for f in files:
                if f.lower().endswith(f'.{ext}'):
                    test_files.append(os.path.join(root, f))
    
    info["uploads_temp_files"] = os.listdir(UPLOAD_FOLDER) if os.path.exists(UPLOAD_FOLDER) else []
    
    if found and test_files:
        try:
            tmp_dir = tempfile.mkdtemp()
            user_profile = tempfile.mkdtemp()
            abs_file = os.path.abspath(test_files[0])
            env = os.environ.copy()
            env['PATH'] = '/usr/bin:/usr/local/bin:' + env.get('PATH','')
            cmd = [found[0],
                   f'-env:UserInstallation=file://{user_profile}',
                   '--headless','--norestore','--nofirststartwizard',
                   '--convert-to','docx','--outdir',tmp_dir, abs_file]
            info["test_cmd"] = ' '.join(cmd)
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
            info["test_returncode"] = r.returncode
            info["test_stdout"] = r.stdout
            info["test_stderr"] = r.stderr[:500]
            info["test_output_files"] = os.listdir(tmp_dir)
        except Exception as e:
            info["test_error"] = str(e)
    
    return jsonify(info)
