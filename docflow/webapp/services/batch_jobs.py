"""批量测试任务服务。"""

import json
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Optional

from docflow.paths import PROJECT_ROOT, REPORTS_DIR, SAMPLE_DATA_DIR, SCRIPTS_DIR
from docflow.settings import DEFAULT_PDF_MODE
from docflow.settings import normalize_pdf_mode as _normalize_pdf_mode
from docflow.support import build_error_info

BATCH_SUITE_ALIASES = {
    "test_documents": SAMPLE_DATA_DIR / "test_documents",
    "test_documents_edge_cases": SAMPLE_DATA_DIR / "test_documents_edge_cases",
}
BATCH_TEST_JOBS = {}
BATCH_TEST_PROCESSES = {}
BATCH_TEST_LOCK = threading.Lock()


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
    match = re.search(r"\[(INFO|WARNING|ERROR|DEBUG)\]", line)
    if not match:
        return "INFO"
    return {"WARNING": "WARN"}.get(match.group(1), match.group(1))


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

    reports_dir = REPORTS_DIR
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
            cwd=str(PROJECT_ROOT),
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
            match = re.search(r"\[(\d+)/(\d+)\]\s+(.*?)\s+->\s+(.+)$", line)
            if match:
                index = int(match.group(1))
                current_total = int(match.group(2))
                suite_name = match.group(3).strip()
                current_file = match.group(4).strip()
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
