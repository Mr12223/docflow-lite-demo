"""公共路由。"""

import os
import threading
import time
from pathlib import Path

from flask import jsonify, request, send_from_directory

from docflow.paths import FRONTEND_DIR, IMAGE_OCR_CACHE_DIR, UPLOADS_DIR
from docflow.support import collect_dependency_status, install_missing_dependencies

from ..core import REPORTS_FOLDER, UPLOAD_FOLDER, app, logger


@app.route("/")
def index():
    response = send_from_directory(FRONTEND_DIR, "doc_tool.html", conditional=False, max_age=0)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.route("/reports/<path:report_path>")
def serve_report(report_path):
    return send_from_directory(REPORTS_FOLDER, report_path)


@app.route("/system/dependencies", methods=["GET"])
def get_system_dependencies():
    return jsonify({"success": True, **collect_dependency_status()})


@app.route("/system/dependencies/install", methods=["POST"])
def install_system_dependencies():
    from flask import request

    payload = request.get_json(silent=True) or {}
    include_optional = bool(payload.get("include_optional", True))
    result = install_missing_dependencies(include_optional=include_optional)
    return jsonify({"success": True, **result})


@app.route("/system/cache/image-ocr", methods=["DELETE"])
def clear_system_image_ocr_cache():
    if not _is_local_request():
        return jsonify({"success": False, "error": "只允许本机清理缓存"}), 403

    from ..services.ocr_engines import clear_image_ocr_cache

    result = clear_image_ocr_cache()
    logger.info(
        "Image OCR cache cleared: files=%s bytes=%s memory=%s",
        result.get("deleted_files", 0),
        result.get("deleted_bytes", 0),
        result.get("cleared_memory_items", 0),
    )
    return jsonify({"success": True, **result})


@app.route("/system/cache/upload-buffer", methods=["DELETE"])
def clear_system_upload_buffer():
    if not _is_local_request():
        return jsonify({"success": False, "error": "只允许本机清理缓冲文件"}), 403

    result = _clear_upload_buffer_files()
    logger.info(
        "Upload buffer cleared: files=%s dirs=%s bytes=%s skipped_active=%s",
        result.get("deleted_files", 0),
        result.get("deleted_dirs", 0),
        result.get("deleted_bytes", 0),
        result.get("skipped_active_files", 0),
    )
    return jsonify({"success": True, **result})


def _get_active_upload_paths() -> set[Path]:
    from ..services.process_jobs import PROCESS_JOBS, PROCESS_JOB_LOCK

    active_states = {"queued", "running", "cancelling"}
    active_paths: set[Path] = set()
    with PROCESS_JOB_LOCK:
        jobs = list(PROCESS_JOBS.values())

    for job in jobs:
        if str(job.get("state") or "") not in active_states:
            continue
        save_path = str(job.get("save_path") or "").strip()
        if not save_path:
            continue
        try:
            active_paths.add(Path(save_path).resolve())
        except OSError:
            continue
    return active_paths


def _clear_upload_buffer_files() -> dict:
    upload_dir = Path(UPLOAD_FOLDER).resolve()
    expected_dir = UPLOADS_DIR.resolve()
    if upload_dir != expected_dir:
        raise RuntimeError(f"Unsafe upload buffer path: {upload_dir}")

    upload_dir.mkdir(parents=True, exist_ok=True)
    ocr_cache_dir = IMAGE_OCR_CACHE_DIR.resolve()
    active_paths = _get_active_upload_paths()
    deleted_files = 0
    deleted_dirs = 0
    deleted_bytes = 0
    skipped_active_files = 0
    skipped_cache_files = 0

    items = sorted(upload_dir.rglob("*"), key=lambda item: len(item.parts), reverse=True)
    for item in items:
        try:
            resolved = item.resolve()
        except OSError:
            continue

        if resolved == ocr_cache_dir or ocr_cache_dir in resolved.parents:
            if item.is_file() or item.is_symlink():
                skipped_cache_files += 1
            continue

        if resolved in active_paths:
            skipped_active_files += 1
            continue

        try:
            if item.is_file() or item.is_symlink():
                try:
                    deleted_bytes += item.stat().st_size
                except OSError:
                    pass
                item.unlink()
                deleted_files += 1
            elif item.is_dir():
                try:
                    item.rmdir()
                    deleted_dirs += 1
                except OSError:
                    pass
        except OSError:
            continue

    upload_dir.mkdir(parents=True, exist_ok=True)
    IMAGE_OCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return {
        "buffer_dir": str(upload_dir),
        "deleted_files": deleted_files,
        "deleted_dirs": deleted_dirs,
        "deleted_bytes": deleted_bytes,
        "skipped_active_files": skipped_active_files,
        "skipped_cache_files": skipped_cache_files,
    }


def _is_local_request() -> bool:
    remote_addr = str(request.remote_addr or "").strip()
    return remote_addr in {"127.0.0.1", "::1", "localhost"} or remote_addr.startswith("::ffff:127.")


@app.route("/system/shutdown", methods=["POST"])
def shutdown_system():
    """关闭当前本地开发服务进程。"""

    if not _is_local_request():
        return jsonify({"success": False, "error": "只允许本机关闭服务"}), 403

    logger.info("Shutdown requested from web UI")
    werkzeug_shutdown = request.environ.get("werkzeug.server.shutdown")

    def stop_server() -> None:
        time.sleep(0.35)
        if callable(werkzeug_shutdown):
            try:
                werkzeug_shutdown()
                return
            except Exception:
                pass
        parent_pid = os.getppid()
        if os.environ.get("WERKZEUG_RUN_MAIN") == "true" and parent_pid and parent_pid != os.getpid():
            try:
                import subprocess

                subprocess.run(
                    ["taskkill", "/PID", str(parent_pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    timeout=5,
                )
                return
            except Exception:
                pass
        os._exit(0)

    threading.Thread(target=stop_server, daemon=True).start()
    return jsonify({"success": True, "message": "DocFlow 服务正在关闭"})


@app.route("/debug-doc", methods=["GET"])
def debug_doc():
    import glob
    import os
    import platform
    import shutil
    import subprocess
    import tempfile

    info = {}
    info["platform"] = platform.system()
    info["python"] = platform.python_version()

    candidates = [
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    found = []
    for candidate in candidates:
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            found.append(candidate)

    info["soffice_found"] = found
    info["which_soffice"] = shutil.which("soffice")
    info["PATH"] = os.environ.get("PATH", "")

    if found:
        try:
            result = subprocess.run(
                [found[0], "--version"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            info["soffice_version"] = result.stdout.strip() or result.stderr.strip()
            info["soffice_version_code"] = result.returncode
        except Exception as exc:
            info["soffice_version_err"] = str(exc)

    info["lock_files"] = glob.glob(
        os.path.expanduser("~/.config/libreoffice/**/.~lock*"),
        recursive=True,
    )

    test_files = []
    for ext in ["docx", "doc"]:
        for root, _, files in os.walk(UPLOAD_FOLDER):
            for filename in files:
                if filename.lower().endswith(f".{ext}"):
                    test_files.append(os.path.join(root, filename))

    info["uploads_temp_files"] = os.listdir(UPLOAD_FOLDER) if os.path.exists(UPLOAD_FOLDER) else []

    if found and test_files:
        try:
            tmp_dir = tempfile.mkdtemp()
            user_profile = tempfile.mkdtemp()
            abs_file = os.path.abspath(test_files[0])
            env = os.environ.copy()
            env["PATH"] = "/usr/bin:/usr/local/bin:" + env.get("PATH", "")
            cmd = [
                found[0],
                f"-env:UserInstallation=file://{user_profile}",
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--convert-to",
                "docx",
                "--outdir",
                tmp_dir,
                abs_file,
            ]
            info["test_cmd"] = " ".join(cmd)
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=90,
                env=env,
            )
            info["test_returncode"] = result.returncode
            info["test_stdout"] = result.stdout
            info["test_stderr"] = result.stderr[:500]
            info["test_output_files"] = os.listdir(tmp_dir)
        except Exception as exc:
            info["test_error"] = str(exc)

    return jsonify(info)
