"""公共路由。"""

from flask import jsonify, send_from_directory

from docflow.paths import FRONTEND_DIR
from docflow_support import collect_dependency_status, install_missing_dependencies

from ..core import REPORTS_FOLDER, UPLOAD_FOLDER, app


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
    from flask import request

    payload = request.get_json(silent=True) or {}
    include_optional = bool(payload.get("include_optional", True))
    result = install_missing_dependencies(include_optional=include_optional)
    return jsonify({"success": True, **result})


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

