"""依赖检查与自动安装工具。"""
from __future__ import annotations

import importlib
import importlib.util
import os
import platform
import subprocess
import sys
from datetime import datetime
from importlib import metadata
from typing import Any

from .tesseract import TOOL_CANDIDATES, resolve_tool_path


DEPENDENCY_SPECS: list[dict[str, Any]] = [
    {"id": "pdfplumber", "label": "PDF 文本解析", "module": "pdfplumber", "package": "pdfplumber", "install": "pip install pdfplumber", "required_for": ["pdf"], "optional": False},
    {"id": "pymupdf", "label": "PDF 文本备用解析", "module": "fitz", "package": "PyMuPDF", "install": "pip install PyMuPDF", "required_for": ["pdf"], "optional": True},
    {"id": "pypdfium2", "label": "PDF 栅格化兜底", "module": "pypdfium2", "package": "pypdfium2", "install": "pip install pypdfium2", "required_for": ["pdf", "ocr"], "optional": True},
    {"id": "numpy", "label": "数值计算基础库", "module": "numpy", "package": "numpy", "install": "pip install numpy", "required_for": ["pdf", "ocr", "image"], "optional": True},
    {"id": "easyocr", "label": "OCR 引擎", "module": "easyocr", "package": "easyocr", "install": "pip install easyocr", "required_for": ["pdf", "image", "ocr"], "optional": True},
    {"id": "onnxruntime", "label": "ONNXRuntime 推理引擎", "module": "onnxruntime", "package": "onnxruntime", "install": "pip install onnxruntime", "required_for": ["image", "ocr"], "optional": True},
    {"id": "rapidocr", "label": "RapidOCR 引擎", "module": "rapidocr", "package": "rapidocr", "install": "pip install rapidocr onnxruntime", "required_for": ["image", "ocr"], "optional": True},
    {"id": "pytesseract", "label": "Tesseract Python 接口", "module": "pytesseract", "package": "pytesseract", "install": "pip install pytesseract", "required_for": ["image", "ocr"], "optional": True},
    {"id": "pillow", "label": "图像读写", "module": "PIL", "package": "Pillow", "install": "pip install Pillow", "required_for": ["image", "pdf"], "optional": False},
    {"id": "python_docx", "label": "Word 解析", "module": "docx", "package": "python-docx", "install": "pip install python-docx", "required_for": ["docx", "doc"], "optional": False},
    {"id": "openpyxl", "label": "Excel 解析", "module": "openpyxl", "package": "openpyxl", "install": "pip install openpyxl", "required_for": ["xlsx", "xls"], "optional": False},
    {"id": "python_pptx", "label": "PPT 解析", "module": "pptx", "package": "python-pptx", "install": "pip install python-pptx", "required_for": ["pptx"], "optional": False},
]


def _module_exists(module_name: str) -> bool:
    try:
        return importlib.util.find_spec(module_name) is not None
    except Exception:
        return False


def _try_import_module(module_name: str) -> tuple[bool, str]:
    try:
        importlib.import_module(module_name)
        return True, ""
    except Exception as exc:
        return False, f"{exc.__class__.__name__}: {exc}"


def _probe_python_dependency(spec: dict[str, Any]) -> tuple[bool, str]:
    module_name = str(spec.get("module") or "").strip()
    dep_id = str(spec.get("id") or "").strip()

    if not module_name:
        return False, "未提供模块名"

    if dep_id == "rapidocr":
        ok, error = _try_import_module("onnxruntime")
        if not ok:
            return False, f"onnxruntime 导入失败：{error}"
        ok, error = _try_import_module("cv2")
        if not ok:
            return False, f"cv2 导入失败：{error}"
        ok, error = _try_import_module(module_name)
        if not ok:
            return False, f"rapidocr 导入失败：{error}"
        try:
            module = importlib.import_module(module_name)
            if getattr(module, "RapidOCR", None) is None:
                return False, "rapidocr 包内未找到 RapidOCR 类"
        except Exception as exc:
            return False, f"{exc.__class__.__name__}: {exc}"
        return True, ""

    return _try_import_module(module_name)


def _safe_version(package_name: str | None) -> str:
    if not package_name:
        return ""
    try:
        return metadata.version(package_name)
    except metadata.PackageNotFoundError:
        return ""
    except Exception:
        return ""


def _item_available(item: dict[str, Any]) -> bool:
    return bool(item.get("available", item.get("installed")))


def _derive_profile_status(dep_map: dict[str, dict[str, Any]], tool_map: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    rapid_ready = _item_available(dep_map["onnxruntime"]) and _item_available(dep_map["rapidocr"])
    easyocr_ready = _item_available(dep_map["easyocr"])
    pytesseract_ready = _item_available(dep_map["pytesseract"])
    tesseract_ready = _item_available(tool_map["tesseract"])
    pdfplumber_ready = _item_available(dep_map["pdfplumber"])
    pymupdf_ready = _item_available(dep_map["pymupdf"])
    pdf_fallback_ready = (
        _item_available(dep_map["pypdfium2"])
        and _item_available(dep_map["pillow"])
        and _item_available(dep_map["numpy"])
        and (rapid_ready or easyocr_ready or (pytesseract_ready and tesseract_ready))
    )

    def pack(label: str, status: str, reason: str) -> dict[str, Any]:
        return {"label": label, "status": status, "reason": reason}

    return {
        "pdf": pack(
            "PDF",
            "ready" if pdfplumber_ready or pymupdf_ready else ("degraded" if pdf_fallback_ready else "missing"),
            "pdfplumber 可直接解析文本层" if pdfplumber_ready else (
                "PyMuPDF 可作为文本层备用解析器" if pymupdf_ready else (
                    "主解析器缺失，但 OCR 兜底可用" if pdf_fallback_ready else "缺少 PDF 文本解析器，且 OCR 兜底链路不完整"
                )
            ),
        ),
        "word": pack("Word", "ready" if _item_available(dep_map["python_docx"]) else "missing", "python-docx 已安装" if _item_available(dep_map["python_docx"]) else "缺少 python-docx"),
        "excel": pack("Excel", "ready" if _item_available(dep_map["openpyxl"]) else "missing", "openpyxl 已安装" if _item_available(dep_map["openpyxl"]) else "缺少 openpyxl"),
        "powerpoint": pack("PPT", "ready" if _item_available(dep_map["python_pptx"]) else "missing", "python-pptx 已安装" if _item_available(dep_map["python_pptx"]) else "缺少 python-pptx"),
        "image_ocr": pack(
            "图片 OCR",
            "ready" if rapid_ready or easyocr_ready or (pytesseract_ready and tesseract_ready) else "missing",
            "RapidOCR 可用" if rapid_ready else ("EasyOCR 可用" if easyocr_ready else ("pytesseract + tesseract 可用" if pytesseract_ready and tesseract_ready else "OCR 引擎链路不完整")),
        ),
        "legacy_doc_fallback": pack("旧版 DOC 兜底", "ready" if _item_available(tool_map["soffice"]) else "degraded", "LibreOffice/soffice 可用" if _item_available(tool_map["soffice"]) else "未检测到 soffice，仅使用内置 DOC 解析"),
        "text_json_csv": pack("文本/JSON/CSV", "ready", "内置支持"),
    }


def collect_dependency_status() -> dict[str, Any]:
    items: list[dict[str, Any]] = []
    dep_map: dict[str, dict[str, Any]] = {}
    tool_map: dict[str, dict[str, Any]] = {}

    for spec in DEPENDENCY_SPECS:
        installed = _module_exists(spec["module"])
        runtime_ready = False
        runtime_error = ""
        if installed:
            runtime_ready, runtime_error = _probe_python_dependency(spec)

        available = installed and runtime_ready
        item = {
            "id": spec["id"], "type": "python", "label": spec["label"],
            "module": spec["module"], "package": spec["package"], "install": spec["install"],
            "required_for": spec["required_for"], "optional": spec["optional"],
            "installed": installed, "available": available,
            "runtime_ready": runtime_ready, "runtime_error": runtime_error,
            "status": ("ready" if available else ("degraded" if installed else ("optional_missing" if spec["optional"] else "missing"))),
            "version": _safe_version(spec["package"]) if installed else "",
            "message": ("已安装并可用" if available else (f"已安装但运行失败：{runtime_error}" if installed else ("未安装（可选）" if spec["optional"] else "未安装"))),
        }
        dep_map[spec["id"]] = item
        items.append(item)

    for tool_id, candidates in TOOL_CANDIDATES.items():
        tool_path = resolve_tool_path(tool_id)
        available = bool(tool_path)
        tool_item = {
            "id": tool_id, "type": "tool",
            "label": "LibreOffice / soffice" if tool_id == "soffice" else "Tesseract 可执行程序",
            "path": tool_path, "installed": available, "available": available,
            "status": "ready" if available else "optional_missing",
            "required_for": ["doc"] if tool_id == "soffice" else ["ocr", "image"],
            "optional": True,
            "install": "安装 LibreOffice 并确保 soffice 可执行" if tool_id == "soffice" else "安装 Tesseract OCR 并加入 PATH",
            "version": "", "message": tool_path or "未检测到可执行程序",
        }
        tool_map[tool_id] = tool_item
        items.append(tool_item)

    profiles = _derive_profile_status(dep_map, tool_map)
    critical_missing = sum(1 for item in items if item["type"] == "python" and not item["optional"] and not _item_available(item))
    optional_missing = sum(1 for item in items if item["optional"] and not _item_available(item))
    overall_status = "ready" if critical_missing == 0 and optional_missing == 0 else ("degraded" if critical_missing == 0 else "missing")

    recommendations = []
    for item in items:
        if item["status"] in {"missing", "optional_missing", "degraded"}:
            prefix = "建议修复" if item["status"] == "degraded" else ("建议安装" if item["status"] == "missing" else "可选安装")
            recommendations.append(f"{prefix}：{item['label']}（{item['install']}）")

    return {
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "environment": {"python": sys.version.split()[0], "platform": platform.platform(), "executable": sys.executable},
        "summary": {"overall_status": overall_status, "total_items": len(items), "ready_count": sum(1 for item in items if _item_available(item)), "critical_missing": critical_missing, "optional_missing": optional_missing},
        "profiles": profiles,
        "items": items,
        "recommendations": recommendations[:8],
    }


def _find_dependency_spec(dep_id: str) -> dict[str, Any] | None:
    for spec in DEPENDENCY_SPECS:
        if spec["id"] == dep_id:
            return spec
    return None


def _tail_text(text: str, limit: int = 1600) -> str:
    value = (text or "").strip()
    if len(value) <= limit:
        return value
    return value[-limit:]


def _build_pip_install_attempts(package_name: str) -> list[list[str]]:
    base = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-input", package_name]
    fallback = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-input", "--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org", "-i", "https://pypi.org/simple", package_name]
    tuna = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-input", "--trusted-host", "pypi.tuna.tsinghua.edu.cn", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", package_name]
    return [base, fallback, tuna]


def _run_pip_install(package_name: str, timeout_seconds: int = 900) -> dict[str, Any]:
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PIP_DISABLE_PIP_VERSION_CHECK", "1")
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    env["PIP_NO_PROXY"] = "*"
    for proxy_key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy", "PIP_PROXY", "REQUESTS_CA_BUNDLE", "CURL_CA_BUNDLE"):
        env.pop(proxy_key, None)

    attempts: list[dict[str, Any]] = []
    for idx, command in enumerate(_build_pip_install_attempts(package_name), start=1):
        started_at = datetime.now()
        try:
            completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=timeout_seconds, env=env)
            attempt = {"index": idx, "command": " ".join(command), "return_code": completed.returncode, "stdout_tail": _tail_text(completed.stdout), "stderr_tail": _tail_text(completed.stderr), "started_at": started_at.isoformat(timespec="seconds"), "finished_at": datetime.now().isoformat(timespec="seconds"), "timeout": False}
            attempts.append(attempt)
            if completed.returncode == 0:
                return {"ok": True, "command": attempt["command"], "return_code": completed.returncode, "stdout_tail": attempt["stdout_tail"], "stderr_tail": attempt["stderr_tail"], "attempts": attempts}
        except subprocess.TimeoutExpired as exc:
            attempts.append({"index": idx, "command": " ".join(command), "return_code": None, "stdout_tail": _tail_text(exc.stdout or ""), "stderr_tail": _tail_text(exc.stderr or ""), "started_at": started_at.isoformat(timespec="seconds"), "finished_at": datetime.now().isoformat(timespec="seconds"), "timeout": True})

    last_attempt = attempts[-1] if attempts else {}
    return {"ok": False, "command": last_attempt.get("command", ""), "return_code": last_attempt.get("return_code"), "stdout_tail": last_attempt.get("stdout_tail", ""), "stderr_tail": last_attempt.get("stderr_tail", ""), "attempts": attempts}


def install_missing_dependencies(include_optional: bool = True) -> dict[str, Any]:
    before = collect_dependency_status()
    target_items = [item for item in before["items"] if item["type"] == "python" and not item["installed"] and (include_optional or not item["optional"])]

    if not target_items:
        remaining_tools = [item for item in before["items"] if item["type"] == "tool" and not item["installed"]]
        return {
            "install_ok": True, "changed": False, "include_optional": include_optional,
            "message": "当前没有需要自动安装的 Python 依赖",
            "before": before, "after": before, "results": [],
            "summary": {"requested_count": 0, "installed_count": 0, "failed_count": 0, "still_missing_required": before["summary"].get("critical_missing", 0), "still_missing_optional": before["summary"].get("optional_missing", 0), "remaining_manual_tools": len(remaining_tools)},
            "remaining": {"python": [], "tools": remaining_tools},
            "manual_actions": [f"{item['label']}：{item['install']}" for item in remaining_tools],
        }

    results: list[dict[str, Any]] = []
    for item in target_items:
        package_name = item.get("package") or item.get("module") or item["id"]
        install_result = _run_pip_install(package_name)
        importlib.invalidate_caches()
        results.append({"id": item["id"], "label": item["label"], "package": package_name, "optional": item["optional"], "required_for": item.get("required_for", []), "requested_install": item.get("install", ""), "attempt_count": len(install_result["attempts"]), "command": install_result["command"], "return_code": install_result["return_code"], "stdout_tail": install_result["stdout_tail"], "stderr_tail": install_result["stderr_tail"], "attempts": install_result["attempts"], "pip_ok": install_result["ok"]})

    importlib.invalidate_caches()
    after = collect_dependency_status()
    after_python_items = {item["id"]: item for item in after["items"] if item["type"] == "python"}

    installed_count = 0
    failed_count = 0
    for result in results:
        installed_now = after_python_items.get(result["id"], {}).get("installed", False)
        result["installed"] = installed_now
        result["status"] = "installed" if installed_now else "failed"
        result["version"] = after_python_items.get(result["id"], {}).get("version", "")
        if installed_now:
            installed_count += 1
            result["message"] = "安装成功"
        else:
            failed_count += 1
            result["message"] = result["stderr_tail"] or result["stdout_tail"] or "安装失败"

    remaining_python = [item for item in after["items"] if item["type"] == "python" and not item["installed"] and (include_optional or not item["optional"])]
    remaining_tools = [item for item in after["items"] if item["type"] == "tool" and not item["installed"]]
    install_ok = failed_count == 0
    message = f"已安装 {installed_count} 项 Python 依赖" if install_ok else (f"已安装 {installed_count} 项，仍有 {failed_count} 项安装失败" if installed_count else "未能安装缺失依赖")

    return {
        "install_ok": install_ok, "changed": installed_count > 0, "include_optional": include_optional, "message": message,
        "before": before, "after": after, "results": results,
        "summary": {"requested_count": len(target_items), "installed_count": installed_count, "failed_count": failed_count, "still_missing_required": after["summary"].get("critical_missing", 0), "still_missing_optional": after["summary"].get("optional_missing", 0), "remaining_manual_tools": len(remaining_tools)},
        "remaining": {"python": remaining_python, "tools": remaining_tools},
        "manual_actions": [f"{item['label']}：{item['install']}" for item in remaining_tools],
        "requested": [{"id": item["id"], "label": item["label"], "package": item.get("package") or item.get("module") or item["id"], "optional": item["optional"], "required_for": item.get("required_for", [])} for item in target_items],
    }
