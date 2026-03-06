from __future__ import annotations

import importlib
import importlib.util
import os
import platform
import re
import shutil
import subprocess
import sys
from collections import Counter
from datetime import datetime
from importlib import metadata
from pathlib import Path
from typing import Any


DEPENDENCY_SPECS = [
    {
        "id": "pdfplumber",
        "label": "PDF 文本解析",
        "module": "pdfplumber",
        "package": "pdfplumber",
        "install": "pip install pdfplumber",
        "required_for": ["pdf"],
        "optional": False,
    },
    {
        "id": "pymupdf",
        "label": "PDF 文本备用解析",
        "module": "fitz",
        "package": "PyMuPDF",
        "install": "pip install PyMuPDF",
        "required_for": ["pdf"],
        "optional": True,
    },
    {
        "id": "pypdfium2",
        "label": "PDF 栅格化兜底",
        "module": "pypdfium2",
        "package": "pypdfium2",
        "install": "pip install pypdfium2",
        "required_for": ["pdf", "ocr"],
        "optional": True,
    },
    {
        "id": "numpy",
        "label": "数值计算基础库",
        "module": "numpy",
        "package": "numpy",
        "install": "pip install numpy",
        "required_for": ["pdf", "ocr", "image"],
        "optional": True,
    },
    {
        "id": "easyocr",
        "label": "OCR 引擎",
        "module": "easyocr",
        "package": "easyocr",
        "install": "pip install easyocr",
        "required_for": ["pdf", "image", "ocr"],
        "optional": True,
    },
    {
        "id": "onnxruntime",
        "label": "ONNXRuntime 推理引擎",
        "module": "onnxruntime",
        "package": "onnxruntime",
        "install": "pip install onnxruntime",
        "required_for": ["image", "ocr"],
        "optional": True,
    },
    {
        "id": "rapidocr",
        "label": "RapidOCR 引擎",
        "module": "rapidocr",
        "package": "rapidocr",
        "install": "pip install rapidocr onnxruntime",
        "required_for": ["image", "ocr"],
        "optional": True,
    },
    {
        "id": "pytesseract",
        "label": "Tesseract Python 接口",
        "module": "pytesseract",
        "package": "pytesseract",
        "install": "pip install pytesseract",
        "required_for": ["image", "ocr"],
        "optional": True,
    },
    {
        "id": "pillow",
        "label": "图像读写",
        "module": "PIL",
        "package": "Pillow",
        "install": "pip install Pillow",
        "required_for": ["image", "pdf"],
        "optional": False,
    },
    {
        "id": "python_docx",
        "label": "Word 解析",
        "module": "docx",
        "package": "python-docx",
        "install": "pip install python-docx",
        "required_for": ["docx", "doc"],
        "optional": False,
    },
    {
        "id": "openpyxl",
        "label": "Excel 解析",
        "module": "openpyxl",
        "package": "openpyxl",
        "install": "pip install openpyxl",
        "required_for": ["xlsx", "xls"],
        "optional": False,
    },
    {
        "id": "python_pptx",
        "label": "PPT 解析",
        "module": "pptx",
        "package": "python-pptx",
        "install": "pip install python-pptx",
        "required_for": ["pptx"],
        "optional": False,
    },
]

TOOL_CANDIDATES = {
    "soffice": [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ],
    "tesseract": [
        "tesseract",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "/usr/bin/tesseract",
        "/opt/homebrew/bin/tesseract",
    ],
}


def _module_exists(module_name: str) -> bool:
    try:
        return importlib.util.find_spec(module_name) is not None
    except Exception:
        return False


def _safe_version(package_name: str | None) -> str:
    if not package_name:
        return ""
    try:
        return metadata.version(package_name)
    except metadata.PackageNotFoundError:
        return ""
    except Exception:
        return ""


def _find_tool_path(candidates: list[str]) -> str:
    for candidate in candidates:
        if Path(candidate).is_file():
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return ""


def resolve_tool_path(tool_id: str) -> str:
    return _find_tool_path(TOOL_CANDIDATES.get(tool_id, []))


def _is_tessdata_dir(path: Path) -> bool:
    try:
        return path.is_dir() and any(path.glob("*.traineddata"))
    except Exception:
        return False


def resolve_tessdata_dir(tool_path: str = "") -> str:
    candidates: list[Path] = []

    env_prefix = (os.environ.get("TESSDATA_PREFIX") or "").strip().strip('"')
    if env_prefix:
        env_path = Path(env_prefix).expanduser()
        candidates.extend([env_path, env_path / "tessdata"])

    if tool_path:
        install_root = Path(tool_path).resolve().parent
        candidates.extend(
            [
                install_root / "tessdata",
                install_root,
                install_root.parent / "share" / "tessdata",
                install_root.parent / "share" / "tesseract-ocr" / "5" / "tessdata",
                install_root.parent / "share" / "tesseract-ocr" / "4.00" / "tessdata",
            ]
        )

    candidates.extend(
        [
            Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tessdata"),
            Path("/usr/share/tessdata"),
            Path("/usr/share/tesseract-ocr/5/tessdata"),
            Path("/usr/share/tesseract-ocr/4.00/tessdata"),
            Path("/opt/homebrew/share/tessdata"),
        ]
    )

    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except Exception:
            resolved = candidate
        key = os.path.normcase(str(resolved))
        if not key or key in seen:
            continue
        seen.add(key)
        if _is_tessdata_dir(resolved):
            return str(resolved)
    return ""


def configure_pytesseract_command() -> str:
    try:
        import pytesseract
    except ImportError:
        return ""

    current_cmd = getattr(pytesseract.pytesseract, "tesseract_cmd", "") or ""
    candidates = []
    if current_cmd:
        candidates.append(current_cmd)
    candidates.extend(TOOL_CANDIDATES.get("tesseract", []))
    tool_path = _find_tool_path(candidates)
    if not tool_path:
        return ""

    pytesseract.pytesseract.tesseract_cmd = tool_path
    tessdata_dir = resolve_tessdata_dir(tool_path)
    if tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = tessdata_dir
    elif os.environ.get("TESSDATA_PREFIX"):
        os.environ.pop("TESSDATA_PREFIX", None)
    return tool_path


def build_tesseract_ocr_config(extra: str = "", tool_path: str = "") -> str:
    extra = (extra or "").strip()
    return extra


def prepare_pytesseract(preferred_languages: tuple[str, ...] = ("chi_sim", "eng")) -> dict[str, Any]:
    try:
        import pytesseract
    except ImportError as exc:
        raise RuntimeError("未安装 pytesseract") from exc

    tool_path = configure_pytesseract_command()
    if not tool_path:
        raise RuntimeError("未检测到 Tesseract 可执行程序")

    tesseract_config = build_tesseract_ocr_config(tool_path=tool_path)
    try:
        available_languages = [
            lang.strip()
            for lang in pytesseract.get_languages(config=tesseract_config)
            if str(lang).strip()
        ]
    except Exception as exc:
        raise RuntimeError(f"Tesseract 初始化失败：{exc}") from exc

    available_set = set(available_languages)
    selected_languages = [lang for lang in preferred_languages if lang in available_set]
    if not selected_languages:
        if available_languages:
            selected_languages = [available_languages[0]]
        else:
            raise RuntimeError("Tesseract 未检测到任何可用语言包")

    return {
        "command": tool_path,
        "config": tesseract_config,
        "tessdata_dir": resolve_tessdata_dir(tool_path),
        "available_languages": available_languages,
        "lang": "+".join(selected_languages),
    }


def _derive_profile_status(dep_map: dict[str, dict[str, Any]], tool_map: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    rapid_ready = dep_map["onnxruntime"]["installed"] and dep_map["rapidocr"]["installed"]
    easyocr_ready = dep_map["easyocr"]["installed"]
    pytesseract_ready = dep_map["pytesseract"]["installed"]
    tesseract_ready = tool_map["tesseract"]["installed"]
    pdfplumber_ready = dep_map["pdfplumber"]["installed"]
    pymupdf_ready = dep_map["pymupdf"]["installed"]
    pdf_fallback_ready = dep_map["pypdfium2"]["installed"] and dep_map["pillow"]["installed"] and dep_map["numpy"]["installed"] and (
        rapid_ready or easyocr_ready or (pytesseract_ready and tesseract_ready)
    )

    def pack(label: str, status: str, reason: str) -> dict[str, Any]:
        return {"label": label, "status": status, "reason": reason}

    return {
        "pdf": pack(
            "PDF",
            "ready" if pdfplumber_ready or pymupdf_ready else ("degraded" if pdf_fallback_ready else "missing"),
            "pdfplumber 可直接解析文本" if pdfplumber_ready else (
                "PyMuPDF 可作为文本层备用解析器" if pymupdf_ready else (
                    "主解析缺失，但 OCR 兜底可用" if pdf_fallback_ready else "缺少 PDF 文本解析器，且 OCR 兜底链路不完整"
                )
            ),
        ),
        "word": pack(
            "Word",
            "ready" if dep_map["python_docx"]["installed"] else "missing",
            "python-docx 已安装" if dep_map["python_docx"]["installed"] else "缺少 python-docx",
        ),
        "excel": pack(
            "Excel",
            "ready" if dep_map["openpyxl"]["installed"] else "missing",
            "openpyxl 已安装" if dep_map["openpyxl"]["installed"] else "缺少 openpyxl",
        ),
        "powerpoint": pack(
            "PPT",
            "ready" if dep_map["python_pptx"]["installed"] else "missing",
            "python-pptx 已安装" if dep_map["python_pptx"]["installed"] else "缺少 python-pptx",
        ),
        "image_ocr": pack(
            "图片 OCR",
            "ready" if rapid_ready or easyocr_ready or (pytesseract_ready and tesseract_ready) else "missing",
            "RapidOCR 可用" if rapid_ready else (
                "EasyOCR 可用" if easyocr_ready else (
                    "pytesseract + tesseract 可用" if pytesseract_ready and tesseract_ready else "OCR 引擎链路不完整"
                )
            ),
        ),
        "legacy_doc_fallback": pack(
            "旧版 DOC 兜底",
            "ready" if tool_map["soffice"]["installed"] else "degraded",
            "LibreOffice/soffice 可用" if tool_map["soffice"]["installed"] else "未检测到 soffice，仅使用内置 DOC 解析",
        ),
        "text_json_csv": pack("文本/JSON/CSV", "ready", "内置支持"),
    }


def collect_dependency_status() -> dict[str, Any]:
    items: list[dict[str, Any]] = []
    dep_map: dict[str, dict[str, Any]] = {}
    tool_map: dict[str, dict[str, Any]] = {}

    for spec in DEPENDENCY_SPECS:
        installed = _module_exists(spec["module"])
        item = {
            "id": spec["id"],
            "type": "python",
            "label": spec["label"],
            "module": spec["module"],
            "package": spec["package"],
            "install": spec["install"],
            "required_for": spec["required_for"],
            "optional": spec["optional"],
            "installed": installed,
            "status": "ready" if installed else ("optional_missing" if spec["optional"] else "missing"),
            "version": _safe_version(spec["package"]) if installed else "",
            "message": "已安装" if installed else ("未安装（可选）" if spec["optional"] else "未安装"),
        }
        dep_map[spec["id"]] = item
        items.append(item)

    for tool_id, candidates in TOOL_CANDIDATES.items():
        tool_path = resolve_tool_path(tool_id)
        tool_item = {
            "id": tool_id,
            "type": "tool",
            "label": "LibreOffice / soffice" if tool_id == "soffice" else "Tesseract 可执行程序",
            "path": tool_path,
            "installed": bool(tool_path),
            "status": "ready" if tool_path else "optional_missing",
            "required_for": ["doc"] if tool_id == "soffice" else ["ocr", "image"],
            "optional": True,
            "install": "安装 LibreOffice 并确保 soffice 可执行" if tool_id == "soffice" else "安装 Tesseract OCR 并加入 PATH",
            "version": "",
            "message": tool_path or "未检测到可执行程序",
        }
        tool_map[tool_id] = tool_item
        items.append(tool_item)

    profiles = _derive_profile_status(dep_map, tool_map)
    critical_missing = sum(
        1
        for item in items
        if item["type"] == "python" and not item["optional"] and not item["installed"]
    )
    optional_missing = sum(1 for item in items if item["optional"] and not item["installed"])

    if critical_missing == 0 and optional_missing == 0:
        overall_status = "ready"
    elif critical_missing == 0:
        overall_status = "degraded"
    else:
        overall_status = "missing"

    recommendations = []
    for item in items:
        if item["status"] in {"missing", "optional_missing"}:
            prefix = "建议安装" if item["status"] == "missing" else "可选安装"
            recommendations.append(f"{prefix}：{item['label']}（{item['install']}）")

    return {
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "environment": {
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "executable": sys.executable,
        },
        "summary": {
            "overall_status": overall_status,
            "total_items": len(items),
            "ready_count": sum(1 for item in items if item["installed"]),
            "critical_missing": critical_missing,
            "optional_missing": optional_missing,
        },
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
    base = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        "--no-input",
        package_name,
    ]
    fallback = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        "--no-input",
        "--trusted-host",
        "pypi.org",
        "--trusted-host",
        "files.pythonhosted.org",
        "-i",
        "https://pypi.org/simple",
        package_name,
    ]
    tuna = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        "--no-input",
        "--trusted-host",
        "pypi.tuna.tsinghua.edu.cn",
        "-i",
        "https://pypi.tuna.tsinghua.edu.cn/simple",
        package_name,
    ]
    return [base, fallback, tuna]


def _run_pip_install(package_name: str, timeout_seconds: int = 900) -> dict[str, Any]:
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PIP_DISABLE_PIP_VERSION_CHECK", "1")
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    env["PIP_NO_PROXY"] = "*"
    for proxy_key in (
        "HTTP_PROXY",
        "HTTPS_PROXY",
        "ALL_PROXY",
        "http_proxy",
        "https_proxy",
        "all_proxy",
        "PIP_PROXY",
        "REQUESTS_CA_BUNDLE",
        "CURL_CA_BUNDLE",
    ):
        env.pop(proxy_key, None)

    attempts: list[dict[str, Any]] = []
    for idx, command in enumerate(_build_pip_install_attempts(package_name), start=1):
        started_at = datetime.now()
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=timeout_seconds,
                env=env,
            )
            attempt = {
                "index": idx,
                "command": " ".join(command),
                "return_code": completed.returncode,
                "stdout_tail": _tail_text(completed.stdout),
                "stderr_tail": _tail_text(completed.stderr),
                "started_at": started_at.isoformat(timespec="seconds"),
                "finished_at": datetime.now().isoformat(timespec="seconds"),
                "timeout": False,
            }
            attempts.append(attempt)
            if completed.returncode == 0:
                return {
                    "ok": True,
                    "command": attempt["command"],
                    "return_code": completed.returncode,
                    "stdout_tail": attempt["stdout_tail"],
                    "stderr_tail": attempt["stderr_tail"],
                    "attempts": attempts,
                }
        except subprocess.TimeoutExpired as exc:
            attempts.append(
                {
                    "index": idx,
                    "command": " ".join(command),
                    "return_code": None,
                    "stdout_tail": _tail_text(exc.stdout or ""),
                    "stderr_tail": _tail_text(exc.stderr or ""),
                    "started_at": started_at.isoformat(timespec="seconds"),
                    "finished_at": datetime.now().isoformat(timespec="seconds"),
                    "timeout": True,
                }
            )

    last_attempt = attempts[-1] if attempts else {}
    return {
        "ok": False,
        "command": last_attempt.get("command", ""),
        "return_code": last_attempt.get("return_code"),
        "stdout_tail": last_attempt.get("stdout_tail", ""),
        "stderr_tail": last_attempt.get("stderr_tail", ""),
        "attempts": attempts,
    }


def install_missing_dependencies(include_optional: bool = True) -> dict[str, Any]:
    before = collect_dependency_status()
    target_items = [
        item
        for item in before["items"]
        if item["type"] == "python"
        and not item["installed"]
        and (include_optional or not item["optional"])
    ]

    if not target_items:
        remaining_tools = [
            item
            for item in before["items"]
            if item["type"] == "tool" and not item["installed"]
        ]
        return {
            "install_ok": True,
            "changed": False,
            "include_optional": include_optional,
            "message": "当前没有需要自动安装的 Python 依赖",
            "before": before,
            "after": before,
            "results": [],
            "summary": {
                "requested_count": 0,
                "installed_count": 0,
                "failed_count": 0,
                "still_missing_required": before["summary"].get("critical_missing", 0),
                "still_missing_optional": before["summary"].get("optional_missing", 0),
                "remaining_manual_tools": len(remaining_tools),
            },
            "remaining": {
                "python": [],
                "tools": remaining_tools,
            },
            "manual_actions": [f"{item['label']}：{item['install']}" for item in remaining_tools],
        }

    results: list[dict[str, Any]] = []
    for item in target_items:
        package_name = item.get("package") or item.get("module") or item["id"]
        install_result = _run_pip_install(package_name)
        importlib.invalidate_caches()
        results.append(
            {
                "id": item["id"],
                "label": item["label"],
                "package": package_name,
                "optional": item["optional"],
                "required_for": item.get("required_for", []),
                "requested_install": item.get("install", ""),
                "attempt_count": len(install_result["attempts"]),
                "command": install_result["command"],
                "return_code": install_result["return_code"],
                "stdout_tail": install_result["stdout_tail"],
                "stderr_tail": install_result["stderr_tail"],
                "attempts": install_result["attempts"],
                "pip_ok": install_result["ok"],
            }
        )

    importlib.invalidate_caches()
    after = collect_dependency_status()
    after_python_items = {
        item["id"]: item
        for item in after["items"]
        if item["type"] == "python"
    }

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

    remaining_python = [
        item
        for item in after["items"]
        if item["type"] == "python"
        and not item["installed"]
        and (include_optional or not item["optional"])
    ]
    remaining_tools = [
        item
        for item in after["items"]
        if item["type"] == "tool" and not item["installed"]
    ]

    install_ok = failed_count == 0
    if install_ok:
        message = f"已安装 {installed_count} 项 Python 依赖"
    elif installed_count:
        message = f"已安装 {installed_count} 项，仍有 {failed_count} 项安装失败"
    else:
        message = "未能安装缺失依赖"

    return {
        "install_ok": install_ok,
        "changed": installed_count > 0,
        "include_optional": include_optional,
        "message": message,
        "before": before,
        "after": after,
        "results": results,
        "summary": {
            "requested_count": len(target_items),
            "installed_count": installed_count,
            "failed_count": failed_count,
            "still_missing_required": after["summary"].get("critical_missing", 0),
            "still_missing_optional": after["summary"].get("optional_missing", 0),
            "remaining_manual_tools": len(remaining_tools),
        },
        "remaining": {
            "python": remaining_python,
            "tools": remaining_tools,
        },
        "manual_actions": [f"{item['label']}：{item['install']}" for item in remaining_tools],
        "requested": [
            {
                "id": item["id"],
                "label": item["label"],
                "package": item.get("package") or item.get("module") or item["id"],
                "optional": item["optional"],
                "required_for": item.get("required_for", []),
            }
            for item in target_items
        ],
    }


def extract_install_command(message: str) -> str:
    if not message:
        return ""

    pip_match = re.search(r"(pip install [\w\-.]+)", message, re.IGNORECASE)
    if pip_match:
        return pip_match.group(1)

    mod_match = re.search(r"No module named ['\"]?([\w\-.]+)['\"]?", message, re.IGNORECASE)
    if mod_match:
        module_name = mod_match.group(1)
        package_name = {
            "docx": "python-docx",
            "pptx": "python-pptx",
            "PIL": "Pillow",
            "fitz": "PyMuPDF",
        }.get(module_name, module_name)
        return f"pip install {package_name}"

    return ""


def build_error_info(
    error_message: str | None,
    file_name: str = "",
    file_ext: str = "",
    metadata_dict: dict[str, Any] | None = None,
    source: str = "runtime",
) -> dict[str, Any] | None:
    message = (error_message or "").strip()
    if not message:
        return None

    metadata_dict = metadata_dict or {}
    lower = message.lower()
    install_command = extract_install_command(message)
    dependency = ""
    code = "parse_failure"
    category = "解析失败"
    title = "文档解析失败"
    severity = "error"
    hint = "请检查文件内容、格式和依赖环境后重试。"

    dep_match = re.search(r"(?:No module named ['\"]?([\w\-.]+)['\"]?)|(?:pip install ([\w\-.]+))", message, re.IGNORECASE)
    if dep_match:
        dependency = dep_match.group(1) or dep_match.group(2) or ""

    if "缺少依赖" in message or "no module named" in lower or install_command:
        code = "missing_dependency"
        category = "依赖缺失"
        title = "运行依赖未安装"
        severity = "warning"
        hint = install_command or "请安装对应 Python 依赖后再重试。"
    elif "不支持的文件格式" in message or "unsupported" in lower:
        code = "unsupported_format"
        category = "格式不支持"
        title = "文件格式暂不支持"
        hint = "请改用系统支持的格式，或先转换为 PDF/DOCX/XLSX/PPTX/TXT/CSV/图片/JSON。"
    elif "编码" in message or "unicode" in lower or "decode" in lower:
        code = "encoding_error"
        category = "编码异常"
        title = "文件编码无法识别"
        hint = "建议将文本文件另存为 UTF-8 编码后重试。"
    elif any(token in lower for token in ["corrupt", "badzipfile", "not a zip file", "ole2", "piece table", "fib", "clx"]) or "损坏" in message:
        code = "corrupt_file"
        category = "文件损坏"
        title = "文件可能损坏或伪装格式"
        hint = "请确认文件能被原始办公软件正常打开，必要时重新导出。"
    elif "ocr" in lower and ("未安装" in message or "不可用" in message or "未识别到内容" in message):
        code = "ocr_unavailable"
        category = "OCR 不可用"
        title = "OCR 依赖或识别链路不可用"
        hint = install_command or "建议安装 EasyOCR，或补齐 Tesseract 运行环境。"
    elif "文件不存在" in message or "no such file" in lower:
        code = "file_not_found"
        category = "文件缺失"
        title = "待处理文件不存在"
        hint = "请确认上传成功或文件路径有效。"
    elif "权限" in message or "permission" in lower:
        code = "permission_error"
        category = "权限不足"
        title = "文件访问权限不足"
        hint = "请关闭占用文件的程序，或检查读写权限。"

    return {
        "code": code,
        "category": category,
        "title": title,
        "message": message,
        "hint": hint,
        "severity": severity,
        "dependency": dependency,
        "install_command": install_command,
        "file_name": file_name,
        "file_ext": file_ext,
        "source": source,
        "parser_hint": metadata_dict.get("解析方式", ""),
    }


def augment_result_payload(
    payload: dict[str, Any],
    file_name: str = "",
    file_ext: str = "",
    source: str = "runtime",
) -> dict[str, Any]:
    result = dict(payload)
    result["error_info"] = build_error_info(
        result.get("error", ""),
        file_name=file_name or result.get("file", ""),
        file_ext=file_ext,
        metadata_dict=result.get("metadata") or {},
        source=source,
    )
    return result


def summarize_error_records(records: list[dict[str, Any]]) -> dict[str, Any]:
    categories = Counter()
    samples = []

    for record in records:
        if record.get("success"):
            continue
        info = record.get("error_info") or build_error_info(
            record.get("error", ""),
            file_name=record.get("filename", ""),
            file_ext=record.get("extension", ""),
            metadata_dict=record.get("metadata") or {},
            source="batch",
        )
        if not info:
            continue
        categories[info["category"]] += 1
        if len(samples) < 8:
            samples.append(
                {
                    "filename": record.get("filename", ""),
                    "category": info["category"],
                    "message": info["message"],
                    "hint": info["hint"],
                }
            )

    return {
        "error_category_counts": dict(categories),
        "error_samples": samples,
    }
