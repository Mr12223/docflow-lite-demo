"""错误信息构建与增强工具。"""
from __future__ import annotations

import re
from collections import Counter
from typing import Any


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
        "total_errors": sum(categories.values()),
    }
