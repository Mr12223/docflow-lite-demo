"""Tesseract OCR 工具路径解析与配置。"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Any


TOOL_CANDIDATES: dict[str, list[str]] = {
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
