"""docflow.support — 依赖检查、Tesseract 配置和错误处理工具。"""
from .tesseract import (
    TOOL_CANDIDATES,
    resolve_tool_path,
    resolve_tessdata_dir,
    configure_pytesseract_command,
    build_tesseract_ocr_config,
    prepare_pytesseract,
)
from .deps import (
    DEPENDENCY_SPECS,
    collect_dependency_status,
    install_missing_dependencies,
    _item_available,
    _module_exists,
    _safe_version,
    _find_dependency_spec,
)
from .errors import (
    extract_install_command,
    build_error_info,
    augment_result_payload,
    summarize_error_records,
)

__all__ = [
    "TOOL_CANDIDATES",
    "DEPENDENCY_SPECS",
    "resolve_tool_path",
    "resolve_tessdata_dir",
    "configure_pytesseract_command",
    "build_tesseract_ocr_config",
    "prepare_pytesseract",
    "collect_dependency_status",
    "install_missing_dependencies",
    "extract_install_command",
    "build_error_info",
    "augment_result_payload",
    "summarize_error_records",
]
