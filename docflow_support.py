"""兼容垫片：实际代码已移至 docflow/support/"""
# ruff: noqa: F401, F403
from docflow.support import *
from docflow.support import (
    TOOL_CANDIDATES,
    DEPENDENCY_SPECS,
    resolve_tool_path,
    resolve_tessdata_dir,
    configure_pytesseract_command,
    build_tesseract_ocr_config,
    prepare_pytesseract,
    collect_dependency_status,
    install_missing_dependencies,
    extract_install_command,
    build_error_info,
    augment_result_payload,
    summarize_error_records,
)
