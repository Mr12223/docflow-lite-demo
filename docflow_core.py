"""兼容垫片：实际代码已移至 docflow/core/"""
# ruff: noqa: F401, F403
from docflow.core import *
from docflow.core import (
    DocFlowCancelledError,
    ExtractionResult,
    BaseParser,
    PDFParser,
    WordParser,
    ExcelParser,
    PPTXParser,
    TextParser,
    TextAnalyzer,
    OutputFormatter,
    DocFlowProcessor,
    import_with_base_fallback,
)
