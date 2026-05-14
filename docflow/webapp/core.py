"""DocFlow Web 核心共享模块。"""

from flask import Flask
from flask_cors import CORS

from docflow.paths import REPORTS_DIR, UPLOADS_DIR, ensure_runtime_directories
from docflow.runtime import (
    configure_logging,
    configure_native_thread_env,
    configure_windows_utf8_stdio,
)

configure_native_thread_env()
configure_windows_utf8_stdio()
ensure_runtime_directories()

logger = configure_logging("DocFlow")

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = str(UPLOADS_DIR)
REPORTS_FOLDER = str(REPORTS_DIR)

from docflow.invoice.db import InvoiceDB

invoice_db = InvoiceDB()

__all__ = ["app", "logger", "UPLOAD_FOLDER", "REPORTS_FOLDER", "invoice_db"]
