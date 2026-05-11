"""DocFlow Flask 启动入口。"""

from docflow.webapp import app, main
from docflow.webapp.services.ocr import IMAGE_EXTS, process_image_ocr

__all__ = ["app", "main", "IMAGE_EXTS", "process_image_ocr"]


if __name__ == "__main__":
    main()
