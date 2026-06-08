"""DocFlow Web 应用包入口。"""

from .core import app
from .routes import register_routes
from .services.ocr import IMAGE_EXTS, process_image_ocr

register_routes()


def main() -> None:
    """启动本地开发服务。"""
    print("=" * 45)
    print("  DocFlow 服务已启动！")
    print("  请用浏览器打开：http://127.0.0.1:5000")
    print("=" * 45)
    app.run(debug=True, port=5000, use_reloader=False)


__all__ = ["app", "main", "IMAGE_EXTS", "process_image_ocr"]
