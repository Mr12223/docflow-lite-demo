"""DocFlow 路由注册。"""

from . import batch, common, invoice, process


def register_routes() -> None:
    """触发各路由模块导入，完成装饰器注册。"""
    _ = (common, process, batch, invoice)

