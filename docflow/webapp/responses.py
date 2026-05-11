"""Web 响应工具。"""

from flask import jsonify

from docflow_support import build_error_info


def error_response(
    message: str,
    status_code: int = 400,
    file_name: str = "",
    file_ext: str = "",
):
    """返回统一结构的错误响应。"""
    payload = {
        "success": False,
        "error": message,
        "error_info": build_error_info(
            message,
            file_name=file_name,
            file_ext=file_ext,
            metadata_dict={},
            source="api",
        ),
    }
    return jsonify(payload), status_code
