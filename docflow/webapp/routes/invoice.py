"""发票记录路由。"""

from flask import Response, jsonify, request

from ..core import app, invoice_db
from ..responses import error_response


@app.route("/invoices", methods=["GET"])
def list_invoices():
    page = request.args.get("page", 1, type=int)
    per_page = request.args.get("per_page", 20, type=int)
    search = request.args.get("search", "").strip()
    data = invoice_db.list_records(page=page, per_page=per_page, search=search)
    return jsonify({"success": True, **data})


@app.route("/invoices", methods=["DELETE"])
def clear_invoices():
    deleted_count = invoice_db.delete_all_records()
    return jsonify({"success": True, "deleted": True, "deleted_count": deleted_count})


@app.route("/invoices/<int:record_id>", methods=["GET"])
def get_invoice(record_id: int):
    record = invoice_db.get_record(record_id)
    if not record:
        return error_response("发票记录不存在", status_code=404)
    return jsonify({"success": True, "record": record})


@app.route("/invoices/<int:record_id>", methods=["DELETE"])
def delete_invoice(record_id: int):
    ok = invoice_db.delete_record(record_id)
    if not ok:
        return error_response("发票记录不存在", status_code=404)
    return jsonify({"success": True, "deleted": True})


@app.route("/invoices/export", methods=["GET"])
def export_invoices():
    csv_content = invoice_db.export_csv()
    if not csv_content:
        return error_response("暂无发票记录可导出")
    return Response(
        "\ufeff" + csv_content,
        content_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=invoice_records.csv"},
    )
