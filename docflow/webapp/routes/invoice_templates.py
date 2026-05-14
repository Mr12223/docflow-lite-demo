"""Invoice template management routes."""

from flask import jsonify, request

from docflow.invoice.schema import (
    delete_custom_invoice_schema,
    get_field_labels,
    get_invoice_field_order,
    list_invoice_schema_payloads,
    upsert_custom_invoice_schema,
    validate_invoice_schema_payload,
)

from ..core import app
from ..responses import error_response


@app.route("/invoice-templates", methods=["GET"])
def list_invoice_templates():
    return jsonify(
        {
            "success": True,
            "templates": list_invoice_schema_payloads(),
            "field_labels": get_field_labels(),
            "field_order": list(get_invoice_field_order()),
        }
    )


@app.route("/invoice-templates", methods=["POST"])
def save_invoice_template():
    payload = request.get_json(silent=True) or {}
    normalized, error_message = validate_invoice_schema_payload(payload)
    if error_message:
        return error_response(error_message, status_code=400)
    templates = upsert_custom_invoice_schema(normalized)
    return jsonify(
        {
            "success": True,
            "templates": templates,
            "field_labels": get_field_labels(),
            "field_order": list(get_invoice_field_order()),
        }
    )


@app.route("/invoice-templates/<template_key>", methods=["DELETE"])
def delete_invoice_template(template_key: str):
    before = {
        item["key"]
        for item in list_invoice_schema_payloads()
        if not item.get("builtin")
    }
    if template_key not in before:
        return error_response("自定义模板不存在，或内置模板不能删除", status_code=404)
    templates = delete_custom_invoice_schema(template_key)
    return jsonify(
        {
            "success": True,
            "deleted": True,
            "templates": templates,
            "field_labels": get_field_labels(),
            "field_order": list(get_invoice_field_order()),
        }
    )
