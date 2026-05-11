"""
发票记录 SQLite 持久化模块
提供发票提取结果的存储、查询、删除和导出功能。
"""

import csv
import io
import os
import sqlite3
import time
from typing import Optional

from docflow.paths import INVOICE_DB_PATH


# 数据库文件路径
_DEFAULT_DB_PATH = INVOICE_DB_PATH

# 发票字段列表（与 invoice_extractor 保持一致）
INVOICE_COLUMNS = [
    "invoice_code", "invoice_number", "invoice_date",
    "amount", "tax", "total",
    "buyer_name", "seller_name",
    "buyer_tax_id", "seller_tax_id",
    "invoice_type", "machine_number", "check_code",
]


class InvoiceDB:
    """发票记录数据库管理"""

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = str(db_path or _DEFAULT_DB_PATH)
        # 确保目录存在
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        conn = self._get_conn()
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS invoice_records (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name     TEXT NOT NULL,
                    file_type     TEXT NOT NULL DEFAULT '',
                    file_size     INTEGER DEFAULT 0,
                    invoice_code   TEXT,
                    invoice_number TEXT,
                    invoice_date   TEXT,
                    amount         TEXT,
                    tax            TEXT,
                    total          TEXT,
                    buyer_name     TEXT,
                    seller_name    TEXT,
                    buyer_tax_id   TEXT,
                    seller_tax_id  TEXT,
                    invoice_type   TEXT,
                    machine_number TEXT,
                    check_code     TEXT,
                    confidence     TEXT DEFAULT 'low',
                    field_count    INTEGER DEFAULT 0,
                    raw_text       TEXT,
                    created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()
        finally:
            conn.close()

    # ──────────────────────────────────────
    #  CRUD 操作
    # ──────────────────────────────────────

    def save_record(
        self,
        fields: dict,
        file_name: str,
        file_type: str = "",
        file_size: int = 0,
        confidence: str = "low",
        field_count: int = 0,
        raw_text: str = "",
    ) -> int:
        """
        保存一条发票记录，返回新记录 ID。

        :param fields: InvoiceExtractor 提取的 fields 字典
                       格式: {"字段名": {"value": "...", ...}, ...}
        """
        # 从 fields 中提取各列的值
        col_values = {}
        for col in INVOICE_COLUMNS:
            f = fields.get(col)
            col_values[col] = f["value"] if f and isinstance(f, dict) else None

        conn = self._get_conn()
        try:
            cursor = conn.execute(
                """
                INSERT INTO invoice_records
                    (file_name, file_type, file_size,
                     invoice_code, invoice_number, invoice_date,
                     amount, tax, total,
                     buyer_name, seller_name,
                     buyer_tax_id, seller_tax_id,
                     invoice_type, machine_number, check_code,
                     confidence, field_count, raw_text)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    file_name, file_type, file_size,
                    col_values["invoice_code"], col_values["invoice_number"], col_values["invoice_date"],
                    col_values["amount"], col_values["tax"], col_values["total"],
                    col_values["buyer_name"], col_values["seller_name"],
                    col_values["buyer_tax_id"], col_values["seller_tax_id"],
                    col_values["invoice_type"], col_values["machine_number"], col_values["check_code"],
                    confidence, field_count, raw_text[:2000] if raw_text else "",
                ),
            )
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def list_records(
        self,
        page: int = 1,
        per_page: int = 20,
        search: str = "",
    ) -> dict:
        """
        分页查询发票记录。

        返回: {"records": [...], "total": int, "page": int, "per_page": int}
        """
        conn = self._get_conn()
        try:
            where_clause = ""
            params: list = []
            if search:
                where_clause = """
                    WHERE file_name LIKE ? OR invoice_number LIKE ?
                    OR buyer_name LIKE ? OR seller_name LIKE ?
                    OR invoice_code LIKE ?
                """
                like = f"%{search}%"
                params = [like, like, like, like, like]

            # 总数
            total_row = conn.execute(
                f"SELECT COUNT(*) as cnt FROM invoice_records {where_clause}",
                params,
            ).fetchone()
            total = total_row["cnt"] if total_row else 0

            # 分页
            offset = (max(1, page) - 1) * per_page
            rows = conn.execute(
                f"""SELECT * FROM invoice_records {where_clause}
                    ORDER BY created_at DESC
                    LIMIT ? OFFSET ?""",
                params + [per_page, offset],
            ).fetchall()

            records = [dict(r) for r in rows]
            return {
                "records": records,
                "total": total,
                "page": page,
                "per_page": per_page,
            }
        finally:
            conn.close()

    def get_record(self, record_id: int) -> Optional[dict]:
        """获取单条记录"""
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT * FROM invoice_records WHERE id = ?", (record_id,)
            ).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def delete_record(self, record_id: int) -> bool:
        """删除一条记录，返回是否成功"""
        conn = self._get_conn()
        try:
            cursor = conn.execute(
                "DELETE FROM invoice_records WHERE id = ?", (record_id,)
            )
            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    def delete_all_records(self) -> int:
        conn = self._get_conn()
        try:
            cursor = conn.execute("DELETE FROM invoice_records")
            conn.commit()
            return max(0, cursor.rowcount or 0)
        finally:
            conn.close()

    # ──────────────────────────────────────
    #  导出
    # ──────────────────────────────────────

    def export_csv(self) -> str:
        """导出所有记录为 CSV 字符串"""
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM invoice_records ORDER BY created_at DESC"
            ).fetchall()

            if not rows:
                return ""

            output = io.StringIO()
            writer = csv.writer(output)

            # 表头（中文）
            header_map = {
                "id": "ID", "file_name": "文件名", "file_type": "文件类型",
                "invoice_code": "发票代码", "invoice_number": "发票号码",
                "invoice_date": "开票日期", "amount": "金额",
                "tax": "税额", "total": "价税合计",
                "buyer_name": "购买方", "seller_name": "销售方",
                "buyer_tax_id": "购买方税号", "seller_tax_id": "销售方税号",
                "invoice_type": "发票类型", "confidence": "置信度",
                "field_count": "识别字段数", "created_at": "创建时间",
            }
            # 选择导出的列
            export_cols = [
                "id", "file_name", "file_type",
                "invoice_code", "invoice_number", "invoice_date",
                "amount", "tax", "total",
                "buyer_name", "seller_name",
                "buyer_tax_id", "seller_tax_id",
                "invoice_type", "confidence", "field_count", "created_at",
            ]
            writer.writerow([header_map.get(c, c) for c in export_cols])

            for row in rows:
                d = dict(row)
                writer.writerow([d.get(c, "") for c in export_cols])

            return output.getvalue()
        finally:
            conn.close()
