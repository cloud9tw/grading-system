"""
access_to_mariadb.py
批次將多個 Access 檔案資料寫入 Synology MariaDB

使用方式:
    python access_to_mariadb.py                        # 匯入 ACCESS_DIR 下所有 .accdb/.mdb
    python access_to_mariadb.py path/to/file.accdb     # 只匯入指定檔案

設定方式: 修改下方 CONFIG 區塊，或建立 .env 並加入對應的環境變數
"""

import os
import sys
import logging
import re
from pathlib import Path
from typing import Any

import pyodbc
import pymysql
import pymysql.cursors
from dotenv import load_dotenv

load_dotenv()

# ─── CONFIG ───────────────────────────────────────────────────────────────────
ACCESS_DIR = os.getenv("ACCESS_DIR", r"C:\Access_files")   # Access 檔案資料夾
DB_HOST    = os.getenv("MARIA_HOST", "192.168.1.x")        # Synology IP
DB_PORT    = int(os.getenv("MARIA_PORT", "3306"))
DB_NAME    = os.getenv("MARIA_DB",   "mydb")
DB_USER    = os.getenv("MARIA_USER", "root")
DB_PASS    = os.getenv("MARIA_PASS", "password")

# 若為 True，相同 table 的資料會先 TRUNCATE 再插入；False 則直接 INSERT（可能重複）
TRUNCATE_BEFORE_INSERT = os.getenv("TRUNCATE_BEFORE_INSERT", "false").lower() == "true"

# 若 Access table 在 MariaDB 不存在，是否自動建表
AUTO_CREATE_TABLE = True
# ──────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)


# ── Access 型別 → MariaDB 型別 ────────────────────────────────────────────────
# pyodbc type_code 對應：https://github.com/mkleehammer/pyodbc/wiki/Data-Types
ACCESS_TYPE_MAP: dict[int, str] = {
    # SQL_CHAR / SQL_WCHAR / SQL_VARCHAR / SQL_WVARCHAR / SQL_LONGVARCHAR
    1: "VARCHAR(255)",
    -8: "VARCHAR(255)",
    12: "TEXT",
    -9: "TEXT",
    -1: "LONGTEXT",
    # SQL_SMALLINT / SQL_INTEGER / SQL_BIGINT
    5: "SMALLINT",
    4: "INT",
    -5: "BIGINT",
    # SQL_REAL / SQL_FLOAT / SQL_DOUBLE
    7: "FLOAT",
    6: "DOUBLE",
    8: "DOUBLE",
    # SQL_DECIMAL / SQL_NUMERIC
    3: "DECIMAL(18,4)",
    2: "DECIMAL(18,4)",
    # SQL_BIT
    -7: "TINYINT(1)",
    # SQL_TINYINT
    -6: "TINYINT",
    # SQL_TYPE_DATE / SQL_TYPE_TIME / SQL_TYPE_TIMESTAMP
    91: "DATE",
    92: "TIME",
    93: "DATETIME",
    # SQL_BINARY / SQL_VARBINARY / SQL_LONGVARBINARY
    -2: "BLOB",
    -3: "BLOB",
    -4: "LONGBLOB",
}


def get_mariadb_type(type_code: int, precision: int, scale: int) -> str:
    if type_code in (12, -9, 1, -8):
        # 可變長度字串，根據精度決定用 VARCHAR 還是 TEXT
        if 0 < precision <= 16383:
            return f"VARCHAR({precision})"
        return "TEXT"
    if type_code in (3, 2):
        return f"DECIMAL({precision or 18},{scale or 4})"
    return ACCESS_TYPE_MAP.get(type_code, "TEXT")


def safe_table_name(name: str) -> str:
    """把 Access table 名稱轉成合法的 MariaDB backtick 名稱"""
    return "`" + name.replace("`", "``") + "`"


def get_access_conn(accdb_path: str) -> pyodbc.Connection:
    """建立 Access 連線，自動嘗試 accdb / mdb driver"""
    drivers = [
        "{Microsoft Access Driver (*.mdb, *.accdb)}",
        "{Microsoft Access Driver (*.mdb)}",
    ]
    for drv in drivers:
        try:
            conn_str = (
                f"DRIVER={drv};"
                f"DBQ={accdb_path};"
                "ExtendedAnsiSQL=1;"
            )
            return pyodbc.connect(conn_str, autocommit=False)
        except pyodbc.Error:
            continue
    raise RuntimeError(
        "找不到 Access ODBC Driver，請安裝 Microsoft Access Database Engine:\n"
        "https://www.microsoft.com/en-us/download/details.aspx?id=54920"
    )


def get_user_tables(acc_conn: pyodbc.Connection) -> list[str]:
    """取得 Access 中所有使用者自建的 table（排除系統表）"""
    cursor = acc_conn.cursor()
    tables = [
        row.table_name
        for row in cursor.tables(tableType="TABLE")
        if not row.table_name.startswith("MSys")
    ]
    return tables


def ensure_table(maria_cur: pymysql.cursors.Cursor, table_name: str, columns: list) -> None:
    """若 MariaDB 中無此表則自動建立"""
    col_defs = []
    for col in columns:
        # col: (name, type_code, display_size, internal_size, precision, scale, null_ok)
        col_name  = col[0]
        type_code = col[1]
        precision = col[4] or 0
        scale     = col[5] or 0
        nullable  = "NULL" if col[6] else "NOT NULL"
        maria_type = get_mariadb_type(type_code, precision, scale)
        col_defs.append(f"  `{col_name}` {maria_type} {nullable}")

    ddl = (
        f"CREATE TABLE IF NOT EXISTS {safe_table_name(table_name)} (\n"
        + ",\n".join(col_defs)
        + "\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;"
    )
    log.debug("DDL: %s", ddl)
    maria_cur.execute(ddl)


def import_table(
    acc_conn: pyodbc.Connection,
    maria_conn: pymysql.connections.Connection,
    table_name: str,
    file_label: str,
) -> int:
    """匯入單張 table，回傳插入筆數"""
    acc_cur = acc_conn.cursor()
    acc_cur.execute(f"SELECT * FROM [{table_name}]")
    columns  = acc_cur.description  # list of 7-tuples
    col_names = [col[0] for col in columns]
    rows = acc_cur.fetchall()

    if not rows:
        log.info("  [%s] %s — 空表，略過", file_label, table_name)
        return 0

    maria_cur = maria_conn.cursor()

    if AUTO_CREATE_TABLE:
        ensure_table(maria_cur, table_name, columns)

    if TRUNCATE_BEFORE_INSERT:
        maria_cur.execute(f"TRUNCATE TABLE {safe_table_name(table_name)}")
        log.info("  [%s] %s — TRUNCATE 完成", file_label, table_name)

    placeholders = ", ".join(["%s"] * len(col_names))
    col_list     = ", ".join([f"`{c}`" for c in col_names])
    sql = (
        f"INSERT INTO {safe_table_name(table_name)} ({col_list}) "
        f"VALUES ({placeholders})"
    )

    # 將 pyodbc Row 轉成 tuple，確保 pymysql 相容
    data = [tuple(r) for r in rows]
    maria_cur.executemany(sql, data)
    maria_conn.commit()

    log.info("  [%s] %s — 插入 %d 筆", file_label, table_name, len(data))
    return len(data)


def process_access_file(accdb_path: Path, maria_conn: pymysql.connections.Connection) -> None:
    label = accdb_path.name
    log.info("▶ 處理檔案: %s", accdb_path)

    try:
        acc_conn = get_access_conn(str(accdb_path))
    except Exception as e:
        log.error("  無法開啟 %s: %s", label, e)
        return

    tables = get_user_tables(acc_conn)
    if not tables:
        log.warning("  %s 沒有使用者資料表", label)
        acc_conn.close()
        return

    log.info("  找到 %d 張表: %s", len(tables), tables)
    total = 0
    for tbl in tables:
        try:
            total += import_table(acc_conn, maria_conn, tbl, label)
        except Exception as e:
            log.error("  匯入 %s.%s 失敗: %s", label, tbl, e)
            maria_conn.rollback()

    acc_conn.close()
    log.info("  ✓ %s 共匯入 %d 筆\n", label, total)


def main() -> None:
    # 判斷是否指定了特定檔案
    if len(sys.argv) > 1:
        files = [Path(p) for p in sys.argv[1:]]
    else:
        base = Path(ACCESS_DIR)
        if not base.exists():
            log.error("ACCESS_DIR 不存在: %s", base)
            sys.exit(1)
        files = sorted(base.glob("**/*.accdb")) + sorted(base.glob("**/*.mdb"))

    if not files:
        log.warning("找不到任何 Access 檔案")
        sys.exit(0)

    log.info("共找到 %d 個 Access 檔案", len(files))

    try:
        maria_conn = pymysql.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASS,
            database=DB_NAME,
            charset="utf8mb4",
            autocommit=False,
        )
        log.info("MariaDB 連線成功 → %s:%d/%s", DB_HOST, DB_PORT, DB_NAME)
    except Exception as e:
        log.error("MariaDB 連線失敗: %s", e)
        sys.exit(1)

    for f in files:
        process_access_file(f, maria_conn)

    maria_conn.close()
    log.info("全部完成")


if __name__ == "__main__":
    main()
