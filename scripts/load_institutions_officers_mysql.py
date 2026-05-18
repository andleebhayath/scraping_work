"""
Load institutions_officers_social.xlsx into MySQL raw_institutions_officers_social_raw.

1) Edit CONFIG below.
2) Run:  python scripts/load_institutions_officers_mysql.py

Creates `RAW_TABLE` (and `PARENT_RUN_TABLE` if set) when missing, then loads Excel.
If raw.run_id FKs to the parent table, a new parent row is inserted each run.
"""

from __future__ import annotations

from pathlib import Path

import mysql.connector
import pandas as pd

# --- edit these ---
HOST = "localhost"
PORT = 3306
USER = "root"
PASSWORD = "root"
DATABASE = "organization_data"

RAW_TABLE = "raw_institutions_officers_social_raw"
PARENT_RUN_TABLE = "core_ingestion_run"  # FK parent for run_id; empty string = skip FK check

EXCEL_PATH = Path(__file__).resolve().parents[1] / "output" / "institutions_officers_social.xlsx"
# ---


def ensure_tables(cur) -> None:
    """Create parent run table and raw table if they do not exist."""
    parent = str(PARENT_RUN_TABLE).strip()

    if parent:
        cur.execute(
            f"""
            CREATE TABLE IF NOT EXISTS `{parent}` (
                run_id             BIGINT AUTO_INCREMENT PRIMARY KEY,
                pipeline_name      VARCHAR(200) NOT NULL,
                source_system      VARCHAR(100) NOT NULL,
                started_at_utc     DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
                finished_at_utc    DATETIME(3) NULL,
                status             VARCHAR(30) NOT NULL DEFAULT 'running',
                notes              VARCHAR(2000) NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
            """
        )

    fk = ""
    if parent:
        fk = f",\n                CONSTRAINT fk_raw_run FOREIGN KEY (run_id) REFERENCES `{parent}`(run_id) ON DELETE RESTRICT ON UPDATE CASCADE"

    cur.execute(
        f"""
        CREATE TABLE IF NOT EXISTS `{RAW_TABLE}` (
            raw_id               BIGINT AUTO_INCREMENT PRIMARY KEY,
            run_id               BIGINT NOT NULL,
            source_row_number    INT NOT NULL,
            iau_id               VARCHAR(100) NULL,
            university_name      VARCHAR(500) NULL,
            officer_role         VARCHAR(200) NULL,
            officer_name         VARCHAR(300) NULL,
            job_title            VARCHAR(300) NULL,
            search_query         VARCHAR(1000) NULL,
            search_engine        VARCHAR(100) NULL,
            error_text           VARCHAR(2000) NULL,
            linkedin_url         VARCHAR(2000) NULL,
            other_social_urls    LONGTEXT NULL,
            university_page_url  VARCHAR(2000) NULL,
            official_www         VARCHAR(1000) NULL,
            top_urls             LONGTEXT NULL,
            loaded_at_utc        DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3)
            {fk}
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
        """
    )


# Excel column name -> MySQL column name (only columns both sides need)
COL_MAP = [
    ("iau_id", "iau_id"),
    ("university_name", "university_name"),
    ("officer_role", "officer_role"),
    ("officer_name", "officer_name"),
    ("job_title", "job_title"),
    ("search_query", "search_query"),
    ("search_engine", "search_engine"),
    ("error", "error_text"),
    ("linkedin_url", "linkedin_url"),
    ("other_social_urls", "other_social_urls"),
    ("university_page_url", "university_page_url"),
    ("official_www", "official_www"),
    ("top_urls", "top_urls"),
]


def clean(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none", "<na>"}:
        return None
    return s


def get_run_id(cur) -> int:
    """New parent row each run (satisfies FK on raw.run_id). Set PARENT_RUN_TABLE = '' to use 1 without insert."""
    if not str(PARENT_RUN_TABLE).strip():
        return 1

    cur.execute(
        f"""
        INSERT INTO `{PARENT_RUN_TABLE}` (pipeline_name, source_system, status)
        VALUES (%s, %s, 'success')
        """,
        ("excel_raw_load", "WHED"),
    )
    return int(cur.lastrowid)


def main() -> None:
    if not EXCEL_PATH.is_file():
        raise FileNotFoundError(f"Excel not found: {EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")

    for excel_col, _ in COL_MAP:
        if excel_col not in df.columns:
            raise ValueError(f"Excel missing column: {excel_col}")

    conn = mysql.connector.connect(
        host=HOST,
        port=PORT,
        user=USER,
        password=PASSWORD,
        database=DATABASE,
    )
    cur = conn.cursor()
    try:
        ensure_tables(cur)
        conn.commit()
        print("Tables OK (created if missing).")

        run_id = get_run_id(cur)

        sql_cols = ["run_id", "source_row_number"] + [db_col for _, db_col in COL_MAP]
        placeholders = ", ".join(["%s"] * len(sql_cols))
        col_list = ", ".join(f"`{c}`" for c in sql_cols)
        sql = f"INSERT INTO `{RAW_TABLE}` ({col_list}) VALUES ({placeholders})"

        n = 0
        for i, row in df.iterrows():
            excel_row = int(i) + 2  # header is row 1
            vals: list = [run_id, excel_row]
            for excel_col, _db_col in COL_MAP:
                vals.append(clean(row.get(excel_col)))
            cur.execute(sql, vals)
            n += 1

        conn.commit()
        print(f"Inserted {n} rows into {RAW_TABLE} (run_id={run_id}).")
    except Exception:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
